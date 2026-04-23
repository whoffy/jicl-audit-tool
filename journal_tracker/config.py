import csv
import re
from difflib import SequenceMatcher
from pathlib import Path
import yaml


CONFIG_DIR = Path(__file__).parent / "config"
RANGE_PATTERN = re.compile(r"(?:Fn\.?\s*)?(\d+)\s*(?:-|to)\s*(\d+)")

SECTION_KEYWORDS = {
    "source_pull_1": [r"source\s*pull\s*(?:round\s*)?1", r"round\s*1\s*source\s*pull"],
    "source_pull_2": [r"source\s*pull\s*(?:round\s*)?2", r"round\s*2\s*source\s*pull"],
    "editing_1": [r"editing\s*round\s*1", r"round\s*1\s*editing"],
    "editing_2": [r"editing\s*round\s*2", r"round\s*2\s*editing"],
}

# Ambiguous keywords that need context (could be either phase)
ROUND_2_PATTERN = re.compile(r"^rnd\s*2$", re.IGNORECASE)


def load_name_aliases(path=None):
    path = path or CONFIG_DIR / "name_aliases.yaml"
    variant_to_canonical = {}
    canonical_variants = {}

    if not Path(path).exists():
        return variant_to_canonical, canonical_variants

    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f)

    if not data:
        return variant_to_canonical, canonical_variants

    for canonical, variants in data.get("aliases", {}).items():
        canonical_variants[canonical] = variants
        variant_to_canonical[canonical.lower()] = canonical
        for v in variants:
            variant_to_canonical[v.lower()] = canonical

    return variant_to_canonical, canonical_variants


def normalize_name(name, variant_to_canonical):
    if not name:
        return name
    key = name.strip().lower()
    return variant_to_canonical.get(key, name.strip())


def auto_detect_aliases(names, name_counts=None, threshold=0.85):
    """Find names that are likely the same person based on string similarity.

    Returns dict mapping variant → canonical name.
    If name_counts provided, canonical is the most frequent form.
    Otherwise falls back to longest name.
    """
    name_list = sorted(set(n for n in names if n and len(n) > 1))
    groups = []
    used = set()

    for i in range(len(name_list)):
        if name_list[i] in used:
            continue
        group = {name_list[i]}
        for j in range(i + 1, len(name_list)):
            if name_list[j] in used:
                continue
            ratio = SequenceMatcher(
                None, name_list[i].lower(), name_list[j].lower()
            ).ratio()
            if ratio >= threshold:
                group.add(name_list[j])
                used.add(name_list[j])
        if len(group) > 1:
            groups.append(group)
            used.add(name_list[i])

    counts = name_counts or {}
    auto_map = {}
    for group in groups:
        canonical = max(group, key=lambda n: (counts.get(n, 0), len(n)))
        for name in group:
            if name != canonical:
                auto_map[name] = canonical

    return auto_map


def _detect_section(line, default_phase="editing"):
    """Detect which section a header line belongs to.

    default_phase is used to resolve ambiguous headers like 'Rnd 2'.
    """
    text = line.lower().strip()
    for section_key, patterns in SECTION_KEYWORDS.items():
        for pat in patterns:
            if re.search(pat, text):
                return section_key
    # Handle ambiguous "Rnd 2" style headers
    if ROUND_2_PATTERN.match(text):
        return f"{default_phase}_2"
    return None


def _parse_range(text):
    """Parse a footnote range string like '1-52', 'Fn. 1-76', '1 to 52'."""
    if not text:
        return None, None
    m = RANGE_PATTERN.search(str(text))
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _parse_assignment_csv(csv_path):
    """Parse a single assignment CSV, handling all known formats.

    Handles:
    - Single-column layout (sections stacked vertically)
    - Side-by-side layout (e.g., Source Pull left, Editing right)

    Returns list of dicts: {section, person, fn_start, fn_end}
    """
    with open(csv_path, encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        rows = list(reader)

    # Skip metadata CSVs (source pull file listings)
    for row in rows:
        cells = [c.strip().lower() for c in row]
        if "team" in cells and "footnote" in cells and "modified by" in cells:
            return []

    # Infer default section from filename if no header is found
    filename_lower = str(csv_path).lower()
    if "source pull" in filename_lower:
        default_section = "source_pull_1"
    else:
        default_section = "editing_1"

    # Detect side-by-side layout: look for rows with section headers in multiple columns
    panels = _detect_panels(rows)
    if panels:
        return _parse_multi_panel(rows, panels)
    return _parse_single_column(rows, default_section)


def _detect_panels(rows):
    """Check if the CSV has a side-by-side layout with multiple section columns.

    Returns list of (col_index, section_key) tuples if multi-panel, else None.
    """
    for row in rows:
        sections_found = []
        for col_idx, cell in enumerate(row):
            if not cell:
                continue
            section = _detect_section(cell.strip())
            if section:
                sections_found.append((col_idx, section))
        if len(sections_found) >= 2:
            return sections_found
    return None


def _parse_multi_panel(rows, initial_panels):
    """Parse a CSV with side-by-side panels (e.g., SP left, Editing right)."""
    assignments = []

    # Track current section per panel column position
    panel_cols = [col for col, _ in initial_panels]
    current_sections = {col: None for col in panel_cols}

    for row in rows:
        if not row or all(not (cell.strip() if cell else "") for cell in row):
            continue

        # Check for section headers in any panel column
        new_sections = False
        for col_idx, cell in enumerate(row):
            if not cell:
                continue
            section = _detect_section(cell.strip())
            if section:
                # Find which panel this belongs to (closest panel col)
                closest = min(panel_cols, key=lambda c: abs(c - col_idx))
                current_sections[closest] = section
                new_sections = True

        if new_sections:
            continue

        # Skip header/meta rows
        first_cell = row[0].strip() if row[0] else ""
        if first_cell.lower() in ("staff editor", ""):
            continue
        if any(kw in first_cell.lower() for kw in ["start", "footnotes", "note:", "tentatively"]):
            continue

        # Try to extract person+range from each panel
        for panel_col in panel_cols:
            section = current_sections.get(panel_col)
            if not section:
                continue
            if panel_col >= len(row):
                continue

            person_cell = row[panel_col].strip() if row[panel_col] else ""
            if not person_cell or person_cell.lower() in ("staff editor", ""):
                continue
            if any(kw in person_cell.lower() for kw in ["start", "footnotes", "note:", "tentatively"]):
                continue

            # Look for range in columns after this panel's person column
            fn_start, fn_end = None, None
            for search_col in range(panel_col + 1, min(panel_col + 3, len(row))):
                cell = row[search_col].strip() if row[search_col] else ""
                if cell.lower() in ("y", "n", ""):
                    continue
                fn_start, fn_end = _parse_range(cell)
                if fn_start is not None:
                    break

            if fn_start is not None:
                assignments.append({
                    "section": section,
                    "person": person_cell,
                    "fn_start": fn_start,
                    "fn_end": fn_end,
                })

    return assignments


def _parse_single_column(rows, default_section="editing_1"):
    """Parse a single-column (vertically stacked sections) CSV."""
    assignments = []
    current_section = None
    default_phase = default_section.rsplit("_", 1)[0]  # "source_pull" or "editing"

    for row in rows:
        if not row or all(not (cell.strip() if cell else "") for cell in row):
            continue

        first_cell = row[0].strip() if row[0] else ""

        section = _detect_section(first_cell, default_phase)
        if section:
            current_section = section
            continue

        # Skip header rows
        if first_cell.lower() in ("staff editor", "staff editor,", "name", ""):
            continue
        if any(kw in first_cell.lower() for kw in ["start", "end", "footnotes", "note:", "tentatively", "?"]):
            continue

        # Try to find person + range
        if not first_cell or first_cell[0].isdigit():
            continue

        person = first_cell
        fn_range_text = None

        for cell in row[1:]:
            cell = cell.strip() if cell else ""
            if cell.lower() in ("y", "n", ""):
                continue
            fn_start, fn_end = _parse_range(cell)
            if fn_start is not None:
                fn_range_text = cell
                break

        if person and fn_range_text and current_section:
            fn_start, fn_end = _parse_range(fn_range_text)
            if fn_start is not None:
                assignments.append({
                    "section": current_section,
                    "person": person,
                    "fn_start": fn_start,
                    "fn_end": fn_end,
                })

    # If we found some assignments but the first entries have no section
    # (e.g., file starts with "Staff editor,,Footnotes" before any header),
    # those pre-header entries were skipped. Re-parse using default_section for
    # entries before the first detected section header.
    if not assignments or not any(a["section"] == default_section for a in assignments):
        pre_section = []
        for row in rows:
            if not row or all(not (cell.strip() if cell else "") for cell in row):
                continue
            first_cell = row[0].strip() if row[0] else ""
            if _detect_section(first_cell):
                break
            if first_cell.lower() in ("staff editor", "staff editor,", "name", ""):
                continue
            if any(kw in first_cell.lower() for kw in ["start", "end", "footnotes", "note:", "tentatively", "?"]):
                continue
            if not first_cell or first_cell[0].isdigit():
                continue
            person = first_cell
            for cell in row[1:]:
                cell = cell.strip() if cell else ""
                if cell.lower() in ("y", "n", ""):
                    continue
                fn_start, fn_end = _parse_range(cell)
                if fn_start is not None:
                    pre_section.append({
                        "section": default_section,
                        "person": person,
                        "fn_start": fn_start,
                        "fn_end": fn_end,
                    })
                    break

        if pre_section:
            assignments = pre_section + assignments

    return assignments


def load_team_assignments(input_dir, variant_to_canonical):
    """Load all assignments from per-team CSV files under input_dir.

    Returns dict: {team_name: [assignment_dicts]}
    Each assignment dict: {phase, round, person, fn_start, fn_end}
    """
    input_path = Path(input_dir)
    all_assignments = {}

    for team_dir in sorted(input_path.iterdir()):
        if not team_dir.is_dir():
            continue
        team_name = team_dir.name

        team_assignments = []
        for csv_file in team_dir.glob("*.csv"):
            raw = _parse_assignment_csv(csv_file)
            for a in raw:
                section = a["section"]
                if "source_pull" in section:
                    phase = "Source Pull"
                    rnd = 1 if section.endswith("1") else 2
                else:
                    phase = "Editing"
                    rnd = 1 if section.endswith("1") else 2

                team_assignments.append({
                    "team": team_name,
                    "phase": phase,
                    "round": rnd,
                    "person": normalize_name(a["person"], variant_to_canonical),
                    "fn_start": a["fn_start"],
                    "fn_end": a["fn_end"],
                })

        all_assignments[team_name] = team_assignments

    return all_assignments


def get_source_pull_assignments(all_assignments):
    """From all team assignments, extract source pull round 1, keyed by person."""
    result = {}
    for team_name, assignments in all_assignments.items():
        for a in assignments:
            if a["phase"] == "Source Pull" and a["round"] == 1:
                result[a["person"]] = a
    return result


def get_editing_assignments(all_assignments):
    """From all team assignments, extract editing round 1, keyed by person."""
    result = {}
    for team_name, assignments in all_assignments.items():
        for a in assignments:
            if a["phase"] == "Editing" and a["round"] == 1:
                result[a["person"]] = a
    return result
