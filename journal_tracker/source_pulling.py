import csv
from pathlib import Path
import pandas as pd
from .config import normalize_name


def load_metadata(input_dir, variant_to_canonical):
    """Load source pull metadata from per-team CSVs in each team folder.

    Looks for CSV files matching the metadata format (Team,Footnote,Filename,...).

    Returns a DataFrame with columns: team, footnote, filename, modified_by
    """
    input_path = Path(input_dir)
    all_rows = []

    for team_dir in sorted(input_path.iterdir()):
        if not team_dir.is_dir():
            continue
        team_name = team_dir.name

        for csv_file in team_dir.glob("*.csv"):
            rows = _try_parse_metadata_csv(csv_file, team_name, variant_to_canonical)
            if rows:
                all_rows.extend(rows)

    if not all_rows:
        return pd.DataFrame(columns=["team", "footnote", "filename", "modified_by"])
    return pd.DataFrame(all_rows)


def _try_parse_metadata_csv(csv_path, team_name, variant_to_canonical):
    """Try to parse a CSV as source pull metadata.

    Returns list of row dicts if it matches the metadata format, else empty list.
    """
    with open(csv_path, encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        all_rows = list(reader)

    # Find the header row: look for a row with "Team", "Footnote", "Modified By"
    header_idx = None
    for i, row in enumerate(all_rows):
        cells = [c.strip().lower() for c in row]
        if "team" in cells and "footnote" in cells and "modified by" in cells:
            header_idx = i
            break

    if header_idx is None:
        return []

    header = [c.strip().lower() for c in all_rows[header_idx]]
    try:
        team_col = header.index("team")
        fn_col = header.index("footnote")
        filename_col = header.index("filename")
        modified_by_col = header.index("modified by")
    except ValueError:
        return []

    rows = []
    for row in all_rows[header_idx + 1:]:
        if len(row) <= modified_by_col:
            continue
        modified_by = row[modified_by_col].strip()
        if not modified_by:
            continue
        try:
            fn_num = int(row[fn_col].strip())
        except (ValueError, TypeError):
            continue

        rows.append({
            "team": row[team_col].strip() or team_name,
            "footnote": fn_num,
            "filename": row[filename_col].strip() if len(row) > filename_col else "",
            "modified_by": normalize_name(modified_by, variant_to_canonical),
        })

    return rows


def compute_source_pull_metrics(metadata_df, sp_assignments):
    """Compute per-person source pull metrics.

    Only includes people with source pull assignments.
    People with 0 contributions still appear.

    Returns a DataFrame with columns:
        team, person, total, in_assigned, outside_assigned,
        other_edits_to_assigned, quality_ratio
    """
    results = []
    assigned_people = set(sp_assignments.keys())

    for team in metadata_df["team"].unique():
        team_df = metadata_df[metadata_df["team"] == team]
        team_people = set(team_df["modified_by"].unique()) & assigned_people

        for person, assignment in sp_assignments.items():
            if person in team_df["modified_by"].values:
                team_people.add(person)

        for person in team_people:
            person_files = team_df[team_df["modified_by"] == person]
            total = len(person_files)

            assignment = sp_assignments.get(person)
            if assignment:
                fn_start = assignment["fn_start"]
                fn_end = assignment["fn_end"]
                in_assigned = len(person_files[
                    (person_files["footnote"] >= fn_start) &
                    (person_files["footnote"] <= fn_end)
                ])
                outside_assigned = total - in_assigned

                others_in_range = team_df[
                    (team_df["footnote"] >= fn_start) &
                    (team_df["footnote"] <= fn_end) &
                    (team_df["modified_by"] != person)
                ]
                other_edits_to_assigned = len(others_in_range)
            else:
                in_assigned = total
                outside_assigned = 0
                other_edits_to_assigned = 0

            quality_ratio = (
                other_edits_to_assigned / total if total > 0 else 0.0
            )

            results.append({
                "team": team,
                "person": person,
                "total": total,
                "in_assigned": in_assigned,
                "outside_assigned": outside_assigned,
                "other_edits_to_assigned": other_edits_to_assigned,
                "quality_ratio": quality_ratio,
            })

    return pd.DataFrame(results)


def compute_leaderboards(metrics_df):
    """Produce the summary leaderboards from per-person metrics."""
    agg = metrics_df.groupby("person", as_index=False).agg({
        "total": "sum",
        "in_assigned": "sum",
        "outside_assigned": "sum",
        "other_edits_to_assigned": "sum",
    })
    agg["quality_ratio"] = agg["other_edits_to_assigned"] / agg["total"].replace(0, 1)

    highest = agg.nlargest(5, "total")[["person", "total"]].reset_index(drop=True)
    fewest = agg.nsmallest(5, "total")[["person", "total"]].reset_index(drop=True)
    best_quality = agg.nsmallest(10, "quality_ratio")[["person", "quality_ratio"]].reset_index(drop=True)
    worst_quality = agg.nlargest(5, "quality_ratio")[["person", "quality_ratio"]].reset_index(drop=True)

    difficulty = metrics_df.groupby("team", as_index=False)["total"].sum()
    difficulty = difficulty.sort_values("total", ascending=False).reset_index(drop=True)
    difficulty.columns = ["team", "total_contributions"]

    return {
        "highest_contributions": highest,
        "fewest_contributions": fewest,
        "best_quality": best_quality,
        "worst_quality": worst_quality,
        "most_difficult_articles": difficulty,
    }


def build_detail_table(metadata_df, variant_to_canonical):
    """Build the per-team detail table (footnote + person, every row)."""
    detail = metadata_df[["team", "footnote", "modified_by"]].copy()
    detail = detail.sort_values(["team", "footnote", "modified_by"])
    return detail
