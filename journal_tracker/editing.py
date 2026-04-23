"""Parse .docx files for tracked changes and comments, attributing each to an author and footnote."""

import zipfile
from collections import defaultdict
from lxml import etree
import pandas as pd

from .config import normalize_name

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def prescan_authors(docx_path):
    """Quick extraction of all unique author names from a docx."""
    z = zipfile.ZipFile(docx_path)
    authors = set()
    for xml_name in ["word/document.xml", "word/footnotes.xml", "word/comments.xml"]:
        try:
            tree = etree.fromstring(z.read(xml_name))
            for elem in tree.iter():
                author = elem.get(f"{W}author")
                if author and author.strip():
                    authors.add(author.strip())
        except (KeyError, etree.XMLSyntaxError):
            pass
    z.close()
    return authors


def parse_docx(docx_path, variant_to_canonical):
    """Parse a .docx file and return all editing contributions.

    Returns a DataFrame with columns: footnote, person, location, type, date, text
      - location: 'below_line' (footnotes.xml), 'above_line' (document.xml body), or 'comment'
      - type: 'ins', 'del', or 'comment'
      - date: ISO timestamp (for edits/comments that have one)
      - text: comment text (only for comments, None otherwise)
    """
    z = zipfile.ZipFile(docx_path)

    # Build footnote reference map from document.xml (for locating above-line edits and comments)
    fn_ref_positions = _build_footnote_ref_map(z)

    below_line = _parse_footnote_edits(z, variant_to_canonical)
    above_line = _parse_body_edits(z, fn_ref_positions, variant_to_canonical)
    comments = _parse_comments(z, fn_ref_positions, variant_to_canonical)

    z.close()

    rows = below_line + above_line + comments
    if not rows:
        return pd.DataFrame(columns=["footnote", "person", "location", "type", "date", "text"])
    return pd.DataFrame(rows)


def _build_footnote_ref_map(z):
    """Walk document.xml body in order, recording position of each footnoteReference.

    Returns a dict mapping element position index → footnote id.
    Also stores the ordered list so we can assign elements to their nearest preceding footnote.
    """
    xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")

    # Walk all elements in document order, track footnote references
    fn_refs = []  # list of (element_index, fn_id)
    all_elements = list(body.iter())

    for idx, elem in enumerate(all_elements):
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag == "footnoteReference":
            fn_id = elem.get(f"{W}id")
            if fn_id:
                fn_refs.append((idx, int(fn_id)))

    return {"elements": all_elements, "fn_refs": fn_refs}


def _get_footnote_for_position(fn_ref_positions, element):
    """Given an element, find which footnote it's associated with (nearest preceding footnoteReference)."""
    all_elements = fn_ref_positions["elements"]
    fn_refs = fn_ref_positions["fn_refs"]

    try:
        elem_idx = all_elements.index(element)
    except ValueError:
        return 0

    current_fn = 0
    for ref_idx, fn_id in fn_refs:
        if ref_idx <= elem_idx:
            current_fn = fn_id
        else:
            break
    return current_fn


def _get_footnote_for_ancestor(fn_ref_positions, element):
    """Walk up the tree from element to find its position in the document, then map to footnote."""
    all_elements = fn_ref_positions["elements"]
    fn_refs = fn_ref_positions["fn_refs"]

    # Try the element itself, then ancestors
    node = element
    while node is not None:
        try:
            elem_idx = all_elements.index(node)
            current_fn = 0
            for ref_idx, fn_id in fn_refs:
                if ref_idx <= elem_idx:
                    current_fn = fn_id
                else:
                    break
            return current_fn
        except ValueError:
            node = node.getparent()
    return 0


def _parse_footnote_edits(z, variant_to_canonical):
    """Extract tracked changes from footnotes.xml."""
    xml = z.read("word/footnotes.xml")
    tree = etree.fromstring(xml)

    rows = []
    for fn in tree.findall(f".//{W}footnote"):
        fn_id = fn.get(f"{W}id")
        if fn_id in ("0", "-1"):
            continue
        fn_num = int(fn_id)

        for ins in fn.findall(f".//{W}ins"):
            author = ins.get(f"{W}author")
            date = ins.get(f"{W}date")
            if author:
                rows.append({
                    "footnote": fn_num,
                    "person": normalize_name(author, variant_to_canonical),
                    "location": "below_line",
                    "type": "ins",
                    "date": date,
                    "text": None,
                })

        for d in fn.findall(f".//{W}del"):
            author = d.get(f"{W}author")
            date = d.get(f"{W}date")
            if author:
                rows.append({
                    "footnote": fn_num,
                    "person": normalize_name(author, variant_to_canonical),
                    "location": "below_line",
                    "type": "del",
                    "date": date,
                    "text": None,
                })

    return rows


def _parse_body_edits(z, fn_ref_positions, variant_to_canonical):
    """Extract tracked changes from document.xml body (above-the-line)."""
    all_elements = fn_ref_positions["elements"]
    fn_refs = fn_ref_positions["fn_refs"]

    rows = []
    for idx, elem in enumerate(all_elements):
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag not in ("ins", "del"):
            continue

        author = elem.get(f"{W}author")
        if not author:
            continue
        date = elem.get(f"{W}date")

        # Determine footnote by position
        current_fn = 0
        for ref_idx, fn_id in fn_refs:
            if ref_idx <= idx:
                current_fn = fn_id
            else:
                break

        rows.append({
            "footnote": current_fn,
            "person": normalize_name(author, variant_to_canonical),
            "location": "above_line",
            "type": tag,
            "date": date,
            "text": None,
        })

    return rows


def _parse_comments(z, fn_ref_positions, variant_to_canonical):
    """Extract comments and map them to footnotes via their anchors in document.xml."""
    xml = z.read("word/comments.xml")
    tree = etree.fromstring(xml)

    comment_data = {}
    for c in tree.findall(f".//{W}comment"):
        cid = c.get(f"{W}id")
        author = c.get(f"{W}author")
        date = c.get(f"{W}date")
        texts = c.findall(f".//{W}t")
        text = " ".join(t.text for t in texts if t.text)
        if cid and author:
            comment_data[cid] = {"author": author, "date": date, "text": text}

    # Map comment IDs to footnotes via commentRangeStart in document.xml
    all_elements = fn_ref_positions["elements"]
    fn_refs = fn_ref_positions["fn_refs"]

    rows = []
    for idx, elem in enumerate(all_elements):
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag != "commentRangeStart":
            continue

        cid = elem.get(f"{W}id")
        if not cid or cid not in comment_data:
            continue

        current_fn = 0
        for ref_idx, fn_id in fn_refs:
            if ref_idx <= idx:
                current_fn = fn_id
            else:
                break

        cd = comment_data[cid]
        rows.append({
            "footnote": current_fn,
            "person": normalize_name(cd["author"], variant_to_canonical),
            "location": "comment",
            "type": "comment",
            "date": cd["date"],
            "text": cd["text"],
        })

    return rows


def compute_editing_metrics(edits_df, editing_assignments, team_name):
    """Compute per-person editing metrics for one team's article.

    Only includes people listed in editing_assignments for this team.
    People with 0 contributions still appear (explicitly showing they did nothing).

    Returns a DataFrame with columns:
        team, person, total, in_assigned, outside_assigned,
        other_edits_to_assigned, quality_ratio, above_line, below_line,
        comments, share_of_team
    """
    # Only count edits from assigned team members
    assigned_people = set(editing_assignments.keys())
    if not assigned_people:
        return pd.DataFrame()

    edits_df = edits_df[edits_df["person"].isin(assigned_people)].copy()
    team_total = len(edits_df)
    people = assigned_people

    results = []
    for person in people:
        person_edits = edits_df[edits_df["person"] == person]
        total = len(person_edits)
        above_line = len(person_edits[person_edits["location"] == "above_line"])
        below_line = len(person_edits[person_edits["location"] == "below_line"])
        comments = len(person_edits[person_edits["location"] == "comment"])

        assignment = editing_assignments.get(person)
        if assignment:
            fn_start = assignment["fn_start"]
            fn_end = assignment["fn_end"]
            in_assigned = len(person_edits[
                (person_edits["footnote"] >= fn_start) &
                (person_edits["footnote"] <= fn_end)
            ])
            outside_assigned = total - in_assigned

            others_in_range = edits_df[
                (edits_df["footnote"] >= fn_start) &
                (edits_df["footnote"] <= fn_end) &
                (edits_df["person"] != person)
            ]
            other_edits_to_assigned = len(others_in_range)
        else:
            in_assigned = total
            outside_assigned = 0
            other_edits_to_assigned = 0

        quality_ratio = other_edits_to_assigned / total if total > 0 else 0.0
        share_of_team = total / team_total if team_total > 0 else 0.0

        results.append({
            "team": team_name,
            "person": person,
            "total": total,
            "in_assigned": in_assigned,
            "outside_assigned": outside_assigned,
            "other_edits_to_assigned": other_edits_to_assigned,
            "quality_ratio": quality_ratio,
            "above_line": above_line,
            "below_line": below_line,
            "comments": comments,
            "share_of_team": share_of_team,
        })

    return pd.DataFrame(results)
