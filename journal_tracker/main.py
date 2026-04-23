"""Journal Editing Effort Tracker.

Usage:
    python -m journal_tracker.main --input input/ --output "output/JICL Audit.xlsx"
    python -m journal_tracker.main --input input/ --sharepoint   # pull SP metadata from SharePoint
"""

import argparse
from pathlib import Path
import pandas as pd

from .config import (
    auto_detect_aliases,
    load_name_aliases,
    load_team_assignments,
    get_source_pull_assignments,
    get_editing_assignments,
)
from .source_pulling import (
    load_metadata,
    compute_source_pull_metrics,
    compute_leaderboards,
    build_detail_table,
)
from .editing import prescan_authors, parse_docx, compute_editing_metrics
from .analytics import (
    compute_work_timeline,
    compute_edit_velocity,
    compute_footnote_heatmap,
    compute_overlap_matrix,
    score_comments,
    compute_comment_quality_summary,
)
from .report import write_full_report


def main():
    parser = argparse.ArgumentParser(description="Journal Effort Tracker")
    parser.add_argument("--input", default="input", help="Input directory")
    parser.add_argument("--output", default="output/JICL Audit.xlsx", help="Output path")
    parser.add_argument("--aliases", default=None, help="Path to name aliases YAML")
    parser.add_argument("--sharepoint", action="store_true",
                        help="Fetch source pull metadata from SharePoint instead of local CSVs")
    parser.add_argument("--sp-config", default=None,
                        help="Path to sharepoint.yaml config file")
    args = parser.parse_args()

    input_dir = Path(args.input)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    print("Loading name aliases...")
    variant_to_canonical, _ = load_name_aliases(args.aliases)

    print("Scanning for name variants...")
    from collections import Counter
    name_counts = Counter()
    for team_dir in sorted(input_dir.iterdir()):
        if not team_dir.is_dir():
            continue
        for docx_file in team_dir.glob("*.docx"):
            for author in prescan_authors(docx_file):
                name_counts[author] += 1
        for csv_file in team_dir.glob("*.csv"):
            import csv as csv_mod
            with open(csv_file, encoding="utf-8", errors="replace") as f:
                rows = list(csv_mod.reader(f))
            for i, row in enumerate(rows):
                cells = [c.strip().lower() for c in row]
                if "modified by" in cells:
                    col = cells.index("modified by")
                    for r in rows[i + 1:]:
                        if len(r) > col and r[col].strip():
                            name_counts[r[col].strip()] += 1
                    break

    auto_aliases = auto_detect_aliases(set(name_counts.keys()), name_counts)
    if auto_aliases:
        print(f"  Auto-detected {len(auto_aliases)} name variant(s):")
        for variant, canonical in sorted(auto_aliases.items()):
            print(f"    '{variant}' -> '{canonical}'")
            variant_to_canonical[variant.lower()] = canonical
    else:
        print("  No new variants detected.")

    print("Loading assignments from team CSVs...")
    all_assignments = load_team_assignments(input_dir, variant_to_canonical)
    for team, assignments in sorted(all_assignments.items()):
        sp = [a for a in assignments if a["phase"] == "Source Pull"]
        ed = [a for a in assignments if a["phase"] == "Editing"]
        print(f"  {team}: {len(sp)} source pull, {len(ed)} editing")

    # --- Source Pull ---
    sp_assignments_by_team = {}
    sp_metrics_df = None
    sp_leaderboards = None
    sp_detail_df = None

    if args.sharepoint:
        from .sharepoint import load_sharepoint_config, fetch_and_write_csvs
        print("\nFetching source pull metadata from SharePoint...")
        sp_config = load_sharepoint_config(args.sp_config)
        fetch_and_write_csvs(sp_config, input_dir, variant_to_canonical)

    print("\nLoading source pull metadata from team folders...")
    metadata_df = load_metadata(input_dir, variant_to_canonical)
    if not metadata_df.empty:
        print(f"  {len(metadata_df)} file records across {metadata_df['team'].nunique()} teams")

        all_sp_metrics = []
        for team_name in sorted(metadata_df["team"].unique()):
            team_metadata = metadata_df[metadata_df["team"] == team_name]
            metadata_people = set(team_metadata["modified_by"].unique())

            best_match_count = 0
            best_assigns = {}
            for t_name, t_assignments in all_assignments.items():
                sp_assigns = {a["person"]: a for a in t_assignments
                              if a["phase"] == "Source Pull" and a["round"] == 1}
                overlap = len(set(sp_assigns.keys()) & metadata_people)
                if overlap > best_match_count:
                    best_match_count = overlap
                    best_assigns = sp_assigns

            sp_assignments_by_team[team_name] = best_assigns
            if best_assigns:
                print(f"  {team_name}: matched SP assignments ({best_match_count} people)")
            metrics = compute_source_pull_metrics(team_metadata, best_assigns)
            if not metrics.empty:
                all_sp_metrics.append(metrics)

        if all_sp_metrics:
            sp_metrics_df = pd.concat(all_sp_metrics, ignore_index=True)
            sp_leaderboards = compute_leaderboards(sp_metrics_df)
        sp_detail_df = build_detail_table(metadata_df, variant_to_canonical)
    else:
        print("  No source pull metadata found.")

    # --- Editing ---
    print("\nParsing .docx files for tracked changes...")
    all_editing_metrics = []
    all_editing_details = {}
    all_edits_combined = []
    ed_assignments_by_team = {}

    for team_dir in sorted(input_dir.iterdir()):
        if not team_dir.is_dir():
            continue
        team_name = team_dir.name

        docx_files = list(team_dir.glob("*.docx"))
        if not docx_files:
            print(f"  {team_name}: no .docx found, skipping")
            continue

        docx_path = docx_files[0]
        print(f"  {team_name}: parsing {docx_path.name}...")
        edits_df = parse_docx(docx_path, variant_to_canonical)
        edits_df["team"] = team_name
        print(f"    {len(edits_df)} contributions found")

        # Determine which team's editing assignments apply to this article.
        # Some teams cross-assign (Team Kyle members edit Team Cameron's article).
        # Strategy: find which team's assigned editors have the most overlap
        # with people who actually edited this docx, then include ALL of
        # that team's assigned members (even those with 0 contributions).
        docx_people = set(edits_df["person"].unique())
        best_match_team = None
        best_match_count = 0
        team_editing_candidates = {}

        for t_name, t_assignments in all_assignments.items():
            ed_assigns = {a["person"]: a for a in t_assignments
                         if a["phase"] == "Editing" and a["round"] == 1}
            if not ed_assigns:
                continue
            overlap = len(set(ed_assigns.keys()) & docx_people)
            if overlap > best_match_count:
                best_match_count = overlap
                best_match_team = t_name
                team_editing_candidates = ed_assigns

        team_editing = team_editing_candidates
        ed_assignments_by_team[team_name] = team_editing

        metrics = compute_editing_metrics(edits_df, team_editing, team_name)
        if not metrics.empty:
            all_editing_metrics.append(metrics)
            assigned_people = set(team_editing.keys())
            team_edits = edits_df[edits_df["person"].isin(assigned_people)].copy()
            all_editing_details[team_name] = team_edits
            all_edits_combined.append(team_edits)

    ed_metrics_df = pd.concat(all_editing_metrics, ignore_index=True) if all_editing_metrics else None
    combined_edits = pd.concat(all_edits_combined, ignore_index=True) if all_edits_combined else pd.DataFrame()

    # --- Analytics ---
    analytics = {}
    if not combined_edits.empty:
        print("\nRunning analytics...")

        print("  Work timeline...")
        analytics["timeline"] = compute_work_timeline(combined_edits)

        print("  Edit velocity...")
        analytics["velocity"] = compute_edit_velocity(combined_edits)

        print("  Footnote heatmap...")
        analytics["heatmap"] = compute_footnote_heatmap(combined_edits)

        print("  Overlap matrix...")
        analytics["overlap"] = compute_overlap_matrix(combined_edits)

        print("  Comment scoring...")
        comments_only = combined_edits[combined_edits["type"] == "comment"].copy()
        scored = score_comments(comments_only)
        analytics["comment_scores"] = scored
        analytics["comment_quality"] = compute_comment_quality_summary(scored)

    # --- Report ---
    print(f"\nWriting report to {output_path}...")
    write_full_report(
        sp_metrics_df, sp_leaderboards, sp_detail_df,
        ed_metrics_df, all_editing_details,
        analytics,
        sp_assignments_by_team, ed_assignments_by_team,
        output_path,
    )

    # --- Summary ---
    if sp_leaderboards:
        print("\n=== Source Pull ===")
        print("Highest:")
        for _, row in sp_leaderboards["highest_contributions"].iterrows():
            print(f"  {row['person']}: {row['total']}")

    if ed_metrics_df is not None:
        print("\n=== Editing ===")
        agg = ed_metrics_df.groupby("person", as_index=False)["total"].sum()
        top5 = agg.nlargest(5, "total")
        print("Highest:")
        for _, row in top5.iterrows():
            print(f"  {row['person']}: {row['total']}")

    if "timeline" in analytics and not analytics["timeline"].empty:
        print("\n=== Work Timeline (Top 5 Active Days) ===")
        tl = analytics["timeline"].nlargest(5, "active_days")
        for _, row in tl.iterrows():
            print(f"  {row['person']}: {row['active_days']} days, {row['edits_per_active_day']} edits/day")

    if "comment_quality" in analytics and not analytics["comment_quality"].empty:
        print("\n=== Comment Quality (Top 5) ===")
        cq = analytics["comment_quality"].head(5)
        for _, row in cq.iterrows():
            print(f"  {row['person']}: {row['total_comments']} comments, {row['rule_citations']} cite rules, avg score {row['avg_score']}")

    print(f"\nDone. Report saved to {output_path}")


if __name__ == "__main__":
    main()
