"""Advanced analytics for editing data: timelines, heatmaps, overlap, comment scoring."""

import re
from collections import defaultdict
from datetime import datetime
import pandas as pd


RULE_PATTERNS = [
    re.compile(r"\bBB\b", re.IGNORECASE),
    re.compile(r"\bBluebook\b", re.IGNORECASE),
    re.compile(r"\bCMOS\b", re.IGNORECASE),
    re.compile(r"\bChicago Manual\b", re.IGNORECASE),
    re.compile(r"\b\d+\.\d+", re.IGNORECASE),  # Rule references like "15.4"
    re.compile(r"\bsupra\b", re.IGNORECASE),
    re.compile(r"\bid\.\b", re.IGNORECASE),
    re.compile(r"\bhereinafter\b", re.IGNORECASE),
]

SUBSTANCE_KEYWORDS = [
    re.compile(r"\bcitation\b", re.IGNORECASE),
    re.compile(r"\bformat", re.IGNORECASE),
    re.compile(r"\bitalic", re.IGNORECASE),
    re.compile(r"\bcapitaliz", re.IGNORECASE),
    re.compile(r"\bspacing\b", re.IGNORECASE),
    re.compile(r"\bpunctuat", re.IGNORECASE),
    re.compile(r"\babbreviat", re.IGNORECASE),
]


def compute_work_timeline(edits_df):
    """Compute per-person work timeline: first edit, last edit, total days active, edits per day.

    Returns DataFrame with: person, team, first_edit, last_edit, days_span, active_days, edits_per_active_day
    """
    if "date" not in edits_df.columns or edits_df["date"].isna().all():
        return pd.DataFrame()

    df = edits_df.dropna(subset=["date"]).copy()
    df["date_parsed"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date_parsed"])
    df["day"] = df["date_parsed"].dt.date

    results = []
    for (person, team), group in df.groupby(["person", "team"]):
        first = group["date_parsed"].min()
        last = group["date_parsed"].max()
        days_span = (last - first).days + 1
        active_days = group["day"].nunique()
        total_edits = len(group)
        edits_per_active_day = total_edits / active_days if active_days > 0 else 0

        results.append({
            "person": person,
            "team": team,
            "first_edit": first,
            "last_edit": last,
            "days_span": days_span,
            "active_days": active_days,
            "total_edits": total_edits,
            "edits_per_active_day": round(edits_per_active_day, 1),
        })

    return pd.DataFrame(results).sort_values("total_edits", ascending=False).reset_index(drop=True)


def compute_edit_velocity(edits_df):
    """Compute edits per day per person (for charting).

    Returns DataFrame with: person, team, date, edit_count
    """
    if "date" not in edits_df.columns:
        return pd.DataFrame()

    df = edits_df.dropna(subset=["date"]).copy()
    df["date_parsed"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date_parsed"])
    df["day"] = df["date_parsed"].dt.date

    velocity = df.groupby(["person", "team", "day"]).size().reset_index(name="edit_count")
    velocity = velocity.sort_values(["team", "person", "day"]).reset_index(drop=True)
    return velocity


def compute_footnote_heatmap(edits_df):
    """Compute edit density per footnote per team.

    Returns DataFrame with: team, footnote, total_edits, num_editors, edits_per_editor
    """
    df = edits_df[edits_df["footnote"] > 0].copy()

    heatmap = df.groupby(["team", "footnote"]).agg(
        total_edits=("person", "count"),
        num_editors=("person", "nunique"),
    ).reset_index()
    heatmap["edits_per_editor"] = round(heatmap["total_edits"] / heatmap["num_editors"], 1)
    heatmap = heatmap.sort_values(["team", "total_edits"], ascending=[True, False]).reset_index(drop=True)
    return heatmap


def compute_overlap_matrix(edits_df):
    """Compute which editors overlap on the same footnotes.

    Returns DataFrame with: team, person_a, person_b, shared_footnotes, overlap_edits
    """
    df = edits_df[edits_df["footnote"] > 0].copy()
    results = []

    for team, team_df in df.groupby("team"):
        # Build person → set of footnotes
        person_fns = {}
        for person, pgroup in team_df.groupby("person"):
            person_fns[person] = set(pgroup["footnote"].unique())

        people = sorted(person_fns.keys())
        for i in range(len(people)):
            for j in range(i + 1, len(people)):
                a, b = people[i], people[j]
                shared = person_fns[a] & person_fns[b]
                if shared:
                    # Count total edits in shared footnotes
                    shared_edits = len(team_df[
                        (team_df["footnote"].isin(shared)) &
                        (team_df["person"].isin([a, b]))
                    ])
                    results.append({
                        "team": team,
                        "person_a": a,
                        "person_b": b,
                        "shared_footnotes": len(shared),
                        "overlap_edits": shared_edits,
                    })

    return pd.DataFrame(results).sort_values("shared_footnotes", ascending=False).reset_index(drop=True)


def score_comments(comments_df):
    """Score comments by substance: whether they cite rules, provide reasoning, or are low-effort.

    Returns DataFrame with: person, team, footnote, text, score, category
    Categories: 'rule_citation', 'substantive', 'low_effort'
    """
    if comments_df.empty or "text" not in comments_df.columns:
        return pd.DataFrame()

    results = []
    for _, row in comments_df.iterrows():
        text = row.get("text", "") or ""
        score = 0
        category = "low_effort"

        # Check for rule citations (highest value)
        for pat in RULE_PATTERNS:
            if pat.search(text):
                score += 2
                category = "rule_citation"
                break

        # Check for substantive keywords
        if category != "rule_citation":
            for pat in SUBSTANCE_KEYWORDS:
                if pat.search(text):
                    score += 1
                    category = "substantive"
                    break

        # Length bonus
        if len(text) > 50:
            score += 1

        results.append({
            "person": row["person"],
            "team": row["team"],
            "footnote": row["footnote"],
            "text": text[:200],
            "score": score,
            "category": category,
        })

    df = pd.DataFrame(results)
    return df


def compute_comment_quality_summary(scored_comments):
    """Aggregate comment scores per person.

    Returns DataFrame with: person, team, total_comments, rule_citations, substantive, low_effort, avg_score
    """
    if scored_comments.empty:
        return pd.DataFrame()

    results = []
    for (person, team), group in scored_comments.groupby(["person", "team"]):
        total = len(group)
        rule_citations = len(group[group["category"] == "rule_citation"])
        substantive = len(group[group["category"] == "substantive"])
        low_effort = len(group[group["category"] == "low_effort"])
        avg_score = group["score"].mean()

        results.append({
            "person": person,
            "team": team,
            "total_comments": total,
            "rule_citations": rule_citations,
            "substantive": substantive,
            "low_effort": low_effort,
            "avg_score": round(avg_score, 2),
        })

    return pd.DataFrame(results).sort_values("avg_score", ascending=False).reset_index(drop=True)


def compute_deadline_adherence(edits_df, all_assignments):
    """Check if edits fall within the assignment window.

    NOTE: Many CSVs don't have clean date parsing, so this returns what it can.
    Returns DataFrame with: person, team, edits_before_window, edits_in_window, edits_after_window
    """
    # For now, return a simple first/last edit analysis
    # Full deadline parsing would need the CSV dates to be standardized
    return compute_work_timeline(edits_df)
