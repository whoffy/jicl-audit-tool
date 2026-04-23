# JICL Editing Audit Tool

Automated report builder that analyzes journal editing effort across teams and generates a formatted Excel audit report (`JICL Audit.xlsx`).

## How to run

```bash
python -m journal_tracker.main --input input --output "output/JICL Audit.xlsx"
```

With SharePoint metadata fetch (requires Azure AD setup + `config/sharepoint.yaml`):
```bash
python -m journal_tracker.main --input input --sharepoint
```

## Folder structure

```
journal-workflow/
├── input/                          ← One subfolder per team
│   ├── Team A/
│   │   ├── article.docx            ← Tracked-changes doc (editing data)
│   │   ├── Editing Assignments.csv ← Assignment CSVs (who edits which FNs)
│   │   └── Team A - Sheet1.csv     ← SP metadata (Team, Footnote, Filename, Modified By)
│   ├── Team B/
│   └── ...
├── output/                         ← Generated JICL Audit.xlsx
├── journal_tracker/                ← Python package
│   ├── main.py                     ← CLI entry point, orchestrates everything
│   ├── config.py                   ← Assignment CSV parsing, name normalization, section detection
│   ├── source_pulling.py           ← SP metadata loading, per-person metrics, leaderboards
│   ├── editing.py                  ← .docx XML parsing (tracked changes + comments)
│   ├── analytics.py                ← Timeline, velocity, heatmap, overlap, comment scoring
│   ├── report.py                   ← Excel report generation (openpyxl)
│   ├── sharepoint.py               ← SharePoint/Graph API metadata fetcher (scaffold, not yet active)
│   └── config/
│       ├── name_aliases.yaml       ← Canonical name → variant mappings
│       └── sharepoint.yaml         ← Azure AD credentials + team folder paths (not yet created)
└── CLAUDE.md
```

## Key concepts

- **Cross-team matching**: One team's members may edit another team's article. Assignment matching uses overlap between assigned editors and actual editors in the docx.
- **Name aliases**: People appear under different names in docx metadata vs CSVs. `config/name_aliases.yaml` maps variants to canonical names.
- **Section detection**: Assignment CSVs have headers like "Source Pull Round 1", "Editing Round 2". Regex patterns in `config.py` detect these to categorize assignments.
- **Metadata CSVs vs assignment CSVs**: Metadata CSVs have columns (Team, Footnote, Filename, Modified By) — these are SP file listings. Assignment CSVs list who is assigned to which footnote ranges. The parser distinguishes them automatically.

## Report output (JICL Audit.xlsx)

- **Combined Rankings** — SP + editing leaderboards, quality ratios, carries, fewest contributions, most difficult articles
- **Team Source Pull sheets** — per-team FN detail with color-coded ranges, leaderboard, breakdown
- **Team Editing sheets** — same layout plus comment quality, activity metrics, stacked bar chart
- **Edit Timeline** — line chart of daily edit activity for top contributors
- **Analytics sheets** — Work Timeline, Edit Velocity, Footnote Heatmap, Editor Overlap, Comment Quality, Comment Detail

## Formatting

Report matches JICL Editing Audit format: Arial font, size 13 bold section headers, C9DAF8 blue fill for sections, D9D2E9 purple for "Most Difficult Articles", thin borders on sub-headers, percentage formatting on ratios, freeze panes, color-coded team tabs.

## Dependencies

- pandas, openpyxl, lxml, pyyaml, requests (requests only needed for --sharepoint)
