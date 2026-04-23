# JICL Editing Audit Tool

This tool automatically builds the JICL Editing Audit Excel report. You give it the team folders with the Word documents and assignment spreadsheets, and it generates the full audit with leaderboards, breakdowns, charts, and analytics.

No coding knowledge is needed. You can run everything through Claude Code or Codex without touching the terminal yourself.

---

## Getting the Source Pull Metadata

The hardest part of the whole process is getting the source pull data — who last modified each file in each team's source pull folder. The article `.docx` and assignment CSVs are easy to grab, but the source pull metadata isn't sitting in a spreadsheet anywhere. Here's how to get it.

### Recommended: Claude Screen Extraction

**Use Claude with computer use (Claude in Chrome, Claude Desktop, etc.) to read the metadata directly off the OneDrive screen.** This requires zero setup, takes about 20–30 minutes for all teams, and produces a ready-to-paste table.

**What you need:** A Claude Pro, Max, or Team subscription ($20+/month) that includes computer use / the Chrome extension. If Anthropic changes the subscription tiers or removes the browser extension, this approach may stop working — but as of early 2025, it works well.

**How it works:** You open the OneDrive folder in your browser, give Claude a prompt telling it to read each file's name and "Modified By" column, and it walks through the folder and produces a tab-separated table you paste into a spreadsheet. You do one team at a time.

Here's the prompt. Swap out the team names for your year's teams:

---

> **TASK**
>
> Extract upload metadata for every file in each team's "Source Pulls" folder in OneDrive and produce a spreadsheet-ready table. I'll run you once per team.
>
> **TEAMS TO PROCESS**
>
> Team Alice, Team Bob, Team Carol *(list your teams here)*
>
> Start with Team Alice. After you finish a team, stop and wait for me to say "next" before moving on.
>
> **NAVIGATION**
>
> From the current OneDrive view, click into the team folder. Click into the Source Pulls subfolder. Check for another folder layer — if Source Pulls contains a subfolder instead of files, click into it. Keep going until you see the actual PDFs. Record the full path you ended up at.
>
> Switch to **List view** (not Tiles) so Modified and Modified By columns are visible. If "Modified By" isn't showing, right-click a column header → Column options → add "Modified By."
>
> **DATA TO CAPTURE (one row per file)**
>
> | Column | How to get it |
> |--------|--------------|
> | Team | The team name, e.g., Team Alice |
> | Footnote | The leading number in the filename. `001_Elias_...pdf` → 1. Strip leading zeros. |
> | Filename | Full filename with extension |
> | Folder Path | Display path with spaces around slashes, e.g., `Team Alice / Source Pulls / filename.pdf` |
> | Modified | Month and year only, e.g., `March 2026`. If OneDrive shows "Yesterday" or "2 days ago," hover to get the actual date. |
> | Modified By | Full name of the person shown in the Modified By column |
> | Notes | Leave blank |
>
> **OUTPUT FORMAT**
>
> After each team, output a tab-separated table with this header:
>
> `Team    Footnote    Filename    Folder Path    Modified    Modified By    Notes`
>
> One row per file. No extra commentary inside the table.
>
> **CRITICAL RULES**
>
> - Get **every** file. If the folder has 50+ files, scroll all the way to the bottom. If the list loads lazily, keep scrolling until the count stops increasing. Tell me the final file count so I can verify.
> - **Do not invent or guess data.** If Modified By is blank, put UNKNOWN.
> - **Do not shorten filenames.** Copy them verbatim.
> - **Exact dates only.** No "Yesterday," no dates without a year.
> - **One team per run.** Don't start the next team until I confirm.
>
> **WHEN YOU FINISH A TEAM**
>
> End with: ✅ Team [X] complete — N files captured. Uploaders found: [Name 1] (x files), [Name 2] (y files), ... Ready for next team?

---

**Gotchas from experience:**
- The prompt above is a refined version — the first attempt needed tweaking halfway through (mainly around date formatting and making sure Claude didn't skip files in long folders).
- Long folders that lazy-load can trip it up. Telling it to scroll and report the file count helps you verify nothing was missed.
- Folder structures aren't always consistent across teams — some have an extra subfolder layer. The prompt handles this but Claude will sometimes ask you to confirm.
- After Claude outputs each table, paste it into a Google Sheet or Excel, then download as CSV and drop it into the team's input folder. The tool needs columns: **Team, Footnote, Filename, Modified By** (the other columns are fine to leave in — the tool ignores them).

### If You're Feeling Ambitious: SharePoint Automation

If JICL ever moves its source pull files from OneDrive to SharePoint (or if they're already there), the tool can pull this metadata automatically with zero manual work each year. This requires a one-time Azure AD setup by someone with admin access — about an afternoon of work. After that, you just add `--sharepoint` to the run command and the metadata extraction is fully automated.

If someone on the board has admin access and wants to set this up, the full step-by-step instructions are in the **SharePoint Integration** section at the bottom of this README. It's a one-time investment that pays off every year after.

If JICL migrates from OneDrive to SharePoint for other reasons, this becomes the obvious default path and the Claude screen extraction is no longer needed.

---

## Step 1: Install the Tool (One-Time)

You only need to do this once, or when setting up on a new computer.

### What is Python?

Python is a programming language. This tool is written in Python, so you need it installed to run the tool. You don't need to know Python — you just need it on your computer.

### Do I have Python?

Open a terminal and type:

```
python --version
```

If you see something like `Python 3.11.5`, you're good. If you get an error, install it:

1. Go to https://www.python.org/downloads/
2. Click the big yellow **Download Python** button
3. Run the installer
4. **IMPORTANT:** Check the box that says **"Add Python to PATH"** before clicking Install
5. Restart your terminal after installing

### How do I open a terminal?

- **Windows:** Press the Windows key, type `cmd`, and hit Enter.
- **Mac:** Press Cmd+Space, type `Terminal`, and hit Enter.

### Install the dependencies

1. Open a terminal.
2. Navigate to the `journal-workflow` folder:
   ```
   cd Desktop/journal-workflow
   ```
   On Windows it might be:
   ```
   cd C:\Users\YourName\Desktop\journal-workflow
   ```
3. Run:
   ```
   pip install -r requirements.txt
   ```
   You'll see text scroll by. As long as it doesn't end with an error, you're fine.

---

## Step 2: Set Up Your Input Folders

Inside the `input/` folder, create one subfolder for each team. **The folder name becomes the team name in the report** — whatever you name them is what shows up.

Each team folder needs **three files**:

| File | What is it | Where to get it |
|------|-----------|----------------|
| `article.docx` | The Word document with Track Changes on | The article your team edited |
| Assignment CSV | Spreadsheet listing who's assigned to which footnote ranges | Download from your shared Google Sheet as CSV |
| Source pull metadata CSV | Spreadsheet showing who last modified each source pull file | Use Claude screen extraction (see above) |

Your folder should look like this:

```
input/
  Team Alice/
    article.docx
    Editing Assignments.csv
    Team Alice - Sheet1.csv
  Team Bob/
    article.docx
    Editing Assignments.csv
    Team Bob - Sheet1.csv
```

### About the Assignment CSV

This is the spreadsheet that says things like "Alice: Fn. 1-25, Bob: Fn. 26-50." The tool can read several different layouts:

- Sections stacked vertically (Source Pull Round 1, then Editing Round 1, etc.)
- Side-by-side panels (Source Pull on the left, Editing on the right)
- Different header styles ("Source Pull Round 1", "Rnd 2", etc.)

Just download the assignment sheet as a CSV and drop it in. The tool figures out the format.

---

## Step 3: Run the Report

### Option A: Using Claude Code or Codex (Recommended)

If you don't want to deal with the terminal, use an AI coding assistant. This is the easiest way.

**Using Claude Code:**

1. Install Claude Code: https://claude.ai/download (requires a Claude Pro/Max/Team subscription, $20+/month)
2. Open Claude Code and point it at the `journal-workflow` folder
3. Tell it:

   > "Install the dependencies and run the JICL audit report"

   Claude Code will read this README, understand the project, install anything missing, and run everything. It can also troubleshoot errors for you.

**Using Codex (OpenAI):**

1. Go to https://chatgpt.com/codex or open Codex in your IDE (requires a ChatGPT Plus/Pro subscription)
2. Point it at the `journal-workflow` folder
3. Tell it:

   > "Read the README and run the JICL audit report"

**What to tell the AI if something goes wrong:**

- **"Python not found"** → Ask it to install Python
- **"Module not found"** → Ask it to run `pip install -r requirements.txt`
- **Two people with similar names showing up separately** → "These two names are the same person, fix it"
- **A team's data is showing all zeros** → "Check if the assignment CSV is being parsed correctly"

### Option B: Run it yourself

1. Open a terminal and navigate to the `journal-workflow` folder.
2. Run:
   ```
   python -m journal_tracker.main
   ```
3. Open `output/JICL Audit.xlsx` in Excel. That's your report.

---

## How the Tool Handles Name Misspellings

Word saves author names based on each person's Microsoft account. If someone's name is spelled differently in different places (e.g., "Matt" vs "Matthew", "O'Neil" vs "O'Neill"), the tool handles this two ways:

1. **Automatic detection (no setup needed):** Before processing, the tool scans all files and looks for names that are almost identical. If it finds any, it merges them and tells you:
   ```
   Auto-detected 1 name variant(s):
     'Mathew Smith' -> 'Matthew Smith'
   ```

2. **Manual overrides (if auto-detection misses something):** Edit the file `journal_tracker/config/name_aliases.yaml`:
   ```yaml
   aliases:
     "Matthew Smith":
       - "Matt Smith"
       - "Mathew Smith"
   ```
   The name on the left is the correct version. The names on the right are variants that should be treated as the same person. This file is **optional** — if it's empty or missing, the tool still works.

---

## Year-to-Year Setup

**Nothing in the code needs to change between years.** New teams, new people, new articles — just:

1. Delete the old team folders from `input/` (or move them somewhere else)
2. Create new team folders with the new team names
3. Drop in the new files (article.docx, assignment CSV, source pull metadata CSV)
4. Run the report

The tool discovers everything from folder names and file contents. There's no place in the code where team names or people's names are written in.

If you used `name_aliases.yaml` last year, you can clear it out or just delete it. Old aliases won't cause errors — they just won't match anyone.

---

## What's in the Report

The output Excel file (`JICL Audit.xlsx`) contains these sheets:

| Sheet | What it shows |
|-------|--------------|
| **Combined Rankings** | The big picture — top and bottom contributors across all teams, quality ratios, who carried their team, most difficult articles |
| **Team ___ Source Pull** | Each team's source pull detail: every footnote and who modified it, color-coded by assigned range |
| **Team ___ Editing** | Each team's editing detail: tracked changes and comments per footnote, leaderboard, breakdown of above-line vs. below-line edits, comment quality, stacked bar chart |
| **Edit Timeline** | Line chart showing when each person was active over time |
| **Work Timeline** | Table version: first edit, last edit, days active, edits per day |
| **Edit Velocity** | Daily edit counts per person |
| **Footnote Heatmap** | Which footnotes got the most edits — helps spot problem areas |
| **Editor Overlap** | Which editors worked on the same footnotes |
| **Comment Quality** | Scores each person's comments: rule citations, substantive vs. low-effort |
| **Comment Detail** | Every individual comment with its score |

---

## SharePoint Integration (Optional — One-Time Setup)

If your source pull files are in SharePoint (not OneDrive), the tool can pull the metadata automatically. This replaces the Claude screen extraction step entirely — you just add `--sharepoint` to the command and the tool does the rest.

This requires a **one-time Azure AD setup** by someone with admin access to your organization's Azure portal. It's about an afternoon of work. Once done, it works every year until the secret expires (1-2 years), at which point you just generate a new secret (5 minutes).

### What you're doing

Creating a read-only "service account" that lets the tool see file metadata in your SharePoint site. It can only read — it can't change or delete anything.

### Setup Steps

1. Go to [portal.azure.com](https://portal.azure.com) and sign in with your organization's account (the same one you use for SharePoint/Outlook).

2. In the search bar at the top, type **"App registrations"** and click on it.

3. Click **"+ New registration"** at the top.
   - **Name:** Type `JICL Audit Tool`
   - **Supported account types:** Leave on the default ("Single tenant")
   - **Redirect URI:** Leave blank
   - Click **Register**

4. You're now on the app's overview page. Copy these two values somewhere safe:
   - **Application (client) ID** — a long string like `a1b2c3d4-e5f6-7890-abcd-ef1234567890`
   - **Directory (tenant) ID** — same format, different value

5. In the left sidebar, click **"Certificates & secrets"**.
   - Click **"+ New client secret"**
   - **Description:** `JICL Audit`
   - **Expires:** Pick 24 months
   - Click **Add**
   - **IMPORTANT:** Copy the **Value** column (NOT the Secret ID) immediately. You can never see it again after leaving this page.

6. In the left sidebar, click **"API permissions"**.
   - Click **"+ Add a permission"**
   - Click **"Microsoft Graph"** (the big one at the top)
   - Click **"Application permissions"** (NOT "Delegated permissions")
   - Search for `Sites.Read.All`, check the box, click **"Add permissions"**
   - Back on the permissions page, click **"Grant admin consent for [Your Org]"** and confirm
   - If this button is grayed out, you need your IT admin to click it for you

### Create the Config File

Create a file at `journal_tracker/config/sharepoint.yaml`:

```yaml
tenant_id: "paste-your-tenant-id-here"
client_id: "paste-your-client-id-here"
client_secret: "paste-your-secret-value-here"

# Your SharePoint URL looks like: https://yourorg.sharepoint.com/sites/JICLEditing
# Take the part after https:// and format it like this:
site_name: "yourorg.sharepoint.com:/sites/JICLEditing"

# List each team and their source pull folder path.
# Update these each year to match your team names.
team_folders:
  Team Alice: "Source Pull/Team Alice"
  Team Bob: "Source Pull/Team Bob"
  Team Carol: "Source Pull/Team Carol"
```

### Running with SharePoint

```
python -m journal_tracker.main --sharepoint
```

This pulls file metadata from SharePoint, then generates the report. You still need the `.docx` files and assignment CSVs in the input folders — SharePoint only replaces the metadata CSV.

### When the Secret Expires

You'll get an authentication error. To fix it:

1. Go to [portal.azure.com](https://portal.azure.com) → App registrations → JICL Audit Tool → Certificates & secrets
2. Create a new secret (same steps as above)
3. Update `client_secret` in `sharepoint.yaml`

---

## Troubleshooting

**"Python is not recognized" / "python: command not found"**
- Python isn't installed or wasn't added to PATH. Reinstall and check "Add Python to PATH." On some systems, try `python3` instead.

**"No module named 'pandas'" or similar**
- Run `pip install -r requirements.txt`

**A team's numbers are all zeros**
- Check that the assignment CSV is in the right team folder and has sections like "Editing Round 1" with names and footnote ranges.

**Two entries for the same person**
- Auto-detection uses ~85% name similarity. If the names are too different (e.g., a nickname vs. full name), add a manual alias in `name_aliases.yaml`, or tell Claude Code: "These two names are the same person, fix it."

**"No .docx found, skipping"**
- Make sure the Word document is in the team folder and has a `.docx` extension (not `.doc`).

**Report looks wrong or has missing data**
- Run the tool and read the console output. It prints what it finds for each team. If a team shows "0 source pull, 0 editing," the assignment CSV isn't being parsed correctly. Ask Claude Code or Codex to debug it.
