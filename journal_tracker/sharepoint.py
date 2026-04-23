"""Pull source-pull file metadata from SharePoint via Microsoft Graph API.

Setup (one-time):
    1. Register an app in Azure AD (portal.azure.com → App registrations)
    2. Add API permission: Microsoft Graph → Sites.Read.All (application)
       or Files.Read.All (delegated)
    3. Create a client secret under Certificates & secrets
    4. Copy tenant ID, client ID, and client secret into config/sharepoint.yaml

Config file (config/sharepoint.yaml):
    tenant_id: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    client_id: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    client_secret: "your-secret-here"
    site_name: "YourOrg.sharepoint.com:/sites/JICLEditing"
    # Map each team to its SharePoint folder path within the document library
    team_folders:
        Team Cameron: "Source Pull/Team Cameron"
        Team Camille: "Source Pull/Team Camille"
        Team Kyle: "Source Pull/Team Kyle"
        Team Mariko: "Source Pull/Team Mariko"
        Team Ryne: "Source Pull/Team Ryne"
        Team Tristan: "Source Pull/Team Tristan"
"""

import csv
import re
from pathlib import Path

import requests
import yaml

from .config import normalize_name, CONFIG_DIR


FN_FROM_FILENAME = re.compile(r"(?:fn|footnote)\s*(\d+)", re.IGNORECASE)


def load_sharepoint_config(path=None):
    path = path or CONFIG_DIR / "sharepoint.yaml"
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def _get_access_token(config):
    url = f"https://login.microsoftonline.com/{config['tenant_id']}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type": "client_credentials",
        "client_id": config["client_id"],
        "client_secret": config["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
    })
    resp.raise_for_status()
    return resp.json()["access_token"]


def _get_site_id(token, site_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_name}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.json()["id"]


def _get_drive_id(token, site_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    drives = resp.json()["value"]
    if not drives:
        raise ValueError("No document libraries found on this site")
    return drives[0]["id"]


def _list_files(token, drive_id, folder_path):
    encoded = folder_path.replace(" ", "%20")
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded}:/children"
    headers = {"Authorization": f"Bearer {token}"}

    files = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        files.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return files


def _parse_footnote_from_name(filename):
    m = FN_FROM_FILENAME.search(filename)
    if m:
        return int(m.group(1))
    digits = re.search(r"(\d+)", filename)
    if digits:
        return int(digits.group(1))
    return 0


def fetch_metadata(config, variant_to_canonical):
    token = _get_access_token(config)
    site_id = _get_site_id(token, config["site_name"])
    drive_id = _get_drive_id(token, site_id)

    rows = []
    for team_name, folder_path in config.get("team_folders", {}).items():
        print(f"  SharePoint: listing {team_name} → {folder_path}")
        files = _list_files(token, drive_id, folder_path)

        for f in files:
            if f.get("folder"):
                continue
            name = f.get("name", "")
            modified_by = (
                f.get("lastModifiedBy", {})
                .get("user", {})
                .get("displayName", "")
            )
            if not modified_by:
                continue

            rows.append({
                "team": team_name,
                "footnote": _parse_footnote_from_name(name),
                "filename": name,
                "modified_by": normalize_name(modified_by, variant_to_canonical),
            })

    return rows


def fetch_and_write_csvs(config, output_dir, variant_to_canonical):
    rows = fetch_metadata(config, variant_to_canonical)

    output_path = Path(output_dir)
    by_team = {}
    for r in rows:
        by_team.setdefault(r["team"], []).append(r)

    written = []
    for team_name, team_rows in sorted(by_team.items()):
        team_dir = output_path / team_name
        team_dir.mkdir(parents=True, exist_ok=True)
        csv_path = team_dir / f"{team_name} - SharePoint.csv"

        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["Team", "Footnote", "Filename", "Modified By"])
            writer.writeheader()
            for r in sorted(team_rows, key=lambda x: x["footnote"]):
                writer.writerow({
                    "Team": r["team"],
                    "Footnote": r["footnote"],
                    "Filename": r["filename"],
                    "Modified By": r["modified_by"],
                })

        written.append(csv_path)
        print(f"  Wrote {len(team_rows)} records → {csv_path}")

    return written
