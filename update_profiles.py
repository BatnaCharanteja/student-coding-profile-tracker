"""
Student Coding Profile Updater
================================
Reads an Excel file, fetches live data from LeetCode, HackerRank, and GitHub
for every student, and writes the results back into the same file.

Usage:
    python update_profiles.py

Libraries:
    pip install openpyxl requests beautifulsoup4
"""

import sys
import time
import random
import re
from pathlib import Path

import openpyxl
import requests
from bs4 import BeautifulSoup
from config import GITHUB_TOKEN


# ─── Constants ────────────────────────────────────────────────────────────────

EXPECTED_HEADERS = [
    "HackerRank Profile URL",
    "Problem Solving",
    "C",
    "Python",
    "Java",
    "SQL",
    "LeetCode Profile URL",
    "Total Leetcode Problems Solved",
    "No.of Easy Problems Solved",
    "No.of Medium Problems Solved",
    "No.of Hard Problems Solved",
    "GitHub Profile URL",
    "No.of Repos in GitHub",
]

HACKERRANK_BADGES = ["Problem Solving", "C", "Python", "Java", "SQL"]

REQUEST_TIMEOUT = 15  # seconds

HEADERS_BROWSER = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}



# ─── Header Detection ────────────────────────────────────────────────────────

def detect_columns(ws):
    """
    Scan Row 1 and return a dict mapping header name → column index (1-based).
    Only headers that appear in EXPECTED_HEADERS are included.
    """
    col_map = {}
    for col_idx in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col_idx).value
        if cell_value is None:
            continue
        header = str(cell_value).strip().replace("\n", " ").replace("  ", " ")
        if header in EXPECTED_HEADERS:
            col_map[header] = col_idx
    return col_map


# ─── Username Extractors ─────────────────────────────────────────────────────

def extract_leetcode_username(url: str) -> str | None:
    """Extract username from https://leetcode.com/u/username/ or /username/"""
    if not url:
        return None
    url = url.strip().rstrip("/")
    # Pattern: /u/username  OR  /username
    match = re.search(r"leetcode\.com/u/([^/?#]+)", url)
    if match:
        return match.group(1)
    match = re.search(r"leetcode\.com/([^/?#]+)", url)
    if match:
        return match.group(1)
    return None


def extract_hackerrank_username(url: str) -> str | None:
    """Extract username from https://www.hackerrank.com/profile/username"""
    if not url:
        return None
    url = url.strip().rstrip("/")
    match = re.search(r"hackerrank\.com/profile/([^/?#]+)", url)
    if match:
        return match.group(1)
    return None


def extract_github_username(url: str) -> str | None:
    """Extract username from https://github.com/username"""
    if not url:
        return None
    url = url.strip().rstrip("/")
    match = re.search(r"github\.com/([^/?#]+)", url)
    if match:
        return match.group(1)
    return None


# ─── LeetCode Fetcher ────────────────────────────────────────────────────────

def fetch_leetcode(username: str) -> dict:
    """
    Query LeetCode's public GraphQL API.
    Returns {"total": int, "easy": int, "medium": int, "hard": int}
    or raises on failure.
    """
    query = """
    {
        matchedUser(username: "%s") {
            submitStats {
                acSubmissionNum {
                    difficulty
                    count
                }
            }
        }
    }
    """ % username

    resp = requests.post(
        "https://leetcode.com/graphql",
        json={"query": query},
        headers={
            **HEADERS_BROWSER,
            "Content-Type": "application/json",
            "Referer": f"https://leetcode.com/u/{username}/",
        },
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()

    data = resp.json()
    matched = data.get("data", {}).get("matchedUser")
    if matched is None:
        raise ValueError(f"LeetCode user '{username}' not found")

    ac_list = matched["submitStats"]["acSubmissionNum"]
    result = {"total": 0, "easy": 0, "medium": 0, "hard": 0}
    for item in ac_list:
        diff = item["difficulty"].lower()
        if diff == "all":
            result["total"] = item["count"]
        elif diff == "easy":
            result["easy"] = item["count"]
        elif diff == "medium":
            result["medium"] = item["count"]
        elif diff == "hard":
            result["hard"] = item["count"]
    return result


# ─── HackerRank Fetcher ──────────────────────────────────────────────────────

def fetch_hackerrank(username: str) -> dict:
    """
    Scrape HackerRank public profile page.
    Returns a dict like {"Problem Solving": 3, "C": 5, "Python": 0, ...}
    Stars range 0–5.  Missing badge → 0.
    """
    url = f"https://www.hackerrank.com/profile/{username}"
    resp = requests.get(url, headers=HEADERS_BROWSER, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    # Build a mapping from badge name → star count
    badge_stars = {}

    # Try multiple selectors that HackerRank has used over time
    # Method 1: Look for badge containers with hacker-badge class
    badge_containers = soup.select(".hacker-badge")
    for container in badge_containers:
        # Get badge title
        title_el = container.select_one(".badge-title") or container.select_one("text") or container.select_one("p")
        if not title_el:
            continue
        title = title_el.get_text(strip=True)

        # Count filled stars
        stars = len(container.select(".badge-star, .star-filled, .full-star"))
        if stars == 0:
            # Try counting via star SVGs or similar patterns
            star_section = container.select_one(".stars") or container.select_one(".badge-stars")
            if star_section:
                stars = len(star_section.select(".star-filled, .full-star, svg.star"))

        badge_stars[title] = stars

    # Method 2: Look for the badge cards used in newer layout
    if not badge_stars:
        badge_cards = soup.select("[class*='badge'], [class*='Badge']")
        for card in badge_cards:
            text = card.get_text(" ", strip=True)
            for badge_name in HACKERRANK_BADGES:
                if badge_name.lower() in text.lower():
                    # Count stars - look for filled star indicators
                    stars = 0
                    star_els = card.select("[class*='star'], [class*='Star']")
                    for star_el in star_els:
                        classes = " ".join(star_el.get("class", []))
                        if "full" in classes.lower() or "filled" in classes.lower() or "active" in classes.lower():
                            stars += 1
                    # If still no stars found, try parsing from text
                    if stars == 0:
                        star_match = re.search(r"(\d)\s*(?:star|⭐)", text, re.IGNORECASE)
                        if star_match:
                            stars = int(star_match.group(1))
                    badge_stars[badge_name] = stars

    # Method 3: Parse from the embedded JSON / script tags (most reliable)
    if not badge_stars:
        scripts = soup.find_all("script")
        for script in scripts:
            if script.string and "badges" in script.string:
                # Try to extract badge data from inline scripts
                for badge_name in HACKERRANK_BADGES:
                    pattern = rf'"{re.escape(badge_name)}"[^}}]*?"stars"\s*:\s*(\d+)'
                    match = re.search(pattern, script.string, re.IGNORECASE)
                    if match:
                        badge_stars[badge_name] = int(match.group(1))

    # Method 4: Use the HackerRank REST API (most reliable fallback)
    if not badge_stars:
        try:
            api_url = f"https://www.hackerrank.com/rest/hackers/{username}/badges"
            api_resp = requests.get(api_url, headers=HEADERS_BROWSER, timeout=REQUEST_TIMEOUT)
            if api_resp.status_code == 200:
                badges_data = api_resp.json()
                if isinstance(badges_data, dict) and "models" in badges_data:
                    for badge in badges_data["models"]:
                        badge_name = badge.get("badge_name", "") or badge.get("domain", "")
                        stars = badge.get("stars", 0) or badge.get("current_points", 0)
                        badge_stars[badge_name] = stars
                elif isinstance(badges_data, list):
                    for badge in badges_data:
                        badge_name = badge.get("badge_name", "") or badge.get("domain", "")
                        stars = badge.get("stars", 0) or badge.get("current_points", 0)
                        badge_stars[badge_name] = stars
        except Exception:
            pass  # Fallback failed, will return 0s

    # Alias map: normalise the many names HackerRank may return
    # Keys are lowercase; values are our canonical HACKERRANK_BADGES names.
    BADGE_ALIASES = {
        "problem solving":  "Problem Solving",
        "problem_solving":  "Problem Solving",
        "problemsolving":   "Problem Solving",
        "c":                "C",
        "c programming":    "C",
        "python":           "Python",
        "python3":          "Python",
        "java":             "Java",
        "sql":              "SQL",
        "mysql":            "SQL",
    }

    # Build final result — missing badges → 0
    result = {b: 0 for b in HACKERRANK_BADGES}

    for raw_key, stars in badge_stars.items():
        normalised = raw_key.strip().lower().replace("-", "_")
        # Check alias map first
        if normalised in BADGE_ALIASES:
            result[BADGE_ALIASES[normalised]] = stars
        else:
            # Fallback: partial / substring match against expected names
            for badge_name in HACKERRANK_BADGES:
                if badge_name.lower() in normalised or normalised in badge_name.lower():
                    result[badge_name] = stars
                    break

    return result


# ─── GitHub Fetcher ───────────────────────────────────────────────────────────

def fetch_github(username: str) -> int:
    """
    Call the GitHub public API and return public_repos count.
    Uses GITHUB_TOKEN for authentication if provided.
    """
    url = f"https://api.github.com/users/{username}"
    headers = {**HEADERS_BROWSER}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data.get("public_repos", 0)


# ─── Main Processing Logic ───────────────────────────────────────────────────

def process_workbook(filepath: Path):
    """Load workbook, process every student row, save."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    # Step 1 — Detect columns
    col_map = detect_columns(ws)
    print(f"\nDetected columns: {col_map}\n")

    missing = [h for h in EXPECTED_HEADERS if h not in col_map]
    if missing:
        print(f"⚠  Warning: These headers were NOT found in Row 1 and will be skipped:")
        for m in missing:
            print(f"   • {m}")
        print()

    total_rows = ws.max_row

    # Step 2 — Iterate over every student (Row 2 onwards)
    for row in range(2, total_rows + 1):
        # Try to show a student identifier (col A or B usually has a name/ID)
        student_label = ws.cell(row=row, column=1).value or ws.cell(row=row, column=2).value or ""
        print(f"Processing row {row}: {student_label} ...")

        # ── LeetCode ──────────────────────────────────────────────────────
        lc_url_col = col_map.get("LeetCode Profile URL")
        if lc_url_col:
            lc_url = ws.cell(row=row, column=lc_url_col).value
            if lc_url and str(lc_url).strip():
                username = extract_leetcode_username(str(lc_url))
                if username:
                    try:
                        lc = fetch_leetcode(username)
                        if "Total Leetcode Problems Solved" in col_map:
                            ws.cell(row=row, column=col_map["Total Leetcode Problems Solved"]).value = lc["total"]
                        if "No.of Easy Problems Solved" in col_map:
                            ws.cell(row=row, column=col_map["No.of Easy Problems Solved"]).value = lc["easy"]
                        if "No.of Medium Problems Solved" in col_map:
                            ws.cell(row=row, column=col_map["No.of Medium Problems Solved"]).value = lc["medium"]
                        if "No.of Hard Problems Solved" in col_map:
                            ws.cell(row=row, column=col_map["No.of Hard Problems Solved"]).value = lc["hard"]
                        print(f"  ✓ LeetCode — Total: {lc['total']}, Easy: {lc['easy']}, Med: {lc['medium']}, Hard: {lc['hard']}")
                    except Exception as e:
                        print(f"  ✗ LeetCode error: {e}")
                        for hdr in ["Total Leetcode Problems Solved", "No.of Easy Problems Solved",
                                    "No.of Medium Problems Solved", "No.of Hard Problems Solved"]:
                            if hdr in col_map:
                                ws.cell(row=row, column=col_map[hdr]).value = "Error"
                else:
                    print(f"  ⚠ Could not extract LeetCode username from URL")
            # else: URL is empty → skip, leave cells blank

        # ── HackerRank ────────────────────────────────────────────────────
        hr_url_col = col_map.get("HackerRank Profile URL")
        if hr_url_col:
            hr_url = ws.cell(row=row, column=hr_url_col).value
            if hr_url and str(hr_url).strip():
                username = extract_hackerrank_username(str(hr_url))
                if username:
                    try:
                        hr = fetch_hackerrank(username)
                        for badge_name in HACKERRANK_BADGES:
                            if badge_name in col_map:
                                ws.cell(row=row, column=col_map[badge_name]).value = hr[badge_name]
                        stars_str = ", ".join(f"{k}: {v}★" for k, v in hr.items())
                        print(f"  ✓ HackerRank — {stars_str}")
                    except Exception as e:
                        print(f"  ✗ HackerRank error: {e}")
                        for badge_name in HACKERRANK_BADGES:
                            if badge_name in col_map:
                                ws.cell(row=row, column=col_map[badge_name]).value = "Error"
                else:
                    print(f"  ⚠ Could not extract HackerRank username from URL")

        # ── GitHub ────────────────────────────────────────────────────────
        gh_url_col = col_map.get("GitHub Profile URL")
        if gh_url_col:
            gh_url = ws.cell(row=row, column=gh_url_col).value
            if gh_url and str(gh_url).strip():
                username = extract_github_username(str(gh_url))
                if username:
                    try:
                        repos = fetch_github(username)
                        if "No.of Repos in GitHub" in col_map:
                            ws.cell(row=row, column=col_map["No.of Repos in GitHub"]).value = repos
                        print(f"  ✓ GitHub — Repos: {repos}")
                    except Exception as e:
                        print(f"  ✗ GitHub error: {e}")
                        if "No.of Repos in GitHub" in col_map:
                            ws.cell(row=row, column=col_map["No.of Repos in GitHub"]).value = "Error"
                else:
                    print(f"  ⚠ Could not extract GitHub username from URL")

        # Small delay to avoid rate-limiting
        time.sleep(random.uniform(1.0, 2.0))

    # Step 3 — Save
    wb.save(filepath)
    print(f"\n✅ Done! File saved as {filepath.name}")


# ─── Entry Point ──────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Student Coding Profile Updater")
    print("=" * 60)

    filename = input("\nEnter the Excel filename (e.g. students.xlsx): ").strip()
    if not filename:
        print("No filename provided. Exiting.")
        sys.exit(1)

    # Resolve relative to script's own directory
    script_dir = Path(__file__).resolve().parent
    filepath = script_dir / filename

    if not filepath.exists():
        print(f"Error: File '{filepath}' not found.")
        sys.exit(1)

    if not filepath.suffix.lower() in (".xlsx", ".xlsm"):
        print("Error: Only .xlsx / .xlsm files are supported.")
        sys.exit(1)

    process_workbook(filepath)


if __name__ == "__main__":
    main()
