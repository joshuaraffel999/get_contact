import os
import time
import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

load_dotenv()
TOKEN = os.environ["GITHUB_TOKEN"]
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/vnd.github+json",
}

# ---------- SETTINGS: change these ----------
LOCATION = "New York"      # GitHub location search term
LANGUAGE = ""             # e.g. "python" (leave "" for any)
MIN_FOLLOWERS = 0         # e.g. 10 to skip inactive profiles
MAX_USERS = 100           # GitHub search returns at most 1000
OUTPUT_FILE = "github_users.xlsx"
# --------------------------------------------


def get(url, params=None):
    """GET with simple rate-limit handling."""
    while True:
        r = requests.get(url, headers=HEADERS, params=params, timeout=30)
        if r.status_code in (403, 429) and r.headers.get("X-RateLimit-Remaining") == "0":
            reset = int(r.headers.get("X-RateLimit-Reset", time.time() + 60))
            wait = max(reset - int(time.time()), 1) + 1
            print(f"Rate limit hit, waiting {wait}s...")
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json()


def search_users():
    q = f'location:"{LOCATION}"'
    if LANGUAGE:
        q += f" language:{LANGUAGE}"
    if MIN_FOLLOWERS:
        q += f" followers:>={MIN_FOLLOWERS}"

    logins, page = [], 1
    while len(logins) < MAX_USERS:
        data = get(
            "https://api.github.com/search/users",
            {"q": q, "per_page": 100, "page": page},
        )
        items = data.get("items", [])
        if not items:
            break
        logins += [i["login"] for i in items]
        page += 1
        time.sleep(2)  # be gentle with the search API (~30 req/min)
    return logins[:MAX_USERS]


def main():
    logins = search_users()
    print(f"Found {len(logins)} users, fetching profiles...")

    rows = []
    for i, login in enumerate(logins, 1):
        u = get(f"https://api.github.com/users/{login}")
        rows.append([
            u.get("login"),
            u.get("name"),
            u.get("location"),
            u.get("company"),
            u.get("email"),   # only filled if the user made it public
            u.get("blog"),
            u.get("bio"),
            u.get("followers"),
            u.get("public_repos"),
            u.get("html_url"),
        ])
        print(f"  {i}/{len(logins)} {login}")

    headers = ["Username", "Name", "Location", "Company", "Public email",
               "Website", "Bio", "Followers", "Public repos", "Profile URL"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Users"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)

    # Column widths, freeze header, add filter
    for idx, h in enumerate(headers, 1):
        longest = max([len(str(h))] + [len(str(r[idx - 1] or "")) for r in rows])
        ws.column_dimensions[get_column_letter(idx)].width = min(longest + 2, 50)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(OUTPUT_FILE)
    with_email = sum(1 for r in rows if r[4])
    print(f"Saved {len(rows)} users to {OUTPUT_FILE} ({with_email} with a public email)")


if __name__ == "__main__":
    main()