"""
fills_core.py - the engine behind Stonebridge Fills Maker.

Logs into WellSky, reads the live weekly schedule, matches clients to regions
using the client list, and builds one printable Excel calendar per region.
Everything is done in memory (no files are written to disk).
"""

import io
from datetime import datetime, timedelta
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIG
# ==========================================

WEEKLY_URL = "https://stonebridgesouth.clearcareonline.com/dashboard/live/weekly/?weeks=1"

# Event classes that count as "needs coverage"
SHIFT_CLASSES = ("OPEN_SHIFT", "CANCELLED_BY_OFFICE")

# Keyed by the label shown in the app
REGIONS = {
    "Weber/Davis/Logan": {
        "title": "Weber/Davis/Logan Fills",
        "tags": ["Weber/Davis Counties", "Cache/Box Elder Counties"],
        "filename": "Weber-Davis-Logan Fills.xlsx",
    },
    "Salt Lake County": {
        "title": "Salt Lake County Fills",
        "tags": ["Salt Lake County"],
        "filename": "Salt Lake County Fills.xlsx",
    },
    "Utah County": {
        "title": "Utah County Fills",
        "tags": ["Utah County"],
        "filename": "Utah County Fills.xlsx",
    },
}

DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
MIN_GRID_ROWS = 6
REQUIRED_CLIENT_COLUMNS = ["Last Name", "First Name", "City", "Tags"]


class FillsError(Exception):
    """A problem the user can understand and act on."""


# ==========================================
# HELPERS
# ==========================================

def ordinal(n: int) -> str:
    if 11 <= n % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def short_name(last: str, first: str) -> str:
    """'Smith', 'Jane' -> 'Jane S.'"""
    return f"{first} {last[0]}." if last else first


def get_regions(tags: str, selected: list[str]) -> list[str]:
    return [
        name
        for name in selected
        if any(phrase in tags for phrase in REGIONS[name]["tags"])
    ]


# ==========================================
# WELLSKY LOGIN + FETCH
# ==========================================

def looks_like_schedule(resp) -> bool:
    return "sched_row" in resp.text and "login" not in resp.url.lower()


def fetch_weekly_html(username: str, password: str) -> str:
    """
    Log into WellSky the way a browser would (load the login page, fill in the
    form including hidden fields, submit) and return the weekly schedule HTML.
    """
    session = requests.Session()
    session.headers["User-Agent"] = "Mozilla/5.0 (StonebridgeFillsMaker)"

    try:
        resp = session.get(WEEKLY_URL, timeout=30)
        resp.raise_for_status()

        if not looks_like_schedule(resp):
            soup = BeautifulSoup(resp.text, "html.parser")
            pw_input = soup.find("input", {"type": "password"})
            form = pw_input.find_parent("form") if pw_input else None
            if not form:
                raise FillsError("Couldn't find the WellSky login form.")

            data = {}
            for inp in form.find_all(["input", "select", "textarea"]):
                name = inp.get("name")
                itype = (inp.get("type") or "").lower()
                if not name or itype in ("submit", "button", "image", "file"):
                    continue
                if itype in ("checkbox", "radio") and not inp.has_attr("checked"):
                    continue
                data[name] = inp.get("value", "")

            user_input = None
            for inp in form.find_all("input"):
                if inp is pw_input:
                    break
                if (inp.get("type") or "text").lower() in ("text", "email") and inp.get("name"):
                    user_input = inp
            if user_input is None:
                raise FillsError("Couldn't find the username field on the WellSky login form.")

            data[user_input["name"]] = username
            data[pw_input["name"]] = password

            action = urljoin(resp.url, form.get("action") or resp.url)
            if (form.get("method") or "post").lower() == "get":
                session.get(action, params=data, timeout=30)
            else:
                session.post(action, data=data, timeout=30)

            resp = session.get(WEEKLY_URL, timeout=30)
            resp.raise_for_status()
    except requests.RequestException as e:
        raise FillsError(f"Couldn't reach WellSky: {e}") from e

    if not looks_like_schedule(resp):
        raise FillsError(
            "Logged into WellSky, but didn't get the schedule back. "
            "The WellSky login saved in the app may be wrong or expired."
        )
    return resp.text


# ==========================================
# CLIENT LIST
# ==========================================

def build_client_lookup(client_file) -> dict:
    """client_file: a path or file-like object. -> {(last, first): {city, tags}}"""
    try:
        df = pd.read_excel(client_file).fillna("")
    except Exception as e:
        raise FillsError(f"Couldn't read the client file as an Excel sheet: {e}") from e

    missing = [c for c in REQUIRED_CLIENT_COLUMNS if c not in df.columns]
    if missing:
        raise FillsError(
            "The client file is missing these columns: " + ", ".join(missing)
            + ". Make sure it's the right report."
        )

    lookup = {}
    for _, row in df.iterrows():
        key = (
            str(row["Last Name"]).strip().lower(),
            str(row["First Name"]).strip().lower(),
        )
        lookup[key] = {
            "city": str(row["City"]).strip(),
            "tags": str(row["Tags"]).strip(),
        }
    return lookup


# ==========================================
# PARSING
# ==========================================

def has_shift_class(css_classes) -> bool:
    return bool(css_classes) and any(c in css_classes for c in SHIFT_CLASSES)


def extract_shifts(html: str, client_lookup: dict, selected: list[str]):
    """Returns (shifts, sorted list of schedule clients missing from the client file)."""
    soup = BeautifulSoup(html, "html.parser")
    rows = soup.find_all("tr", class_="sched_row")
    if not rows:
        raise FillsError("No client rows found on the schedule page.")

    shifts = []
    unmatched = set()

    for row in rows:
        person_div = row.find("div", class_="person-name")
        a_tag = person_div.find("a") if person_div else None
        if not a_tag:
            continue

        client_text = a_tag.get_text(strip=True)
        if "," not in client_text:
            continue
        last, first = (x.strip() for x in client_text.split(",", 1))

        info = client_lookup.get((last.lower(), first.lower()))
        if not info:
            # Only worth flagging if the client actually has a shift needing coverage
            if row.find("div", class_=has_shift_class):
                unmatched.add(client_text)
            continue

        regions = get_regions(info["tags"], selected)
        if not regions:
            continue

        for event in row.find_all("div", class_=has_shift_class):
            start = event.get("data-start")
            if not start:
                continue
            try:
                dt = datetime.strptime(start[:19], "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue

            time_span = event.find("span", class_="time")
            time_text = time_span.get_text(strip=True) if time_span else ""

            for region in regions:
                shifts.append({
                    "region": region,
                    "datetime": dt,
                    "day": dt.strftime("%A"),
                    "client": short_name(last, first),
                    "city": info["city"],
                    "time": time_text,
                })

    return shifts, sorted(unmatched)


# ==========================================
# WORKBOOKS
# ==========================================

def build_headers(shifts: list[dict]) -> list[str]:
    """One header per weekday, Sunday..Saturday, so columns always line up."""
    if not shifts:
        return list(DAYS)
    earliest = min(s["datetime"] for s in shifts).date()
    sunday = earliest - timedelta(days=(earliest.weekday() + 1) % 7)
    headers = []
    for i, day in enumerate(DAYS):
        d = sunday + timedelta(days=i)
        headers.append(f"{day}\n{d.strftime('%b')}. {ordinal(d.day)}")
    return headers


def build_workbook_bytes(region: str, shifts: list[dict], headers: list[str]) -> bytes:
    day_data = {d: [] for d in DAYS}
    for s in shifts:
        if s["region"] == region:
            day_data[s["day"]].append(s)
    for items in day_data.values():
        items.sort(key=lambda x: x["datetime"])

    grid_rows = max(MIN_GRID_ROWS, max(len(v) for v in day_data.values()))

    wb = Workbook()
    ws = wb.active
    ws.title = "Open Shifts"
    ws.oddHeader.center.text = f"&B&16{REGIONS[region]['title']}"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")

    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=text)
        cell.font = Font(bold=True, size=14)
        cell.alignment = center
        cell.fill = header_fill
        cell.border = border
    ws.row_dimensions[1].height = 65

    for col, day in enumerate(DAYS, start=1):
        for i in range(grid_rows):
            row_num = i + 2
            cell = ws.cell(row=row_num, column=col)
            cell.alignment = center
            cell.font = Font(size=12)
            cell.border = border
            if i < len(day_data[day]):
                s = day_data[day][i]
                cell.value = f"{s['client']}\n{s['city']}\n{s['time']}"
            ws.row_dimensions[row_num].height = 85

    for col in range(1, len(DAYS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 24

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    ws.print_options.horizontalCentered = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==========================================
# ONE CALL FOR THE APP
# ==========================================

def make_fills(client_file, selected: list[str], wellsky_username: str, wellsky_password: str) -> dict:
    """
    Returns {
        "files": {region_label: (filename, xlsx_bytes)},
        "counts": {region_label: number_of_open_shifts},
        "unmatched": [clients on the schedule but not in the client file],
    }
    """
    lookup = build_client_lookup(client_file)
    html = fetch_weekly_html(wellsky_username, wellsky_password)
    shifts, unmatched = extract_shifts(html, lookup, selected)
    headers = build_headers(shifts)

    files, counts = {}, {}
    for region in selected:
        files[region] = (
            REGIONS[region]["filename"],
            build_workbook_bytes(region, shifts, headers),
        )
        counts[region] = sum(1 for s in shifts if s["region"] == region)

    return {"files": files, "counts": counts, "unmatched": unmatched}