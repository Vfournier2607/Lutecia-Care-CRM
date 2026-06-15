import azure.functions as func
import json
import logging
import re
from datetime import datetime, date

import requests
from bs4 import BeautifulSoup

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

# ─── HTTP headers ────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9,fr;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# ─── Interest keywords (titre de l'événement) ────────────────────────────────
KW = {
    "theranostique": [
        "theranostic", "theranostics", "theranostique", "lutetium", "177lu",
        "lu-177", "actinium", "225ac", "radioligand", "radiopharmaceutical therapy",
        "psma", "dotatate", "dotatoc", "peptide receptor",
    ],
    "medecine-nucleaire": [
        "nuclear medicine", "nuclear imaging", "radiopharmaceutical", "pet-ct",
        "pet/ct", "spect", "spect/ct", "scintigraphy", "dosimetry", "radionuclide",
        "nuclear oncology", "eanm", "snmmi", "jfmn", "aimn", "dgn", "semnim",
        "molecular imaging", "pet scan",
    ],
    "radiotherapie": [
        "radiotherapy", "radiation therapy", "radiation oncology", "estro",
        "astro", "degro", "proton therapy", "brachytherapy", "sbrt",
    ],
    "formation": [
        "training course", "training programme", "workshop", "summer school",
        "school on", "hands-on", "practical course", "acoramen",
    ],
    "oncologie": [
        "oncolog", "cancer", "tumor", "tumour", "malign", "esmo", "asco",
        "aiom", "carcinoma", "lymphoma",
    ],
    "imagerie": [
        "radiology", "radiolog", "imaging", "rsna", "ecr", "mri", "ct scan",
        "magnetic resonance", "ultrasound", "mammograph",
    ],
    "medtech": [
        "medtech", "medical device", "medica", "digital health", "healthtech",
        "medical technology", "health innovation", "hlth",
    ],
}

# All interest keywords flat for fast pre-filter
ALL_KW = [kw for lst in KW.values() for kw in lst]

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


# ─── Helpers ─────────────────────────────────────────────────────────────────

def infer_category(title: str, notes: str = "") -> str:
    text = (title + " " + notes).lower()
    for cat, keywords in KW.items():
        if any(kw in text for kw in keywords):
            return cat
    return "healthcare"


def _parse_iaea_date(raw: str):
    """Return (startDate, endDate) as 'YYYY-MM-DD' strings, or (None, None)."""
    s = raw.strip().replace("–", "-").replace("—", "-").replace("–", "-")
    # "28 - 30 Apr 2025"
    m = re.match(r"(\d{1,2})\s*-\s*(\d{1,2})\s+(\w{3})\w*\s+(\d{4})", s)
    if m:
        d1, d2, mon, yr = m.groups()
        mo = MONTHS.get(mon.lower())
        if mo:
            return f"{yr}-{mo:02d}-{int(d1):02d}", f"{yr}-{mo:02d}-{int(d2):02d}"
    # "28 Apr - 2 May 2025"
    m = re.match(r"(\d{1,2})\s+(\w{3})\w*\s*-\s*(\d{1,2})\s+(\w{3})\w*\s+(\d{4})", s)
    if m:
        d1, mon1, d2, mon2, yr = m.groups()
        m1, m2 = MONTHS.get(mon1.lower()), MONTHS.get(mon2.lower())
        if m1 and m2:
            return f"{yr}-{m1:02d}-{int(d1):02d}", f"{yr}-{m2:02d}-{int(d2):02d}"
    # "28 Apr 2025 - 3 Apr 2026"
    m = re.match(
        r"(\d{1,2})\s+(\w{3})\w*\s+(\d{4})\s*-\s*(\d{1,2})\s+(\w{3})\w*\s+(\d{4})", s
    )
    if m:
        d1, mon1, yr1, d2, mon2, yr2 = m.groups()
        m1, m2 = MONTHS.get(mon1.lower()), MONTHS.get(mon2.lower())
        if m1 and m2:
            return f"{yr1}-{m1:02d}-{int(d1):02d}", f"{yr2}-{m2:02d}-{int(d2):02d}"
    return None, None


def _is_relevant(title: str) -> bool:
    t = title.lower()
    return any(kw in t for kw in ALL_KW)


# ─── IAEA scraper ────────────────────────────────────────────────────────────

def scrape_iaea(max_pages: int = 10) -> tuple[list, list]:
    events, errors = [], []
    today_str = str(date.today())

    for page in range(max_pages):
        url = "https://www.iaea.org/events" if page == 0 else f"https://www.iaea.org/events?page={page}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code != 200:
                errors.append(f"IAEA page {page}: HTTP {r.status_code}")
                break

            soup = BeautifulSoup(r.text, "lxml")

            # Try to find structured event containers
            items = (
                soup.find_all("div", class_=re.compile(r"event[_-]?(item|card|listing|row)", re.I))
                or soup.find_all("article", class_=re.compile(r"event", re.I))
            )

            if items:
                for item in items:
                    text = item.get_text(separator="|", strip=True)
                    parts = [p.strip() for p in text.split("|") if p.strip()]
                    # Heuristic: first non-empty part with a year is the date
                    date_part, title_part, loc_part = "", "", ""
                    for p in parts:
                        if re.search(r"\d{4}", p) and not date_part:
                            date_part = p
                        elif len(p) > 15 and not any(c in p for c in ["Event Code", "EVT"]):
                            if not title_part:
                                title_part = p
                        elif "," in p and not loc_part:
                            loc_part = p

                    start, end = _parse_iaea_date(date_part)
                    if start and start >= today_str and title_part and _is_relevant(title_part):
                        city, _, country = loc_part.partition(",")
                        events.append({
                            "title": title_part,
                            "startDate": start,
                            "endDate": end or start,
                            "city": city.strip(),
                            "country": country.strip(),
                            "category": infer_category(title_part),
                            "notes": f"Source IAEA. {loc_part}".strip(". "),
                            "source": "IAEA",
                        })
            else:
                # Fallback: parse plain text line by line
                lines = [l.strip() for l in soup.get_text(separator="\n").split("\n") if l.strip()]
                i = 0
                while i < len(lines):
                    ln = lines[i]
                    date_m = re.match(r"(\d{1,2}[\s––\-].*?\d{4})", ln)
                    if date_m:
                        start, end = _parse_iaea_date(date_m.group(1))
                        if start and start >= today_str and i + 2 < len(lines):
                            loc = lines[i + 1]
                            j = i + 2
                            # skip "Event Code" or "Virtual component" lines
                            while j < len(lines) and (
                                "Event Code" in lines[j] or "Virtual" in lines[j] or re.match(r"EVT\d+", lines[j])
                            ):
                                j += 1
                            title = lines[j] if j < len(lines) else ""
                            if title and _is_relevant(title):
                                city, _, country = loc.partition(",")
                                events.append({
                                    "title": title,
                                    "startDate": start,
                                    "endDate": end or start,
                                    "city": city.strip(),
                                    "country": country.strip(),
                                    "category": infer_category(title),
                                    "notes": f"Source IAEA. {loc}".strip(". "),
                                    "source": "IAEA",
                                })
                            i = j + 1
                            continue
                    i += 1

            # Stop if no "next" page link
            if not soup.find(string=re.compile(r"next", re.I)):
                break

        except requests.exceptions.RequestException as exc:
            errors.append(f"IAEA page {page}: {exc}")
            break

    return events, errors


# ─── Dedup ───────────────────────────────────────────────────────────────────

def dedup(events: list) -> list:
    seen, out = set(), []
    for ev in events:
        key = re.sub(r"[^a-z0-9]", "", ev["title"].lower())[:35]
        if key not in seen:
            seen.add(key)
            out.append(ev)
    return out


# ─── Azure Function entrypoint ───────────────────────────────────────────────

@app.route(route="scrape", methods=["GET", "OPTIONS"])
def scrape(req: func.HttpRequest) -> func.HttpResponse:
    cors = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
    }

    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers=cors)

    logging.info("Scraper triggered")
    all_events, all_errors = [], []

    # IAEA
    try:
        evs, errs = scrape_iaea(max_pages=10)
        all_events.extend(evs)
        all_errors.extend(errs)
        logging.info(f"IAEA: {len(evs)} relevant events")
    except Exception as exc:
        all_errors.append(f"IAEA: {exc}")

    unique = dedup(all_events)

    payload = json.dumps(
        {
            "events": unique,
            "count": len(unique),
            "errors": all_errors,
            "scraped_at": datetime.utcnow().isoformat() + "Z",
        },
        ensure_ascii=False,
        indent=2,
    )

    return func.HttpResponse(payload, mimetype="application/json", headers=cors)
