import os
import re
import sys
import hashlib
import posixpath
import zipfile
import zoneinfo
import xml.etree.ElementTree as ET
from datetime import datetime, date, time, timedelta
from urllib.parse import urljoin, quote, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup
import openpyxl

# ---------------------------------------------------------------- KONFIGURACJA
# Numer albumu: najlepiej ustaw jako sekret INDEX_NUMBER w GitHub Actions.
INDEX_NUMBER = os.environ.get("INDEX_NUMBER") or "21459"

OUTPUT_ICS = "plan_zajec_IOSI.ics"
EXCEL_FILE = "pobrany_plan.xlsx"
URL_STRONY_PLANU = "https://ans-elblag.pl/iis-plany-zajec.html"

SEMESTER_START_MONDAY = date(2026, 10, 5)
SEARCH_LIMIT = date(2027, 2, 15)      # koniec semestru - dalej nie generujemy zajec
SESSIONS_WEEKLY = 15                  # 15w = 15 wykladow, 30l = 15 spotkan po 2h

# Dni bez zajec (ZALOZENIE - popraw, jesli uczelnia ma inaczej)
OFF_DATES = {date(2026, 11, 11), date(2027, 1, 6)}
d = date(2026, 12, 23)
while d <= date(2027, 1, 1):          # przerwa swiateczna
    OFF_DATES.add(d)
    d += timedelta(days=1)

# W tym dniu obowiazuje plan z innego dnia (0=pon ... 4=pt).
# Plan: "22 grudnia 2026 r. zajecia wg planu ze srody"
SWAP_DAYS = {date(2026, 12, 22): 2}

TZ = zoneinfo.ZoneInfo("Europe/Warsaw")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

DAY_NAMES = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek"]
DEFAULT_DAY_COLS = [1, 16, 40, 55, 72]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def log(msg):
    print(msg, flush=True)


# ------------------------------------------------------------------ POBIERANIE
def download_excel_from_ans():
    if os.environ.get("SKIP_DOWNLOAD"):
        return False
    try:
        res = requests.get(URL_STRONY_PLANU, headers=HEADERS, timeout=10)
    except Exception as e:
        log(f"[BLAD] Nie mozna pobrac strony planow: {type(e).__name__}")
        log("[INFO] Wgraj plik recznie do repo jako pobrany_plan.xlsx")
        return False

    log(f"[INFO] Status strony planow: {res.status_code}")
    if res.status_code != 200:
        return False

    soup = BeautifulSoup(res.text, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.lower().endswith((".xlsx", ".xls")):
            links.append((href, a.get_text(" ", strip=True)))
    log(f"[INFO] Linki do Excela: {len(links)}")
    for href, text in links:
        log(f"   - {href} | {text}")

    for href, text in links:
        if "iosi" in href.lower() or "iosi" in text.lower():
            full = urljoin(URL_STRONY_PLANU, href)
            parts = urlsplit(full)
            full = urlunsplit(parts._replace(path=quote(parts.path)))
            try:
                r = requests.get(full, headers=HEADERS, timeout=30)
            except Exception as e:
                log(f"[BLAD] Pobieranie pliku: {type(e).__name__}")
                return False
            if r.status_code != 200:
                log(f"[BLAD] Status pliku: {r.status_code}")
                return False
            with open(EXCEL_FILE, "wb") as f:
                f.write(r.content)
            log("[OK] Plik Excel pobrany")
            return True
    log("[BLAD] Nie znaleziono linku z 'IOSI'")
    return False


# ---------------------------------------------------------------------- GRUPY
def get_student_group(wb, student_index):
    """Zwraca numer grupy (int) z arkusza 'Podział na grupy' albo None."""
    idx = str(student_index).strip()
    name = next((n for n in wb.sheetnames if n.strip().lower() == "podział na grupy"), None)
    if not name:
        log("[UWAGA] Brak arkusza 'Podział na grupy'. Arkusze: " + ", ".join(wb.sheetnames))
        return None
    ws = wb[name]
    for col in range(1, ws.max_column + 1):
        group = None
        for row in range(1, ws.max_row + 1):
            raw = ws.cell(row=row, column=col).value
            val = str(int(raw)) if isinstance(raw, float) and raw == int(raw) else str(raw or "").strip()
            if re.match(r"(?i)gr\.?\s*\d", val):
                group = int(re.search(r"\d+", val).group())
            if val == idx and group:
                return group
    return None


# -------------------------------------------------- POLA TEKSTOWE Z ARKUSZA
def find_plan_sheet(wb):
    for n in wb.sheetnames:
        if n == "Plan":
            return n
    for n in wb.sheetnames:
        if n.strip().lower() == "plan":
            return n
    return None


def drawing_path_for_sheet(z, sheet_name):
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    rid = None
    for s in wb_xml.iter(f"{{{NS_MAIN}}}sheet"):
        if s.get("name") == sheet_name:
            rid = s.get(f"{{{NS_R}}}id")
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = next((r.get("Target") for r in rels if r.get("Id") == rid), None)
    if not target:
        return None
    sheet_path = target[1:] if target.startswith("/") else "xl/" + target
    rels_path = posixpath.join(
        posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels"
    )
    if rels_path not in z.namelist():
        return None
    for r in ET.fromstring(z.read(rels_path)):
        if r.get("Type", "").endswith("/drawing"):
            return posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), r.get("Target")))
    return None


def read_text_boxes(excel_path, sheet_name):
    """Zwraca liste (srodek_kolumny, [akapity]) dla kazdego pola tekstowego."""
    boxes = []
    with zipfile.ZipFile(excel_path) as z:
        dpath = drawing_path_for_sheet(z, sheet_name)
        if not dpath:
            return boxes
        root = ET.fromstring(z.read(dpath))
        for anchor in root:
            if not anchor.tag.endswith("Anchor"):
                continue
            frm = anchor.find(f"{{{NS_XDR}}}from")
            to = anchor.find(f"{{{NS_XDR}}}to")
            if frm is None or to is None:
                continue
            c1 = int(frm.find(f"{{{NS_XDR}}}col").text)
            c2 = int(to.find(f"{{{NS_XDR}}}col").text)
            tb = next(anchor.iter(f"{{{NS_XDR}}}txBody"), None)
            if tb is None:
                continue
            paras = []
            for p in tb.iter(f"{{{NS_A}}}p"):
                t = "".join(x.text or "" for x in p.iter(f"{{{NS_A}}}t"))
                if t.strip():
                    paras.append(t)
            if paras:
                boxes.append(((c1 + c2) / 2, paras))
    return boxes


def day_columns(wb, sheet_name):
    ws = wb[sheet_name]
    found = {}
    for row in ws.iter_rows():
        for c in row:
            v = str(c.value or "").strip()
            if v in DAY_NAMES:
                found[DAY_NAMES.index(v)] = c.column - 1
    if len(found) == 5:
        return [found[i] for i in range(5)]
    log("[UWAGA] Nie znaleziono naglowkow dni, uzywam domyslnych kolumn")
    return DEFAULT_DAY_COLS


# ----------------------------------------------------------------- PARSOWANIE
TIME_RE = re.compile(r"godz\.?\s*(\d{1,2})\s*:\s*(\d{2,3})\s*-\s*(\d{1,2})\s*:\s*(\d{2})", re.I)
GROUP_RE = re.compile(r"\bgr\.\s*(II|I|2|1)\b")
GROUP_MAP = {"I": 1, "1": 1, "II": 2, "2": 2}
TYPE_RE = re.compile(r"(?<![\w.])(\d{1,2})\s*([wlpć])(?!\w)")
TYPE_NAMES = {"w": "wykład", "l": "laboratorium", "p": "projekt", "ć": "ćwiczenia"}


def mk_time(h, m):
    return time(int(h), int(str(m)[-2:]))  # "130" (literowka w planie) -> 30


def parse_box(paras):
    text = " ".join(paras)
    times = [(mk_time(a, b), mk_time(c, d)) for a, b, c, d in TIME_RE.findall(text)]
    if not times:
        return None  # np. PRZERWA OBIADOWA

    # tytul
    title = paras[0]
    title = TYPE_RE.sub(" ", title)
    title = re.sub(r"\s*/\s*(?=\s|$)", " ", title)
    title = re.sub(r"CTS\+IOSI\+ISP", " ", title)
    title = GROUP_RE.sub(" ", title)
    title = re.sub(r"\s+", " ", title).strip(" -/")

    types = []
    for _, t in TYPE_RE.findall(text):
        if TYPE_NAMES[t] not in types:
            types.append(TYPE_NAMES[t])

    gm = GROUP_RE.search(text)
    group = GROUP_MAP[gm.group(1)] if gm else None

    # prowadzacy
    lecturer = ""
    for p in paras:
        ps = p.strip()
        if re.match(r"(dr|mgr|prof)\b", ps):
            lecturer = re.split(r"\s{2,}|\s+s\.\s*\d", ps)[0].strip()
            break

    # sala i adres
    room = None
    m = re.search(r"\bs\.\s*(\d[\d/]*)", text)
    if m:
        room = m.group(1)
    else:
        m = re.search(r"(?<![\d:])(\d{2,3}/\d{2,3})\b", text)
        if m:
            room = m.group(1)
            if re.match(r"^\d{2}/", room):  # "09/110" -> literowka za "109/110"
                room = "1" + room
    addr = ""
    if "Wojska Polskiego" in text:
        addr = "ul. Wojska Polskiego"
    elif "Grunwaldzka" in text:
        addr = "Al. Grunwaldzka"
    location = ""
    if room:
        location = f"s. {room}"
    if addr:
        location = (location + ", " if location else "") + addr + ", Elbląg"

    # terminy
    start = None
    m = re.search(r"\bod\s+(\d{1,2})\.(\d{2})", text)
    if m:
        mo, da = int(m.group(2)), int(m.group(1))
        start = date(2026 if mo >= 8 else 2027, mo, da)
    interval = 2 if re.search(r"co\s*2\s*tyg", text, re.I) else 1
    special = "pierwsze" if re.search(r"pierwsze", text, re.I) else (
        "ostatnie" if re.search(r"ostatnie", text, re.I) else None)
    weeks = None
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*tyg\b", text)
    if m:
        weeks = float(m.group(1).replace(",", "."))
    if weeks is not None:
        n = int(weeks) + (1 if (special or weeks % 1) else 0)
    else:
        n = SESSIONS_WEEKLY

    return {
        "title": title, "types": types, "group": group, "lecturer": lecturer,
        "location": location, "time": times[0],
        "alt_time": times[1] if len(times) > 1 else None,
        "special": special, "start": start, "interval": interval, "sessions": n,
    }


def eff_weekday(d):
    return SWAP_DAYS.get(d, d.weekday())


def session_dates(weekday, start, interval, n):
    out = []
    if interval == 1:
        d = start or SEMESTER_START_MONDAY
        d = max(d, SEMESTER_START_MONDAY)
        while len(out) < n and d <= SEARCH_LIMIT:
            if eff_weekday(d) == weekday and d not in OFF_DATES:
                out.append(d)
            d += timedelta(days=1)
    else:
        d = start or (SEMESTER_START_MONDAY + timedelta(days=weekday))
        while len(out) < n and d <= SEARCH_LIMIT:
            if d not in OFF_DATES:
                out.append(d)
            d += timedelta(days=7 * interval)
    return out


def build_events(excel_path, my_group):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = find_plan_sheet(wb)
    if not sheet:
        log("[BLAD] Brak arkusza 'Plan'. Arkusze: " + ", ".join(wb.sheetnames))
        return []
    cols = day_columns(wb, sheet)
    boxes = read_text_boxes(excel_path, sheet)
    log(f"[INFO] Pol tekstowych w arkuszu '{sheet}': {len(boxes)}")

    events = []
    for center, paras in boxes:
        info = parse_box(paras)
        if not info:
            continue
        weekday = max(i for i in range(5) if cols[i] <= center)
        if my_group and info["group"] and info["group"] != my_group:
            continue

        label = info["title"] + (f" ({' / '.join(info['types'])})" if info["types"] else "")
        desc = []
        if info["lecturer"]:
            desc.append(f"Prowadzący: {info['lecturer']}")
        if info["group"]:
            desc.append(f"Grupa: {info['group']}")
        dates = session_dates(weekday, info["start"], info["interval"], info["sessions"])
        log(f"   {DAY_NAMES[weekday][:3]} {info['time'][0]:%H:%M}-{info['time'][1]:%H:%M} "
            f"{label} | {info['location']} | {len(dates)} spotkan, "
            f"{dates[0] if dates else '-'} .. {dates[-1] if dates else '-'}")

        for i, dt in enumerate(dates):
            tm = info["time"]
            if info["alt_time"] and (
                (info["special"] == "pierwsze" and i == 0)
                or (info["special"] == "ostatnie" and i == len(dates) - 1)
            ):
                tm = info["alt_time"]
            events.append({
                "summary": label,
                "location": info["location"],
                "description": "\n".join(desc),
                "start": datetime.combine(dt, tm[0]),
                "end": datetime.combine(dt, tm[1]),
            })
    return events


# ------------------------------------------------------------------------ ICS
def esc(s):
    return (s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r", "").replace("\n", "\\n"))


def fold(line):
    if len(line.encode("utf-8")) <= 75:
        return line
    out, cur, cur_len, limit = [], "", 0, 75
    for ch in line:
        b = len(ch.encode("utf-8"))
        if cur_len + b > limit:
            out.append(cur)
            cur, cur_len, limit = ch, b, 74
        else:
            cur += ch
            cur_len += b
    out.append(cur)
    return "\r\n ".join(out)


VTIMEZONE = [
    "BEGIN:VTIMEZONE", "TZID:Europe/Warsaw",
    "BEGIN:STANDARD", "DTSTART:19701025T030000", "TZOFFSETFROM:+0200", "TZOFFSETTO:+0100",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU", "TZNAME:CET", "END:STANDARD",
    "BEGIN:DAYLIGHT", "DTSTART:19700329T020000", "TZOFFSETFROM:+0100", "TZOFFSETTO:+0200",
    "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU", "TZNAME:CEST", "END:DAYLIGHT",
    "END:VTIMEZONE",
]


def generate_ics(events, path):
    now_utc = datetime.now(zoneinfo.ZoneInfo("UTC")).strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        "PRODID:-//ANS Elblag//Plan Zajec IOSI//PL",
        "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
        "X-WR-CALNAME:Plan Zajęć ANS IOSI",
        "X-WR-TIMEZONE:Europe/Warsaw",
        "REFRESH-INTERVAL;VALUE=DURATION:PT12H",
        "X-PUBLISHED-TTL:PT12H",
    ] + VTIMEZONE
    for e in sorted(events, key=lambda x: x["start"]):
        key = f"{e['summary']}|{e['start']}|{e['end']}"
        uid = hashlib.md5(key.encode("utf-8")).hexdigest() + "@ans-elblag.pl"
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_utc}",
            f"DTSTART;TZID=Europe/Warsaw:{e['start']:%Y%m%dT%H%M%S}",
            f"DTEND;TZID=Europe/Warsaw:{e['end']:%Y%m%dT%H%M%S}",
            f"SUMMARY:{esc(e['summary'])}",
            f"LOCATION:{esc(e['location'])}",
            f"DESCRIPTION:{esc(e['description'])}",
            "STATUS:CONFIRMED",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(fold(l) for l in lines) + "\r\n")


# ------------------------------------------------------------------------ MAIN
def main():
    ok = download_excel_from_ans()
    if not os.path.exists(EXCEL_FILE):
        log("[BLAD] Brak pliku Excel - koncze.")
        sys.exit(1)
    if not ok:
        log("[UWAGA] Uzywam pliku Excel z repozytorium (pobrany_plan.xlsx).")

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    my_group = get_student_group(wb, INDEX_NUMBER)
    if my_group:
        log(f"[INFO] Twoja grupa: {my_group}")
    else:
        log("[UWAGA] Nie znaleziono indeksu na listach grup - dodaje zajecia WSZYSTKICH grup")

    events = build_events(EXCEL_FILE, my_group)
    log(f"[INFO] Liczba wydarzen: {len(events)}")
    if not events:
        log("[BLAD] Brak zajec - nie nadpisuje ICS.")
        sys.exit(1)

    generate_ics(events, OUTPUT_ICS)
    log(f"[OK] Zapisano {OUTPUT_ICS}")


if __name__ == "__main__":
    main()
