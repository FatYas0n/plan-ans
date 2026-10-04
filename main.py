import os
import re
import sys
import hashlib
import zoneinfo
from datetime import datetime, date, time, timedelta
from urllib.parse import urljoin, quote, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup
import openpyxl

# Numer indeksu: najlepiej ustaw w GitHub -> Settings -> Secrets -> INDEX_NUMBER
INDEX_NUMBER = os.environ.get("INDEX_NUMBER") or "21459"

OUTPUT_ICS = "plan_zajec_IOSI.ics"
EXCEL_FILE = "pobrany_plan.xlsx"
URL_STRONY_PLANU = "https://ans-elblag.pl/iis-plany-zajec.html"

# Pierwszy poniedzialek semestru i koniec powtarzania zajec
SEMESTER_START_MONDAY = date(2026, 10, 5)
RRULE_UNTIL = "20270215T235959Z"

TZ = zoneinfo.ZoneInfo("Europe/Warsaw")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def log(msg):
    print(msg, flush=True)


def download_excel_from_ans():
    """Pobiera plan. Zwraca True gdy sie udalo, a gdy nie - wypisuje dlaczego."""
    try:
        res = requests.get(URL_STRONY_PLANU, headers=HEADERS, timeout=30)
    except Exception as e:
        log(f"[BLAD] Nie mozna pobrac strony planow: {e}")
        return False

    log(f"[INFO] Status strony planow: {res.status_code}")
    if res.status_code != 200:
        return False

    soup = BeautifulSoup(res.text, "html.parser")
    excel_links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.lower().endswith((".xlsx", ".xls")):
            excel_links.append((href, a.get_text(" ", strip=True)))

    log(f"[INFO] Znalezione linki do Excela: {len(excel_links)}")
    for href, text in excel_links:
        log(f"   - {href}  |  {text}")

    for href, text in excel_links:
        if "iosi" in href.lower() or "iosi" in text.lower():
            full = urljoin(URL_STRONY_PLANU, href)
            parts = urlsplit(full)
            full = urlunsplit(parts._replace(path=quote(parts.path)))
            log(f"[INFO] Pobieram: {full}")
            try:
                r = requests.get(full, headers=HEADERS, timeout=60)
            except Exception as e:
                log(f"[BLAD] Pobieranie pliku nie powiodlo sie: {e}")
                return False
            if r.status_code != 200:
                log(f"[BLAD] Status pliku: {r.status_code}")
                return False
            with open(EXCEL_FILE, "wb") as f:
                f.write(r.content)
            log("[OK] Plik Excel pobrany")
            return True

    log("[BLAD] Nie znaleziono linku zawierajacego 'IOSI'")
    return False


def get_student_groups(wb, student_index):
    groups = []
    str_idx = str(student_index).strip()
    if "Podział na grupy" in wb.sheetnames:
        ws = wb["Podział na grupy"]
        for col in range(1, ws.max_column + 1):
            group_name = None
            for row in range(1, ws.max_row + 1):
                val = str(ws.cell(row=row, column=col).value or "").strip()
                if "gr." in val.lower() or "grupa" in val.lower():
                    group_name = val
                if val == str_idx and group_name:
                    groups.append(group_name)
    else:
        log("[UWAGA] Brak arkusza 'Podział na grupy'. Arkusze: " + ", ".join(wb.sheetnames))
    return sorted(set(groups))


def parse_schedule_directly(excel_path, my_groups):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    plan_sheet = "Plan" if "Plan" in wb.sheetnames else "plan "
    if plan_sheet not in wb.sheetnames:
        log("[BLAD] Brak arkusza z planem. Arkusze: " + ", ".join(wb.sheetnames))
        return []
    ws = wb[plan_sheet]

    day_columns = {col: i for i, col in enumerate([2, 17, 41, 56, 73])}
    time_pattern = re.compile(
        r"godz\.?\s*(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})", re.IGNORECASE
    )

    events = []
    for rng in ws.merged_cells.ranges:
        text = str(ws.cell(row=rng.min_row, column=rng.min_col).value or "").strip()
        if not text:
            continue
        match = time_pattern.search(text)
        if not match:
            continue

        start_h, start_m = map(int, match.group(1).split(":"))
        end_h, end_m = map(int, match.group(2).split(":"))

        day_idx = None
        for col_start in sorted(day_columns.keys()):
            if rng.min_col >= col_start:
                day_idx = day_columns[col_start]
        if day_idx is None:
            continue

        lines = [
            l.strip()
            for l in text.split("\n")
            if l.strip() and not time_pattern.search(l)
        ]
        title = lines[0] if lines else "Zajęcia"
        lecturer = lines[1] if len(lines) > 1 else ""
        room = lines[2] if len(lines) > 2 else ""

        event_date = SEMESTER_START_MONDAY + timedelta(days=day_idx)
        events.append(
            {
                "summary": title,
                "location": room,
                "description": f"Prowadzący: {lecturer}\nGrupy: {', '.join(my_groups)}",
                "start": datetime.combine(event_date, time(start_h, start_m)),
                "end": datetime.combine(event_date, time(end_h, end_m)),
            }
        )
    return events


def esc(s):
    return (
        s.replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\r", "")
        .replace("\n", "\\n")
    )


def fold(line):
    """RFC 5545: linie maks. 75 bajtow, kontynuacja zaczyna sie spacja."""
    raw = line.encode("utf-8")
    if len(raw) <= 75:
        return line
    out, cur, cur_len = [], "", 0
    limit = 75
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
    "BEGIN:VTIMEZONE",
    "TZID:Europe/Warsaw",
    "BEGIN:STANDARD",
    "DTSTART:19701025T030000",
    "TZOFFSETFROM:+0200",
    "TZOFFSETTO:+0100",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
    "TZNAME:CET",
    "END:STANDARD",
    "BEGIN:DAYLIGHT",
    "DTSTART:19700329T020000",
    "TZOFFSETFROM:+0100",
    "TZOFFSETTO:+0200",
    "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
    "TZNAME:CEST",
    "END:DAYLIGHT",
    "END:VTIMEZONE",
]


def generate_ics(events_list, output_path):
    now_utc = datetime.now(zoneinfo.ZoneInfo("UTC")).strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//ANS Elblag//Plan Zajec IOSI//PL",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Plan Zajęć ANS IOSI",
        "X-WR-TIMEZONE:Europe/Warsaw",
        "REFRESH-INTERVAL;VALUE=DURATION:PT12H",
        "X-PUBLISHED-TTL:PT12H",
    ]
    lines += VTIMEZONE

    for item in events_list:
        # Stale UID -> ponowny import/odswiezenie nie tworzy duplikatow
        key = f"{item['summary']}|{item['start']}|{item['end']}|{item['location']}"
        uid = hashlib.md5(key.encode("utf-8")).hexdigest() + "@ans-elblag.pl"
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_utc}",
            f"DTSTART;TZID=Europe/Warsaw:{item['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND;TZID=Europe/Warsaw:{item['end'].strftime('%Y%m%dT%H%M%S')}",
            f"RRULE:FREQ=WEEKLY;UNTIL={RRULE_UNTIL}",
            f"SUMMARY:{esc(item['summary'])}",
            f"LOCATION:{esc(item['location'])}",
            f"DESCRIPTION:{esc(item['description'])}",
            "STATUS:CONFIRMED",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")

    with open(output_path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(fold(l) for l in lines) + "\r\n")


def main():
    ok = download_excel_from_ans()
    if not os.path.exists(EXCEL_FILE):
        log("[BLAD] Brak pliku Excel - koncze.")
        sys.exit(1)
    if not ok:
        log("[UWAGA] Uzywam STAREGO pliku Excel z repozytorium.")

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    my_groups = get_student_groups(wb, INDEX_NUMBER)
    log(f"[INFO] Grupy dla indeksu: {my_groups}")

    events = parse_schedule_directly(EXCEL_FILE, my_groups)
    log(f"[INFO] Liczba zajec: {len(events)}")
    if not events:
        log("[BLAD] Nie znaleziono zadnych zajec - nie nadpisuje ICS.")
        sys.exit(1)

    generate_ics(events, OUTPUT_ICS)
    log(f"[OK] Zapisano {OUTPUT_ICS}")


if __name__ == "__main__":
    main()
