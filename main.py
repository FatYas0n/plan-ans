import os
import re
import requests
from bs4 import BeautifulSoup
import openpyxl
from datetime import datetime, date, time, timedelta
import zoneinfo
import uuid

INDEX_NUMBER = "21459"
OUTPUT_ICS = "plan_zajec_IOSI.ics"
EXCEL_FILE = "pobrany_plan.xlsx"
URL_STRONY_PLANU = "https://ans-elblag.pl/iis-plany-zajec.html"

# Pierwszy poniedziałek semestru
SEMESTER_START_MONDAY = date(2026, 10, 5)

def download_excel_from_ans():
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    try:
        res = requests.get(URL_STRONY_PLANU, headers=headers, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if ('IOSI' in href or 'ios' in href.lower() or 'IOSI' in a.get_text()) and (href.endswith('.xlsx') or href.endswith('.xls')):
                    link = href if href.startswith('http') else f"https://ans-elblag.pl/{href.lstrip('/')}"
                    r = requests.get(link, headers=headers)
                    with open(EXCEL_FILE, 'wb') as f:
                        f.write(r.content)
                    return True
    except Exception as e:
        print(f"Błąd: {e}")
    return False

def get_student_groups(wb, student_index):
    groups = []
    str_idx = str(student_index).strip()
    if 'Podział na grupy' in wb.sheetnames:
        ws = wb['Podział na grupy']
        for col in range(1, ws.max_column + 1):
            group_name = None
            for row in range(1, ws.max_row + 1):
                val = str(ws.cell(row=row, column=col).value or '').strip()
                if 'gr.' in val.lower() or 'grupa' in val.lower():
                    group_name = val
                if val == str_idx and group_name:
                    groups.append(group_name)
    return list(set(groups))

def parse_schedule_directly(excel_path, my_groups):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    plan_sheet = 'Plan' if 'Plan' in wb.sheetnames else 'plan '
    ws = wb[plan_sheet]
    day_columns = {col: i for i, col in enumerate([2, 17, 41, 56, 73])}
    time_pattern = re.compile(r'godz\.?\s*(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})', re.IGNORECASE)
    tz = zoneinfo.ZoneInfo("Europe/Warsaw")
    events = []

    for rng in ws.merged_cells.ranges:
        text = str(ws.cell(row=rng.min_row, column=rng.min_col).value or '').strip()
        if not text:
            continue

        match = time_pattern.search(text)
        if not match:
            continue

        start_str, end_str = match.group(1), match.group(2)
        start_h, start_m = map(int, start_str.split(':'))
        end_h, end_m = map(int, end_str.split(':'))

        day_idx = None
        for col_start in sorted(day_columns.keys()):
            if rng.min_col >= col_start:
                day_idx = day_columns[col_start]

        if day_idx is None:
            continue

        lines = [line.strip() for line in text.split('\n') if line.strip() and not time_pattern.search(line)]
        title = lines[0] if lines else "Zajęcia"
        lecturer = lines[1] if len(lines) > 1 else ""
        room = lines[2] if len(lines) > 2 else ""

        event_date = SEMESTER_START_MONDAY + timedelta(days=day_idx)
        start_dt = datetime.combine(event_date, time(hour=start_h, minute=start_m), tzinfo=tz)
        end_dt = datetime.combine(event_date, time(hour=end_h, minute=end_m), tzinfo=tz)

        events.append({
            'summary': title,
            'location': room,
            'description': f"Prowadzący: {lecturer}\nGrupy: {', '.join(my_groups)}",
            'start': start_dt,
            'end': end_dt
        })
    return events

def generate_apple_valid_ics(events_list, output_path):
    now_utc = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//ANS Elblag//Plan Zajec IOSI//PL",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Plan Zajęć ANS IOSI",
        "X-WR-TIMEZONE:Europe/Warsaw"
    ]

    for item in events_list:
        uid = f"{uuid.uuid4()}@ans-elblag.pl"
        dtstart_utc = item['start'].astimezone(zoneinfo.ZoneInfo("UTC")).strftime('%Y%m%dT%H%M%SZ')
        dtend_utc = item['end'].astimezone(zoneinfo.ZoneInfo("UTC")).strftime('%Y%m%dT%H%M%SZ')
        
        summary = item['summary'].replace('\n', ' ').replace(',', '\\,')
        location = item['location'].replace('\n', ' ').replace(',', '\\,')
        description = item['description'].replace('\n', '\\n').replace(',', '\\,')

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_utc}",
            f"DTSTART:{dtstart_utc}",
            f"DTEND:{dtend_utc}",
            "RRULE:FREQ=WEEKLY;UNTIL=20270215T235959Z",
            f"SUMMARY:{summary}",
            f"LOCATION:{location}",
            f"DESCRIPTION:{description}",
            "STATUS:CONFIRMED",
            "END:VEVENT"
        ])

    lines.append("END:VCALENDAR")

    with open(output_path, 'w', encoding='utf-8', newline='\r\n') as f:
        f.write("\r\n".join(lines))

def main():
    download_excel_from_ans()
    if not os.path.exists(EXCEL_FILE):
        return
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    my_groups = get_student_groups(wb, INDEX_NUMBER)
    events = parse_schedule_directly(EXCEL_FILE, my_groups)
    generate_apple_valid_ics(events, OUTPUT_ICS)

if __name__ == "__main__":
    main()
