import os
import re
import requests
from bs4 import BeautifulSoup
import openpyxl
from datetime import datetime, date, time, timedelta
import uuid

# ==================== KONFIGURACJA ====================
INDEX_NUMBER = "21459"
OUTPUT_ICS = "plan_zajec_IOSI.ics"
EXCEL_FILE = "pobrany_plan.xlsx"

# Poniedziałek pierwszego tygodnia semestru zimowego
SEMESTER_START_MONDAY = date(2026, 10, 5) 
# ======================================================

def download_excel_from_ans():
    """
    Pobiera najnowszy plik planu ze strony instytutu.
    """
    urls_to_try = [
        "https://ans-elblag.pl/iis-plany-zajec.html"
    ]
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    
    for page_url in urls_to_try:
        try:
            res = requests.get(page_url, headers=headers, timeout=10)
            if res.status_code == 200:
                soup = BeautifulSoup(res.text, 'html.parser')
                for a in soup.find_all('a', href=True):
                    if 'IOSI' in a['href'] or 'ios' in a['href'].lower():
                        link = a['href']
                        if not link.startswith('http'):
                            base = "/".join(page_url.split('/')[:3])
                            link = f"{base}{link}"
                        print(f"Znaleziono link do pliku: {link}")
                        r = requests.get(link, headers=headers)
                        with open(EXCEL_FILE, 'wb') as f:
                            f.write(r.content)
                        return True
        except Exception as e:
            print(f"Błąd sprawdzania {page_url}: {e}")
            
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

    if 'PODZIAŁ' in wb.sheetnames:
        ws = wb['PODZIAŁ']
        for col in range(1, ws.max_column + 1):
            for row in range(1, ws.max_row + 1):
                val = str(ws.cell(row=row, column=col).value or '').strip()
                if str_idx in val:
                    header = str(ws.cell(row=3, column=col).value or '').strip()
                    if header:
                        groups.append(header)

    return list(set(groups))

def parse_schedule_directly(excel_path, my_groups):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    plan_sheet = 'Plan' if 'Plan' in wb.sheetnames else 'plan '
    ws = wb[plan_sheet]

    day_columns = {col: i for i, col in enumerate([2, 17, 41, 56, 73])}
    time_pattern = re.compile(r'godz\.?\s*(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})', re.IGNORECASE)

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
        start_dt = datetime.combine(event_date, time(hour=start_h, minute=start_m))
        end_dt = datetime.combine(event_date, time(hour=end_h, minute=end_m))

        events.append({
            'summary': title,
            'location': room,
            'description': f"Prowadzący: {lecturer}\nGrupy: {', '.join(my_groups)}",
            'start': start_dt,
            'end': end_dt
        })

    return events

def generate_ics_standard(events_list, output_path):
    now_str = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    
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
        dtstart = item['start'].strftime('%Y%m%dT%H%M%S')
        dtend = item['end'].strftime('%Y%m%dT%H%M%S')
        
        summary = item['summary'].replace('\n', ' ').replace(',', '\\,')
        location = item['location'].replace('\n', ' ').replace(',', '\\,')
        description = item['description'].replace('\n', '\\n').replace(',', '\\,')

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_str}",
            f"DTSTART;TZID=Europe/Warsaw:{dtstart}",
            f"DTEND;TZID=Europe/Warsaw:{dtend}",
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
    print("Próba pobrania planu ze strony ANS...")
    success = download_excel_from_ans()
    
    if not success and not os.path.exists(EXCEL_FILE):
        print("Nie udało się pobrać pliku i brak pliku lokalnego.")
        return

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    my_groups = get_student_groups(wb, INDEX_NUMBER)
    print(f"Znalezione grupy: {my_groups}")

    events = parse_schedule_directly(EXCEL_FILE, my_groups)
    print(f"Przetworzono {len(events)} zajęć.")

    generate_ics_standard(events, OUTPUT_ICS)
    print(f"Zapisano plik {OUTPUT_ICS}")

if __name__ == "__main__":
    main()
