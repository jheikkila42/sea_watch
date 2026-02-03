# -*- coding: utf-8 -*-
"""
STCW-yhteensopiva työvuorogeneraattori
Versio: Uudelleenkirjoitettu jatkuvan operaation tuella
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# Värit
YELLOW = PatternFill("solid", fgColor="FFFF00")
GREEN = PatternFill("solid", fgColor="90EE90")
BLUE = PatternFill("solid", fgColor="ADD8E6")
ORANGE = PatternFill("solid", fgColor="FFA500")
GRAY = PatternFill("solid", fgColor="D3D3D3")
WHITE = PatternFill("solid", fgColor="FFFFFF")

def time_to_index(h, m):
    """Muuntaa ajan (h, m) indeksiksi (0-47)"""
    return h * 2 + (1 if m >= 30 else 0)

def index_to_time_str(idx):
    """Muuntaa indeksin ajaksi HH:MM"""
    h = idx // 2
    m = "30" if idx % 2 else "00"
    return f"{h:02d}:{m}"

def analyze_stcw_from_work_starts(work_slots_48h):
    """
    Analysoi STCW-lepoajat 48h (2 päivän) työvuorolistasta.
    Palauttaa dict: total_rest, rest_period_count, longest_rest, status, issues
    """
    # Laske lepo 24h ikkunassa (slotit 0-47 = päivä 1 klo 00 -> päivä 2 klo 00)
    rest_slots = [not w for w in work_slots_48h[:48]]
    
    # Laske lepojaksot (ignoroi alle 1h tauot kuten lounastauko)
    rest_periods = []
    current_rest = 0
    for is_rest in rest_slots:
        if is_rest:
            current_rest += 1
        else:
            if current_rest > 0:
                hours = current_rest / 2
                if hours >= 1.0:  # Vain yli 1h jaksot lasketaan
                    rest_periods.append(hours)
            current_rest = 0
    if current_rest > 0:
        hours = current_rest / 2
        if hours >= 1.0:
            rest_periods.append(hours)
    
    # Yhdistä yölepo jos se jatkuu päivien yli
    if len(rest_periods) >= 2:
        if not work_slots_48h[47] and not work_slots_48h[0]:
            combined = rest_periods[-1] + rest_periods[0]
            rest_periods = [combined] + rest_periods[1:-1]
    
    total_rest = sum(rest_periods)
    longest_rest = max(rest_periods) if rest_periods else 0
    rest_period_count = len(rest_periods)
    
    issues = []
    if total_rest < 10:
        issues.append(f"Lepoa vain {total_rest}h (min 10h)")
    if rest_period_count > 2:
        issues.append(f"Lepo {rest_period_count} osassa (max 2)")
    if longest_rest < 6:
        issues.append(f"Pisin lepo {longest_rest}h (min 6h)")
    
    return {
        'total_rest': total_rest,
        'rest_period_count': rest_period_count,
        'longest_rest': longest_rest,
        'status': 'OK' if not issues else 'RIKE',
        'issues': issues
    }


def generate_schedule(days_data):
    """
    Generoi työvuorot kaikille päiville.
    
    days_data: lista dictionaryja, joissa:
        - arrival_hour, arrival_minute (tai None)
        - departure_hour, departure_minute (tai None)  
        - port_op_start_hour, port_op_start_minute
        - port_op_end_hour, port_op_end_minute
    """
    
    NORMAL_START = time_to_index(8, 0)   # 16
    NORMAL_END = time_to_index(17, 0)     # 34
    LUNCH_START = time_to_index(11, 30)   # 23
    LUNCH_END = time_to_index(12, 0)      # 24
    TARGET_SLOTS = 17  # 8.5h
    MAX_SLOTS = 18     # 9h
    
    workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
               'Watchman 1', 'Watchman 2', 'Watchman 3']
    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
    
    all_days = {w: [] for w in workers}
    num_days = len(days_data)
    
    # ========================================
    # VAIHE 1: Analysoi koko jakso etukäteen
    # ========================================
    
    # Tunnista jatkuvat operaatiot (päivä N loppuu 00:00, päivä N+1 alkaa 00:00)
    continuous_nights = []  # Lista: (day_index, night_worker)
    
    for d in range(num_days - 1):
        curr = days_data[d]
        next_day = days_data[d + 1]
        
        curr_op_end = curr.get('port_op_end_hour')
        next_op_start = next_day.get('port_op_start_hour')
        
        if curr_op_end == 0 and next_op_start == 0:
            # Jatkuva operaatio yön yli
            # PH1 tekee yleensä illan, joten PH2 tekee yön jatkon
            continuous_nights.append((d, 'Dayman PH2'))
    
    # ========================================
    # VAIHE 2: Laske vuorot päivä kerrallaan
    # ========================================
    
    for d, info in enumerate(days_data):
        
        # Hae operaation ajat
        op_start_h = info.get('port_op_start_hour')
        op_start_m = info.get('port_op_start_minute', 0)
        op_end_h = info.get('port_op_end_hour')
        op_end_m = info.get('port_op_end_minute', 0)
        
        arrival_h = info.get('arrival_hour')
        arrival_m = info.get('arrival_minute', 0)
        departure_h = info.get('departure_hour')
        departure_m = info.get('departure_minute', 0)
        
        # Muunna indekseiksi
        if op_start_h is not None:
            op_start = time_to_index(op_start_h, op_start_m)
            if op_end_h < op_start_h:
                op_end = time_to_index(op_end_h, op_end_m) + 48
            elif op_end_h == 0 and op_start_h > 0:
                op_end = 48  # Keskiyöhön
            else:
                op_end = time_to_index(op_end_h, op_end_m)
        else:
            op_start = NORMAL_START
            op_end = NORMAL_END
        
        arrival_start = time_to_index(arrival_h, arrival_m) if arrival_h is not None else None
        arrival_end = arrival_start + 2 if arrival_start is not None else None  # 1h tulo
        
        departure_start = time_to_index(departure_h, departure_m) if departure_h is not None else None
        departure_end = departure_start + 2 if departure_start is not None else None  # 1h lähtö
        
        # Onko tämä päivä jatkuvan yön jälkeen?
        continues_from_night = False
        night_worker = None
        for (night_day, worker) in continuous_nights:
            if night_day == d - 1:
                continues_from_night = True
                night_worker = worker
                break
        
        # Onko tämä päivä jatkuvan yön alussa?
        starts_night = False
        for (night_day, worker) in continuous_nights:
            if night_day == d:
                starts_night = True
                break
        
        # ========================================
        # BOSUN
        # ========================================
        bosun_work = [False] * 48
        bosun_arr = [False] * 48
        bosun_dep = [False] * 48
        bosun_ops = [False] * 48
        
        # Normaali päivävuoro
        slot = NORMAL_START
        slots_worked = 0
        while slots_worked < TARGET_SLOTS and slot < 48:
            if LUNCH_START <= slot < LUNCH_END:
                slot += 1
                continue
            bosun_work[slot] = True
            if op_start <= slot < min(op_end, 48):
                bosun_ops[slot] = True
            slots_worked += 1
            slot += 1
        
        # Tulo
        if arrival_start is not None:
            for i in range(arrival_start, min(arrival_end, 48)):
                bosun_work[i] = True
                bosun_arr[i] = True
        
        # Lähtö
        if departure_start is not None:
            for i in range(departure_start, min(departure_end, 48)):
                bosun_work[i] = True
                bosun_dep[i] = True
        
        all_days['Bosun'].append({
            'work_slots': bosun_work,
            'arrival_slots': bosun_arr,
            'departure_slots': bosun_dep,
            'port_op_slots': bosun_ops,
            'notes': []
        })
        
        # ========================================
        # DAYMANIT
        # ========================================
        
        for dayman in daymen:
            work = [False] * 48
            arr = [False] * 48
            dep = [False] * 48
            ops = [False] * 48
            notes = []
            
            # ---- JATKUVAN YÖN KÄSITTELY ----
            
            if continues_from_night and dayman == night_worker:
                # Tämä dayman tekee yövuoron jatkon (00:00 -> ~08:00)
                notes.append('Yövuoron jatko')
                
                # Työskentele 00:00 alkaen kunnes tarpeeksi tunteja
                slot = 0
                slots_worked = 0
                target = TARGET_SLOTS
                
                while slots_worked < target and slot < NORMAL_START + 2:  # Max klo 09:00
                    work[slot] = True
                    if slot < min(op_end, 48):
                        ops[slot] = True
                    slots_worked += 1
                    slot += 1
                
                # Lähtö
                if departure_start is not None:
                    for i in range(departure_start, min(departure_end, 48)):
                        work[i] = True
                        dep[i] = True
                
                all_days[dayman].append({
                    'work_slots': work,
                    'arrival_slots': arr,
                    'departure_slots': dep,
                    'port_op_slots': ops,
                    'notes': notes
                })
                continue
            
            if continues_from_night and dayman != night_worker:
                # Muut daymanit: normaali päivävuoro, mutta myöhempi aloitus
                # jotta yötyöntekijällä on aikaa
                
                if dayman == 'Dayman PH1':
                    # PH1 teki illan edellisenä päivänä -> aloittaa myöhemmin
                    start_slot = NORMAL_START + 12  # 14:00
                    notes.append('Lepo iltavuoron jälkeen')
                else:
                    # EU aloittaa normaalisti
                    start_slot = NORMAL_START
                
                slot = start_slot
                slots_worked = 0
                while slots_worked < TARGET_SLOTS and slot < 48:
                    if LUNCH_START <= slot < LUNCH_END:
                        slot += 1
                        continue
                    work[slot] = True
                    if op_start <= slot < min(op_end, 48):
                        ops[slot] = True
                    slots_worked += 1
                    slot += 1
                
                # Lähtö
                if departure_start is not None:
                    for i in range(departure_start, min(departure_end, 48)):
                        work[i] = True
                        dep[i] = True
                
                all_days[dayman].append({
                    'work_slots': work,
                    'arrival_slots': arr,
                    'departure_slots': dep,
                    'port_op_slots': ops,
                    'notes': notes
                })
                continue
            
            # ---- NORMAALI PÄIVÄ TAI ILTA/YÖ ----
            
            # Onko iltavuoro tarpeen?
            # op_end > 48 tarkoittaa että operaatio jatkuu keskiyön yli -> iltavuoro tarvitaan
            needs_evening = (op_end > NORMAL_END and op_end <= 48) or op_end > 48
            needs_night_today = starts_night
            
            # Iltavuoro jatkuu lähtöön asti jos lähtö on operaation jälkeen
            evening_extends_to_departure = departure_start is not None and departure_start > min(op_end, 48)
            
            if dayman == 'Dayman PH1' and (needs_evening or needs_night_today):
                # PH1 tekee iltavuoron
                notes.append('Iltavuoro')
                
                # Laske iltavuoron alku ja loppu
                if needs_night_today:
                    evening_end = 48  # Keskiyöhön
                elif evening_extends_to_departure:
                    evening_end = departure_start  # Jatka lähtöön asti
                else:
                    evening_end = min(op_end, 48)
                
                evening_start = max(op_start, NORMAL_END) if op_start > NORMAL_END else NORMAL_END
                evening_slots = evening_end - evening_start
                
                # Tarvitaanko aamuvuoro?
                if evening_slots < TARGET_SLOTS:
                    # Jaettu vuoro: aamu + ilta
                    morning_slots = TARGET_SLOTS - evening_slots
                    
                    # Aamuvuoro
                    slot = NORMAL_START
                    slots_worked = 0
                    while slots_worked < morning_slots and slot < evening_start:
                        if LUNCH_START <= slot < LUNCH_END:
                            slot += 1
                            continue
                        work[slot] = True
                        if op_start <= slot < min(op_end, 48):
                            ops[slot] = True
                        slots_worked += 1
                        slot += 1
                
                # Iltavuoro
                for i in range(evening_start, evening_end):
                    work[i] = True
                    if op_start <= i < min(op_end, 48):
                        ops[i] = True
                
                # Tulo (jos on)
                if arrival_start is not None:
                    for i in range(arrival_start, min(arrival_end, 48)):
                        work[i] = True
                        arr[i] = True
                
                # Lähtö
                if departure_start is not None:
                    for i in range(departure_start, min(departure_end, 48)):
                        work[i] = True
                        dep[i] = True
                
            elif dayman == 'Dayman PH2' and starts_night:
                # PH2 lepää yötä varten - lyhyempi päivä, ei tuloa
                notes.append('Lepää yövuoroa varten')
                
                slot = NORMAL_START
                slots_worked = 0
                # Lyhyempi päivä: max 8h jotta riittää lepo
                max_slots = 16  # 8h
                while slots_worked < max_slots and slot < NORMAL_END:
                    if LUNCH_START <= slot < LUNCH_END:
                        slot += 1
                        continue
                    work[slot] = True
                    if op_start <= slot < min(op_end, 48):
                        ops[slot] = True
                    slots_worked += 1
                    slot += 1
                
                # EI tuloa - PH2 lepää
                
            else:
                # Normaali päivävuoro (EU tai PH2 normaalisti)
                
                # Aikainen aloitus?
                if op_start < NORMAL_START and dayman == 'Dayman EU':
                    start_slot = op_start
                    notes.append('Aikainen aamuvuoro')
                else:
                    start_slot = NORMAL_START
                
                slot = start_slot
                slots_worked = 0
                while slots_worked < TARGET_SLOTS and slot < 48:
                    if LUNCH_START <= slot < LUNCH_END:
                        slot += 1
                        continue
                    work[slot] = True
                    if op_start <= slot < min(op_end, 48):
                        ops[slot] = True
                    slots_worked += 1
                    slot += 1
                
                # Tulo
                if arrival_start is not None and not (dayman == 'Dayman PH2' and starts_night):
                    for i in range(arrival_start, min(arrival_end, 48)):
                        work[i] = True
                        arr[i] = True
                
                # Lähtö
                if departure_start is not None:
                    for i in range(departure_start, min(departure_end, 48)):
                        work[i] = True
                        dep[i] = True
            
            all_days[dayman].append({
                'work_slots': work,
                'arrival_slots': arr,
                'departure_slots': dep,
                'port_op_slots': ops,
                'notes': notes
            })
        
        # ========================================
        # WATCHMANIT (4-on-8-off)
        # ========================================
        
        watch_schedules = {
            'Watchman 1': [(0, 4), (12, 16)],    # 00-04, 12-16
            'Watchman 2': [(4, 8), (16, 20)],    # 04-08, 16-20
            'Watchman 3': [(8, 12), (20, 24)]    # 08-12, 20-24
        }
        
        for watchman, shifts in watch_schedules.items():
            work = [False] * 48
            arr = [False] * 48
            dep = [False] * 48
            ops = [False] * 48
            
            for (start_h, end_h) in shifts:
                start_slot = start_h * 2
                end_slot = end_h * 2
                for i in range(start_slot, end_slot):
                    work[i] = True
            
            all_days[watchman].append({
                'work_slots': work,
                'arrival_slots': arr,
                'departure_slots': dep,
                'port_op_slots': ops,
                'notes': []
            })
    
    # ========================================
    # EXCEL-GENEROINTI
    # ========================================
    
    wb = Workbook()
    wb.remove(wb.active)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for d in range(num_days):
        ws = wb.create_sheet(title=f"Päivä {d+1}")
        
        # Otsikkorivi
        ws.cell(row=1, column=1, value="Nimi")
        for col in range(48):
            h = col // 2
            m = "00" if col % 2 == 0 else "30"
            ws.cell(row=1, column=col+2, value=f"{h:02d}:{m}")
            ws.cell(row=1, column=col+2).alignment = Alignment(textRotation=90)
        ws.cell(row=1, column=50, value="Tunnit")
        
        row = 2
        for w in workers:
            ws.cell(row=row, column=1, value=w)
            
            data = all_days[w][d]
            work = data['work_slots']
            arr = data['arrival_slots']
            dep = data['departure_slots']
            ops = data['port_op_slots']
            
            hours = sum(work) / 2
            
            for col in range(48):
                cell = ws.cell(row=row, column=col+2)
                cell.border = thin_border
                
                if work[col]:
                    if arr[col]:
                        cell.fill = GREEN
                        cell.value = "T"
                    elif dep[col]:
                        cell.fill = BLUE
                        cell.value = "L"
                    elif ops[col]:
                        cell.fill = YELLOW
                        cell.value = "S"
                    else:
                        cell.fill = ORANGE
                        cell.value = "X"
                else:
                    cell.fill = WHITE
            
            ws.cell(row=row, column=50, value=hours)
            row += 1
        
        # Sarakeleveydet
        ws.column_dimensions['A'].width = 15
        for col in range(2, 50):
            ws.column_dimensions[get_column_letter(col)].width = 3
    
    # STCW-raportti
    report = []
    for w in workers:
        if len(all_days[w]) >= 2:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            combined = work1 + work2
            ana = analyze_stcw_from_work_starts(combined)
            report.append({
                'worker': w,
                'analysis': ana
            })
    
    return wb, all_days, report


# Testaus
if __name__ == "__main__":
    days_data = [
        {
            'arrival_hour': 21, 'arrival_minute': 0,
            'departure_hour': None, 'departure_minute': 0,
            'port_op_start_hour': 22, 'port_op_start_minute': 0,
            'port_op_end_hour': 0, 'port_op_end_minute': 0
        },
        {
            'arrival_hour': None, 'arrival_minute': 0,
            'departure_hour': 19, 'departure_minute': 0,
            'port_op_start_hour': 0, 'port_op_start_minute': 0,
            'port_op_end_hour': 18, 'port_op_end_minute': 0
        }
    ]
    
    wb, all_days, report = generate_schedule(days_data)
    
    print("=== Päivä 1 ===")
    for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
        work = all_days[w][0]['work_slots']
        hours = sum(work) / 2
        
        ranges = []
        start = None
        for i, x in enumerate(work):
            if x and start is None:
                start = i
            elif not x and start is not None:
                ranges.append(f"{index_to_time_str(start)}-{index_to_time_str(i)}")
                start = None
        if start is not None:
            ranges.append(f"{index_to_time_str(start)}-00:00")
        
        print(f"  {w}: {hours}h | {' + '.join(ranges)}")
    
    print("\n=== Päivä 2 ===")
    for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
        work = all_days[w][1]['work_slots']
        hours = sum(work) / 2
        
        ranges = []
        start = None
        for i, x in enumerate(work):
            if x and start is None:
                start = i
            elif not x and start is not None:
                ranges.append(f"{index_to_time_str(start)}-{index_to_time_str(i)}")
                start = None
        if start is not None:
            ranges.append(f"{index_to_time_str(start)}-00:00")
        
        print(f"  {w}: {hours}h | {' + '.join(ranges)}")
    
    print("\n=== STCW ===")
    for r in report:
        if 'Dayman' in r['worker']:
            ana = r['analysis']
            status = '✓' if ana['status'] == 'OK' else '⚠'
            print(f"  {r['worker']}: {ana['total_rest']}h lepo, {ana['rest_period_count']} jaksoa {status}")
