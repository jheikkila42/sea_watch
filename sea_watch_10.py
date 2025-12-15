# -*- coding: utf-8 -*-
"""
Created on Fri Nov 28 13:26:15 2025

@author: OMISTAJA
"""

# -*- coding: utf-8 -*-
"""
Sea Watch 9 – Työvuorogeneraattori Streamlit-sovellukseen
Puhdistettu ja optimoitu versio
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------
# APUFUNKTIOT
# ---------------------------------------------------------------------

def time_to_index(hours, minutes=0):
    return hours * 2 + (1 if minutes >= 30 else 0)

def index_to_time_str(index):
    hours = (index // 2) % 24
    minutes = 30 if index % 2 else 0
    return f"{hours:02d}:{minutes:02d}"

def get_watchman_on_duty(hour):
    if 8 <= hour < 12 or 20 <= hour < 24:
        return 'Watchman 1'
    elif 12 <= hour < 16 or 0 <= hour < 4:
        return 'Watchman 2'
    else:
        return 'Watchman 3'

def parse_time(time_str):
    if not time_str:
        return None, None
    try:
        parts = time_str.replace(".", ":").split(":")
        hour = int(parts[0])
        minute = int(parts[1]) if len(parts) > 1 else 0
        return hour, minute
    except:
        return None, None

# ---------------------------------------------------------------------
# STCW – LEPOAIKALASKENTA
# ---------------------------------------------------------------------

def find_rest_periods(work_slots, min_duration_hours=1.0):
    rest_periods = []
    in_rest = False
    rest_start = 0
    
    for i, is_working in enumerate(work_slots):
        if not is_working and not in_rest:
            in_rest = True
            rest_start = i
        elif is_working and in_rest:
            in_rest = False
            duration = (i - rest_start) / 2
            if duration >= min_duration_hours:
                rest_periods.append((rest_start, i, duration))

    if in_rest:
        duration = (len(work_slots) - rest_start) / 2
        if duration >= min_duration_hours:
            rest_periods.append((rest_start, len(work_slots), duration))

    return rest_periods


def find_work_start_points(work_slots):
    start_points = []
    in_work = False
    
    for i, is_working in enumerate(work_slots):
        if is_working and not in_work:
            start_points.append(i)
            in_work = True
        elif not is_working:
            in_work = False

    return start_points


def analyze_stcw_from_work_starts(work_slots_48h):
    """
    STCW-säännöt (liukuva 24h ikkuna):
    - Vähintään 10h lepoa joka 24h jaksossa
    - Lepo voidaan jakaa max 2 osaan
    - Yksi lepojakso vähintään 6h
    - Työjakso max 14h
    
    Laskenta: Tarkistetaan jokainen hetki päivässä 2 ja lasketaan
    24h taaksepäin siitä hetkestä.
    """
    # Tarkistetaan joka tunti päivässä 2 (slotit 48-95)
    worst_rest = 24
    worst_result = None
    
    for check_point in range(48, 96, 2):  # Joka tunti päivässä 2
        # 24h ikkuna taaksepäin = 48 slottia
        window_start = check_point - 48
        window_end = check_point
        window = work_slots_48h[window_start:window_end]
        
        # Laske työ ja lepo
        total_work = sum(window) / 2
        total_rest = 24 - total_work
        
        # Etsi lepojaksot (vähintään 1h)
        rests = find_rest_periods(window, min_duration_hours=1.0)
        longest = max((r[2] for r in rests), default=0)
        count = len(rests)
        
        # Pisin työjakso
        max_work_period = 0
        work_periods = []
        in_work = False
        work_start = 0
        
        for i, is_working in enumerate(window):
            if is_working and not in_work:
                in_work = True
                work_start = i
            elif not is_working and in_work:
                in_work = False
                work_periods.append((work_start, i))
        if in_work:
            work_periods.append((work_start, len(window)))
        
        for wp in work_periods:
            duration = (wp[1] - wp[0]) / 2
            max_work_period = max(max_work_period, duration)
        
        # Tarkista onko tämä huonoin kohta
        if total_rest < worst_rest:
            worst_rest = total_rest
            
            issues = []
            if total_rest < 10:
                issues.append(f"Lepoa vain {total_rest}h (min 10h)")
            if count > 2:
                issues.append(f"Lepo {count} osassa (max 2)")
            if longest < 6 and count > 0:
                issues.append(f"Pisin lepo {longest}h (min 6h)")
            if max_work_period > 14:
                issues.append(f"Työjakso {max_work_period}h (max 14h)")
            
            worst_result = {
                'total_rest': total_rest,
                'total_work': total_work,
                'longest_rest': longest,
                'rest_period_count': count,
                'max_gap_between_rest': max_work_period,
                'status': "OK" if not issues else "VAROITUS",
                'issues': issues,
                'worst_point': check_point
            }
    
    if worst_result is None:
        return {
            'total_rest': 24, 'total_work': 0, 'longest_rest': 24,
            'rest_period_count': 1, 'max_gap_between_rest': 0,
            'status': 'OK', 'issues': [], 'worst_point': None
        }
    
    return worst_result

# ---------------------------------------------------------------------
# SATAMAOPERAATIOVUOROT
# ---------------------------------------------------------------------

def calculate_port_operation_shifts(op_start_h, op_start_m, op_end_h, op_end_m, arrival_hour=None, departure_hour=None):
    """
    Laskee daymanien vuorot satamaoperaatiolle.
    
    Periaatteet:
    1. Kaikki daymanit osallistuvat tuloon/lähtöön
    2. Satamaoperaatiovuoro merkitään S:llä vain operaatioaikana
    3. Normaali työaika (08-17) priorisoidaan
    4. Yövuoron tekijän vuoroa säädetään niin ettei tule STCW-rikettä
    """
    TARGET = 17  # 8.5h = 17 puolituntia
    LUNCH_START = time_to_index(11, 30)
    LUNCH_END = time_to_index(12, 0)
    NORMAL_START = time_to_index(8, 0)
    NORMAL_END = time_to_index(17, 0)

    op_start = time_to_index(op_start_h, op_start_m)
    if op_end_h < op_start_h:
        op_end = time_to_index(op_end_h, op_end_m) + 48  # Seuraava päivä
    else:
        op_end = time_to_index(op_end_h, op_end_m)

    # Normaali työaika → ei tarvetta erikoisvuoroille
    if op_start >= NORMAL_START and op_end <= NORMAL_END:
        return None

    shifts = {}
    
    # Tallenna operaation todellinen aika (S-merkintää varten)
    op_start_slot = op_start
    op_end_slot = min(op_end, 48)  # Tänään loppuva osa

    # Analysoi tilanne
    needs_early = op_start < NORMAL_START
    needs_night = op_end > 48
    
    # Laske tulon vaikutus (jos tulo on)
    arrival_slots = 4 if arrival_hour is not None else 0  # 2h tulo
    arrival_start = time_to_index(arrival_hour, 0) if arrival_hour is not None else None

    # --- DAYMAN EU: Aamuvuoro tai normaali päivä ---
    if needs_early:
        eu_start = op_start
    else:
        eu_start = NORMAL_START
    eu_end = eu_start + TARGET
    if eu_start < LUNCH_START < eu_end:
        eu_end += 1
    
    shifts['Dayman EU'] = {
        'start': eu_start,
        'end': min(eu_end, 48),
        'next_day_end': None,
        'op_start_slot': op_start_slot,
        'op_end_slot': op_end_slot
    }

    # --- DAYMAN PH1 ja PH2: Yövuoro tai iltavuoro ---
    # PH2:n vuoro riippuu PH1:n tilanteesta, joten ne lasketaan yhdessä
    if needs_night:
        # Yövuoro - laske optimaalinen aloitus
        night_end_next_day = op_end - 48  # Kuinka pitkälle yö jatkuu seuraavana päivänä
        
        TARGET_SLOTS = 17  # 8.5h
        MAX_SLOTS = 18  # 9h max
        
        if arrival_start is not None:
            # Kaikki daymanit ovat tulossa
            # Tulo: arrival_start -> arrival_start + 4 (2h)
            arrival_end = arrival_start + 4
            
            # Operaation pituus tänään (14:00-00:00 = 10h jos op 14-01)
            op_length_today = 48 - op_start
            
            # JAETTU VUORO -malli:
            # 1. PH2 tekee keskiosan (tulosta eteenpäin)
            # 2. PH1 tekee aamun (08:00 -> tulon loppuun) + yövuoron (PH2:n lopusta -> 00:00)
            
            # PH2: aloittaa tulosta
            # Max lopetusaika: 22:00 (slot 44) jotta 10h lepo ennen seuraavan päivän 08:00
            ph2_start = arrival_start
            ph2_max_end = 44  # 22:00 = slot 44, jotta 10h lepoa ennen klo 08
            
            ph2_slots = TARGET_SLOTS  # 8.5h = 17 slottia
            if ph2_start < LUNCH_START:
                ph2_slots += 1  # Lounastauko
            ph2_end = ph2_start + ph2_slots
            ph2_end = min(ph2_end, ph2_max_end, 48)  # Ei yli 22:00 eikä keskiyön
            
            shifts['Dayman PH2'] = {
                'start': ph2_start,
                'end': ph2_end,
                'next_day_end': None,
                'op_start_slot': op_start_slot,
                'op_end_slot': op_end_slot
            }
            
            # PH1: aamuvuoro 08:00 -> operaation alkuun (sisältää tulon)
            ph1_morning_start = NORMAL_START  # 08:00
            ph1_morning_end = op_start  # Operaation alkuun
            
            # PH1: yövuoro alkaa kun PH2 lopettaa
            ph1_night_start = ph2_end
            
            # Laske PH1:n kokonaistunnit
            ph1_morning_slots = ph1_morning_end - ph1_morning_start
            ph1_night_slots = 48 - ph1_night_start
            ph1_total_slots = ph1_morning_slots + ph1_night_slots
            
            # Jos PH1:llä tulee liikaa tunteja, lyhennä aamuvuoroa
            if ph1_total_slots > MAX_SLOTS:
                excess = ph1_total_slots - TARGET_SLOTS
                ph1_morning_end = ph1_morning_end - excess
                ph1_morning_end = max(ph1_morning_end, NORMAL_START)  # Ei alle 08:00
            
            # PH1:n jaettu vuoro
            shifts['Dayman PH1'] = {
                'start': ph1_morning_start,
                'end': ph1_morning_end,
                'night_start': ph1_night_start,  # Yövuoron alku = PH2:n loppu
                'night_end': 48,
                'next_day_end': night_end_next_day,
                'op_start_slot': op_start_slot,
                'op_end_slot': op_end_slot,
                'split_shift': True
            }
        else:
            # Ei tuloa - normaali yövuoro
            ph1_start = min(48 - TARGET_SLOTS, op_start)  # Ei myöhemmin kuin operaatio
            
            shifts['Dayman PH1'] = {
                'start': ph1_start,
                'end': 48,
                'next_day_end': night_end_next_day,
                'op_start_slot': op_start_slot,
                'op_end_slot': op_end_slot
            }
            
            # PH2 normaali päivä
            ph2_start = NORMAL_START
            ph2_end = ph2_start + TARGET_SLOTS
            if ph2_start < LUNCH_START < ph2_end:
                ph2_end += 1
            
            shifts['Dayman PH2'] = {
                'start': ph2_start,
                'end': min(ph2_end, 48),
                'next_day_end': None,
                'op_start_slot': op_start_slot,
                'op_end_slot': op_end_slot
            }
    elif op_end > NORMAL_END:
        # Iltavuoro (ei yötä)
        ph1_end = op_end
        ph1_start = ph1_end - TARGET
        if ph1_start < LUNCH_START < ph1_end:
            ph1_start -= 1
        
        shifts['Dayman PH1'] = {
            'start': max(ph1_start, NORMAL_START),
            'end': ph1_end,
            'next_day_end': None,
            'op_start_slot': op_start_slot,
            'op_end_slot': op_end_slot
        }
        
        # PH2 normaali päivä iltavuorossa
        ph2_start = NORMAL_START
        ph2_end = ph2_start + TARGET
        if ph2_start < LUNCH_START < ph2_end:
            ph2_end += 1
        
        shifts['Dayman PH2'] = {
            'start': ph2_start,
            'end': min(ph2_end, 48),
            'next_day_end': None,
            'op_start_slot': op_start_slot,
            'op_end_slot': op_end_slot
        }
    else:
        # Normaali päivä - kaikki tekevät 08-17
        ph1_start = NORMAL_START
        ph1_end = ph1_start + TARGET
        if ph1_start < LUNCH_START < ph1_end:
            ph1_end += 1
        
        shifts['Dayman PH1'] = {
            'start': ph1_start,
            'end': min(ph1_end, 48),
            'next_day_end': None,
            'op_start_slot': op_start_slot,
            'op_end_slot': op_end_slot
        }
        
        ph2_start = NORMAL_START
        ph2_end = ph2_start + TARGET
        if ph2_start < LUNCH_START < ph2_end:
            ph2_end += 1
        
        shifts['Dayman PH2'] = {
            'start': ph2_start,
            'end': min(ph2_end, 48),
            'next_day_end': None,
            'op_start_slot': op_start_slot,
            'op_end_slot': op_end_slot
        }

    return shifts

# ---------------------------------------------------------------------
# PÄIVÄVUOROT
# ---------------------------------------------------------------------

def calculate_day_shift_for_dayworker(worker, day_info, prev_day_info=None, port_shifts=None):
    TARGET = 8.5
    LUNCH_START = time_to_index(11, 30)
    LUNCH_END = time_to_index(12, 0)
    NORMAL_START = time_to_index(8, 0)
    NORMAL_END = time_to_index(17, 0)
    MIN_REST = 6

    work = [False]*48
    arr = [False]*48
    dep = [False]*48
    ops = [False]*48
    notes = []

    # Tulo (2h) - lisätään AINA jos dayman on töissä tuloaikaan
    ah = day_info['arrival_hour']
    am = day_info['arrival_minute']
    arrival_hours = 0
    arrival_start = None
    arrival_end = None
    if ah is not None:
        arrival_start = time_to_index(ah, am)
        arrival_end = arrival_start + 4
        arrival_hours = 2

    # Lähtö (1h) - lisätään AINA jos dayman on töissä lähtöaikaan
    dh = day_info['departure_hour']
    dm = day_info['departure_minute']
    departure_hours = 0
    departure_start = None
    departure_end = None
    if dh is not None:
        departure_start = time_to_index(dh, dm)
        departure_end = departure_start + 2
        departure_hours = 1

    # Satamaoperaatiovuoro
    if port_shifts and worker in port_shifts:
        sh = port_shifts[worker]
        op_start_slot = sh.get('op_start_slot', 0)  # Operaation todellinen alkuaika
        op_end_slot = sh.get('op_end_slot', 48)     # Operaation todellinen loppuaika
        is_split_shift = sh.get('split_shift', False)  # Jaettu vuoro (aamu + yö)
        
        if is_split_shift:
            # Jaettu vuoro: aamuvuoro + yövuoro
            # Aamuvuoro
            for i in range(sh['start'], sh['end']):
                if 0 <= i < 48:
                    work[i] = True
                    if i >= op_start_slot and i < min(op_end_slot, 48):
                        ops[i] = True
            
            # Yövuoro
            night_start = sh.get('night_start', 48)
            night_end = sh.get('night_end', 48)
            for i in range(night_start, night_end):
                if 0 <= i < 48:
                    work[i] = True
                    if i >= op_start_slot and i < min(op_end_slot, 48):
                        ops[i] = True
            
            # Tulo - lisätään vain aamuvuoron sisälle (ei ylitä sh['end'])
            if arrival_start is not None:
                arr_end_limited = min(arrival_end, sh['end'])
                for i in range(arrival_start, min(arr_end_limited, 48)):
                    work[i] = True
                    arr[i] = True
        else:
            # Normaali yhtenäinen vuoro
            if sh['start'] is not None:
                for i in range(sh['start'], sh['end']):
                    if 0 <= i < 48:
                        work[i] = True
                        # Merkitse S vain jos slot on operaatioajan sisällä
                        if i >= op_start_slot and i < min(op_end_slot, 48):
                            ops[i] = True
        
        # Tulo - lisätään AINA kaikille daymaneille (paitsi split shift, jossa jo lisätty)
        if arrival_start is not None and not is_split_shift:
            for i in range(arrival_start, min(arrival_end, 48)):
                work[i] = True
                arr[i] = True
        
        # Lähtö - lisätään AINA kaikille daymaneille
        if departure_start is not None:
            for i in range(departure_start, min(departure_end, 48)):
                work[i] = True
                dep[i] = True
        
        return {
            'work_slots': work,
            'arrival_slots': arr,
            'departure_slots': dep,
            'port_op_slots': ops,
            'notes': ["Satamaoperaatio"],
            'next_day_end': sh.get('next_day_end')
        }

    # Normaali päivävuoro (ei satamaoperaatiota)
    
    # Lisää tulo
    if arrival_start is not None:
        for i in range(arrival_start, min(arrival_end, 48)):
            work[i] = True
            arr[i] = True

    # Lisää lähtö
    if departure_start is not None:
        for i in range(departure_start, min(departure_end, 48)):
            work[i] = True
            dep[i] = True

    remaining = TARGET - arrival_hours - departure_hours
    slots_needed = int(remaining * 2)

    earliest = NORMAL_START
    latest = 48

    # Esim. myöhäinen lähtö → aikaisempi päätyminen
    if dh is not None and time_to_index(dh, dm) >= NORMAL_END:
        latest = time_to_index(dh, dm)

    # Rakenna vuoro
    slot = earliest
    added = 0

    while added < slots_needed and slot < latest:
        if LUNCH_START <= slot < LUNCH_END:
            slot = LUNCH_END
            continue
        if work[slot]:
            slot += 1
            continue
        work[slot] = True
        added += 1
        slot += 1

    return {
        'work_slots': work,
        'arrival_slots': arr,
        'departure_slots': dep,
        'port_op_slots': ops,
        'notes': notes,
        'next_day_end': None
    }

def calculate_day_shift_for_watchman(worker, day_info):
    work = [False]*48
    arr = [False]*48
    dep = [False]*48

    base = {
        'Watchman 1': [(time_to_index(8), time_to_index(12)),
                       (time_to_index(20), time_to_index(24))],
        'Watchman 2': [(time_to_index(12), time_to_index(16)),
                       (time_to_index(0), time_to_index(4))],
        'Watchman 3': [(time_to_index(16), time_to_index(20)),
                       (time_to_index(4), time_to_index(8))]
    }

    for s,e in base[worker]:
        for i in range(s, min(e, 48)):
            work[i] = True

    ah = day_info['arrival_hour']
    am = day_info['arrival_minute']
    if ah is not None and get_watchman_on_duty(ah) == worker:
        s = time_to_index(ah, am); e = s + 4
        for i in range(s, min(e, 48)):
            work[i] = True
            arr[i] = True

    dh = day_info['departure_hour']
    dm = day_info['departure_minute']
    if dh is not None and get_watchman_on_duty(dh) == worker:
        s = time_to_index(dh, dm); e = s + 2
        for i in range(s, min(e, 48)):
            work[i] = True
            dep[i] = True

    return {
        'work_slots': work,
        'arrival_slots': arr,
        'departure_slots': dep,
        'port_op_slots': [False]*48,
        'notes': [],
        'next_day_end': None
    }

# ---------------------------------------------------------------------
# PÄÄFUNKTIO STREAMLITILLE
# ---------------------------------------------------------------------

def generate_schedule(days_data, output_path=None):
    """
    Luo työvuorot ja STCW-analyysin.
    Palauttaa (Workbook, all_days, report_str)
    """

    workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
               'Watchman 1', 'Watchman 2', 'Watchman 3']

    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']

    all_days = {w: [] for w in workers}
    prev_ops = {}
    num_days = len(days_data)

    # ---- LASKENTA ----
    for d, info in enumerate(days_data):
        prev = days_data[d-1] if d>0 else None
        ops = None

        if info['port_op_start_hour'] is not None:
            ops = calculate_port_operation_shifts(
                info['port_op_start_hour'],
                info['port_op_start_minute'],
                info['port_op_end_hour'],
                info['port_op_end_minute'],
                arrival_hour=info.get('arrival_hour'),
                departure_hour=info.get('departure_hour')
            )

        for w in workers:
            if w == "Bosun":
                res = calculate_day_shift_for_dayworker(w, info, prev, None)
            elif w in daymen:
                if d>0 and w in prev_ops and prev_ops[w].get('next_day_end'):
                    # yövuoron jatko
                    ns = prev_ops[w]['next_day_end']
                    work = [False]*48
                    opsl = [False]*48
                    arrl = [False]*48
                    depl = [False]*48
                    
                    for i in range(0, ns):
                        work[i] = True
                        opsl[i] = True
                    
                    # Laske päivävuoron aloitus:
                    # - Vähintään 6h lepo yövuoron jälkeen
                    # - Mutta ei ennen klo 08:00 (normaali työajan alku)
                    NORMAL_START = time_to_index(8, 0)  # klo 08:00 = slot 16
                    min_rest_slots = 12  # 6h = 12 puolituntia
                    earliest_after_rest = ns + min_rest_slots
                    slot = max(earliest_after_rest, NORMAL_START)  # Ei ennen klo 08!
                    
                    rem = 17 - ns  # Jäljellä olevat työtunnit (slotteina)
                    L1 = time_to_index(11,30)
                    L2 = time_to_index(12,0)
                    a=0
                    while a < rem and slot < 48:
                        if L1 <= slot < L2:
                            slot = L2; continue
                        work[slot] = True
                        a+=1; slot+=1
                    
                    # Lisää tulo jos dayman on töissä tuloaikaan
                    ah = info['arrival_hour']
                    am = info['arrival_minute']
                    if ah is not None:
                        arr_start = time_to_index(ah, am)
                        arr_end = arr_start + 4
                        for i in range(arr_start, min(arr_end, 48)):
                            if work[i]:  # Jos on jo töissä
                                arrl[i] = True
                    
                    # Lisää lähtö jos dayman on töissä lähtöaikaan
                    dh = info['departure_hour']
                    dm = info['departure_minute']
                    if dh is not None:
                        dep_start = time_to_index(dh, dm)
                        dep_end = dep_start + 2
                        for i in range(dep_start, min(dep_end, 48)):
                            if work[i]:  # Jos on jo töissä
                                depl[i] = True
                    
                    res = {
                        'work_slots': work,
                        'arrival_slots': arrl,
                        'departure_slots': depl,
                        'port_op_slots':opsl,
                        'notes':['Yövuoron jälkeen päivävuoro'],
                        'next_day_end': None
                    }
                else:
                    res = calculate_day_shift_for_dayworker(w, info, prev, ops)
            else:
                res = calculate_day_shift_for_watchman(w, info)

            all_days[w].append(res)

        prev_ops = ops or {}

    # ---- EXCEL ----
    wb = Workbook()

    C_WORK  = PatternFill('solid', fgColor='4472C4')
    C_ARR   = PatternFill('solid', fgColor='FFC000')
    C_DEP   = PatternFill('solid', fgColor='9966FF')
    C_OP    = PatternFill('solid', fgColor='00B050')
    C_HDR   = PatternFill('solid', fgColor='D9D9D9')
    C_WARN  = PatternFill('solid', fgColor='FF6B6B')
    C_OK    = PatternFill('solid', fgColor='92D050')

    thin = Border(left=Side(style='thin'),right=Side(style='thin'),
                  top=Side(style='thin'),bottom=Side(style='thin'))

    for d in range(num_days):
        sheet = wb.active if d==0 else wb.create_sheet()
        sheet.title = f"Päivä {d+1}"

        # aikarivi
        for h in range(24):
            col = 2 + h*2
            c = sheet.cell(row=1, column=col)
            c.value = h
            c.fill = C_HDR
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal='center')
            sheet.merge_cells(start_row=1, start_column=col,
                              end_row=1, end_column=col+1)

        sheet.column_dimensions['A'].width = 14
        for col in range(2,50):
            sheet.column_dimensions[get_column_letter(col)].width = 2.5

        # STCW-pylväät
        if d>0:
            hdr = ["Työ (h)","Lepo (h)","Pisin lepo","Status"]
            for i,hname in enumerate(hdr):
                col = 51+i
                cell = sheet.cell(row=1, column=col)
                cell.value=hname
                cell.fill=C_HDR
                cell.font=Font(bold=True)
                cell.alignment=Alignment(horizontal='center')

        # työntekijät
        for r,w in enumerate(workers, start=2):
            data = all_days[w][d]
            work=data['work_slots']; arr=data['arrival_slots']
            dep =data['departure_slots']; ops=data['port_op_slots']

            nm = sheet.cell(row=r, column=1)
            nm.value=w
            nm.font=Font(bold=True)
            nm.border=thin

            for t in range(48):
                c = sheet.cell(row=r, column=2+t)
                c.border=thin
                if arr[t]:
                    c.fill = C_ARR; c.value="B"
                    c.alignment=Alignment(horizontal='center')
                elif dep[t]:
                    c.fill = C_DEP; c.value="C"
                    c.alignment=Alignment(horizontal='center')
                elif ops[t]:
                    c.fill = C_OP;  c.value="S"
                    c.alignment=Alignment(horizontal='center')
                elif work[t]:
                    c.fill = C_WORK

            # STCW
            if d>0:
                prev_work = all_days[w][d-1]['work_slots']
                combined = prev_work + work
                ana = analyze_stcw_from_work_starts(combined)

                sheet.cell(row=r,column=51).value = ana['total_work']
                sheet.cell(row=r,column=52).value = ana['total_rest']
                sheet.cell(row=r,column=53).value = ana['longest_rest']

                stc = sheet.cell(row=r, column=54)
                stc.value = ana['status']
                stc.font = Font(bold=True)
                stc.alignment = Alignment(horizontal='center')
                stc.fill = C_OK if ana['status']=="OK" else C_WARN

        # selite
        base = len(workers) + 4
        sheet.cell(row=base, column=1).value="Selite:"
        sheet.cell(row=base, column=1).font=Font(bold=True)

        sheet.cell(row=base+1, column=1).value="Työ"
        sheet.cell(row=base+1, column=2).fill=C_WORK

        sheet.cell(row=base+2, column=1).value="Satamaan tulo (B)"
        sheet.cell(row=base+2, column=2).fill=C_ARR

        sheet.cell(row=base+3, column=1).value="Satamasta lähtö (C)"
        sheet.cell(row=base+3, column=2).fill=C_DEP

        sheet.cell(row=base+4, column=1).value="Satamaoperaatio (S)"
        sheet.cell(row=base+4, column=2).fill=C_OP

    # tallenna halutessa
    if output_path:
        wb.save(output_path)

    # ---- TEKSTIRAPORTTI ----
    lines=[]
    lines.append("="*60)
    lines.append("TYÖVUOROT JA LEPOAIKA-ANALYYSI")
    lines.append("="*60)

    for d in range(num_days):
        lines.append(f"\n--- PÄIVÄ {d+1} ---")
        for w in workers:
            dat = all_days[w][d]
            h = sum(dat['work_slots'])/2
            notes = f" ({', '.join(dat['notes'])})" if dat['notes'] else ""

            if d==0:
                lines.append(f"  {w}: {h}h työtä{notes}")
            else:
                prev = all_days[w][d-1]['work_slots']
                ana = analyze_stcw_from_work_starts(prev + dat['work_slots'])
                icon = "✓" if ana['status']=="OK" else "⚠"
                lines.append(
                    f"{icon} {w}: {h}h työtä, Lepo {ana['total_rest']}h, "
                    f"Pisin lepo {ana['longest_rest']}h{notes}"
                )
                for issue in ana['issues']:
                    lines.append(f"    - {issue}")

    report = "\n".join(lines)
    return wb, all_days, report
