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
    day2_slots = work_slots_48h[48:96]
    starts2 = find_work_start_points(day2_slots)
    starts_abs = [s + 48 for s in starts2]

    if not starts_abs:
        return {
            'total_rest': 24, 'total_work': 0, 'longest_rest': 24,
            'rest_period_count': 1, 'max_gap_between_rest': 0,
            'status': 'OK', 'issues': [], 'worst_point': None
        }

    worst_rest = 24
    worst = None
    worst_point = None

    for ws in starts_abs:
        window = work_slots_48h[ws - 48: ws]
        total_work = sum(window) / 2
        total_rest = 24 - total_work
        rests = find_rest_periods(window)
        longest = max((r[2] for r in rests), default=0)
        count = len(rests)

        gaps = 0
        if len(rests) >= 2:
            for i in range(len(rests) - 1):
                gaps = max(gaps, (rests[i+1][0] - rests[i][1]) / 2)

        if total_rest < worst_rest:
            worst_rest = total_rest
            worst_point = ws
            worst = {
                'total_rest': total_rest,
                'total_work': total_work,
                'longest_rest': longest,
                'rest_period_count': count,
                'max_gap_between_rest': gaps
            }

    issues = []
    if worst['total_rest'] < 10:
        issues.append(f"Lepoa vain {worst['total_rest']}h (min 10h)")
    if worst['rest_period_count'] > 2:
        issues.append(f"Lepo {worst['rest_period_count']} osassa (max 2)")
    if worst['longest_rest'] < 6:
        issues.append(f"Pisin lepo {worst['longest_rest']}h (min 6h)")
    if worst['max_gap_between_rest'] > 14:
        issues.append(f"Lepojaksojen väli {worst['max_gap_between_rest']}h (max 14h)")

    return {
        **worst,
        'status': "OK" if not issues else "VAROITUS",
        'issues': issues,
        'worst_point': index_to_time_str(worst_point)
    }

# ---------------------------------------------------------------------
# SATAMAOPERAATIOVUOROT
# ---------------------------------------------------------------------

def calculate_port_operation_shifts(op_start_h, op_start_m, op_end_h, op_end_m, arrival_hour=None, departure_hour=None):
    """
    Laskee daymanien vuorot satamaoperaatiolle.
    Priorisoi normaalityöaikaa (08-17) mahdollisuuksien mukaan.
    Optimoi vuorot STCW-lepoaikojen kannalta.
    
    arrival_hour ja departure_hour: Jos tulo/lähtö osuu samalle päivälle,
    optimoidaan vuorot niin ettei yövuoron tekijälle tule liikaa tunteja.
    """
    TARGET = 17  # 8.5h = 17 puolituntia
    MAX_SLOTS = 18  # 9h = maksimi päivän pituus (TARGET + 1 lounastauko)
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

    # Analysoi operaation osat
    needs_early = op_start < NORMAL_START  # Alkaa ennen klo 08
    needs_late = op_end > NORMAL_END       # Jatkuu klo 17 jälkeen (sama päivä tai yö)
    needs_night = op_end > 48              # Jatkuu seuraavaan päivään

    # Laske tulo/lähtö slotit jos annettu
    arrival_slots = 0
    arrival_start_idx = None
    if arrival_hour is not None:
        arrival_start_idx = time_to_index(arrival_hour, 0)
        arrival_slots = 4  # 2h tulo
    
    departure_slots = 0
    departure_start_idx = None
    if departure_hour is not None:
        departure_start_idx = time_to_index(departure_hour, 0)
        departure_slots = 2  # 1h lähtö

    # --- Tapaus 1: Operaatio alkaa aikaisin (ennen 08) ---
    if needs_early:
        # EU aloittaa aikaisin
        eu_start = op_start
        eu_end = eu_start + TARGET
        if eu_start < LUNCH_START < eu_end:
            eu_end += 1
        shifts['Dayman EU'] = {
            'start': eu_start,
            'end': min(eu_end, 48),
            'next_day_end': None if eu_end <= 48 else eu_end - 48
        }
        
        # PH1: yövuoro jos tarpeen
        if needs_night:
            # Laske PH1:n yövuoron pituus
            ph1_night_slots = TARGET  # 8.5h yövuoro
            
            # Tarkista osuuko tulo PH1:n aikaan (yövuoron tekijä)
            # Jos tulo on ja PH1 joutuisi tekemään sen + yövuoron → liikaa tunteja
            ph1_extra_slots = 0
            if arrival_start_idx is not None:
                # Tulo osuu PH1:lle jos se on ennen yövuoron alkua mutta PH1 on töissä
                # PH1 tekee yövuoron, joten hän voi olla tulossa jos tulo on myöhään
                ph1_night_start = 48 - TARGET
                if arrival_start_idx >= NORMAL_START and arrival_start_idx < ph1_night_start:
                    # Tulo on keskellä päivää - PH1 saattaa joutua tulemaan
                    ph1_extra_slots = arrival_slots
            
            # Jos PH1:lle tulisi liikaa tunteja, siirrä PH2:n vuoroa myöhemmäksi
            if ph1_night_slots + ph1_extra_slots > MAX_SLOTS:
                # Siirrä PH2 myöhemmäksi niin että PH1 voi aloittaa myöhemmin
                # ja joku muu kattaa tulon
                shift_amount = ph1_extra_slots
                ph2_start = NORMAL_START + shift_amount
                ph2_end = ph2_start + TARGET
                if ph2_start < LUNCH_START < ph2_end:
                    ph2_end += 1
                shifts['Dayman PH2'] = {
                    'start': ph2_start,
                    'end': min(ph2_end, 48),
                    'next_day_end': None
                }
                
                # PH1:n yövuoro alkaa myöhemmin
                ph1_start = 48 - TARGET
                shifts['Dayman PH1'] = {
                    'start': ph1_start,
                    'end': 48,
                    'next_day_end': op_end - 48
                }
            else:
                # Normaali tilanne
                ph2_start = NORMAL_START
                ph2_end = ph2_start + TARGET
                if ph2_start < LUNCH_START < ph2_end:
                    ph2_end += 1
                shifts['Dayman PH2'] = {
                    'start': ph2_start,
                    'end': min(ph2_end, 48),
                    'next_day_end': None
                }
                
                ph1_start = 48 - TARGET
                shifts['Dayman PH1'] = {
                    'start': ph1_start,
                    'end': 48,
                    'next_day_end': op_end - 48
                }
        elif needs_late:
            ph1_end = op_end
            ph1_start = ph1_end - TARGET
            shifts['Dayman PH1'] = {
                'start': max(ph1_start, 0),
                'end': ph1_end,
                'next_day_end': None
            }
            
            ph2_start = NORMAL_START
            ph2_end = ph2_start + TARGET
            if ph2_start < LUNCH_START < ph2_end:
                ph2_end += 1
            shifts['Dayman PH2'] = {
                'start': ph2_start,
                'end': min(ph2_end, 48),
                'next_day_end': None
            }
        else:
            # Normaali päivä kaikille
            ph1_start = NORMAL_START
            ph1_end = ph1_start + TARGET
            if ph1_start < LUNCH_START < ph1_end:
                ph1_end += 1
            shifts['Dayman PH1'] = {
                'start': ph1_start,
                'end': min(ph1_end, 48),
                'next_day_end': None
            }
            
            ph2_start = NORMAL_START
            ph2_end = ph2_start + TARGET
            if ph2_start < LUNCH_START < ph2_end:
                ph2_end += 1
            shifts['Dayman PH2'] = {
                'start': ph2_start,
                'end': min(ph2_end, 48),
                'next_day_end': None
            }
    
    # --- Tapaus 2: Operaatio alkaa normaalisti/myöhään ja jatkuu iltaan/yöhön ---
    elif needs_late or needs_night:
        
        if needs_night:
            # Yövuoro - tarkista STCW-optimointi
            ph1_night_slots = TARGET  # 8.5h yövuoro
            
            # Laske paljonko PH1:lle tulisi lisätunteja tulosta
            ph1_extra_slots = 0
            ph1_night_start_default = 48 - TARGET  # klo 15:30
            
            if arrival_start_idx is not None:
                # Jos tulo on ja se on ennen yövuoron alkua
                arrival_end_idx = arrival_start_idx + arrival_slots
                if arrival_end_idx > ph1_night_start_default:
                    # Tulo menee päällekkäin yövuoron kanssa - ei lisätunteja
                    pass
                elif arrival_start_idx >= NORMAL_START:
                    # Tulo on keskellä päivää - PH1 saattaa joutua tulemaan
                    # Lasketaan overlap: jos tulo 12-14 ja yövuoro 15:30-00
                    # PH1 tekisi 12-14 (tulo) + 15:30-00 (yö) = 10.5h
                    ph1_extra_slots = arrival_slots
            
            total_ph1_slots = ph1_night_slots + ph1_extra_slots
            
            if total_ph1_slots > MAX_SLOTS and arrival_start_idx is not None:
                # STCW-OPTIMOINTI: Siirrä vuoroja niin että PH1 ei saa liikaa tunteja
                # 
                # Strategia: PH2 siirtyy myöhemmäksi ja kattaa ajan ennen yövuoroa
                # EU kattaa tulon, PH1 tekee vain yövuoron
                
                # EU tekee normaalin päivän (kattaa tulon)
                eu_start = NORMAL_START
                eu_end = eu_start + TARGET
                if eu_start < LUNCH_START < eu_end:
                    eu_end += 1
                shifts['Dayman EU'] = {
                    'start': eu_start,
                    'end': min(eu_end, 48),
                    'next_day_end': None
                }
                
                # PH2 siirtyy myöhemmäksi - kattaa iltapäivän ennen PH1:n yövuoroa
                # Lasketaan: PH1:n yövuoro alkaa 48 - TARGET = 31 (15:30)
                # PH2:n pitää kattaa aika EU:n lopun ja PH1:n alun välillä
                # Siirrä PH2 niin että hän loppuu kun PH1 alkaa (tai vähän ennen)
                
                ph1_start = 48 - TARGET  # 15:30
                # PH2 loppuu kun PH1 alkaa + pieni overlap
                ph2_end = ph1_start + 4  # 2h overlap
                ph2_start = ph2_end - TARGET
                if ph2_start < LUNCH_START < ph2_end:
                    ph2_start -= 1
                
                # Varmista että PH2 ei ala liian aikaisin
                ph2_start = max(ph2_start, NORMAL_START)
                
                shifts['Dayman PH2'] = {
                    'start': ph2_start,
                    'end': min(ph2_end, 48),
                    'next_day_end': None
                }
                
                # PH1 tekee yövuoron (ei tuloa)
                shifts['Dayman PH1'] = {
                    'start': ph1_start,
                    'end': 48,
                    'next_day_end': op_end - 48
                }
            else:
                # Normaali tilanne - ei STCW-ongelmaa
                eu_start = NORMAL_START
                eu_end = eu_start + TARGET
                if eu_start < LUNCH_START < eu_end:
                    eu_end += 1
                shifts['Dayman EU'] = {
                    'start': eu_start,
                    'end': min(eu_end, 48),
                    'next_day_end': None
                }
                
                ph2_start = NORMAL_START
                ph2_end = ph2_start + TARGET
                if ph2_start < LUNCH_START < ph2_end:
                    ph2_end += 1
                shifts['Dayman PH2'] = {
                    'start': ph2_start,
                    'end': min(ph2_end, 48),
                    'next_day_end': None
                }
                
                ph1_start = 48 - TARGET
                shifts['Dayman PH1'] = {
                    'start': ph1_start,
                    'end': 48,
                    'next_day_end': op_end - 48
                }
        else:
            # Iltavuoro (ei yötä)
            eu_start = NORMAL_START
            eu_end = eu_start + TARGET
            if eu_start < LUNCH_START < eu_end:
                eu_end += 1
            shifts['Dayman EU'] = {
                'start': eu_start,
                'end': min(eu_end, 48),
                'next_day_end': None
            }
            
            ph2_start = NORMAL_START
            ph2_end = ph2_start + TARGET
            if ph2_start < LUNCH_START < ph2_end:
                ph2_end += 1
            shifts['Dayman PH2'] = {
                'start': ph2_start,
                'end': min(ph2_end, 48),
                'next_day_end': None
            }
            
            ph1_end = op_end
            ph1_start = ph1_end - TARGET
            if ph1_start < LUNCH_START < ph1_end:
                ph1_start -= 1
            shifts['Dayman PH1'] = {
                'start': max(ph1_start, 0),
                'end': ph1_end,
                'next_day_end': None
            }
    
    # --- Tapaus 3: Operaatio loppuu ennen iltaa ---
    else:
        # Kaikki tekevät normaalin päivän
        for worker in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            start = NORMAL_START
            end = start + TARGET
            if start < LUNCH_START < end:
                end += 1
            shifts[worker] = {
                'start': start,
                'end': min(end, 48),
                'next_day_end': None
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
        if sh['start'] is not None:
            for i in range(sh['start'], sh['end']):
                if 0 <= i < 48:
                    work[i] = True
                    ops[i] = True
        
        # Laske työntekijän kokonaistunnit
        current_slots = sum(work)
        has_night_shift = sh.get('next_day_end') is not None
        
        # Lisää tulo jos se osuu työaikaan tai laajentaa sitä
        # MUTTA: Jos työntekijällä on yövuoro ja tulo lisäisi liikaa tunteja, EI lisätä
        if arrival_start is not None:
            is_working_at_arrival = any(work[i] for i in range(arrival_start, min(arrival_end, 48)) if i < 48)
            would_extend = (sh['start'] is not None and arrival_end >= sh['start'] - 4)
            
            # Laske kuinka monta slottia tulo lisäisi
            extra_slots = 0
            if not is_working_at_arrival and would_extend:
                for i in range(arrival_start, min(arrival_end, 48)):
                    if not work[i]:
                        extra_slots += 1
            
            # Jos yövuoro ja tulo lisäisi liikaa tunteja (yli 9h), EI lisätä tuloa
            max_slots = 18  # 9h
            if has_night_shift and (current_slots + extra_slots) > max_slots:
                # Ei lisätä tuloa - joku muu dayman hoitaa sen
                pass
            elif is_working_at_arrival or would_extend:
                for i in range(arrival_start, min(arrival_end, 48)):
                    work[i] = True
                    arr[i] = True
        
        # Lisää lähtö jos se osuu työaikaan tai laajentaa sitä
        if departure_start is not None:
            is_working_at_departure = any(work[i] for i in range(departure_start, min(departure_end, 48)) if i < 48)
            if is_working_at_departure or (sh['end'] is not None and departure_start <= sh['end'] + 4):
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
