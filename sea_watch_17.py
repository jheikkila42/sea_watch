# -*- coding: utf-8 -*-
"""
STCW-yhteensopiva työvuorogeneraattori
Versio 17: Pilkottu ja refaktoroitu

Vaiheet:
- VAIHE 0: Analysoi jatkuvat yövuorot etukäteen
- VAIHE 1: Pakolliset (tulo, lähtö, slussi, shiftaus, op 08-17 ulkopuolella)
- VAIHE 2: Laske tarvittavat lisätunnit per dayman
- VAIHE 3: Jaa työblokit
- VAIHE 4: Täytä aukot
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ============================================================================
# VAKIOT
# ============================================================================

YELLOW = PatternFill("solid", fgColor="FFFF00")
GREEN = PatternFill("solid", fgColor="90EE90")
BLUE = PatternFill("solid", fgColor="ADD8E6")
ORANGE = PatternFill("solid", fgColor="FFA500")
PURPLE = PatternFill("solid", fgColor="C9A0DC")
PINK = PatternFill("solid", fgColor="FFB6C1")
WHITE = PatternFill("solid", fgColor="FFFFFF")

NORMAL_START = 16   # 08:00 (slotti)
NORMAL_END = 34     # 17:00 (slotti)
LUNCH_START = 23    # 11:30
LUNCH_END = 24      # 12:00
MIN_HOURS = 8
MAX_HOURS = 10

WORKERS = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
           'Watchman 1', 'Watchman 2', 'Watchman 3']
DAYMEN = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']


# ============================================================================
# AIKA- JA SLOTTIFUNKTIOT
# ============================================================================

def time_to_slot(h, m=0):
    """Muuntaa ajan slotiksi (0-47)."""
    return h * 2 + (1 if m >= 30 else 0)


def slot_to_time_str(slot):
    """Muuntaa slotin aikamerkkijonoksi (HH:MM)."""
    h = slot // 2
    m = "30" if slot % 2 else "00"
    return f"{h:02d}:{m}"


def get_work_ranges(work_slots):
    """Palauttaa työjaksot luettavassa muodossa."""
    ranges = []
    start = None
    for i, w in enumerate(work_slots):
        if w and start is None:
            start = i
        elif not w and start is not None:
            ranges.append(f"{slot_to_time_str(start)}-{slot_to_time_str(i)}")
            start = None
    if start is not None:
        ranges.append(f"{slot_to_time_str(start)}-00:00")
    return ranges


def add_block(work_list, start, end, marker_list=None):
    """Lisää työblokin slottilistaan."""
    for i in range(start, min(end, 48)):
        work_list[i] = True
        if marker_list is not None:
            marker_list[i] = True


def get_work_blocks(work_slots):
    """Palauttaa työblokit listana (start, end) -tupleja."""
    blocks = []
    start = None
    for i, w in enumerate(work_slots):
        if w and start is None:
            start = i
        elif not w and start is not None:
            blocks.append((start, i))
            start = None
    if start is not None:
        blocks.append((start, 48))
    return blocks


def parse_time_str(time_str):
    """Parsii aikamerkkijonon (HH:MM) slotiksi."""
    if not time_str:
        return None
    if isinstance(time_str, int):
        return time_str * 2
    time_str = str(time_str).replace(".", ":")
    parts = time_str.split(":")
    h = int(parts[0])
    m = int(parts[1]) if len(parts) > 1 else 0
    return time_to_slot(h, m)


# ============================================================================
# STCW-TARKISTUS
# ============================================================================

def analyze_stcw(work_48h):
    """
    Analysoi STCW-lepoajat 48h (2 päivän) työvuorolistasta.
    """
    rest_slots = [not w for w in work_48h[:48]]
    
    rest_periods = []
    current_rest = 0
    for is_rest in rest_slots:
        if is_rest:
            current_rest += 1
        else:
            if current_rest >= 2:
                rest_periods.append(current_rest / 2)
            current_rest = 0
    if current_rest >= 2:
        rest_periods.append(current_rest / 2)
    
    # Yhdistä yölepo jos se jatkuu keskiyön yli
    if len(rest_periods) >= 2 and len(work_48h) >= 48:
        if not work_48h[47] and not work_48h[0]:
            combined = rest_periods[-1] + rest_periods[0]
            rest_periods = [combined] + rest_periods[1:-1]
    
    total_rest = sum(rest_periods)
    longest_rest = max(rest_periods) if rest_periods else 0
    
    sorted_periods = sorted(rest_periods, reverse=True)
    top_two = sorted_periods[:2] if len(sorted_periods) >= 2 else sorted_periods
    top_two_total = sum(top_two)
    
    return {
        'total_rest': total_rest,
        'longest_rest': longest_rest,
        'rest_periods': rest_periods,
        'top_two_total': top_two_total
    }


def check_stcw_ok(work_slots, prev_day_work=None):
    """
    Tarkistaa täyttääkö työvuoro STCW-vaatimukset.
    """
    if prev_day_work is None:
        prev_day_work = [False] * 48
    
    combined = prev_day_work + work_slots
    analysis = analyze_stcw(combined)
    
    # STCW: 10h lepo/24h, josta vähintään 6h yhtäjaksoinen
    return analysis['total_rest'] >= 10 and analysis['longest_rest'] >= 6


def check_stcw_at_slot(work_48h, end_slot):
    """
    Tarkistaa STCW-statuksen tietyssä kohdassa.
    """
    start_slot = max(0, end_slot - 47)
    window = work_48h[start_slot:end_slot + 1]
    
    while len(window) < 48:
        window = [False] + window
    
    analysis = analyze_stcw(window)
    
    total_ok = analysis['total_rest'] >= 10
    longest_ok = analysis['longest_rest'] >= 6
    
    if total_ok and longest_ok:
        status = "OK"
    elif not total_ok:
        status = "TOTAL_REST_VIOLATION"
    else:
        status = "LONGEST_REST_VIOLATION"
    
    return {
        'status': status,
        'total_rest': analysis['total_rest'],
        'longest_rest': analysis['longest_rest'],
        'rest_periods': analysis['rest_periods']
    }


# ============================================================================
# RAJOITTEIDEN KÄSITTELY
# ============================================================================

def can_work_slot(worker, slot, day_idx, constraints, current_hours=0):
    """Tarkistaa voiko työntekijä tehdä tietyn slotin."""
    if not constraints:
        return True
    
    for c in constraints:
        c_worker = c.get("worker")
        c_type = c.get("type")
        c_day = c.get("day")
        
        if c_worker and c_worker != worker:
            continue
        if c_day is not None and c_day != day_idx + 1:
            continue
        
        if c_type == "no_night_shift":
            if slot < NORMAL_START:
                return False
        elif c_type == "no_evening_shift":
            if slot >= NORMAL_END:
                return False
        elif c_type == "cannot_work_slot":
            start = parse_time_str(c.get("start_time"))
            end = parse_time_str(c.get("end_time"))
            if start is not None and end is not None:
                if start <= slot < end:
                    return False
        elif c_type == "max_hours":
            max_h = c.get("value", MAX_HOURS)
            if current_hours >= max_h:
                return False
        elif c_type == "day_off":
            return False
    
    return True


def must_work_slot(worker, slot, day_idx, constraints):
    """Tarkistaa onko työntekijän pakko tehdä tietty slotti."""
    if not constraints:
        return False
    
    for c in constraints:
        c_worker = c.get("worker")
        c_type = c.get("type")
        c_day = c.get("day")
        
        if c_worker and c_worker != worker:
            continue
        if c_day is not None and c_day != day_idx + 1:
            continue
        
        if c_type == "must_work_slot":
            start = parse_time_str(c.get("start_time"))
            end = parse_time_str(c.get("end_time"))
            if start is not None and end is not None:
                if start <= slot < end:
                    return True
    
    return False


def is_day_off(worker, day_idx, constraints):
    """Tarkistaa onko työntekijällä vapaapäivä."""
    if not constraints:
        return False
    
    for c in constraints:
        if c.get("type") == "day_off":
            if c.get("worker") == worker:
                c_day = c.get("day")
                if c_day is None or c_day == day_idx + 1:
                    return True
    return False


def get_min_hours(worker, constraints):
    """Palauttaa työntekijän minimitunnit."""
    for c in constraints or []:
        if c.get("worker") == worker and c.get("type") == "min_hours":
            return c.get("value", MIN_HOURS)
    return MIN_HOURS


def get_max_hours(worker, constraints):
    """Palauttaa työntekijän maksimitunnit."""
    for c in constraints or []:
        if c.get("worker") == worker and c.get("type") == "max_hours":
            return c.get("value", MAX_HOURS)
    return MAX_HOURS


def get_preferred_night_worker(day_idx, constraints, daymen):
    """Palauttaa yövuoroon määrätyn työntekijän."""
    if not constraints:
        return None
    
    for c in constraints:
        if c.get("type") == "assign_night_shift":
            c_day = c.get("day")
            if c_day is None or c_day == day_idx + 1:
                worker = c.get("worker")
                if worker in daymen:
                    return worker
    return None


# ============================================================================
# VAIHE 0: JATKUVIEN YÖVUOROJEN ANALYYSI
# ============================================================================

def analyze_continuous_nights(days_data):
    """
    Analysoi jatkuvat yövuorot etukäteen.
    Palauttaa listan jatkuvista öistä.
    """
    continuous_nights = []
    num_days = len(days_data)
    
    for d in range(num_days - 1):
        curr = days_data[d]
        next_day = days_data[d + 1]
        
        curr_op_end = curr.get('port_op_end_hour')
        next_op_start = next_day.get('port_op_start_hour')
        
        if curr_op_end == 0 and next_op_start == 0:
            continuous_nights.append({
                'day_index': d,
                'early_worker': 'Dayman PH1',
                'late_worker': 'Dayman PH2'
            })
    
    return continuous_nights


def evaluate_night_split(prev_early, prev_late, split_slot, arrival_start=None, departure_start=None):
    """
    Arvioi yövuoron jakokohdan hyvyyttä.
    """
    early_night = list(range(0, split_slot))
    late_night = list(range(split_slot, NORMAL_START))
    
    early_combined = prev_early + [True if i in early_night else False for i in range(48)]
    late_combined = prev_late + [True if i in late_night else False for i in range(48)]
    
    early_analysis = analyze_stcw(early_combined)
    late_analysis = analyze_stcw(late_combined)
    
    total_issues = 0
    if early_analysis['total_rest'] < 10:
        total_issues += 1
    if early_analysis['longest_rest'] < 6:
        total_issues += 1
    if late_analysis['total_rest'] < 10:
        total_issues += 1
    if late_analysis['longest_rest'] < 6:
        total_issues += 1
    
    min_longest_rest = min(early_analysis['longest_rest'], late_analysis['longest_rest'])
    min_total_rest = min(early_analysis['total_rest'], late_analysis['total_rest'])
    
    return (total_issues, -min_longest_rest, -min_total_rest)


def choose_night_split_slot(prev_early, prev_late, arrival_start=None, departure_start=None):
    """
    Valitsee optimaalisen yövuoron jakokohdan (01:00 - 07:00 väliltä).
    """
    candidate_slots = list(range(time_to_slot(1, 0), time_to_slot(7, 0) + 1))
    best_slot = time_to_slot(3, 0)
    best_score = None
    
    for split_slot in candidate_slots:
        score = evaluate_night_split(prev_early, prev_late, split_slot, 
                                     arrival_start, departure_start)
        if best_score is None or score < best_score:
            best_score = score
            best_slot = split_slot
    
    return best_slot


# ============================================================================
# VAIHE 1: PAKOLLISET SLOTIT
# ============================================================================

def parse_day_times(info):
    """
    Parsii päivän ajat sloteiksi.
    Palauttaa dictin kaikista relevanteista sloteista.
    """
    op_start_h = info.get('port_op_start_hour')
    op_start_m = info.get('port_op_start_minute', 0)
    op_end_h = info.get('port_op_end_hour')
    op_end_m = info.get('port_op_end_minute', 0)
    
    arrival_h = info.get('arrival_hour')
    arrival_m = info.get('arrival_minute', 0)
    departure_h = info.get('departure_hour')
    departure_m = info.get('departure_minute', 0)
    
    sluice_arr_h = info.get('sluice_arrival_hour')
    sluice_arr_m = info.get('sluice_arrival_minute', 0)
    sluice_dep_h = info.get('sluice_departure_hour')
    sluice_dep_m = info.get('sluice_departure_minute', 0)
    shifting_h = info.get('shifting_hour')
    shifting_m = info.get('shifting_minute', 0)
    
    # Op-ajat
    if op_start_h is not None:
        op_start = time_to_slot(op_start_h, op_start_m)
        if op_end_h is not None and op_end_h < op_start_h:
            op_end = 48
        elif op_end_h == 0 and op_start_h > 0:
            op_end = 48
        elif op_end_h is not None:
            op_end = time_to_slot(op_end_h, op_end_m)
        else:
            op_end = NORMAL_END
    else:
        op_start = NORMAL_START
        op_end = NORMAL_END
    
    return {
        'op_start': op_start,
        'op_end': op_end,
        'arrival_start': time_to_slot(arrival_h, arrival_m) if arrival_h is not None else None,
        'departure_start': time_to_slot(departure_h, departure_m) if departure_h is not None else None,
        'sluice_arr_start': time_to_slot(sluice_arr_h, sluice_arr_m) if sluice_arr_h is not None else None,
        'sluice_dep_start': time_to_slot(sluice_dep_h, sluice_dep_m) if sluice_dep_h is not None else None,
        'shifting_start': time_to_slot(shifting_h, shifting_m) if shifting_h is not None else None
    }


def apply_constraint_slots(dm_work, dm_ops, daymen, day_idx, times, constraints):
    """
    Vaihe 0.5: Lisää pakolliset slotit rajoitteista (must_work_slot).
    """
    op_start = times['op_start']
    op_end = times['op_end']
    
    for dm in daymen:
        if is_day_off(dm, day_idx, constraints):
            continue
        for slot in range(48):
            if must_work_slot(dm, slot, day_idx, constraints):
                dm_work[dm][slot] = True
                if op_start <= slot < min(op_end, 48):
                    dm_ops[dm][slot] = True


def apply_arrival_slots(dm_work, dm_arr, active_daymen, day_idx, times, constraints):
    """
    Vaihe 1.1: Tulo - kaikki aktiiviset daymanit (1h).
    """
    arrival_start = times['arrival_start']
    if arrival_start is None:
        return
    
    for dm in active_daymen:
        can_add = True
        for slot in range(arrival_start, arrival_start + 2):
            if not can_work_slot(dm, slot, day_idx, constraints, sum(dm_work[dm])/2):
                can_add = False
                break
        if can_add:
            add_block(dm_work[dm], arrival_start, arrival_start + 2, dm_arr[dm])


def apply_departure_slots(dm_work, dm_dep, active_daymen, day_idx, times, constraints):
    """
    Vaihe 1.2: Lähtö - 2 daymaniä (1h).
    """
    departure_start = times['departure_start']
    if departure_start is None:
        return
    
    scores = {}
    for dm in active_daymen:
        can_do = True
        for slot in range(departure_start, departure_start + 2):
            if not can_work_slot(dm, slot, day_idx, constraints, sum(dm_work[dm])/2):
                can_do = False
                break
        if not can_do:
            continue
        
        hours = sum(dm_work[dm]) / 2
        continuity = 1 if (departure_start > 0 and dm_work[dm][departure_start - 1]) else 0
        scores[dm] = -hours + continuity
    
    selected = sorted(scores.keys(), key=lambda x: scores[x], reverse=True)[:2]
    for dm in selected:
        add_block(dm_work[dm], departure_start, departure_start + 2, dm_dep[dm])


def apply_sluice_arrival_slots(dm_work, dm_sluice, daymen, times):
    """
    Vaihe 1.3: Slussi tulo - 1. tunti 2 dm, 2. tunti 3 dm (2h kokonaan).
    """
    sluice_arr_start = times['sluice_arr_start']
    if sluice_arr_start is None:
        return
    
    scores = {}
    for dm in daymen:
        hours = sum(dm_work[dm]) / 2
        scores[dm] = -hours
    
    first_hour_dm = sorted(daymen, key=lambda x: scores[x], reverse=True)[:2]
    
    for dm in first_hour_dm:
        add_block(dm_work[dm], sluice_arr_start, sluice_arr_start + 2, dm_sluice[dm])
    
    for dm in daymen:
        add_block(dm_work[dm], sluice_arr_start + 2, sluice_arr_start + 4, dm_sluice[dm])


def apply_sluice_departure_slots(dm_work, dm_sluice, daymen, times):
    """
    Vaihe 1.4: Slussi lähtö - 2 daymaniä (2h).
    """
    sluice_dep_start = times['sluice_dep_start']
    if sluice_dep_start is None:
        return
    
    scores = {}
    for dm in daymen:
        hours = sum(dm_work[dm]) / 2
        continuity = 1 if (sluice_dep_start > 0 and dm_work[dm][sluice_dep_start - 1]) else 0
        scores[dm] = -hours + continuity
    
    selected = sorted(daymen, key=lambda x: scores[x], reverse=True)[:2]
    for dm in selected:
        add_block(dm_work[dm], sluice_dep_start, sluice_dep_start + 4, dm_sluice[dm])


def apply_shifting_slots(dm_work, dm_shifting, daymen, times):
    """
    Vaihe 1.5: Shiftaus - kaikki daymanit (1h).
    """
    shifting_start = times['shifting_start']
    if shifting_start is None:
        return
    
    for dm in daymen:
        add_block(dm_work[dm], shifting_start, shifting_start + 2, dm_shifting[dm])


def apply_op_outside_normal_hours(dm_work, dm_ops, active_daymen, day_idx, times, 
                                   constraints, prev_day_work, continuous_night_info):
    """
    Vaihe 1.6: Satamaop 08-17 ULKOPUOLELLA - aina 1 dayman töissä.
    """
    op_start = times['op_start']
    op_end = times['op_end']
    
    op_outside_slots = []
    for slot in range(op_start, min(op_end, 48)):
        if slot < NORMAL_START or slot >= NORMAL_END:
            if slot != LUNCH_START:
                op_outside_slots.append(slot)
    
    # Jos jatkuva yö edellisestä päivästä
    if continuous_night_info is not None:
        early_worker = continuous_night_info['early_worker']
        late_worker = continuous_night_info['late_worker']
        night_split_slot = continuous_night_info['split_slot']
        
        for slot in range(0, night_split_slot):
            if slot in op_outside_slots:
                dm_work[early_worker][slot] = True
                dm_ops[early_worker][slot] = True
        
        for slot in range(night_split_slot, NORMAL_START):
            if slot in op_outside_slots:
                dm_work[late_worker][slot] = True
                dm_ops[late_worker][slot] = True
        
        op_outside_slots = [s for s in op_outside_slots if s >= NORMAL_START]
    
    current_worker = None
    preferred = get_preferred_night_worker(day_idx, constraints, active_daymen)
    
    for slot in op_outside_slots:
        can_continue = False
        
        if current_worker is not None:
            current_hours = sum(dm_work[current_worker]) / 2
            max_h = get_max_hours(current_worker, constraints)
            
            if current_hours < max_h and can_work_slot(current_worker, slot, day_idx, constraints, current_hours):
                test_work = dm_work[current_worker][:]
                test_work[slot] = True
                stcw_ok = check_stcw_ok(test_work, prev_day_work[current_worker])
                
                if stcw_ok:
                    can_continue = True
        
        if can_continue:
            dm_work[current_worker][slot] = True
            dm_ops[current_worker][slot] = True
        else:
            best_dm = _find_best_worker_for_slot(
                slot, active_daymen, current_worker, dm_work, prev_day_work,
                day_idx, constraints, preferred
            )
            
            if best_dm:
                dm_work[best_dm][slot] = True
                dm_ops[best_dm][slot] = True
                current_worker = best_dm
            elif current_worker is not None:
                current_hours = sum(dm_work[current_worker]) / 2
                max_h = get_max_hours(current_worker, constraints)
                if current_hours < max_h:
                    dm_work[current_worker][slot] = True
                    dm_ops[current_worker][slot] = True


def _find_best_worker_for_slot(slot, active_daymen, current_worker, dm_work, 
                                prev_day_work, day_idx, constraints, preferred=None):
    """
    Apufunktio: Etsii parhaan työntekijän tietylle slotille.
    """
    best_dm = None
    best_score = -9999
    
    for dm in active_daymen:
        if dm == current_worker:
            continue
        
        current_hours = sum(dm_work[dm]) / 2
        max_h = get_max_hours(dm, constraints)
        
        if current_hours >= max_h:
            continue
        
        if not can_work_slot(dm, slot, day_idx, constraints, current_hours):
            continue
        
        test_work = dm_work[dm][:]
        test_work[slot] = True
        stcw_ok = check_stcw_ok(test_work, prev_day_work[dm])
        
        if not stcw_ok:
            continue
        
        score = 0
        score += (max_h - current_hours) * 10
        
        if preferred and dm == preferred:
            score += 500
        
        prev_work = prev_day_work[dm]
        last_work_slot = -1
        for s in range(47, -1, -1):
            if prev_work[s]:
                last_work_slot = s
                break
        
        if last_work_slot >= 0:
            rest_slots = (slot + 48) - last_work_slot
            rest_hours = rest_slots / 2
            score += min(rest_hours, 24) * 5
        else:
            score += 24 * 5
        
        if score > best_score:
            best_score = score
            best_dm = dm
    
    return best_dm


# ============================================================================
# VAIHE 3: TYÖBLOKKIEN JAKAMINEN
# ============================================================================

def fill_op_inside_normal_hours(dm_work, dm_ops, active_daymen, day_idx, times, 
                                 constraints, prev_day_work):
    """
    Vaihe 3.1: Op-kattavuus 08-17 välillä.
    """
    op_start = times['op_start']
    op_end = times['op_end']
    
    op_inside_slots = []
    for slot in range(max(op_start, NORMAL_START), min(op_end, NORMAL_END)):
        if LUNCH_START <= slot < LUNCH_END:
            continue
        op_inside_slots.append(slot)
    
    for slot in op_inside_slots:
        workers_in_slot = [dm for dm in active_daymen if dm_work[dm][slot]]
        
        if len(workers_in_slot) >= 1:
            for dm in workers_in_slot:
                dm_ops[dm][slot] = True
            continue
        
        best_dm = _find_best_worker_for_inside_slot(
            slot, active_daymen, dm_work, prev_day_work, day_idx, constraints
        )
        
        if best_dm:
            dm_work[best_dm][slot] = True
            dm_ops[best_dm][slot] = True
    
    return op_inside_slots


def _find_best_worker_for_inside_slot(slot, active_daymen, dm_work, prev_day_work, 
                                       day_idx, constraints):
    """
    Apufunktio: Etsii parhaan työntekijän 08-17 slotille.
    """
    best_dm = None
    best_score = -9999
    
    for dm in active_daymen:
        current_hours = sum(dm_work[dm]) / 2
        max_h = get_max_hours(dm, constraints)
        
        if current_hours >= max_h:
            continue
        
        if not can_work_slot(dm, slot, day_idx, constraints, current_hours):
            continue
        
        score = 0
        
        if slot > 0 and dm_work[dm][slot - 1]:
            score += 200
        if slot < 47 and dm_work[dm][slot + 1]:
            score += 200
        
        test_work = dm_work[dm][:]
        test_work[slot] = True
        stcw_ok = check_stcw_ok(test_work, prev_day_work[dm])
        
        if not stcw_ok:
            continue
        
        score += (max_h - current_hours) * 10
        
        if score > best_score:
            best_score = score
            best_dm = dm
    
    return best_dm


def fill_remaining_hours(dm_work, dm_ops, active_daymen, day_idx, times, constraints):
    """
    Vaihe 3.2: Täytä loput tunnit 08:00 alkaen.
    """
    op_start = times['op_start']
    op_end = times['op_end']
    
    for dm in active_daymen:
        current_hours = sum(dm_work[dm]) / 2
        min_h = get_min_hours(dm, constraints)
        max_h = get_max_hours(dm, constraints)
        
        if current_hours >= min_h:
            continue
        
        night_work_slots = sum(1 for s in range(0, NORMAL_START) if dm_work[dm][s])
        did_night_shift = night_work_slots >= 4
        
        if did_night_shift:
            _extend_night_shift(dm, dm_work, dm_ops, op_start, op_end, min_h, max_h, 
                               day_idx, constraints)
            continue
        
        slot = NORMAL_START
        while current_hours < min_h and slot < NORMAL_END:
            if LUNCH_START <= slot < LUNCH_END:
                slot += 1
                continue
            
            if dm_work[dm][slot]:
                slot += 1
                continue
            
            if not can_work_slot(dm, slot, day_idx, constraints, current_hours):
                slot += 1
                continue
            
            if current_hours >= max_h:
                break
            
            dm_work[dm][slot] = True
            if op_start <= slot < min(op_end, 48):
                dm_ops[dm][slot] = True
            current_hours = sum(dm_work[dm]) / 2
            slot += 1


def _extend_night_shift(dm, dm_work, dm_ops, op_start, op_end, min_h, max_h, 
                        day_idx, constraints):
    """
    Apufunktio: Laajentaa yövuoroa tarvittaessa.
    """
    current_hours = sum(dm_work[dm]) / 2
    
    last_night_slot = -1
    for s in range(NORMAL_START - 1, -1, -1):
        if dm_work[dm][s]:
            last_night_slot = s
            break
    
    if last_night_slot >= 0 and last_night_slot < NORMAL_START - 1:
        for s in range(last_night_slot + 1, NORMAL_START):
            if current_hours >= min_h:
                break
            if current_hours >= max_h:
                break
            if not can_work_slot(dm, s, day_idx, constraints, current_hours):
                break
            dm_work[dm][s] = True
            if op_start <= s < min(op_end, 48):
                dm_ops[dm][s] = True
            current_hours = sum(dm_work[dm]) / 2


def ensure_op_coverage(dm_work, dm_ops, op_inside_slots, active_daymen, day_idx, constraints):
    """
    Vaihe 3.3: Varmista op-kattavuus uudelleen.
    """
    for slot in op_inside_slots:
        workers_in_slot = [dm for dm in active_daymen if dm_work[dm][slot]]
        
        if len(workers_in_slot) >= 1:
            for dm in workers_in_slot:
                dm_ops[dm][slot] = True
            continue
        
        best_dm = None
        best_score = -9999
        
        for dm in active_daymen:
            current_hours = sum(dm_work[dm]) / 2
            max_h = get_max_hours(dm, constraints)
            
            if current_hours >= max_h:
                continue
            
            if not can_work_slot(dm, slot, day_idx, constraints, current_hours):
                continue
            
            score = 0
            if slot > 0 and dm_work[dm][slot - 1]:
                score += 200
            if slot < 47 and dm_work[dm][slot + 1]:
                score += 200
            score += (max_h - current_hours) * 10
            
            if score > best_score:
                best_score = score
                best_dm = dm
        
        if best_dm:
            dm_work[best_dm][slot] = True
            dm_ops[best_dm][slot] = True


def fill_gaps_between_blocks(dm_work, dm_ops, active_daymen, day_idx, times, constraints):
    """
    Vaihe 3.4: Täytä turhat aukot blokkien välissä.
    """
    op_start = times['op_start']
    op_end = times['op_end']
    
    for dm in active_daymen:
        work = dm_work[dm]
        blocks = get_work_blocks(work)
        max_h = get_max_hours(dm, constraints)
        
        for i in range(len(blocks) - 1):
            _, block1_end = blocks[i]
            block2_start, _ = blocks[i + 1]
            
            gap = block2_start - block1_end
            
            if 0 < gap <= 4 and block1_end >= NORMAL_START and block2_start <= NORMAL_END:
                current_hours = sum(work) / 2
                
                for s in range(block1_end, block2_start):
                    if LUNCH_START <= s < LUNCH_END:
                        continue
                    if current_hours >= max_h:
                        break
                    if not can_work_slot(dm, s, day_idx, constraints, current_hours):
                        continue
                    work[s] = True
                    if op_start <= s < min(op_end, 48):
                        dm_ops[dm][s] = True
                    current_hours = sum(work) / 2


# ============================================================================
# VAIHE 4: AUKKOJEN TÄYTTÖ
# ============================================================================

def fill_small_gaps(dm_work, dm_ops, active_daymen, times):
    """
    Vaihe 4: Täytä pienet aukot (max 1h).
    """
    op_start = times['op_start']
    op_end = times['op_end']
    
    for dm in active_daymen:
        work = dm_work[dm]
        blocks = get_work_blocks(work)
        
        for i in range(len(blocks) - 1):
            _, block1_end = blocks[i]
            block2_start, _ = blocks[i + 1]
            
            gap = block2_start - block1_end
            
            if 0 < gap <= 2:
                for s in range(block1_end, block2_start):
                    if LUNCH_START <= s < LUNCH_END:
                        continue
                    work[s] = True
                    if op_start <= s < min(op_end, 48):
                        dm_ops[dm][s] = True


# ============================================================================
# BOSUN JA WATCHMANIT
# ============================================================================

def generate_bosun_schedule(times):
    """
    Generoi bosunin työvuorot (08-17 + tulo/lähtö + slussi + shiftaus).
    """
    bosun_work = [False] * 48
    bosun_arr = [False] * 48
    bosun_dep = [False] * 48
    bosun_sluice = [False] * 48
    bosun_shifting = [False] * 48
    
    arrival_start = times['arrival_start']
    departure_start = times['departure_start']
    sluice_arr_start = times['sluice_arr_start']
    sluice_dep_start = times['sluice_dep_start']
    shifting_start = times['shifting_start']
    
    if arrival_start is not None:
        add_block(bosun_work, arrival_start, arrival_start + 2, bosun_arr)
    
    if departure_start is not None:
        add_block(bosun_work, departure_start, departure_start + 2, bosun_dep)
    
    if sluice_arr_start is not None:
        add_block(bosun_work, sluice_arr_start, sluice_arr_start + 4, bosun_sluice)
    
    if sluice_dep_start is not None:
        add_block(bosun_work, sluice_dep_start, sluice_dep_start + 4, bosun_sluice)
    
    if shifting_start is not None:
        add_block(bosun_work, shifting_start, shifting_start + 2, bosun_shifting)
    
    for slot in range(NORMAL_START, NORMAL_END):
        if LUNCH_START <= slot < LUNCH_END:
            continue
        bosun_work[slot] = True
    
    return {
        'work_slots': bosun_work,
        'arrival_slots': bosun_arr,
        'departure_slots': bosun_dep,
        'port_op_slots': [False] * 48,
        'sluice_slots': bosun_sluice,
        'shifting_slots': bosun_shifting
    }


def generate_watchman_schedule():
    """
    Generoi watchmanin työvuorot (tyhjä).
    """
    return {
        'work_slots': [False] * 48,
        'arrival_slots': [False] * 48,
        'departure_slots': [False] * 48,
        'port_op_slots': [False] * 48,
        'sluice_slots': [False] * 48,
        'shifting_slots': [False] * 48
    }


# ============================================================================
# PÄÄFUNKTIO
# ============================================================================

def generate_schedule(days_data, constraints=None):
    """
    Generoi työvuorot blokkipohjaisella lähestymistavalla.
    
    Args:
        days_data: Lista päivien tiedoista
        constraints: Lista rajoitteista (valinnainen)
    
    Returns:
        (workbook, all_days, report)
    """
    if constraints is None:
        constraints = []
    
    all_days = {w: [] for w in WORKERS}
    num_days = len(days_data)
    
    # VAIHE 0: Analysoi jatkuvat yövuorot
    continuous_nights = analyze_continuous_nights(days_data)
    
    # Generoi päivä kerrallaan
    for day_idx, info in enumerate(days_data):
        # Parsitaan päivän ajat
        times = parse_day_times(info)
        
        # Edellisen päivän työvuorot
        prev_day_work = {}
        for dm in DAYMEN:
            if day_idx > 0:
                prev_day_work[dm] = all_days[dm][day_idx - 1]['work_slots']
            else:
                prev_day_work[dm] = [False] * 48
        
        # Tarkista jatkuva yö
        continuous_night_info = None
        for night_info in continuous_nights:
            if night_info['day_index'] == day_idx - 1:
                prev_early = all_days[night_info['early_worker']][day_idx - 1]['work_slots']
                prev_late = all_days[night_info['late_worker']][day_idx - 1]['work_slots']
                split_slot = choose_night_split_slot(
                    prev_early, prev_late, 
                    times['arrival_start'], times['departure_start']
                )
                continuous_night_info = {
                    'early_worker': night_info['early_worker'],
                    'late_worker': night_info['late_worker'],
                    'split_slot': split_slot
                }
                break
        
        # Alusta työvuorolistat
        dm_work = {dm: [False] * 48 for dm in DAYMEN}
        dm_arr = {dm: [False] * 48 for dm in DAYMEN}
        dm_dep = {dm: [False] * 48 for dm in DAYMEN}
        dm_ops = {dm: [False] * 48 for dm in DAYMEN}
        dm_sluice = {dm: [False] * 48 for dm in DAYMEN}
        dm_shifting = {dm: [False] * 48 for dm in DAYMEN}
        
        # Aktiiviset daymanit
        active_daymen = [dm for dm in DAYMEN if not is_day_off(dm, day_idx, constraints)]
        
        # VAIHE 0.5: Pakolliset slotit rajoitteista
        apply_constraint_slots(dm_work, dm_ops, DAYMEN, day_idx, times, constraints)
        
        # VAIHE 1: Pakolliset
        apply_arrival_slots(dm_work, dm_arr, active_daymen, day_idx, times, constraints)
        apply_departure_slots(dm_work, dm_dep, active_daymen, day_idx, times, constraints)
        apply_sluice_arrival_slots(dm_work, dm_sluice, DAYMEN, times)
        apply_sluice_departure_slots(dm_work, dm_sluice, DAYMEN, times)
        apply_shifting_slots(dm_work, dm_shifting, DAYMEN, times)
        apply_op_outside_normal_hours(
            dm_work, dm_ops, active_daymen, day_idx, times,
            constraints, prev_day_work, continuous_night_info
        )
        
        # VAIHE 3: Jaa työblokit
        op_inside_slots = fill_op_inside_normal_hours(
            dm_work, dm_ops, active_daymen, day_idx, times, constraints, prev_day_work
        )
        fill_remaining_hours(dm_work, dm_ops, active_daymen, day_idx, times, constraints)
        ensure_op_coverage(dm_work, dm_ops, op_inside_slots, active_daymen, day_idx, constraints)
        fill_gaps_between_blocks(dm_work, dm_ops, active_daymen, day_idx, times, constraints)
        
        # VAIHE 4: Täytä pienet aukot
        fill_small_gaps(dm_work, dm_ops, active_daymen, times)
        
        # Tallenna daymanien tulokset
        for dm in DAYMEN:
            all_days[dm].append({
                'work_slots': dm_work[dm],
                'arrival_slots': dm_arr[dm],
                'departure_slots': dm_dep[dm],
                'port_op_slots': dm_ops[dm],
                'sluice_slots': dm_sluice[dm],
                'shifting_slots': dm_shifting[dm]
            })
        
        # Bosun
        all_days['Bosun'].append(generate_bosun_schedule(times))
        
        # Watchmanit
        for w in ['Watchman 1', 'Watchman 2', 'Watchman 3']:
            all_days[w].append(generate_watchman_schedule())
    
    # Rakenna Excel
    wb, report = build_workbook_and_report(all_days, num_days, WORKERS)
    
    return wb, all_days, report


# ============================================================================
# EXCEL-GENEROINTI
# ============================================================================

def build_workbook_and_report(all_days, num_days, workers):
    """
    Rakentaa Excel-työkirjan ja raportin.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Työvuorot"
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    report_lines = []
    current_row = 1
    
    for d in range(num_days):
        ws.cell(row=current_row, column=1, value=f"Päivä {d+1}")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
        current_row += 1
        
        ws.cell(row=current_row, column=1, value="Työntekijä")
        for slot in range(48):
            col = slot + 2
            time_str = slot_to_time_str(slot)
            ws.cell(row=current_row, column=col, value=time_str)
            ws.cell(row=current_row, column=col).alignment = Alignment(horizontal='center')
            ws.cell(row=current_row, column=col).font = Font(size=8)
        ws.cell(row=current_row, column=50, value="Tunnit")
        current_row += 1
        
        for worker in workers:
            ws.cell(row=current_row, column=1, value=worker)
            day_data = all_days[worker][d]
            work = day_data['work_slots']
            arr = day_data['arrival_slots']
            dep = day_data['departure_slots']
            ops = day_data.get('port_op_slots', [False] * 48)
            sluice = day_data.get('sluice_slots', [False] * 48)
            shifting = day_data.get('shifting_slots', [False] * 48)
            
            hours = sum(work) / 2
            
            for slot in range(48):
                col = slot + 2
                cell = ws.cell(row=current_row, column=col)
                cell.border = thin_border
                
                if sluice[slot]:
                    cell.fill = PURPLE
                    cell.value = "SL"
                elif shifting[slot]:
                    cell.fill = PINK
                    cell.value = "SH"
                elif arr[slot]:
                    cell.fill = YELLOW
                    cell.value = "B"
                elif dep[slot]:
                    cell.fill = ORANGE
                    cell.value = "C"
                elif ops[slot]:
                    cell.fill = GREEN
                    cell.value = "OP"
                elif work[slot]:
                    cell.fill = BLUE
                    cell.value = "X"
                else:
                    cell.fill = WHITE
                
                cell.alignment = Alignment(horizontal='center')
                cell.font = Font(size=8)
            
            ws.cell(row=current_row, column=50, value=hours)
            current_row += 1
            
            ranges = get_work_ranges(work)
            report_lines.append(f"Päivä {d+1} - {worker}: {hours}h | {' + '.join(ranges)}")
        
        current_row += 1
    
    ws.column_dimensions['A'].width = 12
    for col in range(2, 50):
        ws.column_dimensions[get_column_letter(col)].width = 4
    ws.column_dimensions[get_column_letter(50)].width = 6
    
    report = "\n".join(report_lines)
    return wb, report


# ============================================================================
# MANUAALINEN PÄIVÄ 1
# ============================================================================

def generate_schedule_with_manual_day1(days_data, manual_day1_slots):
    """
    Generoi työvuorot manuaalisella päivällä 1.
    """
    if not days_data:
        return generate_schedule(days_data)
    
    all_days = {w: [] for w in WORKERS}
    num_days = len(days_data)
    
    for worker in WORKERS:
        if worker in manual_day1_slots:
            work = manual_day1_slots[worker]
        else:
            work = [False] * 48
        
        all_days[worker].append({
            'work_slots': work,
            'arrival_slots': [False] * 48,
            'departure_slots': [False] * 48,
            'port_op_slots': [False] * 48,
            'sluice_slots': [False] * 48,
            'shifting_slots': [False] * 48
        })
    
    if num_days > 1:
        _, rest_days, _ = generate_schedule(days_data[1:])
        for worker in WORKERS:
            all_days[worker].extend(rest_days[worker])
    
    wb, report = build_workbook_and_report(all_days, num_days, WORKERS)
    return wb, all_days, report
