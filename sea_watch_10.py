# -*- coding: utf-8 -*-
"""
STCW-yhteensopiva työvuorogeneraattori
Versio: Uudelleenkirjoitettu jatkuvan operaation tuella aa
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

def parse_time(time_str):
    """Parsii ajan merkkijonosta (HH:MM) tunneiksi ja minuuteiksi"""
    if not time_str or time_str == "None":
        return None, 0
    parts = time_str.split(":")
    return int(parts[0]), int(parts[1]) if len(parts) > 1 else 0

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
    continuous_nights = []  # Lista: {'day_index': int, 'early_worker': str, 'late_worker': str}
    
    for d in range(num_days - 1):
        curr = days_data[d]
        next_day = days_data[d + 1]
        
        curr_op_end = curr.get('port_op_end_hour')
        next_op_start = next_day.get('port_op_start_hour')
        
        if curr_op_end == 0 and next_op_start == 0:
            # Jatkuva operaatio yön yli
            # Yö jaetaan dynaamisesti lepoaikasäädösten mukaan
            continuous_nights.append({
                'day_index': d,
                'early_worker': 'Dayman PH1',
                'late_worker': 'Dayman PH2'
            })

    def evaluate_night_split(prev_early, prev_late, split_slot, arrival_start, arrival_end,
                             departure_start, departure_end):
        early_work = [False] * 48
        late_work = [False] * 48

        for slot in range(0, min(split_slot, 48)):
            early_work[slot] = True
        for slot in range(split_slot, min(NORMAL_START, 48)):
            late_work[slot] = True

        if arrival_start is not None:
            for i in range(arrival_start, min(arrival_end, 48)):
                early_work[i] = True
                late_work[i] = True
        if departure_start is not None:
            for i in range(departure_start, min(departure_end, 48)):
                early_work[i] = True
                late_work[i] = True

        early_analysis = analyze_stcw_from_work_starts(prev_early + early_work)
        late_analysis = analyze_stcw_from_work_starts(prev_late + late_work)

        early_issues = len(early_analysis['issues'])
        late_issues = len(late_analysis['issues'])
        total_issues = early_issues + late_issues

        min_longest_rest = min(early_analysis['longest_rest'], late_analysis['longest_rest'])
        min_total_rest = min(early_analysis['total_rest'], late_analysis['total_rest'])

        return (
            total_issues,
            -min_longest_rest,
            -min_total_rest
        ), early_analysis, late_analysis

    def choose_night_split_slot(prev_early, prev_late, arrival_start, arrival_end,
                                departure_start, departure_end):
        candidate_slots = list(range(time_to_index(1, 0), time_to_index(7, 0) + 1))
        best_slot = None
        best_score = None

        for split_slot in candidate_slots:
            score, _, _ = evaluate_night_split(
                prev_early,
                prev_late,
                split_slot,
                arrival_start,
                arrival_end,
                departure_start,
                departure_end
            )
            if best_score is None or score < best_score:
                best_score = score
                best_slot = split_slot

        return best_slot or time_to_index(3, 0)

    def ensure_min_dayman_hours(work, prev_work, min_slots):
        if sum(work) >= min_slots:
            return

        candidate_slots = [
            slot for slot in range(NORMAL_START, min(NORMAL_END, 48))
            if slot < LUNCH_START or slot >= LUNCH_END
        ]

        for slot in candidate_slots:
            if sum(work) >= min_slots:
                break
            if work[slot]:
                continue

            trial = work[:]
            trial[slot] = True
            combined = prev_work + trial
            analysis = analyze_stcw_from_work_starts(combined)
            if analysis['status'] == 'OK':
                work[slot] = True

    def add_slots(start, end, target, marker=None):
        """Lisää slotit [start, end) työksi (ja halutessa marker-listaan)."""
        if start is None or end is None:
            return
        for slot in range(max(0, start), min(end, 48)):
            target[slot] = True
            if marker is not None:
                marker[slot] = True

    def fill_remaining_hours(work, ops, target_slots, prioritize_op_window=True, mark_ops=True):
        """Täyttää puuttuvat tunnit: ensin 08-17, sitten (halutessa) operaatio, lopuksi muu."""
        if sum(work) >= target_slots:
            return

        preferred = [
            slot for slot in range(NORMAL_START, min(NORMAL_END, 48))
            if slot < LUNCH_START or slot >= LUNCH_END
        ]
        op_window = [slot for slot in range(max(0, op_start), min(op_end, 48))] if prioritize_op_window else []
        fallback = [slot for slot in range(48)]

        for slot in preferred + op_window + fallback:
            if sum(work) >= target_slots:
                break
            if LUNCH_START <= slot < LUNCH_END:
                continue
            if work[slot]:
                continue
            work[slot] = True
            if mark_ops and op_start <= slot < min(op_end, 48):
                ops[slot] = True

    def trim_excess_hours(work, ops, locked_slots, max_slots):
        """Karsii ylimääräiset slotit niin, ettei tunnit karkaa yli maksimin."""
        if sum(work) <= max_slots:
            return

        removable = [
            slot for slot in range(48)
            if work[slot] and not locked_slots[slot]
        ]

        def removal_priority(slot):
            in_day_window = NORMAL_START <= slot < NORMAL_END
            is_op = op_start <= slot < min(op_end, 48)
            # pienempi tuple poistetaan ensin
            return (
                0 if not in_day_window else 1,
                0 if not is_op else 1,
                abs(slot - NORMAL_START)
            )

        for slot in sorted(removable, key=removal_priority):
            if sum(work) <= max_slots:
                break
            work[slot] = False
            ops[slot] = False


    def enforce_rest_continuity(work, ops, prev_work, max_slots):
        """Yrittää vähentää lepojaksojen pirstaloitumista yhdistämällä lyhyitä työaukkoja."""
        def find_work_segments(slots):
            segs = []
            start = None
            for i, val in enumerate(slots):
                if val and start is None:
                    start = i
                elif not val and start is not None:
                    segs.append((start, i))
                    start = None
            if start is not None:
                segs.append((start, 48))
            return segs

        while sum(work) < max_slots:
            analysis = analyze_stcw_from_work_starts(prev_work + work)
            if analysis['rest_period_count'] <= 2:
                break

            segments = find_work_segments(work)
            gap_candidates = []
            for i in range(len(segments) - 1):
                left_end = segments[i][1]
                right_start = segments[i + 1][0]
                gap = right_start - left_end
                if 0 < gap <= 4:  # max 2h aukko
                    gap_candidates.append((gap, left_end, right_start))

            if not gap_candidates:
                break

            _, start, end = min(gap_candidates, key=lambda x: x[0])
            needed = end - start
            if sum(work) + needed > max_slots:
                break

            for slot in range(start, end):
                work[slot] = True
                if op_start <= slot < min(op_end, 48):
                    ops[slot] = True
    
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
        early_worker = None
        late_worker = None
        for night_info in continuous_nights:
            if night_info['day_index'] == d - 1:
                continues_from_night = True
                early_worker = night_info['early_worker']
                late_worker = night_info['late_worker']
                break
        
        night_split_slot = None
        if continues_from_night:
            prev_early = all_days[early_worker][d - 1]['work_slots']
            prev_late = all_days[late_worker][d - 1]['work_slots']
            night_split_slot = choose_night_split_slot(
                prev_early,
                prev_late,
                arrival_start,
                arrival_end,
                departure_start,
                departure_end
            )
        
        # Onko tämä päivä jatkuvan yön alussa?
        starts_night = False
        for night_info in continuous_nights:
            if night_info['day_index'] == d:
                starts_night = True
                break
        
        # ========================================
        # BOSUN
        # ========================================
        bosun_work = [False] * 48
        bosun_arr = [False] * 48
        bosun_dep = [False] * 48
        bosun_ops = [False] * 48
        
        # Bosun: pakolliset tulo/lähtö, ei pakollista cargo-operaatiota.
        add_slots(arrival_start, arrival_end, bosun_work, bosun_arr)
        add_slots(departure_start, departure_end, bosun_work, bosun_dep)

        # Täytä loput tunnit (~8.5h) ensisijaisesti 08-17.
        fill_remaining_hours(
            bosun_work,
            bosun_ops,
            TARGET_SLOTS,
            prioritize_op_window=False,
            mark_ops=False
        )
        
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
        
        # Rakenna cargo-operaation minimikattavuus pisteytyksellä.
        op_slots_today = [slot for slot in range(max(0, op_start), min(op_end + 1, 48))]
        op_slot_set = set(op_slots_today)

        # Valmistellaan dayman-kohtaiset päivädatat ennen täyttövaiheita.
        dayman_data = {}
        for dayman in daymen:
            work = [False] * 48
            arr = [False] * 48
            dep = [False] * 48
            ops = [False] * 48
            notes = []
            locked = [False] * 48

            # Vaihe 1: kaikille daymaneille pakolliset tulo/lähtö.
            add_slots(arrival_start, arrival_end, work, arr)
            add_slots(departure_start, departure_end, work, dep)
            if arrival_start is not None:
                for slot in range(arrival_start, min(arrival_end, 48)):
                    locked[slot] = True
            if departure_start is not None:
                for slot in range(departure_start, min(departure_end, 48)):
                    locked[slot] = True

            # ---- JATKUVAN YÖN KÄSITTELY ----
            if continues_from_night and dayman in (early_worker, late_worker) and night_split_slot is not None:
                if dayman == early_worker:
                    notes.append(f"Yövuoro 00-{index_to_time_str(night_split_slot)}")
                    for slot in range(0, min(night_split_slot, 48)):
                        work[slot] = True
                        locked[slot] = True
                        if op_start <= slot < min(op_end, 48):
                            ops[slot] = True
                else:
                    notes.append(f"Yövuoro {index_to_time_str(night_split_slot)}-08")
                    for slot in range(night_split_slot, min(NORMAL_START, 48)):
                        work[slot] = True
                        locked[slot] = True
                        if op_start <= slot < min(op_end, 48):
                            ops[slot] = True

            # Yön aloituspäivä: PH1 jatkaa iltaan/yöhön, jotta jatkuvuus säilyy.
            if starts_night and dayman == 'Dayman PH1':
                notes.append('Yön aloitus, iltajatko')
                evening_start = max(NORMAL_END, op_start)
                for slot in range(evening_start, 48):
                    work[slot] = True
                    locked[slot] = True
                    if op_start <= slot < min(op_end, 48):
                        ops[slot] = True

            dayman_data[dayman] = {
                'work': work,
                'arr': arr,
                'dep': dep,
                'ops': ops,
                'notes': notes,
                'locked': locked
            }

        # Cargo-ops jako kolmessa vaiheessa:
        # 1) off-hours ensin 4h blokeissa, 2) loput cargo-slotit pisteytyksellä,
        # 3) muu tuntitäyttö tehdään dayman-kohtaisessa viimeistelyssä.
        cargo_owner_per_slot = {}

        def stcw_allows_slot(dayman, slot):
            prev_work = all_days[dayman][d - 1]['work_slots'] if d > 0 else [False] * 48
            trial = dayman_data[dayman]['work'][:]
            trial[slot] = True
            combined = prev_work + trial
            return analyze_stcw_from_work_starts(combined)['status'] == 'OK'

        def rest_slots_before(worker, slot):
            if slot <= 0:
                return 0
            slots = dayman_data[worker]['work']
            rest = 0
            i = slot - 1
            while i >= 0 and not slots[i]:
                rest += 1
                i -= 1
            return rest

        def owner_total_slots(worker):
            return sum(dayman_data[worker]['work'])

        def assign_slot(worker, slot):
            cargo_owner_per_slot[slot] = worker
            dayman_data[worker]['work'][slot] = True
            dayman_data[worker]['ops'][slot] = True
            dayman_data[worker]['locked'][slot] = True

        def slot_score(worker, slot):
            if not stcw_allows_slot(worker, slot):
                return -10**9

            score = 100

            if NORMAL_START <= slot < NORMAL_END:
                score += 35
            else:
                score -= 35

            if dayman_data[worker]['work'][slot]:
                score += 40

            # Lepo painaa enemmän kuin 08-17 bonus.
            rest_before = rest_slots_before(worker, slot)
            if 0 < rest_before < 6:
                score -= 180
            elif rest_before >= 12:
                score += 90
            elif rest_before >= 8:
                score += 50

            # Off-hours: vältä päällekkäisyys erittäin vahvasti.
            if slot < NORMAL_START or slot >= NORMAL_END:
                others = any(dayman_data[o]['work'][slot] for o in daymen if o != worker)
                if others:
                    score -= 260

            # Kevyt kuormatasoitus
            score -= owner_total_slots(worker) * 2
            return score

        if op_slots_today:
            off_hours_slots = [s for s in op_slots_today if s < NORMAL_START or s >= NORMAL_END]
            day_hours_slots = [s for s in op_slots_today if NORMAL_START <= s < NORMAL_END]

            # Vaihe A: off-hours 4h blokeissa (8 slottia)
            block_size = 8
            for block_start in range(0, len(off_hours_slots), block_size):
                block = off_hours_slots[block_start:block_start + block_size]
                if not block:
                    continue

                best_worker = None
                best_score = None
                for w in daymen:
                    # arvioi koko blokki samalle workerille
                    block_score = 0
                    feasible = True
                    for slot in block:
                        s = slot_score(w, slot)
                        if s <= -10**8:
                            feasible = False
                            break
                        block_score += s
                    # jatkuvuusbonus jos blokki jatkaa samaa workeria
                    prev_slot = block[0] - 1
                    if cargo_owner_per_slot.get(prev_slot) == w:
                        block_score += 80

                    if not feasible:
                        continue
                    if best_score is None or block_score > best_score:
                        best_score = block_score
                        best_worker = w

                if best_worker is None:
                    # fallback vähiten kuormitetulle
                    best_worker = min(daymen, key=owner_total_slots)

                for slot in block:
                    assign_slot(best_worker, slot)

            # Vaihe B: loput cargo-slotit pisteytyksellä (slot-by-slot)
            for slot in day_hours_slots:
                if slot in cargo_owner_per_slot:
                    continue
                best_worker = None
                best_score = None
                for w in daymen:
                    s = slot_score(w, slot)
                    if cargo_owner_per_slot.get(slot - 1) == w:
                        s += 35
                    if best_score is None or s > best_score:
                        best_score = s
                        best_worker = w

                if best_worker is None or best_score <= -10**8:
                    feasible = [w for w in daymen if stcw_allows_slot(w, slot)]
                    best_worker = feasible[0] if feasible else daymen[0]

                assign_slot(best_worker, slot)

        for dayman in daymen:
            work = dayman_data[dayman]['work']
            arr = dayman_data[dayman]['arr']
            dep = dayman_data[dayman]['dep']
            ops = dayman_data[dayman]['ops']
            notes = dayman_data[dayman]['notes']
            locked = dayman_data[dayman]['locked']

            # Vaihe 2: täydennä tunnit (~8.5h), painota 08-17.
            prev_work = all_days[dayman][d - 1]['work_slots'] if d > 0 else [False] * 48
            ensure_min_dayman_hours(work, prev_work, time_to_index(8, 0))
            fill_remaining_hours(work, ops, TARGET_SLOTS, prioritize_op_window=True, mark_ops=True)
            trim_excess_hours(work, ops, locked, MAX_SLOTS)
            enforce_rest_continuity(work, ops, prev_work, MAX_SLOTS)

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
