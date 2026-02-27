# -*- coding: utf-8 -*-
"""
Created on Tue Feb 17 10:49:16 2026

@author: OMISTAJA
"""

# -*- coding: utf-8 -*-
"""
STCW-yhteensopiva työvuorogeneraattori
Versio 14: Käyttäjän määrittelemä pisteytys

Pisteytys:
  +5000  Slotti 08-17 ulkopuolella, ei vielä työntekijää
  +200   Jatkumo tulo/lähtöön (alle 2h JA edeltävässä slotissa töitä)
  +100   Edeltävässä slotissa töitä (jatkuvuus)
  +100   Slotti klo 08:00
  +50    Slotti 08-17 välillä
  -1000  08-17 ulkopuolella ja joku muu jo töissä
  -2000  STCW-rike
  -5000  Jo 6h+ yö/iltatyötä (08-17 ulkopuolella)
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# Värit
YELLOW = PatternFill("solid", fgColor="FFFF00")
GREEN = PatternFill("solid", fgColor="90EE90")
BLUE = PatternFill("solid", fgColor="ADD8E6")
ORANGE = PatternFill("solid", fgColor="FFA500")
PURPLE = PatternFill("solid", fgColor="C9A0DC")
PINK = PatternFill("solid", fgColor="FFB6C1")
WHITE = PatternFill("solid", fgColor="FFFFFF")

# Vakiot
NORMAL_START = 16   # 08:00
NORMAL_END = 34     # 17:00
LUNCH_START = 23    # 11:30
LUNCH_END = 24      # 12:00
MIN_HOURS = 8
MAX_HOURS = 10
MAX_OUTSIDE_NORMAL_SLOTS = 12  # 6h = 12 slottia
SHORT_SEGMENT_MAX_SLOTS = 2  # 1h tai lyhyempi pätkä pyritään poistamaan


def time_to_index(h, m):
    return h * 2 + (1 if m >= 30 else 0)


def index_to_time_str(idx):
    h = idx // 2
    m = "30" if idx % 2 else "00"
    return f"{h:02d}:{m}"


# ============================================================================
# STCW-TARKISTUS
# ============================================================================

def check_stcw_at_slot(all_work_slots, slot_index):
    start = max(0, slot_index - 47)
    end = slot_index + 1
    window = all_work_slots[start:end]
    
    if len(window) < 48:
        padding = [False] * (48 - len(window))
        window = padding + window
    
    rest_periods = []
    current_rest = 0
    
    for is_work in window:
        if not is_work:
            current_rest += 1
        else:
            if current_rest > 0:
                hours = current_rest / 2
                if hours >= 1.0:
                    rest_periods.append(hours)
            current_rest = 0
    
    if current_rest > 0:
        hours = current_rest / 2
        if hours >= 1.0:
            rest_periods.append(hours)
    
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
        'status': 'OK' if not issues else 'RIKE',
        'issues': issues,
        'total_rest': total_rest,
        'longest_rest': longest_rest,
        'rest_period_count': rest_period_count
    }




def analyze_stcw_from_work_starts(all_work_slots):
    """Yhteensopivuusapu: analysoi viimeisen 24h STCW-tila.

    Huomioi vuorokauden rajan yli jatkuvan lepojakson yhtenä jaksona
    (ensimmäinen ja viimeinen lepojakso yhdistetään tarvittaessa).
    """
    window = list(all_work_slots[-48:])
    if len(window) < 48:
        window = [False] * (48 - len(window)) + window

    rest_periods = []
    current = 0
    for is_work in window:
        if not is_work:
            current += 1
        elif current > 0:
            if current >= 2:
                rest_periods.append(current / 2)
            current = 0

    if current > 0:
        if current >= 2:
            rest_periods.append(current / 2)

    if rest_periods and not window[0] and not window[-1] and len(rest_periods) >= 2:
        rest_periods[0] += rest_periods[-1]
        rest_periods.pop()

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
        'status': 'OK' if not issues else 'RIKE',
        'issues': issues,
        'total_rest': total_rest,
        'longest_rest': longest_rest,
        'rest_period_count': rest_period_count
    }

def would_cause_stcw_violation(slot, current_work, prev_day_work):
    test_work = current_work[:]
    test_work[slot] = True
    all_work = prev_day_work + test_work
    abs_slot = 48 + slot
    result = check_stcw_at_slot(all_work, abs_slot)
    return result['status'] == 'RIKE', result


def count_work_periods(work):
    periods = 0
    in_work = False
    for w in work:
        if w and not in_work:
            periods += 1
            in_work = True
        elif not w:
            in_work = False
    return periods


def is_during_op_slot(slot, op_start, op_end):
    return op_start <= slot < min(op_end, 48)


def violates_single_outside_worker_rule(dayman, slot, all_dayman_work, daymen, op_start, op_end):
    """
    Jos 08-17 ulkopuolella on satamaoperaatio käynnissä,
    vain yksi dayman saa olla töissä per slotti.
    """
    is_outside_normal = slot < NORMAL_START or slot >= NORMAL_END
    if not is_outside_normal:
        return False
    if not is_during_op_slot(slot, op_start, op_end):
        return False

    return any(other != dayman and all_dayman_work[other][slot] for other in daymen)



# ============================================================================
# PISTEYTYS - VAIN KÄYTTÄJÄN MÄÄRITTELEMÄT SÄÄNNÖT
# ============================================================================

def score_slot(slot, dayman, dayman_work, all_dayman_work, prev_day_work,
               daymen, arrival_start, arrival_end, departure_start, departure_end, op_start, op_end):
    """
    Pisteyttää slotin.
    
    +5000  Slotti 08-17 ulkopuolella, ei vielä työntekijää, cargo op käynnissä
    +200   Jatkumo tulo/lähtöön (alle 2h JA edeltävässä slotissa töitä)
    +100   Edeltävässä slotissa töitä
    +100   Slotti klo 08:00
    +50    Slotti 08-17 välillä
    -1000  08-17 ulkopuolella ja joku muu jo töissä
    -2000  STCW-rike
    -5000  Jo 6h+ yö/iltatyötä (08-17 ulkopuolella)
    """
    
    # Ehdottomat estot (palauttaa heti)
    if dayman_work[slot]:
        return -10000  # Jo täytetty
    
    if LUNCH_START <= slot < LUNCH_END:
        return -10000  # Lounastauko
    
    # Kolmas työjakso - esto (STCW-rike)
    test_work = dayman_work[:]
    test_work[slot] = True
    if count_work_periods(test_work) > 3:
        return -10000
    
    # Aloita pistelasku
    score = 0
    
    # STCW-rike: estä valinta käytännössä kokonaan
    would_violate, _ = would_cause_stcw_violation(slot, dayman_work, prev_day_work)
    if would_violate:
        return -10000
    
    # Laske 08-17 ulkopuoliset tunnit (slotit)
    outside_normal_slots = sum(1 for i in range(48)
                                if dayman_work[i] and (i < NORMAL_START or i >= NORMAL_END))

    # Onko tämä slotti 08-17 ulkopuolella?
    is_outside_normal = slot < NORMAL_START or slot >= NORMAL_END

    if is_outside_normal:
        # Kova sääntö: kun satamaoperaatio käy 08-17 ulkopuolella,
        # slotissa saa olla vain yksi dayman.
        if violates_single_outside_worker_rule(dayman, slot, all_dayman_work, daymen, op_start, op_end):
            return -10000

        # -5000: Jo 6h+ yö/iltatyötä
        if outside_normal_slots >= MAX_OUTSIDE_NORMAL_SLOTS:
            score -= 5000

        is_during_op = is_during_op_slot(slot, op_start, op_end)
        if not is_during_op:
            # -1000: 08-17 ulkopuolella mutta cargo op EI käynnissä
            score -= 1000
        else:
            # +5000: Slotti 08-17 ulkopuolella, ei vielä työntekijää, JA cargo op käynnissä
            score += 5000
    else:
        # +50: Slotti 08-17 välillä
        score += 50
    
    # +100: Slotti klo 08:00
    if slot == NORMAL_START:
        score += 100
    
    # +100: Edeltävässä slotissa töitä (jatkuvuus)
    if slot > 0 and dayman_work[slot - 1]:
        score += 100
        
        # +200: Jatkumo tulo/lähtöön (alle 2h JA edeltävässä slotissa töitä)
        if arrival_start is not None:
            distance_to_arrival = abs(slot - arrival_start)
            if distance_to_arrival <= 4:  # Alle 2h
                score += 200
        
        if departure_start is not None:
            distance_to_departure = abs(slot - departure_start)
            if distance_to_departure <= 4:  # Alle 2h
                score += 200
                
    if prev_day_work:
    # Etsi milloin edellisen päivän työ loppui
        last_work_slot = -1
        for i in range(47, -1, -1):
            if prev_day_work[i]:
                last_work_slot = i
                break
        
        if last_work_slot >= 0:
           # Jos edellinen päivä loppui slottiin 46-47 JA tämä on slotti 0-3,
           # työ jatkuu keskiyön yli - ei tarkisteta lepoaikaa
           if last_work_slot >= 46 and slot <= 3:
               pass  # Salli jatkuva vuoro keskiyön yli
           else:
               # Laske lepoaika: (48 - last_work_slot - 1) + slot
               rest_slots = (48 - last_work_slot - 1) + slot
               rest_hours = rest_slots / 2
           
               if rest_hours < 6:
                   score -= 20000  # Liian vähän lepoa - iso miinus
               elif rest_hours < 8:
                   score -= 500    # Alle 8h lepoa - pieni miinus
    

    
    return score


# ============================================================================
# APUFUNKTIOT
# ============================================================================

def add_slots(start, end, work, marker=None):
    if start is None or end is None:
        return
    for slot in range(max(0, start), min(end, 48)):
        work[slot] = True
        if marker is not None:
            marker[slot] = True


def get_work_ranges(work):
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
    return ranges


# ============================================================================
# PÄÄFUNKTIO
# ============================================================================

def _enforce_departure_lock(daymen, all_dayman_work, all_dayman_dep, departure_start, departure_end):
    if departure_start is None:
        return
    for slot in range(max(0, departure_start), min(departure_end, 48)):
        chosen = [dm for dm in daymen if all_dayman_dep[dm][slot]]
        if len(chosen) < 2:
            extras = [dm for dm in daymen if dm not in chosen]
            extras.sort(key=lambda dm: sum(all_dayman_work[dm]))
            chosen.extend(extras[: 2 - len(chosen)])
        elif len(chosen) > 2:
            chosen = sorted(chosen, key=lambda dm: sum(all_dayman_work[dm]))[:2]

        chosen_set = set(chosen[:2])
        for dm in daymen:
            all_dayman_work[dm][slot] = dm in chosen_set
            all_dayman_dep[dm][slot] = dm in chosen_set




def _trim_redundant_short_segments(daymen, all_dayman_work, all_dayman_ops, mandatory_slots,
                                   op_start, op_end, short_segment_max_slots=SHORT_SEGMENT_MAX_SLOTS):
    """Poistaa tarpeettomat lyhyet työpätkät (esim. 30-60 min), jos kattavuus säilyy.

    Segmentti poistetaan vain jos:
      - segmentissä ei ole pakollisia slotteja (tulo/lähtö/slussi/shiftaus)
      - jokaisessa segmentin slotissa on vähintään yksi muu dayman töissä
      - poistamisen jälkeen daymanilla on edelleen vähintään MIN_HOURS
    """
    for dayman in daymen:
        work = all_dayman_work[dayman]
        segments = []
        seg_start = None
        for slot, is_work in enumerate(work):
            if is_work and seg_start is None:
                seg_start = slot
            elif not is_work and seg_start is not None:
                segments.append((seg_start, slot))
                seg_start = None
        if seg_start is not None:
            segments.append((seg_start, 48))

        for start, end in segments:
            seg_len = end - start
            if seg_len == 0 or seg_len > short_segment_max_slots:
                continue

            seg_slots = list(range(start, end))
            if any(slot in mandatory_slots for slot in seg_slots):
                continue

            if (sum(work) - seg_len) / 2 < MIN_HOURS:
                continue

            has_coverage = all(
                any(other != dayman and all_dayman_work[other][slot] for other in daymen)
                for slot in seg_slots
            )
            if not has_coverage:
                continue

            for slot in seg_slots:
                all_dayman_work[dayman][slot] = False
                if op_start <= slot < min(op_end, 48):
                    all_dayman_ops[dayman][slot] = False

def generate_schedule(days_data):
    workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
               'Watchman 1', 'Watchman 2', 'Watchman 3']
    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
    
    all_days = {w: [] for w in workers}
    num_days = len(days_data)
    
    for d, info in enumerate(days_data):
        
        # Hae ajat
        op_start_h = info.get('port_op_start_hour')
        op_start_m = info.get('port_op_start_minute', 0)
        op_end_h = info.get('port_op_end_hour')
        op_end_m = info.get('port_op_end_minute', 0)
        
        arrival_h = info.get('arrival_hour')
        arrival_m = info.get('arrival_minute', 0)
        departure_h = info.get('departure_hour')
        departure_m = info.get('departure_minute', 0)

        sluice_arr_h = info.get('sluice_arrival_hour', info.get('sluice_hour'))
        sluice_arr_m = info.get('sluice_arrival_minute', info.get('sluice_minute', 0))
        sluice_dep_h = info.get('sluice_departure_hour')
        sluice_dep_m = info.get('sluice_departure_minute', 0)
        shifting_h = info.get('shifting_hour')
        shifting_m = info.get('shifting_minute', 0)
        
        # Muunna indekseiksi
        if op_start_h is not None:
            op_start = time_to_index(op_start_h, op_start_m)
            if op_end_h is not None and op_end_h < op_start_h:
                op_end = time_to_index(op_end_h, op_end_m) + 48
            elif op_end_h == 0 and op_start_h > 0:
                op_end = 48
            elif op_end_h is not None:
                op_end = time_to_index(op_end_h, op_end_m)
            else:
                op_end = NORMAL_END
        else:
            op_start = NORMAL_START
            op_end = NORMAL_END
        
        arrival_start = time_to_index(arrival_h, arrival_m) if arrival_h is not None else None
        arrival_end = arrival_start + 2 if arrival_start is not None else None
        
        departure_start = time_to_index(departure_h, departure_m) if departure_h is not None else None
        departure_end = departure_start + 2 if departure_start is not None else None

        sluice_arrival_start = time_to_index(sluice_arr_h, sluice_arr_m) if sluice_arr_h is not None else None
        sluice_departure_start = time_to_index(sluice_dep_h, sluice_dep_m) if sluice_dep_h is not None else None
        shifting_start = time_to_index(shifting_h, shifting_m) if shifting_h is not None else None
        
        # Edellisen päivän työvuorot
        prev_day_work = {}
        for dayman in daymen:
            if d > 0:
                prev_day_work[dayman] = all_days[dayman][d - 1]['work_slots']
            else:
                prev_day_work[dayman] = [False] * 48
        
        # ========================================
        # BOSUN
        # ========================================
        bosun_work = [False] * 48
        bosun_arr = [False] * 48
        bosun_dep = [False] * 48
        bosun_ops = [False] * 48
        
        add_slots(arrival_start, arrival_end, bosun_work, bosun_arr)
        add_slots(departure_start, departure_end, bosun_work, bosun_dep)
        
        for slot in range(NORMAL_START, NORMAL_END):
            if slot != LUNCH_START:
                bosun_work[slot] = True
        
        all_days['Bosun'].append({
            'work_slots': bosun_work,
            'arrival_slots': bosun_arr,
            'departure_slots': bosun_dep,
            'port_op_slots': bosun_ops,
            'sluice_slots': [False] * 48,
            'shifting_slots': [False] * 48,
            'notes': []
        })
        
        # ========================================
        # DAYMANIT
        # ========================================
        
        all_dayman_work = {dm: [False] * 48 for dm in daymen}
        all_dayman_arr = {dm: [False] * 48 for dm in daymen}
        all_dayman_dep = {dm: [False] * 48 for dm in daymen}
        all_dayman_ops = {dm: [False] * 48 for dm in daymen}
        all_dayman_sluice = {dm: [False] * 48 for dm in daymen}
        all_dayman_shifting = {dm: [False] * 48 for dm in daymen}
        
        # VAIHE 1: Tulot kaikille
        for dayman in daymen:
            add_slots(arrival_start, arrival_end, all_dayman_work[dayman], all_dayman_arr[dayman])
        
        # VAIHE 2: Lähdöt kaikille daymaneille
        if departure_start is not None:
            departure_scores = {}
            for dayman in daymen:
                score = 0
                hours = sum(all_dayman_work[dayman]) / 2
                score += (10 - hours) * 10
                for i in range(max(0, departure_start - 4), departure_start):
                    if all_dayman_work[dayman][i]:
                        score += 50
                departure_scores[dayman] = score

            sorted_daymen = sorted(departure_scores.keys(), key=lambda x: departure_scores[x], reverse=True)[:2]
            for dayman in sorted_daymen:
                add_slots(departure_start, departure_end, all_dayman_work[dayman], all_dayman_dep[dayman])

        # VAIHE 2.1: Slussi - tulo (2h): 1. tunti kaksi daymania, 2. tunti kaikki daymanit
        sluice_two_dayman_slots = {}
        first_hour_slots = []
        second_hour_slots = []
        first_hour_daymen = []
        departure_sluice_slots = []
        departure_daymen = []
        shifting_slots = []

        if sluice_arrival_start is not None:
            first_hour_slots = list(range(max(0, sluice_arrival_start), min(sluice_arrival_start + 2, 48)))
            second_hour_slots = list(range(max(0, sluice_arrival_start + 2), min(sluice_arrival_start + 4, 48)))

            sluice_arr_scores = {}
            for dayman in daymen:
                score = 0
                if sluice_arrival_start > 0 and all_dayman_work[dayman][sluice_arrival_start - 1]:
                    score += 150
                score += (10 - (sum(all_dayman_work[dayman]) / 2)) * 10
                sluice_arr_scores[dayman] = score

            first_hour_daymen = sorted(daymen, key=lambda dm: sluice_arr_scores[dm], reverse=True)[:2]
            for slot in first_hour_slots:
                sluice_two_dayman_slots[slot] = set(first_hour_daymen)
                for dayman in first_hour_daymen:
                    all_dayman_work[dayman][slot] = True
                    all_dayman_sluice[dayman][slot] = True

            for slot in second_hour_slots:
                for dayman in daymen:
                    all_dayman_work[dayman][slot] = True
                    all_dayman_sluice[dayman][slot] = True

        # VAIHE 2.2: Slussi - lähtö (2h): koko ajan kaksi daymania
        if sluice_departure_start is not None:
            departure_sluice_slots = list(range(max(0, sluice_departure_start), min(sluice_departure_start + 4, 48)))

            sluice_dep_scores = {}
            for dayman in daymen:
                score = 0
                if sluice_departure_start > 0 and all_dayman_work[dayman][sluice_departure_start - 1]:
                    score += 150
                score += (10 - (sum(all_dayman_work[dayman]) / 2)) * 10
                sluice_dep_scores[dayman] = score

            departure_daymen = sorted(daymen, key=lambda dm: sluice_dep_scores[dm], reverse=True)[:2]
            for slot in departure_sluice_slots:
                sluice_two_dayman_slots[slot] = set(departure_daymen)
                for dayman in departure_daymen:
                    all_dayman_work[dayman][slot] = True
                    all_dayman_sluice[dayman][slot] = True

        # VAIHE 2.3: Shiftaus (1h): kaikki daymanit paikalla koko ajan
        if shifting_start is not None:
            shifting_slots = list(range(max(0, shifting_start), min(shifting_start + 2, 48)))
            for slot in shifting_slots:
                for dayman in daymen:
                    all_dayman_work[dayman][slot] = True
                    all_dayman_shifting[dayman][slot] = True

        def assign_best_dayman(slot):
            # Jatkuvuussääntö: jos ollaan 08-17 ulkopuolella ja OP käynnissä,
            # jatketaan samaa yövuorolaista kuin edellisessä slotissa aina kun mahdollista.
            is_outside_normal = slot < NORMAL_START or slot >= NORMAL_END
            if is_outside_normal and is_during_op_slot(slot, op_start, op_end) and slot > 0:
                prev_workers = [dm for dm in daymen if all_dayman_work[dm][slot - 1]]
                if len(prev_workers) == 1:
                    carry_dm = prev_workers[0]
                    carry_score = score_slot(
                        slot, carry_dm, all_dayman_work[carry_dm], all_dayman_work,
                        prev_day_work[carry_dm], daymen,
                        arrival_start, arrival_end, departure_start, departure_end, op_start, op_end
                    )
                    if carry_score > -10000:
                        all_dayman_work[carry_dm][slot] = True
                        if op_start <= slot < min(op_end, 48):
                            all_dayman_ops[carry_dm][slot] = True
                        return

            scores = {}
            for dayman in daymen:
                forced_daymen = sluice_two_dayman_slots.get(slot)
                if forced_daymen is not None and dayman not in forced_daymen:
                    scores[dayman] = -10000
                    continue
                scores[dayman] = score_slot(
                    slot, dayman, all_dayman_work[dayman], all_dayman_work,
                    prev_day_work[dayman], daymen,
                    arrival_start, arrival_end, departure_start, departure_end, op_start, op_end
                )
            
            best_dayman = max(scores, key=scores.get)
            best_score = scores[best_dayman]
            
            if best_score > 0:
                all_dayman_work[best_dayman][slot] = True
                if op_start <= slot < min(op_end, 48):
                    all_dayman_ops[best_dayman][slot] = True

        mandatory_slots = set()
        for dm in daymen:
            for i in range(48):
                if (all_dayman_arr[dm][i] or all_dayman_dep[dm][i] or
                        all_dayman_sluice[dm][i] or all_dayman_shifting[dm][i]):
                    mandatory_slots.add(i)

        # VAIHE 3: Täytä satamaop-slotit ensin 08-17 ulkopuolelta
        outside_op_slots = []
        for slot in range(max(0, op_start), min(op_end, 48)):
            if slot == LUNCH_START:
                continue
            if slot < NORMAL_START or slot >= NORMAL_END:
                outside_op_slots.append(slot)

        for slot in outside_op_slots:
            assign_best_dayman(slot)

        # VAIHE 4: Täytä muut satamaop-slotit (esim. 08-17 sisällä)
        other_slots = []
        for slot in range(max(0, op_start), min(op_end, 48)):
            if slot == LUNCH_START:
                continue
            if slot in mandatory_slots:
                continue
            if slot in outside_op_slots:
                continue
            other_slots.append(slot)

        for slot in other_slots:
            assign_best_dayman(slot)

        # VAIHE 5: Täytä aukot (max 2h)
        for dayman in daymen:
            i = 0
            while i < 48:
                if all_dayman_work[dayman][i]:
                    while i < 48 and all_dayman_work[dayman][i]:
                        i += 1
                    gap_start = i
                    while i < 48 and not all_dayman_work[dayman][i]:
                        i += 1
                    if i < 48:
                        gap_end = i
                        gap_slots = gap_end - gap_start
                        if gap_slots <= 4 and gap_slots > 0:
                            for s in range(gap_start, gap_end):
                                if LUNCH_START <= s < LUNCH_END:
                                    continue
                                forced_daymen = sluice_two_dayman_slots.get(s)
                                if forced_daymen is not None and dayman not in forced_daymen:
                                    continue
                                if violates_single_outside_worker_rule(dayman, s, all_dayman_work, daymen, op_start, op_end):
                                    continue
                                all_dayman_work[dayman][s] = True
                                if op_start <= s < min(op_end, 48):
                                    all_dayman_ops[dayman][s] = True
                else:
                    i += 1
        
        # VAIHE 6: Varmista minimi 8h
        for dayman in daymen:
            current_hours = sum(all_dayman_work[dayman]) / 2
    
            while current_hours < MIN_HOURS:
                best_slot = None
                best_score = -99999
        
        # Etsi vain 08-17 väliltä
                for slot in range(NORMAL_START, NORMAL_END):
                    if all_dayman_work[dayman][slot]:
                        continue
                    if LUNCH_START <= slot < LUNCH_END:
                        continue
                    forced_daymen = sluice_two_dayman_slots.get(slot)
                    if forced_daymen is not None and dayman not in forced_daymen:
                        continue

                    score = score_slot(
                        slot, dayman, all_dayman_work[dayman], all_dayman_work,
                        prev_day_work[dayman], daymen,
                        arrival_start, arrival_end, departure_start, departure_end, op_start, op_end
                    )
            
                    if score > best_score:
                        best_score = score
                        best_slot = slot
                    
                if best_slot is not None:
                    all_dayman_work[dayman][best_slot] = True
                    if op_start <= best_slot < min(op_end, 48):
                        all_dayman_ops[dayman][best_slot] = True
                    current_hours = sum(all_dayman_work[dayman]) / 2
                else:
                    break
        # VAIHE 5b: Pakota minimi 8h (vain 08-17 väliltä)
        for dayman in daymen:
            current_hours = sum(all_dayman_work[dayman]) / 2
    
            if current_hours < MIN_HOURS:
        # Lisää slotteja 08-17 väliltä kunnes 8h
                for slot in range(NORMAL_START, NORMAL_END):
                    if current_hours >= MIN_HOURS:
                        break
                    if all_dayman_work[dayman][slot]:
                        continue
                    if LUNCH_START <= slot < LUNCH_END:
                        continue
            
                    all_dayman_work[dayman][slot] = True
                    if op_start <= slot < min(op_end, 48):
                        all_dayman_ops[dayman][slot] = True
                    current_hours = sum(all_dayman_work[dayman]) / 2
    
        
        # VAIHE 7: Täytä aukot uudelleen
        for dayman in daymen:
            i = 0
            while i < 48:
                if all_dayman_work[dayman][i]:
                    while i < 48 and all_dayman_work[dayman][i]:
                        i += 1
                    gap_start = i
                    while i < 48 and not all_dayman_work[dayman][i]:
                        i += 1
                    if i < 48:
                        gap_end = i
                        gap_slots = gap_end - gap_start
                        if gap_slots <= 4 and gap_slots > 0:
                            for s in range(gap_start, gap_end):
                                if LUNCH_START <= s < LUNCH_END:
                                    continue
                                forced_daymen = sluice_two_dayman_slots.get(s)
                                if forced_daymen is not None and dayman not in forced_daymen:
                                    continue
                                if violates_single_outside_worker_rule(dayman, s, all_dayman_work, daymen, op_start, op_end):
                                    continue
                                all_dayman_work[dayman][s] = True
                                if op_start <= s < min(op_end, 48):
                                    all_dayman_ops[dayman][s] = True
                else:
                    i += 1
        
        # VAIHE 8: Siisti lyhyet tarpeettomat työpätkät (esim. 30-60 min)
        _trim_redundant_short_segments(
            daymen,
            all_dayman_work,
            all_dayman_ops,
            mandatory_slots,
            op_start,
            op_end,
        )

        _enforce_departure_lock(daymen, all_dayman_work, all_dayman_dep, departure_start, departure_end)

        # Tallenna
        # Varmistus: erikoisoperaatioiden pakolliset läsnäolot pysyvät aina lopputuloksessa.
        for slot in first_hour_slots:
            forced = set(first_hour_daymen)
            for dayman in daymen:
                if dayman not in forced:
                    all_dayman_work[dayman][slot] = False
                    all_dayman_sluice[dayman][slot] = False
            for dayman in first_hour_daymen:
                all_dayman_work[dayman][slot] = True
                all_dayman_sluice[dayman][slot] = True
        for slot in second_hour_slots:
            for dayman in daymen:
                all_dayman_work[dayman][slot] = True
                all_dayman_sluice[dayman][slot] = True
        for slot in departure_sluice_slots:
            forced = set(departure_daymen)
            for dayman in daymen:
                if dayman not in forced:
                    all_dayman_work[dayman][slot] = False
                    all_dayman_sluice[dayman][slot] = False
            for dayman in departure_daymen:
                all_dayman_work[dayman][slot] = True
                all_dayman_sluice[dayman][slot] = True
        for slot in shifting_slots:
            for dayman in daymen:
                all_dayman_work[dayman][slot] = True
                all_dayman_shifting[dayman][slot] = True

        for dayman in daymen:
            all_days[dayman].append({
                'work_slots': all_dayman_work[dayman],
                'arrival_slots': all_dayman_arr[dayman],
                'departure_slots': all_dayman_dep[dayman],
                'port_op_slots': all_dayman_ops[dayman],
                'sluice_slots': all_dayman_sluice[dayman],
                'shifting_slots': all_dayman_shifting[dayman],
                'notes': []
            })
        
        # WATCHMANIT
        watch_schedules = {
            'Watchman 1': [(0, 4), (12, 16)],
            'Watchman 2': [(4, 8), (16, 20)],
            'Watchman 3': [(8, 12), (20, 24)]
        }
        
        for watchman, shifts in watch_schedules.items():
            work = [False] * 48
            for (start_h, end_h) in shifts:
                for i in range(start_h * 2, end_h * 2):
                    work[i] = True
            
            all_days[watchman].append({
                'work_slots': work,
                'arrival_slots': [False] * 48,
                'departure_slots': [False] * 48,
                'port_op_slots': [False] * 48,
                'sluice_slots': [False] * 48,
                'shifting_slots': [False] * 48,
                'notes': []
            })
    
    # EXCEL
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
            sluice = data.get('sluice_slots', [False] * 48)
            shifting = data.get('shifting_slots', [False] * 48)

            hours = sum(work) / 2
            
            for col in range(48):
                cell = ws.cell(row=row, column=col+2)
                cell.border = thin_border
                
                if work[col]:
                    if sluice[col]:
                        cell.fill = PURPLE
                        cell.value = "SL"
                    elif shifting[col]:
                        cell.fill = PINK
                        cell.value = "SH"
                    elif arr[col]:
                        cell.fill = GREEN
                        cell.value = "T"
                    elif dep[col]:
                        cell.fill = BLUE
                        cell.value = "L"
                    elif sluice[col]:
                        cell.fill = PURPLE
                        cell.value = "SL"
                    elif shifting[col]:
                        cell.fill = PINK
                        cell.value = "SH"
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
            result = check_stcw_at_slot(combined, 95)
            report.append({'worker': w, 'analysis': result})
    
    return wb, all_days, report


if __name__ == "__main__":
    days_data = [
        {
            'arrival_hour': 18, 'arrival_minute': 0,
            'departure_hour': None, 'departure_minute': 0,
            'port_op_start_hour': 19, 'port_op_start_minute': 0,
            'port_op_end_hour': 0, 'port_op_end_minute': 0,
            'sluice_hour': None, 'sluice_minute': 0,
            'shifting_hour': None, 'shifting_minute': 0
        },
        {
            'arrival_hour': None, 'arrival_minute': 0,
            'departure_hour': 20, 'departure_minute': 0,
            'port_op_start_hour': 0, 'port_op_start_minute': 0,
            'port_op_end_hour': 19, 'port_op_end_minute': 0,
            'sluice_hour': None, 'sluice_minute': 0,
            'shifting_hour': None, 'shifting_minute': 0
        }
    ]
    
    wb, all_days, report = generate_schedule(days_data)
    
    print("=== Testi ===\n")
    
    for d in range(2):
        print(f"=== Päivä {d+1} ===")
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            work = all_days[w][d]['work_slots']
            ops = all_days[w][d]['port_op_slots']
            dep = all_days[w][d]['departure_slots']
            hours = sum(work) / 2
            op_hours = sum(ops) / 2
            has_dep = "L" if any(dep) else ""
            ranges = get_work_ranges(work)
            print(f"  {w}: {hours}h (op: {op_hours}h) {has_dep} | {' + '.join(ranges)}")
        print()
    
    print("=== STCW ===")
    for r in report:
        if 'Dayman' in r['worker']:
            ana = r['analysis']
            status = '✓' if ana['status'] == 'OK' else '⚠'
            print(f"  {r['worker']}: {ana['total_rest']}h lepo, {ana['rest_period_count']} jaksoa {status}")
            if ana['issues']:
                for issue in ana['issues']:
                    print(f"    -> {issue}")


def build_workbook_and_report(all_days, num_days, workers):
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

        ws.cell(row=1, column=1, value="Nimi")
        for col in range(48):
            h = col // 2
            m = "00" if col % 2 == 0 else "30"
            ws.cell(row=1, column=col + 2, value=f"{h:02d}:{m}")
            ws.cell(row=1, column=col + 2).alignment = Alignment(textRotation=90)
        ws.cell(row=1, column=50, value="Tunnit")

        row = 2
        for w in workers:
            ws.cell(row=row, column=1, value=w)

            data = all_days[w][d]
            work = data['work_slots']
            arr = data['arrival_slots']
            dep = data['departure_slots']
            ops = data['port_op_slots']
            sluice = data.get('sluice_slots', [False] * 48)
            shifting = data.get('shifting_slots', [False] * 48)

            hours = sum(work) / 2

            for col in range(48):
                cell = ws.cell(row=row, column=col + 2)
                cell.border = thin_border

                if work[col]:
                    if sluice[col]:
                        cell.fill = PURPLE
                        cell.value = "SL"
                    elif shifting[col]:
                        cell.fill = PINK
                        cell.value = "SH"
                    elif arr[col]:
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

        ws.column_dimensions['A'].width = 15
        for col in range(2, 50):
            ws.column_dimensions[get_column_letter(col)].width = 3

    report = []
    for w in workers:
        if len(all_days[w]) >= 2:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            combined = work1 + work2
            result = check_stcw_at_slot(combined, 95)
            report.append({'worker': w, 'analysis': result})

    return wb, report


def generate_schedule_with_manual_day1(days_data, manual_day1_work):
    """
    Generoi vuorot normaalisti, mutta korvaa päivän 1 työslotit manuaalisella taulukolla.
    manual_day1_work: dict {worker_name: [48 x bool]}
    """
    wb, all_days, _ = generate_schedule(days_data)

    workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
               'Watchman 1', 'Watchman 2', 'Watchman 3']

    for worker in workers:
        if worker not in all_days or not all_days[worker]:
            continue
        slots = manual_day1_work.get(worker, [False] * 48)
        slots = list(slots)[:48]
        if len(slots) < 48:
            slots.extend([False] * (48 - len(slots)))

        all_days[worker][0]['work_slots'] = [bool(x) for x in slots]
        all_days[worker][0]['arrival_slots'] = [False] * 48
        all_days[worker][0]['departure_slots'] = [False] * 48
        all_days[worker][0]['port_op_slots'] = [False] * 48
        all_days[worker][0]['sluice_slots'] = [False] * 48
        all_days[worker][0]['shifting_slots'] = [False] * 48

    wb, report = build_workbook_and_report(all_days, len(days_data), workers)
    return wb, all_days, report
