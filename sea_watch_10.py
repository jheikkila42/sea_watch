# -*- coding: utf-8 -*-
"""
STCW-yhteensopiva työvuorogeneraattori
Versio 16: Blokkipohjainen lähestymistapa

VAIHE 1: Pakolliset (tulo, lähtö, slussi, shiftaus, op 08-17 ulkopuolella)
VAIHE 2: Laske tarvittavat lisätunnit per dayman
VAIHE 3: Jaa työblokit liittyen pakollisiin
VAIHE 4: Validointi
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
NORMAL_START = 16   # 08:00 (slotti)
NORMAL_END = 34     # 17:00 (slotti)
LUNCH_START = 23    # 11:30
LUNCH_END = 24      # 12:00
MIN_HOURS = 8
MAX_HOURS = 10
MAX_OUTSIDE_NORMAL_SLOTS = 12  # 6h = 12 slottia
SHORT_SEGMENT_MAX_SLOTS = 2  # 1h tai lyhyempi pätkä pyritään poistamaan


def time_to_slot(h, m=0):
    """Muuntaa ajan slotiksi."""
    return h * 2 + (1 if m >= 30 else 0)


def slot_to_time_str(slot):
    """Muuntaa slotin ajaksi."""
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


# ============================================================================
# STCW-TARKISTUS
# ============================================================================

def check_stcw(work_day1, work_day2):
    """
    Tarkistaa STCW-säännöt kahden päivän välillä.
    
    Returns:
        dict: {ok, rest_periods, total_rest, longest_rest, issues}
    """
    combined = work_day1 + work_day2
    
    # Tarkista 24h ikkuna päivän 2 lopussa (slotti 95)
    slot_index = 95
    start = max(0, slot_index - 47)
    end = slot_index + 1
    window = combined[start:end]
    
    if len(window) < 48:
        padding = [False] * (48 - len(window))
        window = padding + window
    
    # Laske lepojaksot
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
    rest_count = len(rest_periods)
    
    issues = []
    if total_rest < 10:
        issues.append(f"Lepoa vain {total_rest}h (min 10h)")
    if rest_count > 2:
        issues.append(f"Lepo {rest_count} osassa (max 2)")
    if longest_rest < 6:
        issues.append(f"Pisin lepo {longest_rest}h (min 6h)")
    
    return {
        'ok': len(issues) == 0,
        'rest_periods': rest_count,
        'total_rest': total_rest,
        'longest_rest': longest_rest,
        'issues': issues
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
    Laskee lepojaksojen määrän 24h ikkunassa päivän lopussa.
    Käytetään STCW-tarkistukseen generoinnin aikana.
    """
    if prev_day_work is None:
        prev_day_work = [False] * 48
    
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
    
    # Tarkista 24h ikkuna NYKYISEN päivän lopussa (slotti 47 + 48 = 95)
    # Mutta koska generoimme päivää kerrallaan, tarkistetaan slotti 47 (päivän loppu)
    check_slot = 47 + 48  # Päivän 1 loppu combined-listassa
    
    start = max(0, check_slot - 47)
    end = check_slot + 1
    window = combined[start:end]
    
    if len(window) < 48:
        padding = [False] * (48 - len(window))
        window = padding + window
    
    rest_periods = 0
    current_rest = 0
    
    for is_work in window:
        if not is_work:
            current_rest += 1
        else:
            if current_rest >= 2:  # Min 1h lepo
                rest_periods += 1
            current_rest = 0
    
    if current_rest >= 2:
        rest_periods += 1
    
    return rest_periods


# ============================================================================
# APUFUNKTIOT
# ============================================================================

def get_work_blocks(work_slots):
    """Palauttaa työblokit listana (start_slot, end_slot)."""
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


def blocks_would_merge(work_slots, new_start, new_end):
    """Tarkistaa liittyykö uusi blokki olemassa oleviin."""
    # Tarkista onko vierekkäin tai päällekkäin
    for i in range(max(0, new_start - 1), min(48, new_end + 1)):
        if work_slots[i]:
            return True
    return False


def add_block(work_slots, start, end, marker_slots=None):
    """Lisää työblokin."""
    for i in range(max(0, start), min(end, 48)):
        work_slots[i] = True
        if marker_slots is not None:
            marker_slots[i] = True


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
    """
    Generoi työvuorot blokkipohjaisella lähestymistavalla.
    """
    workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
               'Watchman 1', 'Watchman 2', 'Watchman 3']
    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
    
    all_days = {w: [] for w in workers}
    num_days = len(days_data)
    
    for d, info in enumerate(days_data):
        
        # Parsitaan ajat
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
        
        sluice_arr_h = info.get('sluice_arrival_hour')
        sluice_arr_m = info.get('sluice_arrival_minute', 0)
        sluice_dep_h = info.get('sluice_departure_hour')
        sluice_dep_m = info.get('sluice_departure_minute', 0)
        shifting_h = info.get('shifting_hour')
        shifting_m = info.get('shifting_minute', 0)
        
        # Muunna sloiteiksi
        if op_start_h is not None:
            op_start = time_to_slot(op_start_h, op_start_m)
            if op_end_h is not None and op_end_h < op_start_h:
                op_end = 48  # Yli keskiyön
            elif op_end_h == 0 and op_start_h > 0:
                op_end = 48
            elif op_end_h is not None:
                op_end = time_to_slot(op_end_h, op_end_m)
            else:
                op_end = NORMAL_END
        else:
            op_start = NORMAL_START
            op_end = NORMAL_END
        
        arrival_start = time_to_slot(arrival_h, arrival_m) if arrival_h is not None else None
        departure_start = time_to_slot(departure_h, departure_m) if departure_h is not None else None
        sluice_arr_start = time_to_slot(sluice_arr_h, sluice_arr_m) if sluice_arr_h is not None else None
        sluice_dep_start = time_to_slot(sluice_dep_h, sluice_dep_m) if sluice_dep_h is not None else None
        shifting_start = time_to_slot(shifting_h, shifting_m) if shifting_h is not None else None
        
        departure_start = time_to_index(departure_h, departure_m) if departure_h is not None else None
        departure_end = departure_start + 2 if departure_start is not None else None

        sluice_arrival_start = time_to_index(sluice_arr_h, sluice_arr_m) if sluice_arr_h is not None else None
        sluice_departure_start = time_to_index(sluice_dep_h, sluice_dep_m) if sluice_dep_h is not None else None
        shifting_start = time_to_index(shifting_h, shifting_m) if shifting_h is not None else None
        
        # Edellisen päivän työvuorot
        prev_day_work = {}
        for dm in daymen:
            if d > 0:
                prev_day_work[dm] = all_days[dm][d - 1]['work_slots']
            else:
                prev_day_work[dm] = [False] * 48
        
        # Alusta työntekijöiden data
        dm_work = {dm: [False] * 48 for dm in daymen}
        dm_arr = {dm: [False] * 48 for dm in daymen}
        dm_dep = {dm: [False] * 48 for dm in daymen}
        dm_ops = {dm: [False] * 48 for dm in daymen}
        dm_sluice = {dm: [False] * 48 for dm in daymen}
        dm_shifting = {dm: [False] * 48 for dm in daymen}
        
        # ====================================================================
        # VAIHE 1: PAKOLLISET
        # ====================================================================
        
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
        
        # 1.2: Lähtö - 2 daymaniä (1h)
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

            if best_score > -1000:
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
                                if (sum(all_dayman_work[dayman]) + 1) / 2 > MAX_HOURS:
                                    continue
                                all_dayman_work[dayman][s] = True
                                if op_start <= s < min(op_end, 48):
                                    all_dayman_ops[dayman][s] = True
                else:
                    # Pilko osiin
                    pos = block_start
                    while pos < block_end:
                        chunk_end = min(pos + MAX_BLOCK_SLOTS, block_end)
                        split_blocks.append((pos, chunk_end))
                        pos = chunk_end
            
            # Jaa blokit daymanien kesken vuorotellen
            dm_index = 0
            
            for block_start, block_end in split_blocks:
                block_hours = (block_end - block_start) / 2
                
                # Etsi paras dayman tälle blokille
                best_dm = None
                best_score = -9999
                
                # Kokeile jokaista daymaniä alkaen vuorottelujärjestyksestä
                for i in range(len(daymen)):
                    dm = daymen[(dm_index + i) % len(daymen)]
                    
                    current_hours = sum(dm_work[dm]) / 2
                    outside_hours = sum(1 for s in range(48) 
                                       if dm_work[dm][s] and (s < NORMAL_START or s >= NORMAL_END)) / 2
                    
                    # Älä ylitä max tunteja
                    if current_hours + block_hours > MAX_HOURS:
                        continue
                    
                    # Älä ylitä 6h yötyötä
                    if outside_hours + block_hours > 6:
                        continue
                    
                    # STCW-tarkistus: simuloi blokin lisäys
                    test_work = dm_work[dm][:]
                    for s in range(block_start, block_end):
                        test_work[s] = True
                    rest_periods = count_rest_periods_in_day(test_work, prev_day_work[dm])
                    
                    if rest_periods > 2:
                        continue  # STCW-rike, ohita
                    
                    score = 0
                    
                    # Jatkuvuusbonus
                    if block_start > 0 and dm_work[dm][block_start - 1]:
                        score += 100
                    if block_end < 48 and dm_work[dm][block_end]:
                        score += 100
                    
                    # STCW-bonus: vähemmän lepojaksoja on parempi
                    score += (3 - rest_periods) * 50
                    
                    # Tasapainobonus
                    score += (MAX_HOURS - current_hours) * 10
                    
                    # Bonus jos tämä on vuorossa oleva dayman
                    if i == 0:
                        score += 50
                    
                    if score > best_score:
                        best_score = score
                        best_dm = dm
                
                if best_dm:
                    add_block(dm_work[best_dm], block_start, block_end, dm_ops[best_dm])
                    dm_index = (daymen.index(best_dm) + 1) % len(daymen)
        
        # ====================================================================
        # VAIHE 2: LASKE TARVITTAVAT LISÄTUNNIT
        # ====================================================================
        
        needed_hours = {}
        for dm in daymen:
            current = sum(dm_work[dm]) / 2
            needed_hours[dm] = max(0, MIN_HOURS - current)
        
        # ====================================================================
        # VAIHE 3: JAA TYÖBLOKIT
        # Strategia: 
        # 1. Ensin varmista op-kattavuus 08-17 välillä
        # 2. Sitten laajenna blokkeja tarvittaessa
        # ====================================================================
        
        # 3.1: Op-kattavuus 08-17 välillä - varmista että joku on aina töissä
        op_inside_slots = []
        for slot in range(max(op_start, NORMAL_START), min(op_end, NORMAL_END)):
            if LUNCH_START <= slot < LUNCH_END:
                continue
            op_inside_slots.append(slot)
        
        for slot in op_inside_slots:
            # Onko joku jo töissä tässä slotissa?
            workers_in_slot = [dm for dm in daymen if dm_work[dm][slot]]
            
            if len(workers_in_slot) >= 1:
                # Joku jo töissä, merkitään op-slotiksi
                for dm in workers_in_slot:
                    dm_ops[dm][slot] = True
                continue
            
            # Kukaan ei töissä - valitse paras dayman
            best_dm = None
            best_score = -9999
            
            for dm in daymen:
                current_hours = sum(dm_work[dm]) / 2
                
                # Älä ylitä max tunteja
                if current_hours >= MAX_HOURS:
                    continue
                
                score = 0
                
                # Jatkuvuusbonus
                if slot > 0 and dm_work[dm][slot - 1]:
                    score += 200
                if slot < 47 and dm_work[dm][slot + 1]:
                    score += 200
                
                # STCW-tarkistus
                test_work = dm_work[dm][:]
                test_work[slot] = True
                rest_periods = count_rest_periods_in_day(test_work, prev_day_work[dm])
                
                if rest_periods > 2:
                    continue
                
                # Tasapainobonus
                score += (MAX_HOURS - current_hours) * 10
                
                if score > best_score:
                    best_score = score
                    best_dm = dm
            
            if best_dm:
                dm_work[best_dm][slot] = True
                dm_ops[best_dm][slot] = True
        
        # VAIHE 6: Varmista minimi 8h
        for dayman in daymen:
            current_hours = sum(all_dayman_work[dayman]) / 2

            while current_hours < MIN_HOURS:
                best_slot = None
                best_score = -99999

        # Etsi vain 08-17 väliltä
                for slot in range(NORMAL_START, NORMAL_END):
                    if current_hours >= target_hours:
                        break
                    if dm_work[dm][slot]:
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
                                if (sum(all_dayman_work[dayman]) + 1) / 2 > MAX_HOURS:
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
        
        # ====================================================================
        # BOSUN (yksinkertainen: 08-17 + tulo/lähtö)
        # ====================================================================
        
        bosun_work = [False] * 48
        bosun_arr = [False] * 48
        bosun_dep = [False] * 48
        
        if arrival_start is not None:
            add_block(bosun_work, arrival_start, arrival_start + 2, bosun_arr)
        if departure_start is not None:
            add_block(bosun_work, departure_start, departure_start + 2, bosun_dep)
        
        for slot in range(NORMAL_START, NORMAL_END):
            if slot != LUNCH_START:
                bosun_work[slot] = True
        
        all_days['Bosun'].append({
            'work_slots': bosun_work,
            'arrival_slots': bosun_arr,
            'departure_slots': bosun_dep,
            'port_op_slots': [False] * 48,
            'sluice_slots': [False] * 48,
            'shifting_slots': [False] * 48
        })
        
        # ====================================================================
        # WATCHMANIT (4h vuorot)
        # ====================================================================
        
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
    
    # ========================================================================
    # LUO EXCEL JA RAPORTTI
    # ========================================================================
    
    wb = Workbook()
    wb.remove(wb.active)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for d in range(num_days):
        ws = wb.create_sheet(title=f"Päivä {d + 1}")
        
        # Otsikkorivi
        ws.cell(row=1, column=1, value="Työntekijä")
        for col in range(48):
            cell = ws.cell(row=1, column=col + 2, value=slot_to_time_str(col))
            cell.alignment = Alignment(horizontal='center', textRotation=90)
            cell.font = Font(size=8)
            cell.border = thin_border
        ws.cell(row=1, column=50, value="Tunnit")
        
        # Työntekijärivit
        for row, worker in enumerate(workers, start=2):
            ws.cell(row=row, column=1, value=worker)
            
            data = all_days[worker][d]
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
        
        # Sarakkeiden leveys
        ws.column_dimensions['A'].width = 12
        for col in range(2, 50):
            ws.column_dimensions[get_column_letter(col)].width = 3
        ws.column_dimensions[get_column_letter(50)].width = 6
    
    # STCW-raportti
    report = []
    for w in workers:
        if 'Dayman' in w and len(all_days[w]) >= 2:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            stcw = check_stcw(work1, work2)
            report.append({
                'worker': w,
                'analysis': {
                    'status': 'OK' if stcw['ok'] else 'RIKE',
                    'issues': stcw['issues'],
                    'total_rest': stcw['total_rest'],
                    'longest_rest': stcw['longest_rest'],
                    'rest_period_count': stcw['rest_periods']
                }
            })
    
    return wb, all_days, report


# ============================================================================
# TESTAUS
# ============================================================================

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
    
    print("Generoidaan työvuorot (blokkipohjainen)...")
    wb, all_days, report = generate_schedule(days_data)
    
    print("\n" + "=" * 60)
    print("TULOKSET")
    print("=" * 60)
    
    for d in range(len(days_data)):
        print(f"\n=== Päivä {d + 1} ===")
        info = days_data[d]
        arr = f"{info.get('arrival_hour', '-')}:00" if info.get('arrival_hour') else "-"
        dep = f"{info.get('departure_hour', '-')}:00" if info.get('departure_hour') else "-"
        print(f"  Tulo: {arr} | Lähtö: {dep}")
        print()
        
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            work = all_days[w][d]['work_slots']
            hours = sum(work) / 2
            ranges = get_work_ranges(work)
            print(f"  {w}: {hours}h | {' + '.join(ranges)}")
    
    print("\n" + "=" * 60)
    print("STCW-TARKISTUS")
    print("=" * 60)
    
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
