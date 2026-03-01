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


def count_rest_periods_in_day(work_slots, prev_day_work=None):
    """
    Laskee lepojaksojen määrän 24h ikkunassa päivän lopussa.
    Käytetään STCW-tarkistukseen generoinnin aikana.
    """
    if prev_day_work is None:
        prev_day_work = [False] * 48
    
    combined = prev_day_work + work_slots
    
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
        
        # Edellisen päivän työvuorot (STCW-tarkistukseen)
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
        
        # 1.1: Tulo - kaikki daymanit (1h)
        if arrival_start is not None:
            for dm in daymen:
                add_block(dm_work[dm], arrival_start, arrival_start + 2, dm_arr[dm])
        
        # 1.2: Lähtö - 2 daymaniä (1h)
        if departure_start is not None:
            # Valitse 2 daymaniä: priorisoi ne joilla vähiten tunteja
            scores = {}
            for dm in daymen:
                hours = sum(dm_work[dm]) / 2
                # Bonus jos jo töissä lähellä lähtöä
                continuity = 0
                for s in range(max(0, departure_start - 4), departure_start):
                    if dm_work[dm][s]:
                        continuity += 1
                scores[dm] = -hours + continuity * 0.5
            
            selected = sorted(daymen, key=lambda x: scores[x], reverse=True)[:2]
            for dm in selected:
                add_block(dm_work[dm], departure_start, departure_start + 2, dm_dep[dm])
        
        # 1.3: Slussi tulo - 1. tunti 2 dm, 2. tunti 3 dm (yhteensä 2h)
        if sluice_arr_start is not None:
            # Valitse 2 daymaniä ensimmäiselle tunnille
            scores = {}
            for dm in daymen:
                hours = sum(dm_work[dm]) / 2
                continuity = 1 if (sluice_arr_start > 0 and dm_work[dm][sluice_arr_start - 1]) else 0
                scores[dm] = -hours + continuity
            
            first_hour_dm = sorted(daymen, key=lambda x: scores[x], reverse=True)[:2]
            
            # 1. tunti: 2 daymaniä
            for dm in first_hour_dm:
                add_block(dm_work[dm], sluice_arr_start, sluice_arr_start + 2, dm_sluice[dm])
            
            # 2. tunti: kaikki 3 daymaniä
            for dm in daymen:
                add_block(dm_work[dm], sluice_arr_start + 2, sluice_arr_start + 4, dm_sluice[dm])
        
        # 1.4: Slussi lähtö - 2 daymaniä (2h)
        if sluice_dep_start is not None:
            scores = {}
            for dm in daymen:
                hours = sum(dm_work[dm]) / 2
                continuity = 1 if (sluice_dep_start > 0 and dm_work[dm][sluice_dep_start - 1]) else 0
                scores[dm] = -hours + continuity
            
            selected = sorted(daymen, key=lambda x: scores[x], reverse=True)[:2]
            for dm in selected:
                add_block(dm_work[dm], sluice_dep_start, sluice_dep_start + 4, dm_sluice[dm])
        
        # 1.5: Shiftaus - kaikki daymanit (1h)
        if shifting_start is not None:
            for dm in daymen:
                add_block(dm_work[dm], shifting_start, shifting_start + 2, dm_shifting[dm])
        
        # 1.6: Satamaop 08-17 ULKOPUOLELLA - aina 1 dayman töissä
        # Pilko pitkät blokit ja jaa daymanien kesken
        op_outside_slots = []
        for slot in range(op_start, min(op_end, 48)):
            if slot < NORMAL_START or slot >= NORMAL_END:
                if slot != LUNCH_START:
                    op_outside_slots.append(slot)
        
        if op_outside_slots:
            # Ryhmittele peräkkäiset slotit blokeiksi
            outside_blocks = []
            block_start = op_outside_slots[0]
            block_end = op_outside_slots[0] + 1
            
            for slot in op_outside_slots[1:]:
                if slot == block_end:
                    block_end = slot + 1
                else:
                    outside_blocks.append((block_start, block_end))
                    block_start = slot
                    block_end = slot + 1
            outside_blocks.append((block_start, block_end))
            
            # Pilko liian pitkät blokit (max 6h = 12 slottia per dayman)
            MAX_BLOCK_SLOTS = 12
            split_blocks = []
            
            for block_start, block_end in outside_blocks:
                block_len = block_end - block_start
                
                if block_len <= MAX_BLOCK_SLOTS:
                    split_blocks.append((block_start, block_end))
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
        
        # 3.2: Täytä loput tunnit (laajenna blokkeja)
        
        for dm in daymen:
            if needed_hours[dm] <= 0:
                continue
            
            current_hours = sum(dm_work[dm]) / 2
            target_hours = min(current_hours + needed_hours[dm], MAX_HOURS)
            
            # Etsi olemassa olevien blokkien alkukohdat
            blocks = get_work_blocks(dm_work[dm])
            
            if not blocks:
                # Ei blokkeja - luo uusi 08:00 alkaen
                slot = NORMAL_START
                while current_hours < target_hours and slot < NORMAL_END:
                    if not dm_work[dm][slot] and not (LUNCH_START <= slot < LUNCH_END):
                        dm_work[dm][slot] = True
                        if op_start <= slot < min(op_end, 48):
                            dm_ops[dm][slot] = True
                        current_hours = sum(dm_work[dm]) / 2
                    slot += 1
                continue
            
            # Etsi ensimmäisen blokin alkuslotti
            first_block_start = blocks[0][0]
            
            # Laajenna TAAKSEPÄIN ensimmäistä blokkia
            # Aloita heti blokin edestä ja mene kohti 08:00
            slot = first_block_start - 1
            
            while current_hours < target_hours and slot >= NORMAL_START:
                if LUNCH_START <= slot < LUNCH_END:
                    slot -= 1
                    continue
                
                if dm_work[dm][slot]:
                    slot -= 1
                    continue
                
                # Lisää slotti (ei STCW-tarkistusta koska laajennetaan yhtenäistä blokkia)
                dm_work[dm][slot] = True
                if op_start <= slot < min(op_end, 48):
                    dm_ops[dm][slot] = True
                current_hours = sum(dm_work[dm]) / 2
                slot -= 1
            
            # Jos vielä tarvitaan lisää, laajenna ETEENPÄIN viimeistä blokkia
            if current_hours < target_hours:
                last_block_end = blocks[-1][1]
                slot = last_block_end
                
                while current_hours < target_hours and slot < NORMAL_END:
                    if LUNCH_START <= slot < LUNCH_END:
                        slot += 1
                        continue
                    
                    if dm_work[dm][slot]:
                        slot += 1
                        continue
                    
                    dm_work[dm][slot] = True
                    if op_start <= slot < min(op_end, 48):
                        dm_ops[dm][slot] = True
                    current_hours = sum(dm_work[dm]) / 2
                    slot += 1
            
            # Jos VIELÄKIN tarvitaan lisää, täytä aukot blokkien välissä
            if current_hours < target_hours:
                for slot in range(NORMAL_START, NORMAL_END):
                    if current_hours >= target_hours:
                        break
                    if dm_work[dm][slot]:
                        continue
                    if LUNCH_START <= slot < LUNCH_END:
                        continue
                    
                    # Tarkista onko vierekkäin olemassa olevan työn kanssa
                    has_neighbor = False
                    if slot > 0 and dm_work[dm][slot - 1]:
                        has_neighbor = True
                    if slot < 47 and dm_work[dm][slot + 1]:
                        has_neighbor = True
                    
                    if has_neighbor:
                        dm_work[dm][slot] = True
                        if op_start <= slot < min(op_end, 48):
                            dm_ops[dm][slot] = True
                        current_hours = sum(dm_work[dm]) / 2
        
        # ====================================================================
        # VAIHE 4: TÄYTÄ AUKOT (max 1h aukot 08-17 välillä)
        # ====================================================================
        
        for dm in daymen:
            work = dm_work[dm]
            blocks = get_work_blocks(work)
            
            for i in range(len(blocks) - 1):
                _, block1_end = blocks[i]
                block2_start, _ = blocks[i + 1]
                
                gap = block2_start - block1_end
                
                # Täytä max 1h (2 slottia) aukot, paitsi lounas
                if 0 < gap <= 2:
                    for s in range(block1_end, block2_start):
                        if LUNCH_START <= s < LUNCH_END:
                            continue
                        work[s] = True
                        if op_start <= s < min(op_end, 48):
                            dm_ops[dm][s] = True
        
        # ====================================================================
        # TALLENNA TULOKSET
        # ====================================================================
        
        for dm in daymen:
            all_days[dm].append({
                'work_slots': dm_work[dm],
                'arrival_slots': dm_arr[dm],
                'departure_slots': dm_dep[dm],
                'port_op_slots': dm_ops[dm],
                'sluice_slots': dm_sluice[dm],
                'shifting_slots': dm_shifting[dm]
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
                'shifting_slots': [False] * 48
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
            ops = data.get('port_op_slots', [False] * 48)
            sluice = data.get('sluice_slots', [False] * 48)
            shifting = data.get('shifting_slots', [False] * 48)
            
            hours = sum(work) / 2
            
            for col in range(48):
                cell = ws.cell(row=row, column=col + 2)
                cell.border = thin_border
                
                if work[col]:
                    if arr[col]:
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
            'port_op_end_hour': 0, 'port_op_end_minute': 0
        },
        {
            'arrival_hour': None, 'arrival_minute': 0,
            'departure_hour': 20, 'departure_minute': 0,
            'port_op_start_hour': 0, 'port_op_start_minute': 0,
            'port_op_end_hour': 19, 'port_op_end_minute': 0
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
        ana = r['analysis']
        status = '✓' if ana['status'] == 'OK' else '⚠'
        print(f"\n{r['worker']}: {status}")
        print(f"  Lepoa: {ana['total_rest']}h | Pisin: {ana['longest_rest']}h | Jaksoja: {ana['rest_period_count']}")
        if ana['issues']:
            for issue in ana['issues']:
                print(f"  -> {issue}")
    
    # Tallenna Excel
    wb.save("tyovuorot_v16.xlsx")
    print(f"\nExcel tallennettu: tyovuorot_v16.xlsx")
