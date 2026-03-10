# -*- coding: utf-8 -*-
"""
Schedule Analyzer - Työvuorojen analysointi

Analysoi generoituja työvuoroja ja tunnistaa:
1. STCW-rikkeet
2. Tuntitasapaino-ongelmat
3. Op-kattavuusaukot
4. Turhat tauot
5. Liian pitkät vuorot
"""

from typing import Dict, List, Any
from sea_watch_17 import (
    check_stcw_at_slot,
    get_work_ranges,
    slot_to_time_str,
    NORMAL_START,
    NORMAL_END,
    LUNCH_START,
    LUNCH_END,
    MIN_HOURS,
    MAX_HOURS,
)


def analyze_worker_day(worker: str, day_idx: int, day_data: Dict, 
                       prev_day_data: Dict = None) -> Dict[str, Any]:
    """
    Analysoi yhden työntekijän yhden päivän vuorot.
    
    Returns:
        Dict sisältäen:
        - hours: työtunnit
        - work_ranges: työjaksot tekstinä
        - issues: lista ongelmista
        - warnings: lista varoituksista
    """
    work = day_data['work_slots']
    hours = sum(work) / 2
    ranges = get_work_ranges(work)
    
    issues = []
    warnings = []
    
    # 1. Tarkista tuntimäärä
    if 'Dayman' in worker:
        if hours < MIN_HOURS:
            issues.append(f"Liian vähän tunteja: {hours}h (min {MIN_HOURS}h)")
        elif hours > MAX_HOURS:
            issues.append(f"Liian paljon tunteja: {hours}h (max {MAX_HOURS}h)")
    
    # 2. Tarkista STCW (jos edellinen päivä saatavilla)
    if prev_day_data is not None:
        prev_work = prev_day_data['work_slots']
        combined = prev_work + work
        stcw_result = check_stcw_at_slot(combined, len(combined) - 1)

        if stcw_result['status'] != "OK":
            issues.append(
                f"STCW: {stcw_result['status']} (total_rest={stcw_result['total_rest']}h, "
                f"longest_rest={stcw_result['longest_rest']}h)"
            )
    
    # 3. Tarkista turhat tauot (yli 1h aukot työjaksojen välissä)
    gaps = find_work_gaps(work)
    for gap_start, gap_end, gap_hours in gaps:
        if gap_hours > 1 and gap_start >= NORMAL_START and gap_end <= NORMAL_END:
            # Ohita lounastauko
            if not (gap_start == LUNCH_START and gap_end == LUNCH_END):
                warnings.append(
                    f"Tauko {slot_to_time_str(gap_start)}-{slot_to_time_str(gap_end)} ({gap_hours}h)"
                )
    
    # 4. Tarkista yövuoron pituus
    night_hours = count_night_hours(work)
    if night_hours > 8:
        warnings.append(f"Pitkä yövuoro: {night_hours}h")
    
    return {
        'worker': worker,
        'day': day_idx + 1,
        'hours': hours,
        'work_ranges': ranges,
        'issues': issues,
        'warnings': warnings,
        'has_problems': len(issues) > 0,
        'has_warnings': len(warnings) > 0
    }


def find_work_gaps(work_slots: List[bool]) -> List[tuple]:
    """
    Etsii aukot työjaksojen välissä.
    
    Returns:
        Lista tupleista: (gap_start_slot, gap_end_slot, gap_hours)
    """
    gaps = []
    in_work = False
    work_end = None
    
    for i, w in enumerate(work_slots):
        if w:
            if work_end is not None and not in_work:
                # Löytyi aukko
                gap_hours = (i - work_end) / 2
                if gap_hours >= 1:
                    gaps.append((work_end, i, gap_hours))
            in_work = True
            work_end = i + 1
        else:
            in_work = False
    
    return gaps


def count_night_hours(work_slots: List[bool]) -> float:
    """Laskee yötyötunnit (00:00-08:00 ja 17:00-24:00)."""
    night_slots = 0
    for i, w in enumerate(work_slots):
        if w and (i < NORMAL_START or i >= NORMAL_END):
            night_slots += 1
    return night_slots / 2


def analyze_op_coverage(all_days: Dict, day_idx: int, 
                        op_start: int, op_end: int) -> Dict[str, Any]:
    """
    Analysoi operaation kattavuus - onko joku aina töissä op-aikana.
    
    Returns:
        Dict sisältäen:
        - coverage_percent: kattavuusprosentti
        - gaps: lista kattamattomista ajoista
        - issues: lista ongelmista
    """
    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
    
    uncovered_slots = []
    total_op_slots = 0
    covered_slots = 0
    
    for slot in range(max(0, op_start), min(op_end, 48)):
        if LUNCH_START <= slot < LUNCH_END:
            continue
        
        total_op_slots += 1
        
        # Onko joku daymaneista töissä?
        workers_in_slot = [
            dm for dm in daymen 
            if all_days[dm][day_idx]['work_slots'][slot]
        ]
        
        if len(workers_in_slot) >= 1:
            covered_slots += 1
        else:
            uncovered_slots.append(slot)
    
    # Ryhmittele peräkkäiset aukot
    gaps = []
    if uncovered_slots:
        gap_start = uncovered_slots[0]
        gap_end = uncovered_slots[0] + 1
        
        for slot in uncovered_slots[1:]:
            if slot == gap_end:
                gap_end = slot + 1
            else:
                gaps.append((gap_start, gap_end))
                gap_start = slot
                gap_end = slot + 1
        gaps.append((gap_start, gap_end))
    
    coverage_percent = (covered_slots / total_op_slots * 100) if total_op_slots > 0 else 100
    
    issues = []
    for gap_start, gap_end in gaps:
        gap_hours = (gap_end - gap_start) / 2
        issues.append(
            f"Op-aukko {slot_to_time_str(gap_start)}-{slot_to_time_str(gap_end)} ({gap_hours}h)"
        )
    
    return {
        'day': day_idx + 1,
        'coverage_percent': round(coverage_percent, 1),
        'gaps': gaps,
        'issues': issues,
        'has_problems': len(gaps) > 0
    }


def analyze_hour_balance(all_days: Dict, day_idx: int) -> Dict[str, Any]:
    """
    Analysoi tuntitasapaino daymanien välillä.
    
    Returns:
        Dict sisältäen:
        - hours_by_worker: tunnit per työntekijä
        - min_hours, max_hours, diff: tilastot
        - issues: lista ongelmista
    """
    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
    
    hours_by_worker = {}
    for dm in daymen:
        work = all_days[dm][day_idx]['work_slots']
        hours_by_worker[dm] = sum(work) / 2
    
    hours_list = list(hours_by_worker.values())
    min_h = min(hours_list)
    max_h = max(hours_list)
    diff = max_h - min_h
    
    issues = []
    warnings = []
    
    if diff > 2:
        min_worker = [w for w, h in hours_by_worker.items() if h == min_h][0]
        max_worker = [w for w, h in hours_by_worker.items() if h == max_h][0]
        warnings.append(
            f"Epätasainen jako: {max_worker.split()[-1]} {max_h}h vs {min_worker.split()[-1]} {min_h}h (ero {diff}h)"
        )
    
    return {
        'day': day_idx + 1,
        'hours_by_worker': hours_by_worker,
        'min_hours': min_h,
        'max_hours': max_h,
        'diff': diff,
        'issues': issues,
        'warnings': warnings,
        'has_problems': len(issues) > 0,
        'has_warnings': len(warnings) > 0
    }


def analyze_schedule(all_days: Dict, days_data: List[Dict]) -> Dict[str, Any]:
    """
    Analysoi koko työvuorolistan.
    
    Args:
        all_days: Generaattorin palauttama all_days-rakenne
        days_data: Alkuperäinen days_data syöte
    
    Returns:
        Kattava analyysi kaikista päivistä
    """
    num_days = len(days_data)
    daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
    
    analysis = {
        'num_days': num_days,
        'worker_analyses': [],
        'op_coverage_analyses': [],
        'hour_balance_analyses': [],
        'summary': {
            'total_issues': 0,
            'total_warnings': 0,
            'stcw_violations': 0,
            'op_coverage_gaps': 0,
            'hour_imbalances': 0
        }
    }
    
    # Analysoi jokainen työntekijä ja päivä
    for d in range(num_days):
        info = days_data[d]
        
        # Op-ajat
        op_start_h = info.get('port_op_start_hour')
        op_end_h = info.get('port_op_end_hour')
        
        if op_start_h is not None:
            op_start = op_start_h * 2
            if op_end_h is not None and op_end_h < op_start_h:
                op_end = 48
            elif op_end_h == 0 and op_start_h > 0:
                op_end = 48
            elif op_end_h is not None:
                op_end = op_end_h * 2
            else:
                op_end = NORMAL_END
        else:
            op_start = NORMAL_START
            op_end = NORMAL_END
        
        # Työntekijäanalyysit
        for dm in daymen:
            day_data = all_days[dm][d]
            prev_day_data = all_days[dm][d - 1] if d > 0 else None
            
            worker_analysis = analyze_worker_day(dm, d, day_data, prev_day_data)
            analysis['worker_analyses'].append(worker_analysis)
            
            # Päivitä yhteenveto
            analysis['summary']['total_issues'] += len(worker_analysis['issues'])
            analysis['summary']['total_warnings'] += len(worker_analysis['warnings'])
            
            for issue in worker_analysis['issues']:
                if 'STCW' in issue:
                    analysis['summary']['stcw_violations'] += 1
        
        # Op-kattavuusanalyysi
        op_analysis = analyze_op_coverage(all_days, d, op_start, op_end)
        analysis['op_coverage_analyses'].append(op_analysis)
        analysis['summary']['op_coverage_gaps'] += len(op_analysis['gaps'])
        
        # Tuntitasapainoanalyysi
        balance_analysis = analyze_hour_balance(all_days, d)
        analysis['hour_balance_analyses'].append(balance_analysis)
        if balance_analysis['has_warnings']:
            analysis['summary']['hour_imbalances'] += 1
    
    return analysis


def format_analysis_report(analysis: Dict) -> str:
    """
    Muotoilee analyysin luettavaksi raportiksi.
    """
    lines = []
    lines.append("=" * 60)
    lines.append("TYÖVUOROANALYYSI")
    lines.append("=" * 60)
    
    summary = analysis['summary']
    
    # Yhteenveto
    lines.append("\n📊 YHTEENVETO")
    lines.append("-" * 40)
    
    if summary['total_issues'] == 0 and summary['total_warnings'] == 0:
        lines.append("✅ Ei ongelmia havaittu!")
    else:
        if summary['stcw_violations'] > 0:
            lines.append(f"❌ STCW-rikkeitä: {summary['stcw_violations']}")
        if summary['op_coverage_gaps'] > 0:
            lines.append(f"⚠️ Op-kattavuusaukkoja: {summary['op_coverage_gaps']}")
        if summary['hour_imbalances'] > 0:
            lines.append(f"⚠️ Tuntitasapaino-ongelmia: {summary['hour_imbalances']}")
        lines.append(f"\nYhteensä: {summary['total_issues']} ongelmaa, {summary['total_warnings']} varoitusta")
    
    # Yksityiskohtaiset ongelmat
    has_details = False
    
    for wa in analysis['worker_analyses']:
        if wa['issues'] or wa['warnings']:
            if not has_details:
                lines.append("\n\n📋 YKSITYISKOHDAT")
                lines.append("-" * 40)
                has_details = True
            
            lines.append(f"\n{wa['worker']} - Päivä {wa['day']} ({wa['hours']}h)")
            for issue in wa['issues']:
                lines.append(f"  ❌ {issue}")
            for warning in wa['warnings']:
                lines.append(f"  ⚠️ {warning}")
    
    for op in analysis['op_coverage_analyses']:
        if op['issues']:
            if not has_details:
                lines.append("\n\n📋 YKSITYISKOHDAT")
                lines.append("-" * 40)
                has_details = True
            
            lines.append(f"\nOp-kattavuus - Päivä {op['day']} ({op['coverage_percent']}%)")
            for issue in op['issues']:
                lines.append(f"  ❌ {issue}")
    
    for bal in analysis['hour_balance_analyses']:
        if bal['warnings']:
            if not has_details:
                lines.append("\n\n📋 YKSITYISKOHDAT")
                lines.append("-" * 40)
                has_details = True
            
            lines.append(f"\nTuntitasapaino - Päivä {bal['day']}")
            for warning in bal['warnings']:
                lines.append(f"  ⚠️ {warning}")
    
    lines.append("\n" + "=" * 60)
    
    return "\n".join(lines)


def get_analysis_for_llm(analysis: Dict) -> Dict[str, Any]:
    """
    Palauttaa analyysin LLM:lle sopivassa muodossa.
    Sisältää vain oleelliset tiedot korjausehdotuksia varten.
    """
    problems = []
    
    for wa in analysis['worker_analyses']:
        if wa['issues']:
            problems.append({
                'type': 'worker_issue',
                'worker': wa['worker'],
                'day': wa['day'],
                'hours': wa['hours'],
                'issues': wa['issues'],
                'work_ranges': wa['work_ranges']
            })
    
    for op in analysis['op_coverage_analyses']:
        if op['issues']:
            problems.append({
                'type': 'op_coverage',
                'day': op['day'],
                'coverage_percent': op['coverage_percent'],
                'issues': op['issues']
            })
    
    for bal in analysis['hour_balance_analyses']:
        if bal['warnings']:
            problems.append({
                'type': 'hour_balance',
                'day': bal['day'],
                'hours_by_worker': bal['hours_by_worker'],
                'diff': bal['diff'],
                'warnings': bal['warnings']
            })
    
    return {
        'num_days': analysis['num_days'],
        'summary': analysis['summary'],
        'problems': problems,
        'has_problems': len(problems) > 0
    }


# ============================================================================
# TESTAUS
# ============================================================================

if __name__ == "__main__":
    from sea_watch_17 import generate_schedule
    
    # Testitapaus
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
    
    print("Generoidaan työvuorot...")
    wb, all_days, _ = generate_schedule(days_data)
    
    print("Analysoidaan...")
    analysis = analyze_schedule(all_days, days_data)
    
    print(format_analysis_report(analysis))
    
    print("\n\nLLM-muoto:")
    llm_data = get_analysis_for_llm(analysis)
    import json
    print(json.dumps(llm_data, indent=2, ensure_ascii=False))
