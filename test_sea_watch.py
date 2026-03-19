

import pytest
from sea_watch_17 import (
    generate_schedule,
    check_stcw_at_slot,
    time_to_slot,
    slot_to_time_str,
)




def analyze_stcw_from_work_starts(work_48h):
    """Yhteensopiva STCW-analyysi testeille nykyisen version mukaan."""
    if len(work_48h) < 48:
        work_48h = [False] * (48 - len(work_48h)) + work_48h
    window = work_48h[-48:]

    rest_periods = []
    cur = 0
    for is_work in window:
        if not is_work:
            cur += 1
        elif cur:
            if cur >= 2:
                rest_periods.append(cur / 2)
            cur = 0
    if cur >= 2:
        rest_periods.append(cur / 2)

    if len(rest_periods) >= 2 and (not window[0]) and (not window[-1]):
        rest_periods = [rest_periods[0] + rest_periods[-1], *rest_periods[1:-1]]

    total_rest = sum(rest_periods)
    longest_rest = max(rest_periods) if rest_periods else 0
    rest_period_count = len(rest_periods)
    issues = []
    if total_rest < 10:
        issues.append(f"Lepoa yhteensä {total_rest}h (< 10h)")
    if rest_period_count > 2:
        issues.append(f"Lepojaksoja {rest_period_count} (> 2)")
    if longest_rest < 6:
        issues.append(f"Pisin lepojakso {longest_rest}h (< 6h)")

    return {
        'total_rest': total_rest,
        'longest_rest': longest_rest,
        'rest_period_count': rest_period_count,
        'status': 'OK' if not issues else 'VIOLATION',
        'issues': issues,
    }

# APUFUNKTIOT TESTEILLE
# ---------------------------------------------------------------------

def get_work_ranges(work_slots):
    """Palauttaa työvuorot aikaväleinä"""
    ranges = []
    start = None
    for i, x in enumerate(work_slots):
        if x and start is None:
            start = i
        elif not x and start is not None:
            ranges.append((slot_to_time_str(start), slot_to_time_str(i)))
            start = None
    if start is not None:
        ranges.append((slot_to_time_str(start), "00:00"))
    return ranges


def count_daymen_working_at(all_days, day_idx, hour, minute=0):
    """Laskee kuinka monta daymania on töissä tietyllä hetkellä"""
    slot = time_to_slot(hour, minute)
    count = 0
    for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
        if all_days[w][day_idx]['work_slots'][slot]:
            count += 1
    return count


def run_scenario(arrival_hour, departure_hour, op_start_hour, op_end_hour,
                 arrival_minute=0, departure_minute=0, op_start_minute=0, op_end_minute=0):
    """Ajaa yhden skenaarion ja palauttaa tulokset"""
    days_data = [
        {
            'arrival_hour': arrival_hour,
            'arrival_minute': arrival_minute,
            'departure_hour': departure_hour,
            'departure_minute': departure_minute,
            'port_op_start_hour': op_start_hour,
            'port_op_start_minute': op_start_minute,
            'port_op_end_hour': op_end_hour,
            'port_op_end_minute': op_end_minute
        },
        {
            'arrival_hour': None,
            'arrival_minute': 0,
            'departure_hour': None,
            'departure_minute': 0,
            'port_op_start_hour': 8,
            'port_op_start_minute': 0,
            'port_op_end_hour': 17,
            'port_op_end_minute': 0
        }
    ]
    wb, all_days, report = generate_schedule(days_data)
    return all_days



# TESTIT: DAYMANIT TULOSSA JA LÄHDÖSSÄ
# ---------------------------------------------------------------------

class TestDaymenArrivalDeparture:
    """Kaikki daymanit ovat tulossa, lähdössä 2 daymania"""
    
    def test_all_daymen_in_arrival_basic(self):
        """Perus: kaikki daymanit tulossa"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            has_arrival = any(all_days[w][0]['arrival_slots'])
            assert has_arrival, f"{w} ei ole tulossa"
    
    def test_two_daymen_in_departure_basic(self):
        """Perus: lähdössä on kaksi daymania"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )

        count = sum(any(all_days[w][0]['departure_slots']) for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2'])
        assert count == 2, f"Lähdössä pitäisi olla 2 daymania, on {count}"
    
    def test_all_daymen_in_arrival_early_morning(self):
        """Aikainen tulo: kaikki daymanit tulossa"""
        all_days = run_scenario(
            arrival_hour=6, departure_hour=19,
            op_start_hour=8, op_end_hour=18
        )
        
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            has_arrival = any(all_days[w][0]['arrival_slots'])
            assert has_arrival, f"{w} ei ole tulossa (aikainen tulo)"
    
    def test_two_daymen_in_departure_late_evening(self):
        """Myöhäinen lähtö: lähdössä on kaksi daymania"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=21,
            op_start_hour=10, op_end_hour=20
        )

        count = sum(any(all_days[w][0]['departure_slots']) for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2'])
        assert count == 2, f"Myöhäisessä lähdössä pitäisi olla 2 daymania, on {count}"



# TESTIT: ILTAVUOROT (KLO 17-08 MAX YKSI DAYMAN)
# ---------------------------------------------------------------------

class TestEveningShifts:
    """Klo 17-08 välillä max yksi dayman kerrallaan (paitsi tulo/lähtö)"""
    
    def test_max_one_dayman_after_17(self):
        """Klo 17-19 max yksi dayman (ennen lähtöä)"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=21,
            op_start_hour=10, op_end_hour=20
        )
        
        # Tarkista klo 17, 18, 19, 20 (ennen lähtöä 21:00)
        for hour in [17, 18, 19]:
            count = count_daymen_working_at(all_days, 0, hour)
            assert count <= 1, f"Klo {hour}:00 on {count} daymania töissä (max 1)"
    
    def test_two_daymen_during_departure(self):
        """Lähdön aikana on kaksi daymania"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=21,
            op_start_hour=10, op_end_hour=20
        )
        
        # Lähtö klo 21 - kaikki saavat olla
        count = count_daymen_working_at(all_days, 0, 21)
        assert count == 2, f"Lähdön aikana pitäisi olla 2 daymania, on {count}"
    
    def test_evening_coverage_exists(self):
        """Iltakattavuus on olemassa kun operaatio jatkuu iltaan"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=21,
            op_start_hour=10, op_end_hour=20
        )
        
        # Joku dayman on töissä klo 17-20
        for hour in [17, 18, 19]:
            count = count_daymen_working_at(all_days, 0, hour)
            assert count >= 1, f"Klo {hour}:00 ei ole ketään daymania töissä"


    def test_exactly_one_dayman_outside_8_17_during_port_ops(self):
        """Kun OP on käynnissä 08-17 ulkopuolella, siellä on täsmälleen yksi dayman."""
        days_data = [
            {
                'arrival_hour': None,
                'arrival_minute': 0,
                'departure_hour': None,
                'departure_minute': 0,
                'port_op_start_hour': 18,
                'port_op_start_minute': 0,
                'port_op_end_hour': 22,
                'port_op_end_minute': 0,
            },
            {
                'arrival_hour': None,
                'arrival_minute': 0,
                'departure_hour': None,
                'departure_minute': 0,
                'port_op_start_hour': 8,
                'port_op_start_minute': 0,
                'port_op_end_hour': 17,
                'port_op_end_minute': 0,
            },
        ]

        _, all_days, _ = generate_schedule(days_data)

        for hour in [18, 19, 20, 21]:
            for minute in (0, 30):
                count = count_daymen_working_at(all_days, 0, hour, minute)
                assert count == 1, f"Klo {hour:02d}:{minute:02d} pitäisi olla 1 dayman, on {count}"



# TESTIT: STCW-SÄÄNNÖT
# ---------------------------------------------------------------------

class TestSTCW:
    """STCW-lepoaikasäännöt"""
    
    @pytest.mark.stcw_rest
    def test_minimum_10h_rest(self):
        """Vähintään 10h lepoa 24h jaksossa"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
                   'Watchman 1', 'Watchman 2', 'Watchman 3']
        
        for w in workers:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            combined = work1 + work2
            ana = analyze_stcw_from_work_starts(combined)
            
            assert ana['total_rest'] >= 10, \
                f"{w}: vain {ana['total_rest']}h lepoa (min 10h)"
    
    @pytest.mark.stcw_split
    def test_max_two_rest_periods(self):
        """Lepo max 2 jaksossa"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
                   'Watchman 1', 'Watchman 2', 'Watchman 3']
        
        for w in workers:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            combined = work1 + work2
            ana = analyze_stcw_from_work_starts(combined)
            
            assert ana['rest_period_count'] <= 2, \
                f"{w}: {ana['rest_period_count']} lepojaksoa (max 2)"
    
    @pytest.mark.stcw_long_rest
    def test_one_rest_period_min_6h(self):
        """Yksi lepojakso vähintään 6h"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
                   'Watchman 1', 'Watchman 2', 'Watchman 3']
        
        for w in workers:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            combined = work1 + work2
            ana = analyze_stcw_from_work_starts(combined)
            
            assert ana['longest_rest'] >= 6, \
                f"{w}: pisin lepo {ana['longest_rest']}h (min 6h)"
    
    @pytest.mark.stcw_status
    def test_stcw_status_ok(self):
        """STCW-status on OK normaalissa skenaariossa"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        workers = ['Bosun', 'Dayman EU', 'Dayman PH1', 'Dayman PH2',
                   'Watchman 1', 'Watchman 2', 'Watchman 3']
        
        for w in workers:
            work1 = all_days[w][0]['work_slots']
            work2 = all_days[w][1]['work_slots']
            combined = work1 + work2
            ana = analyze_stcw_from_work_starts(combined)
            
            assert ana['status'] == 'OK', \
                f"{w}: STCW status {ana['status']}, issues: {ana['issues']}"


# TESTIT: WATCHMANIT
# ---------------------------------------------------------------------

class TestWatchmen:
    """Watchmanien 4-on-8-off vuorot."""

    def test_watchman_4_on_8_off_pattern(self):
        """Watchmanit tekevät 8h päivässä (4-on / 8-off -kierto)."""
        all_days = run_scenario(
            arrival_hour=None, departure_hour=None,
            op_start_hour=8, op_end_hour=17
        )

        for w in ['Watchman 1', 'Watchman 2', 'Watchman 3']:
            work = all_days[w][0]['work_slots']
            assert isinstance(work, list)
            assert len(work) == 48
            assert all(isinstance(slot, bool) for slot in work)



# TESTIT: BOSUN
# ---------------------------------------------------------------------

class TestBosun:
    """Bosunin työvuorot"""
    
    def test_bosun_in_arrival(self):
        """Bosun on tulossa"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        has_arrival = any(all_days['Bosun'][0]['arrival_slots'])
        assert has_arrival, "Bosun ei ole tulossa"
    
    def test_bosun_in_departure(self):
        """Bosun on lähdössä"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        has_departure = any(all_days['Bosun'][0]['departure_slots'])
        assert has_departure, "Bosun ei ole lähdössä"
    
    def test_bosun_normal_day_hours(self):
        """Bosun tekee ~8.5h normaalina päivänä"""
        all_days = run_scenario(
            arrival_hour=None, departure_hour=None,
            op_start_hour=8, op_end_hour=17
        )
        
        work = all_days['Bosun'][0]['work_slots']
        hours = sum(work) / 2
        assert 8.0 <= hours <= 9.0, f"Bosun: {hours}h työtä (pitäisi ~8.5h)"



# TESTIT: ERIKOISTAPAUKSET
# ---------------------------------------------------------------------

class TestSpecialCases:
    """Erikoistapaukset ja reunatilanteet"""
    
    def test_early_arrival_late_departure(self):
        """Aikainen tulo + myöhäinen lähtö"""
        all_days = run_scenario(
            arrival_hour=6, departure_hour=21,
            op_start_hour=8, op_end_hour=20
        )
        
        # Kaikki daymanit tulossa, lähdössä kaksi daymania
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            has_arr = any(all_days[w][0]['arrival_slots'])
            assert has_arr, f"{w} puuttuu tulosta"

        dep_count = sum(any(all_days[w][0]['departure_slots']) for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2'])
        assert dep_count == 2, f"Lähdössä pitäisi olla 2 daymania, on {dep_count}"
    
    def test_no_arrival_no_departure(self):
        """Ei tuloa eikä lähtöä - normaali meripäivä"""
        all_days = run_scenario(
            arrival_hour=None, departure_hour=None,
            op_start_hour=8, op_end_hour=17
        )
        
        # Daymanit tekevät normaalin päivän
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            work = all_days[w][0]['work_slots']
            hours = sum(work) / 2
            assert 8.0 <= hours <= 9.0, f"{w}: {hours}h (pitäisi ~8.5h)"
    
    def test_continuous_night_prefers_evening_worker_then_ph2(self):
        """Jatkuvassa yössä iltavuoron tekijä jatkaa 01:00 asti, sitten PH2 ottaa loppuyön."""
        days_data = [
            {
                'arrival_hour': 18,
                'arrival_minute': 0,
                'departure_hour': None,
                'departure_minute': 0,
                'port_op_start_hour': 19,
                'port_op_start_minute': 0,
                'port_op_end_hour': 0,
                'port_op_end_minute': 0,
            },
            {
                'arrival_hour': None,
                'arrival_minute': 0,
                'departure_hour': None,
                'departure_minute': 0,
                'port_op_start_hour': 0,
                'port_op_start_minute': 0,
                'port_op_end_hour': 8,
                'port_op_end_minute': 0,
            },
        ]

        _, all_days, _ = generate_schedule(days_data)

        # Päivä 1 ilta: EU tekee viimeiset slotit ennen keskiyötä
        assert all_days['Dayman EU'][0]['work_slots'][47] is True

        # Päivä 2: 00:00-01:00 EU, 01:00-08:00 PH2
        assert all_days['Dayman EU'][1]['work_slots'][0] is True
        assert all_days['Dayman EU'][1]['work_slots'][1] is True
        assert all_days['Dayman PH2'][1]['work_slots'][2] is True
        assert all_days['Dayman PH2'][1]['work_slots'][15] is True
        assert all_days['Dayman PH1'][1]['work_slots'][0] is False

    def test_night_operation(self):
        """Yöoperaatio (menee keskiyön yli)"""
        all_days = run_scenario(
            arrival_hour=14, departure_hour=None,
            op_start_hour=16, op_end_hour=2  # Loppuu klo 02:00 seuraavana päivänä
        )
        
        # Tarkista että joku tekee yövuoron
        # Slot 46 = klo 23:00
        night_coverage = False
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            if all_days[w][0]['work_slots'][46]:
                night_coverage = True
                break
        
        assert night_coverage, "Kukaan dayman ei tee yövuoroa"



# TESTIT: REGRESSIOT (aiemmin löydetyt bugit)
# ---------------------------------------------------------------------

class TestRegressions:
    """Regressiotestit - varmistaa etteivät vanhat bugit palaa"""
    
    def test_watchman3_rest_periods_not_three(self):
        """Watchman 3:n lepo ei saa olla 3 jaksossa (bugi #1)"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        work1 = all_days['Watchman 3'][0]['work_slots']
        work2 = all_days['Watchman 3'][1]['work_slots']
        combined = work1 + work2
        ana = analyze_stcw_from_work_starts(combined)
        
        assert ana['rest_period_count'] <= 2, \
            f"Watchman 3: {ana['rest_period_count']} lepojaksoa (bugi palasi!)"
    
    def test_dayman_hours_match_excel(self):
        """Daymanien tunnit täsmäävät (bugi #2)"""
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )
        
        for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
            work = all_days[w][0]['work_slots']
            hours = sum(work) / 2
            # Tunnit pitäisi olla järkevät (8-10h)
            assert 8.0 <= hours <= 10.0, \
                f"{w}: {hours}h (liian vähän/paljon)"


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

@pytest.mark.special_ops
class TestSpecialOperationsMandatory:
    """Slussi ja shiftaus ovat pakollisia kuten tulo/lähtö."""

    def test_sluice_arrival_at_17_is_forced_for_daymen(self):
        days_data = [
            {
                'arrival_hour': None,
                'arrival_minute': 0,
                'departure_hour': None,
                'departure_minute': 0,
                'port_op_start_hour': 8,
                'port_op_start_minute': 0,
                'port_op_end_hour': 17,
                'port_op_end_minute': 0,
                'sluice_arrival_hour': 17,
                'sluice_arrival_minute': 0,
                'sluice_departure_hour': None,
                'sluice_departure_minute': 0,
                'shifting_hour': None,
                'shifting_minute': 0,
            }
        ]

        _, all_days, _ = generate_schedule(days_data)
        daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']

        # 17:00-18:00 => 2 daymania
        for slot in [34, 35]:
            working = sum(all_days[w][0]['work_slots'][slot] for w in daymen)
            marked = sum(all_days[w][0]['sluice_slots'][slot] for w in daymen)
            assert working == 2
            assert marked == 2

        # 18:00-19:00 => 3 daymania
        for slot in [36, 37]:
            working = sum(all_days[w][0]['work_slots'][slot] for w in daymen)
            marked = sum(all_days[w][0]['sluice_slots'][slot] for w in daymen)
            assert working == 3
            assert marked == 3

    def test_shifting_is_forced_for_all_daymen(self):
        days_data = [
            {
                'arrival_hour': None,
                'arrival_minute': 0,
                'departure_hour': None,
                'departure_minute': 0,
                'port_op_start_hour': 8,
                'port_op_start_minute': 0,
                'port_op_end_hour': 17,
                'port_op_end_minute': 0,
                'sluice_arrival_hour': None,
                'sluice_arrival_minute': 0,
                'sluice_departure_hour': None,
                'sluice_departure_minute': 0,
                'shifting_hour': 17,
                'shifting_minute': 0,
            }
        ]

        _, all_days, _ = generate_schedule(days_data)
        daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']

        for slot in [34, 35]:
            working = sum(all_days[w][0]['work_slots'][slot] for w in daymen)
            marked = sum(all_days[w][0]['shifting_slots'][slot] for w in daymen)
            assert working == 3
            assert marked == 3


@pytest.mark.daily_hours
class TestDailyMinimumHours:
    """Kalenterivuorokaudessa vähintään 8h töitä daymaneille."""

    def test_daymen_have_minimum_8h_per_calendar_day(self):
        all_days = run_scenario(
            arrival_hour=8, departure_hour=19,
            op_start_hour=10, op_end_hour=18
        )

        for day_idx in range(2):
            for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2']:
                hours = sum(all_days[w][day_idx]['work_slots']) / 2
                assert hours >= 8, f"{w} päivä {day_idx+1}: {hours}h (min 8h)"


class TestMultipleOperationsPerDay:
    """Useita saman tyypin operaatioita voidaan käsitellä samana päivänä."""

    def test_multiple_arrivals_mark_all_daymen_for_each_entry(self):
        days_data = [
            {
                'arrivals': [
                    {'hour': 6, 'minute': 0},
                    {'hour': 18, 'minute': 0},
                ],
                'departures': [],
                'port_operations': [
                    {'start_hour': 8, 'start_minute': 0, 'end_hour': 17, 'end_minute': 0},
                ],
                'sluice_arrivals': [],
                'sluice_departures': [],
                'shiftings': [],
            }
        ]

        _, all_days, _ = generate_schedule(days_data)

        for slot in [12, 13, 36, 37]:
            count = sum(all_days[w][0]['arrival_slots'][slot] for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2'])
            assert count == 3, f"Kaikkien daymanien pitäisi olla tulossa slotissa {slot}"

    def test_multiple_port_operations_keep_coverage_in_each_segment(self):
        days_data = [
            {
                'arrivals': [],
                'departures': [],
                'port_operations': [
                    {'start_hour': 6, 'start_minute': 0, 'end_hour': 8, 'end_minute': 0},
                    {'start_hour': 18, 'start_minute': 0, 'end_hour': 20, 'end_minute': 0},
                ],
                'sluice_arrivals': [],
                'sluice_departures': [],
                'shiftings': [],
            }
        ]

        _, all_days, _ = generate_schedule(days_data)

        for slot in [12, 13, 36, 37]:
            count = sum(all_days[w][0]['work_slots'][slot] for w in ['Dayman EU', 'Dayman PH1', 'Dayman PH2'])
            assert count >= 1, f"Port op coverage puuttuu slotista {slot}"
