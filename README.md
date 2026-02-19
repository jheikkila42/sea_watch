# Sea Watch

## Testien ajo komentoriviltä

Voit ajaa kaikki testit:

```bash
python run_tests.py all
```

Tai yksittäiset osiot:

```bash
python run_tests.py stcw-rest
python run_tests.py stcw-split
python run_tests.py stcw-long-rest
python run_tests.py stcw-status
python run_tests.py daily-hours
python run_tests.py special-ops
```

## Projektiin mallinnetut STCW-säännöt

Projektin STCW-tarkistus (`check_stcw_at_slot` / `analyze_stcw_from_work_starts`) käyttää näitä ehtoja 24h ikkunassa:

- lepoa yhteensä vähintään **10 h**
- lepo korkeintaan **2 jaksossa**
- vähintään yksi lepojakso vähintään **6 h**

Näille löytyy nyt erilliset marker-pohjaiset testit.


## Kokeellinen constrained-optimointi (Vaihe 1)

Streamlitissä on valinta **"Kokeellinen constrained-optimointi (daymanit)"** sivupalkissa.

Tämä vaatii OR-Toolsin:

```bash
pip install ortools
```

Ilman sitä sovellus näyttää virheilmoituksen constrained-tilaa käytettäessä.
