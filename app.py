import streamlit as st
import io
import pandas as pd

# importoi esittelyversio
from sea_watch_10 import generate_schedule, parse_time, time_to_index, index_to_time_str


# -------------------------------------------------------------------
# APU: TIME PARSER STREAMLITILLE
# -------------------------------------------------------------------
def parse_optional_time(label: str, key: str):
    """
    Ottaa käyttäjältä HH:MM-merkkijonon ja palauttaa (hour, minute) tai (None, None).
    """
    val = st.text_input(label, key=key)
    val = val.strip()

    if val == "":
        return None, None

    h, m = parse_time(val)
    return h, m


# -------------------------------------------------------------------
# APU: RAKENTAA LISTAN PÄIVÄTIEDOISTA
# -------------------------------------------------------------------
def build_days_data(num_days: int):
    """
    Luo sovelluksen input-kentistä days_data-listan,
    jonka generate_schedule() tarvitsee.
    """
    days = []

    for day in range(1, num_days + 1):

        with st.expander(f"Päivä {day}", expanded=(day == 1)):
            col1, col2 = st.columns(2)

            with col1:
                arr_h, arr_m = parse_optional_time(
                    "Satamaan tuloaika (HH:MM, tyhjä jos ei tuloa)",
                    key=f"arr_{day}"
                )
                dep_h, dep_m = parse_optional_time(
                    "Satamasta lähtöaika (HH:MM, tyhjä jos ei lähtöä)",
                    key=f"dep_{day}"
                )

            with col2:
                op_s_h, op_s_m = parse_optional_time(
                    "Satamaoperaation alku (HH:MM)",
                    key=f"opstart_{day}"
                )
                op_e_h, op_e_m = parse_optional_time(
                    "Satamaoperaation loppu (HH:MM)",
                    key=f"opend_{day}"
                )

            days.append({
                "arrival_hour": arr_h,
                "arrival_minute": arr_m or 0,
                "departure_hour": dep_h,
                "departure_minute": dep_m or 0,
                "port_op_start_hour": op_s_h,
                "port_op_start_minute": op_s_m or 0,
                "port_op_end_hour": op_e_h,
                "port_op_end_minute": op_e_m or 0,
            })

    return days


# -------------------------------------------------------------------
# APU: LUO TAULUKKONÄKYMÄ
# -------------------------------------------------------------------
def create_schedule_table(all_days, day_idx, workers):
    """
    Luo pandas DataFrame työvuorotaulukosta.
    """
    # Aikasarakkeet (00:00 - 23:30)
    time_cols = [f"{h:02d}:{m:02d}" for h in range(24) for m in [0, 30]]
    
    # Rakenna data
    data = []
    for w in workers:
        row = {'Työntekijä': w}
        day_data = all_days[w][day_idx]
        work = day_data['work_slots']
        arr = day_data['arrival_slots']
        dep = day_data['departure_slots']
        
        for i, time_col in enumerate(time_cols):
            if i < len(work):
                if arr[i]:
                    row[time_col] = 'B'
                elif dep[i]:
                    row[time_col] = 'C'
                elif work[i]:
                    row[time_col] = '●'
                else:
                    row[time_col] = ''
            else:
                row[time_col] = ''
        data.append(row)
    
    return pd.DataFrame(data)


def style_schedule_table(df):
    """
    Lisää värit taulukkoon.
    """
    def color_cell(val):
        if val == '●':
            return 'background-color: #4472C4; color: white'
        elif val == 'B':
            return 'background-color: #FFC000; color: black'
        elif val == 'C':
            return 'background-color: #FF6600; color: white'
        else:
            return ''
    
    return df.style.applymap(color_cell, subset=df.columns[1:])


# -------------------------------------------------------------------
# SOVELLUS
# -------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Sea Watch - Työvuorogeneraattori", layout="wide")

    st.title("🛳️ Sea Watch - Työvuorolistageneraattori")
    st.write("Syötä päivien tulo-/lähtöajat ja satamaoperaatiot, niin sovellus "
             "laskee työvuorot ja STCW-lepoajat automaattisesti.")

    # Määrä sivupalkkiin
    num_days = st.sidebar.number_input(
        "Päivien määrä",
        min_value=1,
        max_value=14,
        value=2,
        step=1
    )

    st.sidebar.info("Vinkki: jätä kenttä tyhjäksi jos kyseistä tapahtumaa ei ole.")

    # Kerää käyttäjältä kaikkien päivien datat
    days_data = build_days_data(num_days)

    if st.button("🚀 Generoi työvuorot"):

        try:
            wb, all_days, report = generate_schedule(days_data)
        except Exception as e:
            st.error(f"Virhe generoinnissa: {e}")
            raise

        # Työntekijälista
        workers = ["Bosun", "Dayman EU", "Dayman PH1", "Dayman PH2", 
                   "Watchman 1", "Watchman 2", "Watchman 3"]

        # Näytä taulukot jokaiselle päivälle
        st.subheader("📋 Työvuorot")
        
        for d in range(num_days):
            st.markdown(f"**Päivä {d+1}**")
            
            df = create_schedule_table(all_days, d, workers)
            styled_df = style_schedule_table(df)
            
            # Näytä vain osa sarakkeista kerrallaan (parempi luettavuus)
            st.dataframe(styled_df, use_container_width=True, height=300)
            
            st.markdown("---")

        # STCW-yhteenveto
        st.subheader("📊 STCW-lepoaika-analyysi")
        
        stcw_data = []
        for d in range(1, num_days):  # Päivästä 2 alkaen
            for w in workers:
                dat = all_days[w][d]
                prev = all_days[w][d-1]['work_slots']
                work = dat['work_slots']
                
                from sea_watch_10 import analyze_stcw_from_work_starts
                combined = prev + work
                ana = analyze_stcw_from_work_starts(combined)
                
                hours = sum(work) / 2
                stcw_data.append({
                    'Päivä': d + 1,
                    'Työntekijä': w,
                    'Työtunnit': hours,
                    'Lepo (h)': ana['total_rest'],
                    'Pisin lepo (h)': ana['longest_rest'],
                    'Status': '✓ OK' if ana['status'] == 'OK' else '⚠ VAROITUS'
                })
        
        stcw_df = pd.DataFrame(stcw_data)
        st.dataframe(stcw_df, use_container_width=True)

        # Excel-tiedostoksi muistiin
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Lataa Excel-työvuorolista",
            data=buffer,
            file_name="tyovuorot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
