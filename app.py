import streamlit as st
import io
from sea_watch_10 import generate_schedule, parse_time
import sea_watch_10
print("=== DEBUG: sea_watch_10 ladattu tiedostosta:", sea_watch_10.__file__)

# importoi uusi puhdistettu generaattori
from sea_watch_10 import generate_schedule, parse_time


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
# SOVELLUS
# -------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Sea Watch 9 – Vuorogeneraattori", layout="wide")

    st.title("🛳️ Sea Watch 9 – Työvuorolistageneraattori")
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

        # Tekstiraportti
        st.subheader("📄 Raportti ja STCW-analyysi")
        st.text(report)

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

