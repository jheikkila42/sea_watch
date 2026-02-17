import io
import pandas as pd
import streamlit as st

from sea_watch_10 import (
    check_stcw_at_slot,
    generate_schedule,
    generate_schedule_with_manual_day1,
)

WORKERS = [
    "Bosun",
    "Dayman EU",
    "Dayman PH1",
    "Dayman PH2",
    "Watchman 1",
    "Watchman 2",
    "Watchman 3",
]
TIME_COLS = [f"{h:02d}:{m:02d}" for h in range(24) for m in [0, 30]]


def parse_time(time_str: str):
    parts = time_str.split(":")
    if len(parts) != 2:
        raise ValueError(f"Virheellinen aika: {time_str}")
    h, m = int(parts[0]), int(parts[1])
    if not (0 <= h <= 23 and m in (0, 30)):
        raise ValueError("Ajan pitää olla muodossa HH:00 tai HH:30")
    return h, m


def parse_optional_time(label: str, key: str):
    val = st.text_input(label, key=key).strip()
    if val == "":
        return None, None
    h, m = parse_time(val)
    return h, m


def build_days_data(start_day: int, end_day: int, key_prefix: str):
    days = []
    for day in range(start_day, end_day + 1):
        with st.expander(f"Päivä {day}", expanded=(day == start_day)):
            col1, col2, col3 = st.columns(3)

            with col1:
                arr_h, arr_m = parse_optional_time(
                    "Satamaan tuloaika (HH:MM, tyhjä jos ei tuloa)",
                    key=f"{key_prefix}_arr_{day}",
                )
                dep_h, dep_m = parse_optional_time(
                    "Satamasta lähtöaika (HH:MM, tyhjä jos ei lähtöä)",
                    key=f"{key_prefix}_dep_{day}",
                )

            with col2:
                op_s_h, op_s_m = parse_optional_time(
                    "Satamaoperaation alku (HH:MM)",
                    key=f"{key_prefix}_opstart_{day}",
                )
                op_e_h, op_e_m = parse_optional_time(
                    "Satamaoperaation loppu (HH:MM)",
                    key=f"{key_prefix}_opend_{day}",
                )

            with col3:
                sluice_h, sluice_m = parse_optional_time(
                    "Slussi alku (HH:MM, kesto 2h)",
                    key=f"{key_prefix}_sluice_{day}",
                )
                shifting_h, shifting_m = parse_optional_time(
                    "Shiftaus alku (HH:MM, kesto 2h)",
                    key=f"{key_prefix}_shifting_{day}",
                )

            days.append(
                {
                    "arrival_hour": arr_h,
                    "arrival_minute": arr_m or 0,
                    "departure_hour": dep_h,
                    "departure_minute": dep_m or 0,
                    "port_op_start_hour": op_s_h,
                    "port_op_start_minute": op_s_m or 0,
                    "port_op_end_hour": op_e_h,
                    "port_op_end_minute": op_e_m or 0,
                    "sluice_hour": sluice_h,
                    "sluice_minute": sluice_m or 0,
                    "shifting_hour": shifting_h,
                    "shifting_minute": shifting_m or 0,
                }
            )

    return days


def create_schedule_table(all_days, day_idx, workers):
    data = []
    for w in workers:
        row = {"Työntekijä": w}
        day_data = all_days[w][day_idx]
        work = day_data["work_slots"]
        arr = day_data["arrival_slots"]
        dep = day_data["departure_slots"]
        sluice = day_data.get("sluice_slots", [False] * 48)
        shifting = day_data.get("shifting_slots", [False] * 48)

        for i, time_col in enumerate(TIME_COLS):
            if arr[i]:
                row[time_col] = "B"
            elif dep[i]:
                row[time_col] = "C"
            elif sluice[i]:
                row[time_col] = "SL"
            elif shifting[i]:
                row[time_col] = "SH"
            elif work[i]:
                row[time_col] = "●"
            else:
                row[time_col] = ""
        data.append(row)

    return pd.DataFrame(data)


def style_schedule_table(df):
    def color_cell(val):
        if val == "●":
            return "background-color: #4472C4; color: white"
        if val == "B":
            return "background-color: #FFC000; color: black"
        if val == "C":
            return "background-color: #FF6600; color: white"
        if val == "SL":
            return "background-color: #C9A0DC; color: black"
        if val == "SH":
            return "background-color: #FFB6C1; color: black"
        return ""

    return df.style.applymap(color_cell, subset=df.columns[1:])


def init_manual_day1_df():
    data = {"Työntekijä": WORKERS}
    for col in TIME_COLS:
        data[col] = [False] * len(WORKERS)
    return pd.DataFrame(data)


def convert_manual_df_to_slots(df: pd.DataFrame):
    manual = {}
    for _, row in df.iterrows():
        worker = row["Työntekijä"]
        manual[worker] = [bool(row[t]) for t in TIME_COLS]
    return manual


def render_results(num_days, wb, all_days):
    st.subheader("📋 Työvuorot")
    for d in range(num_days):
        st.markdown(f"**Päivä {d+1}**")
        df = create_schedule_table(all_days, d, WORKERS)
        st.dataframe(style_schedule_table(df), use_container_width=True, height=300)
        st.markdown("---")

    st.subheader("📊 STCW-lepoaika-analyysi")
    stcw_data = []
    for d in range(1, num_days):
        for w in WORKERS:
            prev = all_days[w][d - 1]["work_slots"]
            work = all_days[w][d]["work_slots"]
            ana = check_stcw_at_slot(prev + work, 95)
            stcw_data.append(
                {
                    "Päivä": d + 1,
                    "Työntekijä": w,
                    "Työtunnit": sum(work) / 2,
                    "Lepo (h)": ana["total_rest"],
                    "Pisin lepo (h)": ana["longest_rest"],
                    "Status": "✓ OK" if ana["status"] == "OK" else "⚠ VAROITUS",
                }
            )

    st.dataframe(pd.DataFrame(stcw_data), use_container_width=True)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    st.download_button(
        label="📥 Lataa Excel-työvuorolista",
        data=buffer,
        file_name="tyovuorot.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main():
    st.set_page_config(page_title="Sea Watch - Testivuorogeneraattori", layout="wide")

    st.title("🛳️ Sea Watch - Työvuorolistageneraattori")
    st.write(
        "Syötä päivien tiedot. Voit käyttää automaattista syöttöä kaikille päiville "
        "tai syöttää päivän 1 manuaalisesti maalaamalla tunnit taulukkoon."
    )

    num_days = st.sidebar.number_input(
        "Päivien määrä", min_value=1, max_value=14, value=2, step=1
    )
    st.sidebar.info("Vinkki: jätä kenttä tyhjäksi jos kyseistä tapahtumaa ei ole.")

    tab_auto, tab_manual = st.tabs(["Automaattinen syöttö", "Päivä 1 manuaalinen"])

    with tab_auto:
        st.markdown("Syötä kaikkien päivien tulo-/lähtöajat ja satamaoperaatiot.")
        days_data = build_days_data(1, num_days, key_prefix="auto")

        if st.button("🚀 Generoi työvuorot", key="gen_auto"):
            wb, all_days, _ = generate_schedule(days_data)
            render_results(num_days, wb, all_days)

    with tab_manual:
        st.markdown(
            "Maalaa päivän 1 työtunnit työntekijöille. Päivien 2+ ajat syötetään normaalisti."
        )

        manual_default = init_manual_day1_df()
        manual_df = st.data_editor(
            manual_default,
            hide_index=True,
            use_container_width=True,
            key="manual_day1_editor",
            disabled=["Työntekijä"],
            column_config={
                c: st.column_config.CheckboxColumn(c, default=False)
                for c in TIME_COLS
            },
        )

        if num_days >= 2:
            st.markdown("#### Päivät 2+ (normaali syöttö)")
            days_data_rest = build_days_data(2, num_days, key_prefix="manual")
        else:
            days_data_rest = []

        if st.button("🚀 Generoi työvuorot (manuaalinen päivä 1)", key="gen_manual"):
            # Päivä 1 tapahtumadata jätetään neutraaliksi, koska työvuoro tulee manuaalisesti.
            day1_placeholder = {
                "arrival_hour": None,
                "arrival_minute": 0,
                "departure_hour": None,
                "departure_minute": 0,
                "port_op_start_hour": 8,
                "port_op_start_minute": 0,
                "port_op_end_hour": 17,
                "port_op_end_minute": 0,
                "sluice_hour": None,
                "sluice_minute": 0,
                "shifting_hour": None,
                "shifting_minute": 0,
            }
            days_data = [day1_placeholder] + days_data_rest
            manual_slots = convert_manual_df_to_slots(manual_df)

            wb, all_days, _ = generate_schedule_with_manual_day1(days_data, manual_slots)
            render_results(num_days, wb, all_days)


if __name__ == "__main__":
    main()
