import copy
import io

import pandas as pd
import streamlit as st

import sea_watch_10 as sw

check_stcw_at_slot = getattr(sw, "check_stcw_at_slot", None)
generate_schedule = sw.generate_schedule
generate_schedule_with_manual_day1 = getattr(sw, "generate_schedule_with_manual_day1", None)
build_workbook_and_report = getattr(sw, "build_workbook_and_report", None)

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
HALF_DAY_COLS = [TIME_COLS[:24], TIME_COLS[24:]]



def build_workbook_compat(all_days, num_days, workers):
    if build_workbook_and_report is None:
        return None
    wb, _ = build_workbook_and_report(all_days, num_days, workers)
    return wb





def ensure_sea_watch_time_helpers():
    """Varmistaa, että sea_watch_10-moduulissa on aikamuunnoshelperit käytettävissä.

    Joissain ympäristöissä/versioissa generate_schedule voi viitata näihin nimiin
    module-globaalina, vaikka ne puuttuisivat exporteista.
    """
    if not hasattr(sw, "time_to_index"):
        def _time_to_index(h, m):
            return h * 2 + (1 if m >= 30 else 0)
        sw.time_to_index = _time_to_index

    if not hasattr(sw, "index_to_time_str"):
        def _index_to_time_str(idx):
            h = idx // 2
            m = "30" if idx % 2 else "00"
            return f"{h:02d}:{m}"
        sw.index_to_time_str = _index_to_time_str

def parse_time(time_str: str):
    normalized = time_str.strip().replace(".", ":")

    if ":" in normalized:
        parts = normalized.split(":")
        if len(parts) != 2:
            raise ValueError(f"Virheellinen aika: {time_str}")
        h, m = int(parts[0]), int(parts[1])
    else:
        # Salli myös muodot kuten "22" => 22:00
        h, m = int(normalized), 0

    if not (0 <= h <= 23 and m in (0, 30)):
        raise ValueError("Ajan pitää olla muodossa HH:MM, HH.MM tai pelkkä tunti (esim 22)")
    return h, m


def parse_optional_time(label: str, key: str):
    val = st.text_input(label, key=key).strip()
    if val == "":
        return None, None
    h, m = parse_time(val)
    return h, m




def display_time_label(time_col: str) -> str:
    return time_col.replace(":", ".")


def create_time_column_config(time_cols):
    return {
        display_time_label(c): st.column_config.CheckboxColumn(display_time_label(c), default=False, width="small")
        for c in time_cols
    }

def build_days_data(start_day: int, end_day: int, key_prefix: str):
    days = []
    for day in range(start_day, end_day + 1):
        with st.expander(f"Päivä {day}", expanded=(day == start_day)):
            row1_col1, row1_col2 = st.columns(2)
            row2_col1, row2_col2 = st.columns(2)

            with row1_col1:
                st.markdown("#### Tulo ja lähtö")
                arr_h, arr_m = parse_optional_time(
                    "Satamaan tuloaika (HH:MM, tyhjä jos ei tuloa)",
                    key=f"{key_prefix}_arr_{day}",
                )
                dep_h, dep_m = parse_optional_time(
                    "Satamasta lähtöaika (HH:MM, tyhjä jos ei lähtöä)",
                    key=f"{key_prefix}_dep_{day}",
                )

            with row1_col2:
                st.markdown("#### Satamaoperaatiot")
                op_s_h, op_s_m = parse_optional_time(
                    "Satamaoperaation alku (HH:MM)",
                    key=f"{key_prefix}_opstart_{day}",
                )
                op_e_h, op_e_m = parse_optional_time(
                    "Satamaoperaation loppu (HH:MM)",
                    key=f"{key_prefix}_opend_{day}",
                )

            with row2_col1:
                st.markdown("#### Slussi")
                sluice_arr_h, sluice_arr_m = parse_optional_time(
                    "Slussi - tulo alku (HH:MM, kesto 2h)",
                    key=f"{key_prefix}_sluice_arr_{day}",
                )
                sluice_dep_h, sluice_dep_m = parse_optional_time(
                    "Slussi - lähtö alku (HH:MM, kesto 2h)",
                    key=f"{key_prefix}_sluice_dep_{day}",
                )

            with row2_col2:
                st.markdown("#### Shiftaus")
                shifting_h, shifting_m = parse_optional_time(
                    "Shiftaus alku (HH:MM, kesto 1h)",
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
                    "sluice_arrival_hour": sluice_arr_h,
                    "sluice_arrival_minute": sluice_arr_m or 0,
                    "sluice_departure_hour": sluice_dep_h,
                    "sluice_departure_minute": sluice_dep_m or 0,
                    "shifting_hour": shifting_h,
                    "shifting_minute": shifting_m or 0,
                }
            )

    return days


def create_schedule_table(all_days, day_idx, workers, time_cols):
    indices = [TIME_COLS.index(c) for c in time_cols]
    labels = [display_time_label(c) for c in time_cols]

    data = []
    for w in workers:
        row = {"Työntekijä": w}
        day_data = all_days[w][day_idx]
        work = day_data["work_slots"]
        arr = day_data["arrival_slots"]
        dep = day_data["departure_slots"]
        sluice = day_data.get("sluice_slots", [False] * 48)
        shifting = day_data.get("shifting_slots", [False] * 48)

        for idx, label in zip(indices, labels):
            if sluice[idx]:
                row[label] = "SL"
            elif shifting[idx]:
                row[label] = "SH"
            elif arr[idx]:
                row[label] = "B"
            elif dep[idx]:
                row[label] = "C"
            elif work[idx]:
                row[label] = "●"
            else:
                row[label] = ""
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


def create_editable_work_df(all_days, day_idx, time_cols):
    data = {"Työntekijä": WORKERS}
    for col in time_cols:
        label = display_time_label(col)
        col_idx = TIME_COLS.index(col)
        data[label] = [bool(all_days[w][day_idx]["work_slots"][col_idx]) for w in WORKERS]
    return pd.DataFrame(data)


def apply_edited_work_df(all_days, day_idx, edited_df, time_cols):
    for _, row in edited_df.iterrows():
        worker = row["Työntekijä"]
        if worker not in all_days:
            continue
        for col in time_cols:
            label = display_time_label(col)
            col_idx = TIME_COLS.index(col)
            all_days[worker][day_idx]["work_slots"][col_idx] = bool(row[label])


def render_results(num_days, wb, all_days):
    st.subheader("📋 Työvuorot")
    for d in range(num_days):
        st.markdown(f"**Päivä {d+1}**")
        st.caption("00.00–11.30")
        df1 = create_schedule_table(all_days, d, WORKERS, HALF_DAY_COLS[0])
        st.dataframe(style_schedule_table(df1), use_container_width=True, height=260)

        st.caption("12.00–23.30")
        df2 = create_schedule_table(all_days, d, WORKERS, HALF_DAY_COLS[1])
        st.dataframe(style_schedule_table(df2), use_container_width=True, height=260)
        st.markdown("---")

    st.subheader("📊 STCW-lepoaika-analyysi")
    if check_stcw_at_slot is None:
        st.info("Käytössä oleva sea_watch_10-versio ei sisällä STCW-analyysifunktiota (check_stcw_at_slot).")
    else:
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


def store_generated_result(wb, all_days, num_days):
    st.session_state["generated_wb"] = wb
    st.session_state["generated_all_days"] = all_days
    st.session_state["generated_num_days"] = num_days


def render_post_generation_editor():
    if "generated_all_days" not in st.session_state:
        return

    st.markdown("## ✏️ Muokkaa vuoroja generoinnin jälkeen")
    st.caption(
        "Voit klikata soluja (ja maalata alueita data editorissa) muuttaaksesi työslotteja. "
        "Paina lopuksi 'Generoi uudelleen', niin Excel-linkki päivittyy uuden tilanteen mukaan."
    )

    num_days = st.session_state["generated_num_days"]
    all_days = st.session_state["generated_all_days"]

    with st.form("post_generation_edit_form"):
        edited_dfs = []
        for d in range(num_days):
            st.markdown(f"**Muokattava päivä {d+1}**")
            day_edits = []
            for half_idx, half_cols in enumerate(HALF_DAY_COLS):
                st.caption("00.00–11.30" if half_idx == 0 else "12.00–23.30")
                base_df = create_editable_work_df(all_days, d, half_cols)
                edited_df = st.data_editor(
                    base_df,
                    hide_index=True,
                    use_container_width=True,
                    key=f"post_edit_day_{d}_half_{half_idx}",
                    disabled=["Työntekijä"],
                    column_config=create_time_column_config(half_cols),
                )
                day_edits.append((edited_df, half_cols))
            edited_dfs.append(day_edits)

        regenerate_clicked = st.form_submit_button("🔁 Generoi uudelleen (päivitä Excel)")

    if regenerate_clicked:
        updated_all_days = copy.deepcopy(st.session_state["generated_all_days"])
        for d, day_edits in enumerate(edited_dfs):
            for edited_df, half_cols in day_edits:
                apply_edited_work_df(updated_all_days, d, edited_df, half_cols)

        wb = build_workbook_compat(updated_all_days, num_days, WORKERS)
        if wb is None:
            st.error("Käytössä oleva sea_watch_10-versio ei tue Excelin uudelleenrakennusta muokatuista vuoroista (build_workbook_and_report puuttuu).")
        else:
            store_generated_result(wb, updated_all_days, num_days)
            st.success("Vuorot päivitetty. Excel-linkki käyttää nyt muokattua listaa.")


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
            ensure_sea_watch_time_helpers()
            wb, all_days, _ = generate_schedule(days_data)
            store_generated_result(wb, all_days, num_days)

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
                c: st.column_config.CheckboxColumn(c, default=False) for c in TIME_COLS
            },
        )

        if num_days >= 2:
            st.markdown("#### Päivät 2+ (normaali syöttö)")
            days_data_rest = build_days_data(2, num_days, key_prefix="manual")
        else:
            days_data_rest = []

        if st.button("🚀 Generoi työvuorot (manuaalinen päivä 1)", key="gen_manual"):
            day1_placeholder = {
                "arrival_hour": None,
                "arrival_minute": 0,
                "departure_hour": None,
                "departure_minute": 0,
                "port_op_start_hour": 8,
                "port_op_start_minute": 0,
                "port_op_end_hour": 17,
                "port_op_end_minute": 0,
                "sluice_arrival_hour": None,
                "sluice_arrival_minute": 0,
                "sluice_departure_hour": None,
                "sluice_departure_minute": 0,
                "shifting_hour": None,
                "shifting_minute": 0,
            }
            days_data = [day1_placeholder] + days_data_rest
            manual_slots = convert_manual_df_to_slots(manual_df)

            ensure_sea_watch_time_helpers()
            if generate_schedule_with_manual_day1 is None:
                st.warning("Käytössä oleva sea_watch_10-versio ei tue manuaalista päivä 1 -generaatiota. Käytetään automaattista generaatiota.")
                wb, all_days, _ = generate_schedule(days_data)
            else:
                wb, all_days, _ = generate_schedule_with_manual_day1(days_data, manual_slots)
            store_generated_result(wb, all_days, num_days)

    if "generated_wb" in st.session_state and "generated_all_days" in st.session_state:
        render_results(
            st.session_state["generated_num_days"],
            st.session_state["generated_wb"],
            st.session_state["generated_all_days"],
        )
        render_post_generation_editor()


if __name__ == "__main__":
    main()
