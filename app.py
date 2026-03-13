# -*- coding: utf-8 -*-
"""
Sea Watch - Työvuorolistageneraattori + AI Agent
Yhdistetty versio: kaikki toiminnot + AI-chat rajoitteineen
"""

import copy
import io
import os

import pandas as pd
import streamlit as st

# Sea Watch moduulit
import sea_watch_17 as sw
from schedule_analyzer import analyze_schedule, format_analysis_report, get_analysis_for_llm
from llm_agent import create_agent
from constraint_parser import create_parser

# Funktiot
check_stcw_at_slot = getattr(sw, "check_stcw_at_slot", None)
generate_schedule = sw.generate_schedule
generate_schedule_with_manual_day1 = getattr(sw, "generate_schedule_with_manual_day1", None)
build_workbook_and_report = getattr(sw, "build_workbook_and_report", None)

WORKERS = [
    "Bosun", "Dayman EU", "Dayman PH1", "Dayman PH2",
    "Watchman 1", "Watchman 2", "Watchman 3",
]
DAYMEN = ["Dayman EU", "Dayman PH1", "Dayman PH2"]

TIME_COLS = [f"{h:02d}:{m:02d}" for h in range(24) for m in [0, 30]]
DISPLAY_TIME_COLS = [f"{h:02d}:00" if m == 0 else f"{h:02d}:30" for h in range(24) for m in [0, 30]]

# Käyttöraja (viestejä per sessio)
MAX_MESSAGES_PER_SESSION = 30


# ============================================================================
# APUFUNKTIOT
# ============================================================================

def get_api_key():
    """Hae API-avain Streamlit Secretsistä tai ympäristömuuttujasta."""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except:
        return os.environ.get("ANTHROPIC_API_KEY", "")


def init_session_state():
    """Alusta session state."""
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "message_count" not in st.session_state:
        st.session_state.message_count = 0
    if "generated_all_days" not in st.session_state:
        st.session_state.generated_all_days = None
    if "generated_days_data" not in st.session_state:
        st.session_state.generated_days_data = None
    if "generated_wb" not in st.session_state:
        st.session_state.generated_wb = None
    if "generated_num_days" not in st.session_state:
        st.session_state.generated_num_days = 0
    if "analysis" not in st.session_state:
        st.session_state.analysis = None
    if "agent" not in st.session_state:
        st.session_state.agent = None
    if "parser" not in st.session_state:
        st.session_state.parser = None


def init_agent_and_parser():
    """Alusta agentti ja parseri."""
    api_key = get_api_key()
    if api_key:
        if st.session_state.agent is None:
            st.session_state.agent = create_agent(api_key)
        if st.session_state.parser is None:
            st.session_state.parser = create_parser(api_key)


def build_workbook_compat(all_days, num_days, workers):
    if build_workbook_and_report is None:
        return None
    wb, _ = build_workbook_and_report(all_days, num_days, workers)
    return wb


def parse_time(time_str: str):
    """Parsii ajan HH:MM tai HH.MM muodosta."""
    normalized = time_str.strip().replace(".", ":")
    if ":" in normalized:
        parts = normalized.split(":")
        if len(parts) != 2:
            raise ValueError(f"Virheellinen aika: {time_str}")
        h, m = int(parts[0]), int(parts[1])
    else:
        h, m = int(normalized), 0
    if not (0 <= h <= 23 and m in (0, 30)):
        raise ValueError("Ajan pitää olla muodossa HH:MM, HH.MM tai pelkkä tunti")
    return h, m


def parse_optional_time(label: str, key: str):
    """Ottaa käyttäjältä HH:MM ja palauttaa (hour, minute) tai (None, None)."""
    val = st.text_input(label, key=key).strip()
    if val == "":
        return None, None
    h, m = parse_time(val)
    return h, m


# ============================================================================
# PÄIVIEN DATA
# ============================================================================

def build_days_data(start_day: int, end_day: int, key_prefix: str):
    days = []
    for day in range(start_day, end_day + 1):
        with st.expander(f"Päivä {day}", expanded=(day == start_day)):
            row1_col1, row1_col2 = st.columns(2)
            row2_col1, row2_col2 = st.columns(2)

            with row1_col1:
                st.markdown("#### Tulo ja lähtö")
                arr_h, arr_m = parse_optional_time("Satamaan tuloaika (HH:MM)", key=f"{key_prefix}_arr_{day}")
                dep_h, dep_m = parse_optional_time("Satamasta lähtöaika (HH:MM)", key=f"{key_prefix}_dep_{day}")

            with row1_col2:
                st.markdown("#### Satamaoperaatiot")
                op_s_h, op_s_m = parse_optional_time("Satamaoperaation alku (HH:MM)", key=f"{key_prefix}_opstart_{day}")
                op_e_h, op_e_m = parse_optional_time("Satamaoperaation loppu (HH:MM)", key=f"{key_prefix}_opend_{day}")

            with row2_col1:
                st.markdown("#### Slussi")
                sluice_arr_h, sluice_arr_m = parse_optional_time("Slussi - tulo alku (HH:MM, kesto 2h)", key=f"{key_prefix}_sluice_arr_{day}")
                sluice_dep_h, sluice_dep_m = parse_optional_time("Slussi - lähtö alku (HH:MM, kesto 2h)", key=f"{key_prefix}_sluice_dep_{day}")

            with row2_col2:
                st.markdown("#### Shiftaus")
                shifting_h, shifting_m = parse_optional_time("Shiftaus alku (HH:MM, kesto 1h)", key=f"{key_prefix}_shifting_{day}")

            days.append({
                "arrival_hour": arr_h, "arrival_minute": arr_m or 0,
                "departure_hour": dep_h, "departure_minute": dep_m or 0,
                "port_op_start_hour": op_s_h, "port_op_start_minute": op_s_m or 0,
                "port_op_end_hour": op_e_h, "port_op_end_minute": op_e_m or 0,
                "sluice_arrival_hour": sluice_arr_h, "sluice_arrival_minute": sluice_arr_m or 0,
                "sluice_departure_hour": sluice_dep_h, "sluice_departure_minute": sluice_dep_m or 0,
                "shifting_hour": shifting_h, "shifting_minute": shifting_m or 0,
            })
    return days


# ============================================================================
# TAULUKOT JA NÄYTÖT
# ============================================================================

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

        for i, time_col in enumerate(DISPLAY_TIME_COLS):
            if sluice[i]:
                row[time_col] = "SL"
            elif shifting[i]:
                row[time_col] = "SH"
            elif arr[i]:
                row[time_col] = "B"
            elif dep[i]:
                row[time_col] = "C"
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


def create_editable_work_df(all_days, day_idx):
    data = {"Työntekijä": WORKERS}
    for col_idx, col in enumerate(DISPLAY_TIME_COLS):
        data[col] = [bool(all_days[w][day_idx]["work_slots"][col_idx]) for w in WORKERS]
    return pd.DataFrame(data)


def apply_edited_work_df(all_days, day_idx, edited_df):
    for _, row in edited_df.iterrows():
        worker = row["Työntekijä"]
        if worker not in all_days:
            continue
        all_days[worker][day_idx]["work_slots"] = [bool(row[t]) for t in DISPLAY_TIME_COLS]


def store_generated_result(wb, all_days, days_data, num_days):
    st.session_state.generated_wb = wb
    st.session_state.generated_all_days = all_days
    st.session_state.generated_days_data = days_data
    st.session_state.generated_num_days = num_days
    # Analysoi heti
    st.session_state.analysis = analyze_schedule(all_days, days_data)


# ============================================================================
# RENDER FUNKTIOT
# ============================================================================

def render_results(num_days, wb, all_days):
    st.subheader("📊 STCW-lepoaika-analyysi")
    if check_stcw_at_slot is None:
        st.info("STCW-analyysifunktio ei käytettävissä.")
    else:
        stcw_data = []
        for d in range(1, num_days):
            for w in WORKERS:
                prev = all_days[w][d - 1]["work_slots"]
                work = all_days[w][d]["work_slots"]
                ana = check_stcw_at_slot(prev + work, 95)
                stcw_data.append({
                    "Päivä": d + 1,
                    "Työntekijä": w,
                    "Työtunnit": sum(work) / 2,
                    "Lepo (h)": ana["total_rest"],
                    "Pisin lepo (h)": ana["longest_rest"],
                    "Status": "✓ OK" if ana["status"] == "OK" else "⚠ VAROITUS",
                })
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


def render_post_generation_editor():
    if st.session_state.generated_all_days is None:
        return

    st.markdown("## ✏️ Muokkaa vuoroja generoinnin jälkeen")
    st.caption("Klikkaa soluja muuttaaksesi työslotteja. Paina 'Generoi uudelleen' päivittääksesi.")

    num_days = st.session_state.generated_num_days
    all_days = st.session_state.generated_all_days

    with st.form("post_generation_edit_form"):
        edited_dfs = []
        for d in range(num_days):
            st.markdown(f"**Muokattava päivä {d+1}**")
            base_df = create_editable_work_df(all_days, d)
            edited_df = st.data_editor(
                base_df,
                hide_index=True,
                use_container_width=True,
                key=f"post_edit_day_{d}",
                disabled=["Työntekijä"],
                column_config={c: st.column_config.CheckboxColumn(c, default=False, width="small") for c in DISPLAY_TIME_COLS},
            )
            edited_dfs.append(edited_df)
        regenerate_clicked = st.form_submit_button("🔁 Generoi uudelleen (päivitä Excel)")

    if regenerate_clicked:
        updated_all_days = copy.deepcopy(st.session_state.generated_all_days)
        for d, edited_df in enumerate(edited_dfs):
            apply_edited_work_df(updated_all_days, d, edited_df)
        wb = build_workbook_compat(updated_all_days, num_days, WORKERS)
        if wb is None:
            st.error("Excelin uudelleenrakennus epäonnistui.")
        else:
            store_generated_result(wb, updated_all_days, st.session_state.generated_days_data, num_days)
            st.success("Vuorot päivitetty.")


# ============================================================================
# AI CHAT
# ============================================================================

def add_message(role: str, content: str):
    """Lisää viesti keskusteluun."""
    st.session_state.messages.append({"role": role, "content": content})
    if role == "user":
        st.session_state.message_count += 1


def process_user_message(user_input: str):
    """Käsittele käyttäjän viesti."""
    # Tarkista käyttöraja
    if st.session_state.message_count >= MAX_MESSAGES_PER_SESSION:
        add_message("user", user_input)
        add_message("assistant", f"⚠️ Käyttöraja ({MAX_MESSAGES_PER_SESSION} viestiä) saavutettu tälle sessiolle. Lataa sivu uudelleen aloittaaksesi uuden session.")
        return
    
    add_message("user", user_input)
    
    agent = st.session_state.agent
    parser = st.session_state.parser
    analysis = st.session_state.analysis
    
    if not agent or not agent.is_available():
        add_message("assistant", "⚠️ AI-toiminnot eivät ole käytettävissä.")
        return
    
    lower_input = user_input.lower()
    
    # Tarkista onko rajoite
    is_constraint = any(word in lower_input for word in [
        "ei voi", "ei saa", "max", "min", "vähintään", "enintään",
        "pakollinen", "vapaalla", "yövuoro", "iltavuoro"
    ])
    
    if is_constraint and parser:
        result = parser.parse(user_input)
        
        if result["understood"] and result["constraints"]:
            for c in result["constraints"]:
                parser.add_constraint(c)
            
            response = f"✅ Rajoite lisätty:\n"
            for c in result["constraints"]:
                response += f"  • {parser._describe_constraint(c)}\n"
            response += f"\nAktiivisia rajoitteita: {len(parser.get_constraints())}"
            response += "\n\n*Generoi vuorot uudelleen soveltaaksesi rajoitteita.*"
        else:
            response = f"En ymmärtänyt rajoitetta. {result.get('clarification_needed', '')}"
        
        add_message("assistant", response)
    
    elif "analys" in lower_input or "tarkista" in lower_input or "ongelm" in lower_input:
        if analysis:
            llm_data = get_analysis_for_llm(analysis)
            if llm_data["has_problems"]:
                response = agent.analyze_and_suggest(llm_data)
            else:
                response = "✅ Työvuoroissa ei havaittu ongelmia!"
        else:
            response = "Generoi ensin työvuorot ennen analyysiä."
        
        add_message("assistant", response)
    
    elif "yhteenveto" in lower_input or "tiivistä" in lower_input:
        if st.session_state.generated_all_days and st.session_state.generated_days_data:
            response = agent.get_schedule_summary(
                st.session_state.generated_all_days, 
                st.session_state.generated_days_data
            )
        else:
            response = "Generoi ensin työvuorot."
        
        add_message("assistant", response)
    
    elif "rajoite" in lower_input and ("näytä" in lower_input or "listaa" in lower_input or "mitä" in lower_input):
        if parser:
            response = parser.format_constraints()
        else:
            response = "Rajoite-parseri ei käytettävissä."
        
        add_message("assistant", response)
    
    elif "tyhjennä" in lower_input or "poista rajoitteet" in lower_input or "nollaa" in lower_input:
        if parser:
            parser.clear_constraints()
            response = "✅ Rajoitteet tyhjennetty."
        else:
            response = "Parseri ei käytettävissä."
        
        add_message("assistant", response)
    
    else:
        context = None
        if analysis:
            context = get_analysis_for_llm(analysis)
        
        response = agent.answer_question(user_input, context)
        add_message("assistant", response)


def render_analysis_tab():
    """Renderöi analyysivälilehti."""
    st.markdown("### 📊 Työvuoroanalyysi")
    
    if st.session_state.analysis is None:
        st.info("Generoi ensin työvuorot nähdäksesi analyysin.")
        return
    
    analysis = st.session_state.analysis
    summary = analysis['summary']
    
    # Yhteenveto
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ongelmat", summary['total_issues'])
    col2.metric("Varoitukset", summary['total_warnings'])
    col3.metric("STCW-rikkeet", summary['stcw_violations'])
    col4.metric("Op-aukot", summary['op_coverage_gaps'])
    
    # Yksityiskohdat
    if summary['total_issues'] > 0 or summary['total_warnings'] > 0:
        st.markdown("#### Havaitut ongelmat")
        
        for wa in analysis['worker_analyses']:
            if wa['issues'] or wa['warnings']:
                with st.expander(f"{wa['worker']} - Päivä {wa['day']} ({wa['hours']}h)"):
                    for issue in wa['issues']:
                        st.error(f"❌ {issue}")
                    for warning in wa['warnings']:
                        st.warning(f"⚠️ {warning}")
        
        # LLM-analyysi
        agent = st.session_state.agent
        if agent and agent.is_available():
            if st.button("🤖 Pyydä AI-korjausehdotuksia"):
                with st.spinner("Analysoidaan..."):
                    llm_data = get_analysis_for_llm(analysis)
                    response = agent.analyze_and_suggest(llm_data)
                st.markdown("#### AI-korjausehdotukset")
                st.markdown(response)
    else:
        st.success("✅ Ei ongelmia havaittu!")


def render_chat_tab():
    """Renderöi chat-välilehti."""
    st.markdown("### 💬 Keskustele AI-assistentin kanssa")
    st.caption(f"Voit kysyä vuoroista, pyytää analyysiä tai lisätä rajoitteita. ({st.session_state.message_count}/{MAX_MESSAGES_PER_SESSION} viestiä)")
    
    # Esimerkkejä
    with st.expander("💡 Esimerkkejä"):
        st.markdown("""
        **Rajoitteet:**
        - "EU ei voi tehdä yövuoroa"
        - "PH1 tekee max 8 tuntia"
        - "PH2 on vapaalla päivänä 2"
        
        **Kysymykset:**
        - "Analysoi vuorot"
        - "Miksi PH1:llä on vähän tunteja?"
        - "Tee yhteenveto"
        
        **Hallinta:**
        - "Näytä rajoitteet"
        - "Tyhjennä rajoitteet"
        """)
    
    # Keskusteluhistoria
    chat_container = st.container()
    with chat_container:
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
    
    # Syöte
    if prompt := st.chat_input("Kirjoita viesti..."):
        process_user_message(prompt)
        st.rerun()
    
    # Tyhjennä keskustelu
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("🗑️ Tyhjennä chat"):
            st.session_state.messages = []
            if st.session_state.agent:
                st.session_state.agent.clear_history()
            st.rerun()


# ============================================================================
# MAIN
# ============================================================================

def main():
    st.set_page_config(page_title="Sea Watch - Työvuorogeneraattori", layout="wide")
    
    init_session_state()
    init_agent_and_parser()
    
    # CSS
    st.markdown("""
        <style>
        [data-testid="stDataFrame"] > div {
            max-width: 100% !important;
            overflow-x: auto !important;
        }
        [data-testid="stDataFrame"] table {
            font-size: 10px !important;
        }
        [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
            padding: 2px 3px !important;
            min-width: 25px !important;
            max-width: 35px !important;
        }
        [data-testid="stDataFrame"] th:first-child, [data-testid="stDataFrame"] td:first-child {
            min-width: 80px !important;
            max-width: 100px !important;
        }
        [data-testid="stDataEditor"] {
            font-size: 10px !important;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Sivupalkki
    with st.sidebar:
        st.title("⚙️ Asetukset")
        
        num_days = st.number_input("Päivien määrä", min_value=1, max_value=14, value=2, step=1)
        st.info("Jätä kenttä tyhjäksi jos tapahtumaa ei ole.")
        
        st.divider()
        
        # AI-status
        if st.session_state.agent and st.session_state.agent.is_available():
            st.success("🤖 AI käytettävissä")
        else:
            st.warning("🤖 AI ei käytettävissä")
        
        st.divider()
        
        # Aktiiviset rajoitteet
        st.markdown("**Aktiiviset rajoitteet:**")
        if st.session_state.parser:
            constraints = st.session_state.parser.get_constraints()
            if constraints:
                for i, c in enumerate(constraints):
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.caption(st.session_state.parser._describe_constraint(c))
                    with col2:
                        if st.button("🗑️", key=f"del_constraint_{i}"):
                            st.session_state.parser.remove_constraint(i)
                            st.rerun()
            else:
                st.caption("Ei rajoitteita")
            
            if constraints and st.button("Tyhjennä kaikki rajoitteet"):
                st.session_state.parser.clear_constraints()
                st.rerun()
    
    # Otsikko
    st.title("🛳️ Sea Watch - Työvuorolistageneraattori")
    
    # Välilehdet
    tab_auto, tab_manual, tab_analysis, tab_chat = st.tabs([
        "📝 Automaattinen syöttö", 
        "✋ Päivä 1 manuaalinen",
        "📊 Analyysi",
        "💬 AII Chat"
    ])
    
    # Tab 1: Automaattinen syöttö
    with tab_auto:
        st.markdown("Syötä kaikkien päivien tulo-/lähtöajat ja satamaoperaatiot.")
        days_data = build_days_data(1, num_days, key_prefix="auto")
        
        if st.button("🚀 Generoi työvuorot", key="gen_auto", type="primary"):
            # Hae rajoitteet
            constraints = []
            if st.session_state.parser:
                constraints = st.session_state.parser.get_constraints()
            
            with st.spinner("Generoidaan..."):
                wb, all_days, _ = generate_schedule(days_data, constraints=constraints)
                store_generated_result(wb, all_days, days_data, num_days)
            
            if constraints:
                st.success(f"✅ Työvuorot generoitu {len(constraints)} rajoitteella!")
            else:
                st.success("✅ Työvuorot generoitu!")
        
        # Näytä tulokset
        if st.session_state.generated_all_days is not None:
            render_results(
                st.session_state.generated_num_days,
                st.session_state.generated_wb,
                st.session_state.generated_all_days
            )
            render_post_generation_editor()
    
    # Tab 2: Manuaalinen päivä 1
    with tab_manual:
        st.markdown("Maalaa päivän 1 työtunnit. Päivät 2+ syötetään normaalisti.")
        manual_default = init_manual_day1_df()
        manual_df = st.data_editor(
            manual_default, hide_index=True, use_container_width=True,
            key="manual_day1_editor", disabled=["Työntekijä"],
            column_config={c: st.column_config.CheckboxColumn(c, default=False) for c in TIME_COLS},
        )
        
        if num_days >= 2:
            st.markdown("#### Päivät 2+")
            days_data_rest = build_days_data(2, num_days, key_prefix="manual")
        else:
            days_data_rest = []
        
        if st.button("🚀 Generoi työvuorot (manuaalinen päivä 1)", key="gen_manual"):
            day1_placeholder = {
                "arrival_hour": None, "arrival_minute": 0,
                "departure_hour": None, "departure_minute": 0,
                "port_op_start_hour": 8, "port_op_start_minute": 0,
                "port_op_end_hour": 17, "port_op_end_minute": 0,
                "sluice_arrival_hour": None, "sluice_arrival_minute": 0,
                "sluice_departure_hour": None, "sluice_departure_minute": 0,
                "shifting_hour": None, "shifting_minute": 0,
            }
            days_data = [day1_placeholder] + days_data_rest
            manual_slots = convert_manual_df_to_slots(manual_df)
            
            # Hae rajoitteet
            constraints = []
            if st.session_state.parser:
                constraints = st.session_state.parser.get_constraints()
            
            if generate_schedule_with_manual_day1 is None:
                st.warning("Manuaalinen päivä 1 ei tuettu. Käytetään automaattista.")
                wb, all_days, _ = generate_schedule(days_data, constraints=constraints)
            else:
                wb, all_days, _ = generate_schedule_with_manual_day1(days_data, manual_slots)
            
            store_generated_result(wb, all_days, days_data, num_days)
            st.success("✅ Työvuorot generoitu!")
    
    # Tab 3: Analyysi
    with tab_analysis:
        render_analysis_tab()
    
    # Tab 4: AI Chat
    with tab_chat:
        render_chat_tab()


if __name__ == "__main__":
    main()
