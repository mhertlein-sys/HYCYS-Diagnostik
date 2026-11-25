import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
import datetime
import pwlf
from PIL import Image
import os
import plotly.graph_objects as go
import tempfile
from fpdf import FPDF
import openpyxl

# --- GLOBAL TIME & SPRINT PARSING HELPERS ---
def parse_time_to_seconds(t):
    if pd.isnull(t): return None
    if isinstance(t, datetime.time): return t.hour * 3600 + t.minute * 60 + t.second
    if isinstance(t, str):
        try:
            parts = t.strip().split(':')
            if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + float(parts[2])
            if len(parts) == 2: return int(parts[0])*60 + float(parts[1])
        except: pass
    if isinstance(t, (int, float)):
        if 0 < t < 1: return t * 86400 
        return float(t)
    return 0.0

def parse_sprint_excel(file_contents):
    wb = openpyxl.load_workbook(file_contents, data_only=True)
    sheet_name = "Data Sprint" if "Data Sprint" in wb.sheetnames else wb.sheetnames[0]
    sheet = wb[sheet_name]
    
    sprint_powers = []
    sprint_times = []
    sprint_cadences = []
    
    for r in range(1, sheet.max_row + 1):
        val_a = sheet.cell(row=r, column=1).value
        val_c = sheet.cell(row=r, column=3).value
        val_e = sheet.cell(row=r, column=5).value
        
        if val_a is not None and val_e is not None:
            try:
                p = float(val_a)
                t = parse_time_to_seconds(val_e)
                c = float(val_c) if val_c is not None else 0.0
                if t is not None:
                    sprint_powers.append(p)
                    sprint_times.append(t)
                    sprint_cadences.append(c)
            except ValueError:
                continue

    # Führende und nachfolgende Null-Leistungswerte entfernen und Zeit ab 0.00s starten
    if sprint_powers:
        first_idx = next((i for i, p in enumerate(sprint_powers) if p > 0), None)
        last_idx = next((i for i in range(len(sprint_powers) - 1, -1, -1) if sprint_powers[i] > 0), None)
        if first_idx is not None and last_idx is not None and last_idx >= first_idx:
            sprint_powers = sprint_powers[first_idx: last_idx + 1]
            sprint_times = sprint_times[first_idx: last_idx + 1]
            sprint_cadences = sprint_cadences[first_idx: last_idx + 1]
            t0 = sprint_times[0]
            sprint_times = [round(t - t0, 2) for t in sprint_times]
                
    return sprint_powers, sprint_times, sprint_cadences


SPRINT_PLATEAU_MIN_POINTS = 8


def detect_sprint_end(sprint_powers, sprint_times, min_points=SPRINT_PLATEAU_MIN_POINTS):
    """
    Ermittelt die gewertete Sprint-Gesamtdauer (t_bel).

    Bleibt am Ende des Sprints dieselbe Leistung über MEHR als `min_points`
    Datenpunkte konstant (Ergometer liefert keine neuen Werte mehr), wird dieses
    Plateau nicht mehr zur Sprintdauer gezählt. t_bel endet dann beim ersten
    Datenpunkt des Plateaus. Die Rohdaten selbst bleiben unverändert.

    Returns: (t_bel, plateau_start_idx)  -- plateau_start_idx ist None, wenn kein Plateau erkannt wurde.
    """
    if not sprint_powers or not sprint_times:
        return 0.0, None
    n = len(sprint_powers)
    last_p = sprint_powers[-1]
    start_idx = n - 1
    while start_idx > 0 and sprint_powers[start_idx - 1] == last_p:
        start_idx -= 1
    run_len = n - start_idx
    if run_len > min_points and start_idx > 0:
        return float(sprint_times[start_idx]), start_idx
    return float(max(sprint_times)), None


# Einheitliche Schriftart für alle Matplotlib-Grafiken (Windows lokal & Linux auf Streamlit Cloud)
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = ["Arial", "Liberation Sans", "DejaVu Sans", "Helvetica", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False

# Import PDF generator from dedicated module
from pdf_rad import create_pdf_rad


# 1. APP-LAYOUT & CSS
st.set_page_config(page_title="HYCYS - Diagnostik", layout="wide")

st.markdown("""
    <style>
    /* HYCYS Primary Buttons ("Auswertung starten", etc.) */
    button[kind="primary"],
    .stButton > button[kind="primary"],
    div[data-testid="stFormSubmitButton"] > button {
        background-color: #2cb7b9 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
    }
    button[kind="primary"]:hover,
    .stButton > button[kind="primary"]:hover,
    div[data-testid="stFormSubmitButton"] > button:hover {
        background-color: #259fa0 !important;
        color: white !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #2cb7b9 !important;
        color: white !important;
        border: none !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #259fa0 !important;
        color: white !important;
    }
    /* Schieberegler (st.slider) in HYCYS Türkis (#2cb7b9) statt Standard-Rot */
    div[data-baseweb="slider"] div[role="slider"] {
        background-color: #2cb7b9 !important;
        border-color: #2cb7b9 !important;
        box-shadow: 0 0 0 0.2rem rgba(44, 183, 185, 0.25) !important;
    }
    div[data-baseweb="slider"] div[data-testid="stThumbValue"],
    div[data-baseweb="slider"] div[role="slider"] > div {
        color: #2cb7b9 !important;
        font-weight: 700 !important;
    }
    div[data-baseweb="slider"] > div > div > div:first-child,
    div[data-baseweb="slider"] [data-testid="stSliderTickBar"] + div,
    div[data-baseweb="slider"] div[style*="background-color: rgb(255, 75, 75)"] {
        background: #2cb7b9 !important;
        background-color: #2cb7b9 !important;
    }
    /* Alle Metrik-Deltas (Potenzial-Änderungen) in HYCYS Türkis statt Rot/Grün */
    [data-testid="stMetricDelta"] svg {
        fill: #2cb7b9 !important;
    }
    [data-testid="stMetricDelta"] > div {
        color: #2cb7b9 !important;
        font-weight: 700 !important;
    }
    /* Radio Buttons (z. B. LTP2 vs OBLA 4.0, Ansichtsauswahl) in HYCYS Türkis (#2cb7b9) statt Rot */
    div[data-testid="stRadio"] div[role="radiogroup"] label div:first-child[aria-checked="true"],
    div[data-testid="stRadio"] label div[role="radio"][aria-checked="true"],
    div[data-baseweb="radio"] input:checked + div,
    div[data-baseweb="radio"] div[aria-checked="true"] {
        border-color: #2cb7b9 !important;
        background-color: #2cb7b9 !important;
    }
    div[data-baseweb="radio"] input:checked + div > div,
    div[data-testid="stRadio"] div[role="radiogroup"] label div:first-child[aria-checked="true"] > div {
        background-color: #ffffff !important;
    }
    div[data-baseweb="radio"]:hover input + div,
    div[data-testid="stRadio"] label:hover div[role="radio"] {
        border-color: #2cb7b9 !important;
    }
    div[data-baseweb="radio"] input:focus + div {
        box-shadow: 0 0 0 0.2rem rgba(44, 183, 185, 0.25) !important;
    }
    /* Linke Sidebar in der Defaultansicht ca. 1/5 breiter (ca. 400px statt 336px) */
    section[data-testid="stSidebar"][aria-expanded="true"] {
        width: 400px !important;
        min-width: 400px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- COACH-ACCOUNTS, PASSWORTSCHUTZ & SITZUNGS-PERSISTENZ (RELOAD-SICHER + REMOTE-SUPPORT) ---
import time
import uuid

SESSION_INACTIVITY_TIMEOUT_SEC = 12 * 3600  # 12 Stunden Inaktivitäts-Timeout

# Alle freigeschalteten HYCYS-Coaches (mit Standard-Passwörtern, jederzeit über st.secrets["coach_passwords"] überschreibbar)
DEFAULT_COACH_ACCOUNTS = {
    "Markus Hertlein":   "Markus2026",
    "Marius Trompetter": "MariusT2026",
    "Susanne Traser":    "Susanne2026",
    "Manuel Kuhnle":     "Manuel2026",
    "Jean Surmont":      "Jean2026",
    "Hosea Frick":       "Hosea2026",
    "Björn Geesmann":    "Bjoern2026",
    "Gregor Eichhorn":   "Gregor2026",
    "Marius Rauen":      "MariusR2026",
    "Florian Fichtl":    "Florian2026",
}
HYCYS_COACH_LIST = list(DEFAULT_COACH_ACCOUNTS.keys())
ADMIN_COACHES = {"Markus Hertlein"}


def _get_coach_passwords() -> dict:
    """Liest die Coach-Passwörter aus DEFAULT_COACH_ACCOUNTS und überschreibt sie optional mit st.secrets['coach_passwords']."""
    pw_map = dict(DEFAULT_COACH_ACCOUNTS)
    try:
        if "coach_passwords" in st.secrets:
            for c_name, c_pw in st.secrets["coach_passwords"].items():
                pw_map[str(c_name)] = str(c_pw)
    except Exception:
        pass
    return pw_map


def _get_master_password() -> str:
    expected_pw = "HYCYS2026"
    try:
        if "app_password" in st.secrets:
            expected_pw = str(st.secrets["app_password"])
    except Exception:
        pass
    return expected_pw


@st.cache_resource
def _get_persistent_session_store() -> dict:
    """Serverweiter In-Memory-Speicher für aktive Sitzungen (überlebt F5 / Browser-Reload und ermöglicht Remote-Support)."""
    return {}


_PERSIST_SKIP_PREFIXES = (
    "FormSubmitter:",
    "btn_",
    "_pending_",
    "remote_support_",
)
_PERSIST_SKIP_KEYS = {
    "uploaded_file_key",
    "uploaded_ramp_key",
    "sprint_file_key",
    "spiro_rampe_key",
    "fit_rampe_key",
    "lauftest_lac_editor_key",
    "edited_sf_key",
    "edited_sf_rad_key",
    "sprint_lac_editor_key",
    "hycys_login_form",
}


def save_current_session_snapshot():
    """Speichert den aktuellen Arbeitsstand unter der URL-Session-ID (`?sid=...`), damit bei F5 nichts verloren geht."""
    sid = st.session_state.get("_hycys_sid")
    if not sid or not st.session_state.get("hycys_password_authenticated", False):
        return
    store = _get_persistent_session_store()
    now_ts = time.time()
    logged_coach = st.session_state.get("logged_in_coach", st.session_state.get("coach", "Markus Hertlein"))
    data_snap = {}
    for k, v in st.session_state.items():
        if k in _PERSIST_SKIP_KEYS or any(str(k).startswith(p) for p in _PERSIST_SKIP_PREFIXES):
            continue
        if isinstance(v, (int, float, str, bool, list, dict, tuple, pd.DataFrame, np.ndarray, np.generic)) or v is None:
            data_snap[k] = v
    store[sid] = {
        "last_active": now_ts,
        "authenticated": True,
        "coach_user": logged_coach,
        "athlete_name": str(st.session_state.get("athlete_name", "")),
        "sportart": str(st.session_state.get("sportart", "")),
        "data": data_snap,
    }


def load_session_from_store(target_sid: str, keep_logged_in_coach: str = None) -> bool:
    """Lädt eine bestehende Sitzung aus dem Server-Speicher in den aktuellen Browser-Tab (für F5-Reload oder Remote-Support)."""
    store = _get_persistent_session_store()
    if target_sid not in store:
        return False
    entry = store[target_sid]
    entry["last_active"] = time.time()
    for k_cur in list(st.session_state.keys()):
        if k_cur not in _PERSIST_SKIP_KEYS and not any(str(k_cur).startswith(p) for p in _PERSIST_SKIP_PREFIXES):
            try:
                del st.session_state[k_cur]
            except Exception:
                pass
    for k_saved, v_saved in entry.get("data", {}).items():
        st.session_state[k_saved] = v_saved
    if keep_logged_in_coach:
        st.session_state["logged_in_coach"] = keep_logged_in_coach
    st.session_state["_hycys_sid"] = target_sid
    st.session_state["hycys_password_authenticated"] = True
    try:
        st.query_params["sid"] = target_sid
    except Exception:
        pass
    return True


def check_app_password() -> bool:
    """
    Prüft den Coach-Account & das Passwort und stellt bei einem Seiten-Reload (F5) den kompletten Arbeitsstand
    anhand des URL-Parameters `?sid=...` wieder her, solange die Inaktivitätszeit (12h) nicht abgelaufen ist.
    """
    store = _get_persistent_session_store()
    now_ts = time.time()

    # Abgelaufene Sitzungen aufräumen
    for k_sid in list(store.keys()):
        if now_ts - store[k_sid].get("last_active", 0) > SESSION_INACTIVITY_TIMEOUT_SEC:
            del store[k_sid]

    # URL-Parameter ?sid=... auslesen
    sid_from_url = None
    try:
        sid_from_url = st.query_params.get("sid")
    except Exception:
        pass

    # Falls der Tab neu geladen wurde (st.session_state ist leer), aber eine gültige sid in der URL steht:
    if not st.session_state.get("hycys_password_authenticated", False) and sid_from_url and sid_from_url in store:
        entry = store[sid_from_url]
        if entry.get("authenticated") and (now_ts - entry.get("last_active", 0) <= SESSION_INACTIVITY_TIMEOUT_SEC):
            load_session_from_store(sid_from_url)
            return True

    # Falls bereits in dieser Tab-Session eingeloggt: Inaktivitäts-Timeout prüfen & Timestamp aktualisieren
    if st.session_state.get("hycys_password_authenticated", False):
        sid = st.session_state.get("_hycys_sid") or sid_from_url
        if not sid:
            sid = uuid.uuid4().hex[:16]
            st.session_state["_hycys_sid"] = sid
        try:
            if st.query_params.get("sid") != sid:
                st.query_params["sid"] = sid
        except Exception:
            pass
        if sid in store and (now_ts - store[sid].get("last_active", now_ts) > SESSION_INACTIVITY_TIMEOUT_SEC):
            del store[sid]
            st.session_state["hycys_password_authenticated"] = False
        else:
            save_current_session_snapshot()
            return True

    pw_map = _get_coach_passwords()
    master_pw = _get_master_password()
    coach_names = list(pw_map.keys())

    _, col_pw, _ = st.columns([1.0, 1.25, 1.0])
    with col_pw:
        st.markdown("<div style='margin-top: 50px;'></div>", unsafe_allow_html=True)
        if os.path.exists("pdf_assets/page1_img3.png"):
            st.image("pdf_assets/page1_img3.png", use_container_width=True)
        else:
            st.markdown("<h2 style='text-align:center; color:#595a59;'>HYCYS Diagnostik</h2>", unsafe_allow_html=True)

        st.markdown(
            "<div style='text-align:center; color:#595a59; font-weight:600; margin-bottom:12px;'>"
            "Bitte Coach-Account auswählen und Passwort eingeben:"
            "</div>",
            unsafe_allow_html=True,
        )
        with st.form("hycys_login_form", clear_on_submit=False):
            selected_coach_login = st.selectbox("Coach Account", options=coach_names)
            pw_input = st.text_input("Passwort", type="password", placeholder="Passwort eingeben...")
            resume_existing = st.checkbox("Letzte aktive Sitzung dieses Coaches fortsetzen (falls vorhanden)", value=True)
            login_submitted = st.form_submit_button("Anmelden", type="primary", use_container_width=True)
            if login_submitted:
                expected_coach_pw = pw_map.get(selected_coach_login, master_pw)
                if pw_input.strip() in (expected_coach_pw, master_pw):
                    # Prüfen, ob für diesen Coach bereits eine aktive Sitzung existiert
                    existing_coach_sid = None
                    if resume_existing:
                        best_ts = 0.0
                        for k_s, v_s in store.items():
                            if v_s.get("coach_user") == selected_coach_login and v_s.get("last_active", 0) > best_ts:
                                best_ts = v_s.get("last_active", 0)
                                existing_coach_sid = k_s

                    if existing_coach_sid:
                        load_session_from_store(existing_coach_sid, keep_logged_in_coach=selected_coach_login)
                    else:
                        new_sid = uuid.uuid4().hex[:16]
                        st.session_state["_hycys_sid"] = new_sid
                        st.session_state["hycys_password_authenticated"] = True
                        st.session_state["logged_in_coach"] = selected_coach_login
                        st.session_state["coach"] = selected_coach_login
                        st.session_state["coach_lauft_key"] = selected_coach_login
                        st.session_state["coach_lauft_widget"] = selected_coach_login
                        try:
                            st.query_params["sid"] = new_sid
                        except Exception:
                            pass
                        save_current_session_snapshot()
                    st.rerun()
                else:
                    st.error("❌ Falsches Passwort. Bitte erneut versuchen.")
    return False


if not check_app_password():
    st.stop()

# Initialize session state for metadata
metadata_defaults = [
    ("athlete_name", "Michael Wagner"), 
    ("birthdate", "06.03.1996"), 
    ("coach", "Markus Hertlein"), 
    ("test_date", "22.05.2026"), 
    ("sportart", "Bitte wählen..."), 
    ("kategorie", "Amateur"), 
    ("height", 178), 
    ("weight", 72.3), 
    ("body_fat_pct", 0.0), 
    ("last_uploaded_file", None),
    ("last_test_type", "HYCYS Standart (5min / +0,4m/s)"),
    ("start_v", 2.5),
    ("v_increment", 0.4),
    ("anzahl", 5),
    ("vorlauf", 60),
    ("stufendauer", 5.0),
    ("pausendauer", 30),
    ("ppo_val", 436.0),
    ("hfmax_val", 179.0),
    ("slope_val", 0.11),
    ("intercept_val", 138.0),
    ("carb_intake_factor", 1.0),
    ("vo2_override", 4368.0),
    ("t_bel_manual", 13.2),
    ("t_alak_manual", 3.1),
    ("diagnostik_type", "Running BLUE"),
    ("gender", "männlich"),
    ("lauf_auswertung_gestartet", False)
]
for key, val in metadata_defaults:
    if key not in st.session_state:
        st.session_state[key] = val

def update_sf_rad():
    changes = st.session_state.get("edited_sf_rad_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_sf.loc[int(row_idx), col_name] = new_val

def update_sf_lauf():
    changes = st.session_state.get("edited_sf_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_sf.loc[int(row_idx), col_name] = new_val

def update_sprint_lac():
    changes = st.session_state.get("sprint_lac_editor_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_sprint_lac.loc[int(row_idx), col_name] = new_val

def update_lauftest_lac():
    changes = st.session_state.get("lauftest_lac_editor_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_lauftest_input.loc[int(row_idx), col_name] = new_val

def format_pace_mmss(v_ms):
    try:
        v = float(v_ms)
    except (ValueError, TypeError):
        return "-"
    if pd.isna(v) or v <= 0:
        return "-"
    sec = 1000.0 / v
    m = int(sec // 60)
    s = int(round(sec % 60))
    if s == 60:
        m += 1
        s = 0
    return f"{m:02d}:{s:02d}"


DEFAULT_GDRIVE_ROOT_LINK = "https://drive.google.com/drive/folders/1Deg3OVDz19GLBervdQprPCZ-8MIkVGD5"


def format_hycys_pdf_filename(diagnostik_type: str, athlete_name: str, test_date: str) -> str:
    """
    Erzeugt den standardisierten HYCYS-Dateinamen für das PDF:
    'ArtderDiagnostik_Vorname Nachname_Datum.pdf'
    z.B. 'CyclingBLUE_Markus Hertlein_13.03.2025.pdf' oder 'RunningBLUE_Christoph Muxel_01.10.2026.pdf'
    """
    diag_clean = str(diagnostik_type or "Diagnostik").replace(" ", "").strip()
    name_clean = " ".join(str(athlete_name or "Athlet").strip().split())
    date_clean = str(test_date or "").strip()
    return f"{diag_clean}_{name_clean}_{date_clean}.pdf"


def format_hycys_excel_filename(diagnostik_type: str, athlete_name: str, test_date: str) -> str:
    """
    Erzeugt den standardisierten HYCYS-Dateinamen für die Excel-Rohdaten-Auswertung:
    'ArtderDiagnostik_Vorname Nachname_Datum.xlsx'
    z.B. 'CyclingBLUE_Markus Hertlein_13.03.2025.xlsx'
    """
    diag_clean = str(diagnostik_type or "Diagnostik").replace(" ", "").strip()
    name_clean = " ".join(str(athlete_name or "Athlet").strip().split())
    date_clean = str(test_date or "").strip()
    return f"{diag_clean}_{name_clean}_{date_clean}.xlsx"


def format_athlete_folder_name(athlete_name: str) -> str:
    """
    Erzeugt den Athleten-Ordnernamen im Format 'Nachname_Vorname'
    aus dem Athletennamen (z.B. 'Markus Hertlein' -> 'Hertlein_Markus').
    """
    parts = [p.strip() for p in str(athlete_name or "").strip().split() if p.strip()]
    if len(parts) >= 2:
        vorname = "_".join(parts[:-1])
        nachname = parts[-1]
        return f"{nachname}_{vorname}"
    elif len(parts) == 1:
        return parts[0]
    return "Unbekannt"


def _parse_date_str_to_dt(date_str: str):
    """Konvertiert Datums-Strings wie '13.03.2025' oder '2025-03-13' in ein datetime-Objekt für Excel."""
    import datetime as _dt
    s = str(date_str or "").strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y", "%d/%m/%Y"):
        try:
            return _dt.datetime.strptime(s, fmt)
        except Exception:
            pass
    return s


def _split_vorname_nachname(athlete_name: str):
    """Teilt 'Vorname Nachname' in (vorname, nachname) auf."""
    parts = [p.strip() for p in str(athlete_name or "").strip().split() if p.strip()]
    if len(parts) >= 2:
        return " ".join(parts[:-1]), parts[-1]
    elif len(parts) == 1:
        return "", parts[0]
    return "", "Unbekannt"


def _seconds_to_excel_time(sec_val):
    """Konvertiert Sekunden in ein datetime.time Objekt für die HYCYS-Excel-Vorlagen."""
    import datetime as _dt
    try:
        total_sec = int(round(float(sec_val)))
        if total_sec < 0:
            total_sec = 0
        h = (total_sec // 3600) % 24
        m = (total_sec % 3600) // 60
        s = total_sec % 60
        return _dt.time(h, m, s)
    except Exception:
        return sec_val


def create_hycys_excel_lauf(
    athlete_name: str,
    birthdate: str,
    test_date: str,
    gender_str: str,
    diagnostik_type: str,
    coach: str,
    sportart: str,
    kategorie: str,
    height: float,
    weight: float,
    body_fat_pct: float,
    df_sf: pd.DataFrame,
    ruhe_lac: float,
    ruhe_hr: float,
    speed: list,
    lactate: list,
    hr: list,
    abs_vo2max: float,
    vo2_steady_abs: list,
    vco2_steady_abs: list,
    uploaded_file=None,
    uploaded_ramp=None,
) -> bytes:
    """
    Befüllt die originale 'HYCYS - Auswertung Lauf.xlsx' Vorlage mit allen in der App
    eingegebenen und hochgeladenen Rohdaten und gibt die fertige .xlsx-Datei als Bytes zurück.
    """
    import openpyxl
    import io

    # Robuste Template-Suche (Hauptordner oder Unterordner, Groß-/Kleinschreibung)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "HYCYS - Auswertung Lauf.xlsx"),
        os.path.join(base_dir, "Lauf Auswertungen", "HYCYS - Auswertung Lauf.xlsx"),
        os.path.join(base_dir, "HYCYS_Auswertung_Lauf.xlsx"),
        os.path.join(base_dir, "HYCYS-Auswertung Lauf.xlsx"),
    ]
    template_path = None
    for c in candidates:
        if os.path.isfile(c):
            template_path = c
            break

    if not template_path:
        # Fallback: Ordner nach passender .xlsx durchsuchen
        for root, dirs, files in os.walk(base_dir):
            for f in files:
                if "hycys" in f.lower() and "lauf" in f.lower() and f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    template_path = os.path.join(root, f)
                    break
            if template_path:
                break

    if not template_path or not os.path.isfile(template_path):
        raise FileNotFoundError(
            f"Vorlage 'HYCYS - Auswertung Lauf.xlsx' wurde auf dem Server nicht gefunden. "
            f"Gesucht in: {candidates[0]}"
        )

    wb = openpyxl.load_workbook(template_path)
    ws = wb["Eingabedaten"]

    vorname, nachname = _split_vorname_nachname(athlete_name)
    gender_norm = str(gender_str or "männlich").strip().lower()
    gender_code = "w" if ("weib" in gender_norm or "frau" in gender_norm or gender_norm == "w") else "m"

    ws["B5"] = str(coach or "Markus Hertlein")
    ws["B6"] = _parse_date_str_to_dt(test_date)
    ws["B9"] = nachname
    ws["B10"] = vorname
    ws["B11"] = float(height) if height else 180
    ws["B12"] = float(weight) if weight else 75.0
    ws["B15"] = _parse_date_str_to_dt(birthdate)
    ws["B17"] = gender_code
    ws["B19"] = str(sportart or "Laufen")
    ws["B20"] = str(kategorie or "Amateur")
    diag_str = str(diagnostik_type or "Running BLUE").strip()
    ws["B21"] = diag_str if diag_str.upper().startswith("HYCYS") else f"HYCYS {diag_str}"

    # Parizkova 10-Falten (B28:D37) & Körperfett (G31)
    sf_sum = 0.0
    if df_sf is not None and not df_sf.empty:
        for idx in range(min(10, len(df_sf))):
            r_excel = 28 + idx
            m1 = float(pd.to_numeric(df_sf.iloc[idx].get("M1", 0.0), errors="coerce") or 0.0)
            m2 = float(pd.to_numeric(df_sf.iloc[idx].get("M2", 0.0), errors="coerce") or 0.0)
            m3 = float(pd.to_numeric(df_sf.iloc[idx].get("M3", 0.0), errors="coerce") or 0.0)
            ws.cell(r_excel, 2, m1)
            ws.cell(r_excel, 3, m2)
            ws.cell(r_excel, 4, m3)
            sf_sum += (m1 + m2 + m3) / 3.0
    if sf_sum > 0:
        ws["G31"] = "=(22.32*LOG10(G28)-29.2)"
    else:
        ws["G31"] = float(body_fat_pct)

    # HFmax (M31)
    valid_hrs = [float(h) for h in hr if pd.notna(h) and float(h) > 0]
    if valid_hrs:
        ws["M31"] = int(round(max(valid_hrs)))

    # Stufentest Laktat & HF (A42:C50)
    ws["B42"] = float(ruhe_lac) if pd.notna(ruhe_lac) else 0.8
    ws["C42"] = int(round(float(ruhe_hr))) if pd.notna(ruhe_hr) and float(ruhe_hr) > 0 else 70

    for idx in range(8):
        r_stage = 43 + idx
        if idx < len(speed):
            ws.cell(r_stage, 1, float(speed[idx]))
            ws.cell(r_stage, 2, float(lactate[idx]) if idx < len(lactate) and pd.notna(lactate[idx]) else None)
            ws.cell(r_stage, 3, int(round(float(hr[idx]))) if idx < len(hr) and pd.notna(hr[idx]) else None)
        else:
            ws.cell(r_stage, 1, None)
            ws.cell(r_stage, 2, None)
            ws.cell(r_stage, 3, None)

    # Spiro-Rohdaten Stufentest (B57:G...) & Rampentest (K57:P...)
    for r_clear in range(57, 300):
        for c_clear in (2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16):
            ws.cell(r_clear, c_clear, None)

    def _write_spiro_block(file_obj, start_col: int):
        if file_obj is None:
            return 0
        try:
            file_obj.seek(0)
            df_raw = (
                pd.read_csv(file_obj, header=None, low_memory=False)
                if str(file_obj.name).lower().endswith(".csv")
                else pd.read_excel(file_obj, header=None)
            )
            file_obj.seek(0)
            written = 0
            for r_idx in range(3, len(df_raw)):
                if df_raw.shape[1] <= 14:
                    break
                t_raw = df_raw.iloc[r_idx, 9]
                if pd.isna(t_raw):
                    continue
                t_sec = parse_time_to_seconds(t_raw)
                if pd.isna(t_sec):
                    continue
                target_row = 57 + written
                ws.cell(target_row, start_col, _seconds_to_excel_time(t_sec))
                for offset in range(1, 6):
                    val = pd.to_numeric(df_raw.iloc[r_idx, 9 + offset], errors="coerce")
                    if pd.notna(val):
                        ws.cell(target_row, start_col + offset, float(val))
                written += 1
            return written
        except Exception:
            return 0

    _write_spiro_block(uploaded_file, 2)   # Spalten B..G (Stufentest)
    _write_spiro_block(uploaded_ramp, 11)  # Spalten K..P (Rampentest)

    # Über 30s gemittelte VO2max [ml/min] direkt in V56 hinterlegen
    if abs_vo2max and float(abs_vo2max) > 0:
        ws["V56"] = float(abs_vo2max)

    # Steady-State VO2 & VCO2 je Stufe in 'Calculation Spirometrie' (M8:N15) übertragen,
    # damit 'Calculation Physiologie' exakt mit den App-Stufenfenstern übereinstimmt
    if "Calculation Spirometrie" in wb.sheetnames and vo2_steady_abs:
        ws_cs = wb["Calculation Spirometrie"]
        for idx in range(8):
            r_cs = 8 + idx
            if idx < len(vo2_steady_abs) and vo2_steady_abs[idx] > 0:
                ws_cs.cell(r_cs, 13, float(vo2_steady_abs[idx]))  # Spalte M: VO2
                vco2_v = float(vco2_steady_abs[idx]) if idx < len(vco2_steady_abs) and vco2_steady_abs[idx] > 0 else 0.0
                if vco2_v > 0:
                    ws_cs.cell(r_cs, 14, vco2_v)  # Spalte N: VCO2
            else:
                ws_cs.cell(r_cs, 13, None)
                ws_cs.cell(r_cs, 14, None)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()


def create_hycys_excel_rad(
    athlete_name: str,
    birthdate: str,
    test_date: str,
    diagnostik_type: str,
    coach: str,
    kategorie: str,
    height: float,
    weight: float,
    body_fat_pct: float,
    df_sf: pd.DataFrame,
    df_sprint_lac: pd.DataFrame,
    ppo_val: float,
    abs_vo2max: float,
    t_alak: float,
    t_bel: float,
    slope_val: float,
    intercept_val: float,
    hfmax_val: float,
    sprint_times: list = None,
    sprint_powers: list = None,
    sprint_cadences: list = None,
    uploaded_spiro=None,
) -> bytes:
    """
    Befüllt die originale 'HYCYS - Auswertung Rad.xlsx' Vorlage mit allen in der App
    eingegebenen und hochgeladenen Rohdaten und gibt die fertige .xlsx-Datei als Bytes zurück.
    """
    import openpyxl
    import io

    # Robuste Template-Suche für Radsport-Vorlage
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "Radsport Auswertungen", "HYCYS - Auswertung Rad.xlsx"),
        os.path.join(base_dir, "HYCYS - Auswertung Rad.xlsx"),
        os.path.join(base_dir, "HYCYS_Auswertung_Rad.xlsx"),
        os.path.join(base_dir, "HYCYS-Auswertung Rad.xlsx"),
    ]
    template_path = None
    for c in candidates:
        if os.path.isfile(c):
            template_path = c
            break

    if not template_path:
        for root, dirs, files in os.walk(base_dir):
            for f in files:
                if "hycys" in f.lower() and "rad" in f.lower() and f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    template_path = os.path.join(root, f)
                    break
            if template_path:
                break

    if not template_path or not os.path.isfile(template_path):
        raise FileNotFoundError(
            f"Vorlage 'HYCYS - Auswertung Rad.xlsx' wurde auf dem Server nicht gefunden. "
            f"Gesucht in: {candidates[0]}"
        )

    wb = openpyxl.load_workbook(template_path)
    ws = wb["Eingabedaten"]

    vorname, nachname = _split_vorname_nachname(athlete_name)

    ws["B5"] = str(coach or "Markus Hertlein")
    ws["B6"] = _parse_date_str_to_dt(test_date)
    ws["B9"] = nachname
    ws["B10"] = vorname
    ws["B11"] = float(height) if height else 178
    ws["B12"] = float(weight) if weight else 72.0
    ws["B15"] = _parse_date_str_to_dt(birthdate)
    ws["B18"] = "Radsport"
    ws["B19"] = str(kategorie or "Amateur")
    diag_str = str(diagnostik_type or "Cycling BLUE").strip()
    ws["B20"] = diag_str if diag_str.upper().startswith("HYCYS") else f"HYCYS {diag_str}"

    # Sprinttest Laktatwerte (B26:B38 -> Ruhe 1..3 + 0..10 min)
    if df_sprint_lac is not None and "Laktat [mmol/L]" in df_sprint_lac.columns:
        lac_list = list(df_sprint_lac["Laktat [mmol/L]"].values)
        for idx in range(min(13, len(lac_list))):
            val = pd.to_numeric(lac_list[idx], errors="coerce")
            ws.cell(26 + idx, 2, float(val) if pd.notna(val) else None)

    # VO2max & 30s PPO
    ws["F25"] = float(ppo_val) if ppo_val else 0.0
    ws["F31"] = float(abs_vo2max) if abs_vo2max else 0.0

    # Optional: Spiro-Rohdaten in E35:F76 eintragen, falls hochgeladen
    if uploaded_spiro is not None:
        try:
            uploaded_spiro.seek(0)
            df_sp = (
                pd.read_csv(uploaded_spiro, header=None, low_memory=False)
                if str(uploaded_spiro.name).lower().endswith(".csv")
                else pd.read_excel(uploaded_spiro, header=None)
            )
            uploaded_spiro.seek(0)
            if df_sp.shape[1] > 13:
                sp_rows = []
                for r_idx in range(3, len(df_sp)):
                    t_raw = df_sp.iloc[r_idx, 9]
                    v_raw = pd.to_numeric(df_sp.iloc[r_idx, 13], errors="coerce")
                    t_sec = parse_time_to_seconds(t_raw)
                    if pd.notna(t_sec) and pd.notna(v_raw):
                        sp_rows.append((t_sec, float(v_raw)))
                # Letzte 42 Messpunkte (entspricht E35:F76 = 42 Zeilen) eintragen
                sp_tail = sp_rows[-42:] if len(sp_rows) > 42 else sp_rows
                for idx_r in range(42):
                    r_target = 35 + idx_r
                    if idx_r < len(sp_tail):
                        ws.cell(r_target, 5, _seconds_to_excel_time(sp_tail[idx_r][0]))
                        ws.cell(r_target, 6, sp_tail[idx_r][1])
                    else:
                        ws.cell(r_target, 5, None)
                        ws.cell(r_target, 6, None)
        except Exception:
            pass

    # VLamax Zeiten (t_alak & t_bel Overrides in C42:C43)
    ws["C42"] = float(t_alak) if t_alak else 0.0
    ws["C43"] = float(t_bel) if t_bel else 0.0

    # Herzfrequenz lineare Regression (B50:B52)
    ws["B50"] = float(slope_val)
    ws["B51"] = float(intercept_val)
    ws["B52"] = float(hfmax_val)

    # Parizkova 10-Falten (J26:L35) & Fett % (O29)
    sf_sum = 0.0
    if df_sf is not None and not df_sf.empty:
        for idx in range(min(10, len(df_sf))):
            r_excel = 26 + idx
            m1 = float(pd.to_numeric(df_sf.iloc[idx].get("M1", 0.0), errors="coerce") or 0.0)
            m2 = float(pd.to_numeric(df_sf.iloc[idx].get("M2", 0.0), errors="coerce") or 0.0)
            m3 = float(pd.to_numeric(df_sf.iloc[idx].get("M3", 0.0), errors="coerce") or 0.0)
            ws.cell(r_excel, 10, m1)
            ws.cell(r_excel, 11, m2)
            ws.cell(r_excel, 12, m3)
            sf_sum += (m1 + m2 + m3) / 3.0
    if sf_sum > 0:
        ws["O29"] = "=(22.32*LOG10(O26)-29.2)"
    else:
        ws["O29"] = float(body_fat_pct)

    # Sprint-Rohdaten in 'Data Sprint' (ab Zeile 10: Spalte A = Power, Spalte C = Cadence, Spalte E = Time, Spalte F = Index)
    if "Data Sprint" in wb.sheetnames and sprint_powers and len(sprint_powers) > 0:
        ws_sp = wb["Data Sprint"]
        max_clear = max(250, 10 + len(sprint_powers) + 5)
        for r_cl in range(10, max_clear):
            for c_cl in range(1, 7):
                ws_sp.cell(r_cl, c_cl, None)
        for idx, p_val in enumerate(sprint_powers):
            r_sp = 10 + idx
            cad_val = float(sprint_cadences[idx]) if (sprint_cadences and idx < len(sprint_cadences)) else 0.0
            t_val = float(sprint_times[idx]) if (sprint_times and idx < len(sprint_times)) else round(idx * 0.1, 2)
            t_rounded = round(t_val, 2)
            t_min = int(t_rounded // 60)
            t_rem = t_rounded - t_min * 60
            t_str = f"0:{t_min:02d}:{t_rem:05.2f}"
            ws_sp.cell(r_sp, 1, int(round(float(p_val))))
            ws_sp.cell(r_sp, 3, int(round(cad_val)))
            ws_sp.cell(r_sp, 5, t_str)
            ws_sp.cell(r_sp, 6, idx + 1)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()


def extract_gdrive_folder_id(link_or_id: str) -> str:
    """Extrahiert die Google Drive Folder-ID aus einem Drive-Link oder gibt die ID direkt zurück."""
    import re
    s = str(link_or_id or "").strip()
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    m_id = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", s)
    if m_id:
        return m_id.group(1)
    return s


def upload_files_to_gdrive(
    pdf_bytes: bytes,
    pdf_filename: str,
    excel_bytes: bytes,
    excel_filename: str,
    athlete_folder_name: str,
    gender_str: str,
    root_folder_link_or_id: str,
):
    """
    Speichert SOWOHL das PDF als auch die HYCYS-Excel-Rohdatendatei in Google Drive unter:
      <Root-Ordner> / <Männer bzw. Frauen> / <Nachname_Vorname> / <Dateiname>
    Sucht bestehende Ordner case-insensitiv (z.B. 'Nachname_Vorname') und legt sie nur an,
    falls sie noch nicht existieren.
    Unterstützt:
      1. Lokalen Google-Drive-Desktop-Pfad (falls auf dem PC synchronisiert)
      2. Google Drive API per OAuth2 Desktop-Login (gdrive_oauth_client.json / gdrive_token.json)
      3. Google Drive API per Service Account (gdrive_service_account.json / st.secrets)
    """
    gender_norm = str(gender_str or "männlich").strip().lower()
    is_female = ("weib" in gender_norm) or ("frau" in gender_norm) or (gender_norm == "w")
    gender_folder_candidates = ["Frauen", "Damen", "Weiblich"] if is_female else ["Männer", "Maenner", "Herren", "Männlich"]
    default_gender_folder = "Frauen" if is_female else "Männer"

    raw_target = str(root_folder_link_or_id or DEFAULT_GDRIVE_ROOT_LINK).strip()
    if not raw_target:
        return False, "Bitte einen Google Drive Ordner-Link (oder lokalen Google-Drive-Pfad) hinterlegen."

    files_to_save = [(pdf_filename, pdf_bytes, "application/pdf")]
    if excel_bytes is not None and excel_filename:
        files_to_save.append(
            (excel_filename, excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        )

    # Fall A: Lokaler Google-Drive-Desktop-Pfad (falls auf dem Rechner gemountet)
    if os.path.isdir(raw_target):
        gender_dir = None
        for entry in os.listdir(raw_target):
            full_e = os.path.join(raw_target, entry)
            if os.path.isdir(full_e) and entry.lower() in [c.lower() for c in gender_folder_candidates]:
                gender_dir = full_e
                break
        if gender_dir is None:
            gender_dir = os.path.join(raw_target, default_gender_folder)
            os.makedirs(gender_dir, exist_ok=True)

        athlete_dir = None
        for entry in os.listdir(gender_dir):
            full_a = os.path.join(gender_dir, entry)
            if os.path.isdir(full_a) and entry.strip().lower() == athlete_folder_name.lower():
                athlete_dir = full_a
                break
        created_new = False
        if athlete_dir is None:
            athlete_dir = os.path.join(gender_dir, athlete_folder_name)
            os.makedirs(athlete_dir, exist_ok=True)
            created_new = True

        saved_names = []
        for fname, fbytes, _ in files_to_save:
            target_file = os.path.join(athlete_dir, fname)
            with open(target_file, "wb") as f_out:
                f_out.write(fbytes)
            saved_names.append(fname)
        status_note = "Neuer Athleten-Ordner erstellt" if created_new else "Bestehender Athleten-Ordner genutzt"
        return True, f"Gespeichert in `{athlete_dir}`: **{', '.join(saved_names)}** ({status_note})."

    # Fall B: Google Drive API via google-api-python-client
    root_id = extract_gdrive_folder_id(raw_target)
    try:
        from google.oauth2 import service_account
        from google.oauth2.credentials import Credentials as UserCredentials
        from google.auth.transport.requests import Request as GoogleAuthRequest
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        import io
    except ImportError:
        return (
            False,
            "Für den direkten Cloud-Upload per Google-Drive-Link werden die Pakete "
            "`google-api-python-client`, `google-auth` und `google-auth-oauthlib` benötigt.",
        )

    creds = None
    scopes = ["https://www.googleapis.com/auth/drive"]
    app_dir = os.path.dirname(os.path.abspath(__file__))
    sa_file = os.path.join(app_dir, "gdrive_service_account.json")
    oauth_client_file = os.path.join(app_dir, "gdrive_oauth_client.json")
    token_file = os.path.join(app_dir, "gdrive_token.json")

    # 1. Gespeichertes OAuth2 User-Token prüfen
    if os.path.isfile(token_file):
        try:
            creds = UserCredentials.from_authorized_user_file(token_file, scopes)
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(GoogleAuthRequest())
                with open(token_file, "w", encoding="utf-8") as f_tok:
                    f_tok.write(creds.to_json())
        except Exception:
            creds = None

    # 2. st.secrets prüfen
    if creds is None or not getattr(creds, "valid", True):
        try:
            if "gcp_service_account" in st.secrets:
                creds = service_account.Credentials.from_service_account_info(
                    dict(st.secrets["gcp_service_account"]), scopes=scopes
                )
        except Exception:
            pass

    # 3. Lokale JSON-Datei (Service Account oder OAuth Client) prüfen
    if creds is None and os.path.isfile(sa_file):
        try:
            import json as _json
            with open(sa_file, "r", encoding="utf-8") as _f:
                _jdata = _json.load(_f)
            if _jdata.get("type") == "service_account":
                creds = service_account.Credentials.from_service_account_file(sa_file, scopes=scopes)
            elif "installed" in _jdata or "web" in _jdata:
                flow = InstalledAppFlow.from_client_secrets_file(sa_file, scopes)
                creds = flow.run_local_server(port=0)
                with open(token_file, "w", encoding="utf-8") as f_tok:
                    f_tok.write(creds.to_json())
        except Exception as e_auth:
            return False, f"Fehler beim Laden von `gdrive_service_account.json`: {e_auth}"

    if creds is None and os.path.isfile(oauth_client_file):
        try:
            flow = InstalledAppFlow.from_client_secrets_file(oauth_client_file, scopes)
            creds = flow.run_local_server(port=0)
            with open(token_file, "w", encoding="utf-8") as f_tok:
                f_tok.write(creds.to_json())
        except Exception as e_oauth:
            return False, f"Fehler bei der Google OAuth-Anmeldung: {e_oauth}"

    if creds is None:
        return (
            False,
            "Noch keine Google-Drive-Anmeldedaten hinterlegt. Bitte lade unten im Feld "
            "**'🔑 Einmalige Google-Drive-Verbindung einrichten'** deine Google JSON-Schlüsseldatei "
            "(Service-Account oder OAuth-Desktop-Client) hoch – oder gib einen lokalen Google-Drive-Desktop-Ordnerpfad an.",
        )

    service = build("drive", "v3", credentials=creds)

    def _find_or_create_subfolder(parent_id: str, candidates: list, create_name: str):
        q = (
            f"'{parent_id}' in parents and mimeType = 'application/vnd.google-apps.folder' "
            f"and trashed = false"
        )
        res = service.files().list(
            q=q, fields="files(id, name)", includeItemsFromAllDrives=True, supportsAllDrives=True
        ).execute()
        files = res.get("files", [])
        cand_lower = [c.strip().lower() for c in candidates]
        for f_obj in files:
            if f_obj["name"].strip().lower() in cand_lower:
                return f_obj["id"], f_obj["name"], False
        # Nicht gefunden -> neu anlegen
        meta = {
            "name": create_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        created = service.files().create(body=meta, fields="id, name", supportsAllDrives=True).execute()
        return created["id"], created["name"], True

    try:
        # 1. Geschlechter-Ordner (Männer / Frauen) suchen oder anlegen
        gender_folder_id, gender_folder_actual, _ = _find_or_create_subfolder(
            root_id, gender_folder_candidates, default_gender_folder
        )
        # 2. Athleten-Ordner (Nachname_Vorname) suchen oder anlegen
        athlete_folder_id, athlete_folder_actual, created_athlete_folder = _find_or_create_subfolder(
            gender_folder_id, [athlete_folder_name], athlete_folder_name
        )

        uploaded_links = []
        for fname, fbytes, fmime in files_to_save:
            q_file = (
                f"'{athlete_folder_id}' in parents and name = '{fname}' and trashed = false"
            )
            existing_files = service.files().list(
                q=q_file, fields="files(id, name, webViewLink)", includeItemsFromAllDrives=True, supportsAllDrives=True
            ).execute().get("files", [])

            media = MediaIoBaseUpload(io.BytesIO(fbytes), mimetype=fmime, resumable=False)
            if existing_files:
                file_id = existing_files[0]["id"]
                uploaded = service.files().update(
                    fileId=file_id, media_body=media, fields="id, name, webViewLink", supportsAllDrives=True
                ).execute()
            else:
                file_meta = {"name": fname, "parents": [athlete_folder_id]}
                uploaded = service.files().create(
                    body=file_meta, media_body=media, fields="id, name, webViewLink", supportsAllDrives=True
                ).execute()
            wlink = uploaded.get("webViewLink", "")
            uploaded_links.append(f"[{fname}]({wlink})" if wlink else f"`{fname}`")

        folder_status = "neuer Ordner erstellt" if created_athlete_folder else "bestehender Ordner genutzt"
        return (
            True,
            f"Erfolgreich in **{gender_folder_actual} / {athlete_folder_actual}** ({folder_status}) gespeichert: "
            + " & ".join(uploaded_links),
        )
    except Exception as e_api:
        err_str = str(e_api)
        if "storageQuotaExceeded" in err_str or "Service Accounts do not have storage quota" in err_str:
            return (
                False,
                "⚠️ **Ordner wurde erstellt, aber Datei-Upload durch Google blockiert (`storageQuotaExceeded`):**  \n"
                "Google weist **Service-Accounts** in persönlichen *'Meine Ablage'*-Ordnern **0 Byte eigenen Speicherplatz** zu "
                "(deshalb konnte der Ordner `Muxel_Christoph` mit 0 Byte angelegt werden, die `.pdf`- und `.xlsx`-Dateien aber nicht).  \n\n"
                "**So löst du das in 1 Minute (2 Möglichkeiten):**  \n"
                "1. **Über euer Google Workspace (`hycys.de` – ohne neuen Schlüssel):** Legt den Zielordner in Google Drive unter **'Geteilte Ablagen' (Shared Drives)** statt unter *'Meine Ablage'* an und gebt ihn für die E-Mail-Adresse eures Service-Accounts als **Inhaltsmanager** frei. In *Geteilten Ablagen* zählt das Speicherkontingent der Organisation und der Service-Account funktioniert sofort!  \n"
                "2. **Oder für 'Meine Ablage' (OAuth Desktop-Login):** Erstelle in der Google Cloud Console (*APIs & Dienste ➔ Anmeldedaten ➔ + Anmeldedaten erstellen*) eine **OAuth-Client-ID (Typ: Desktop-App)**, lade die JSON-Datei unten unter **'🔑 Google-Drive-Verbindung einrichten / ändern'** hoch und klicke noch einmal auf Speichern – dann öffnet sich einmalig dein Browser zur Bestätigung und die Dateien nutzen deinen eigenen Speicherplatz (`m.hertlein@hycys.de`).",
            )
        return False, f"Google Drive API Fehler: {e_api}"


def render_gdrive_upload_section(
    pdf_bytes: bytes,
    diagnostik_type: str,
    athlete_name: str,
    test_date: str,
    gender_str: str,
    key_prefix: str,
    excel_bytes: bytes = None,
):
    """Rendert die Download-Buttons (PDF + HYCYS Excel-Rohdaten) sowie die Google-Drive-Direktablage."""
    pdf_filename = format_hycys_pdf_filename(diagnostik_type, athlete_name, test_date)
    excel_filename = format_hycys_excel_filename(diagnostik_type, athlete_name, test_date)
    athlete_folder = format_athlete_folder_name(athlete_name)
    gender_norm = str(gender_str or "männlich").strip().lower()
    gender_subfolder = "Frauen" if ("weib" in gender_norm or "frau" in gender_norm or gender_norm == "w") else "Männer"

    if excel_bytes is not None:
        col_dl_pdf, col_dl_xls, col_gd = st.columns([1.0, 1.0, 1.3])
    else:
        col_dl_pdf, col_gd = st.columns([1.0, 1.4])
        col_dl_xls = None

    with col_dl_pdf:
        st.download_button(
            label="PDF herunterladen",
            data=pdf_bytes,
            file_name=pdf_filename,
            mime="application/pdf",
            type="primary",
            use_container_width=True,
            key=f"{key_prefix}_pdf_download_btn",
        )

    if col_dl_xls is not None and excel_bytes is not None:
        with col_dl_xls:
            st.download_button(
                label="HYCYS Excel herunterladen",
                data=excel_bytes,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"{key_prefix}_excel_download_btn",
            )

    with col_gd:
        default_drive_link = st.session_state.get("hycys_gdrive_root_link", DEFAULT_GDRIVE_ROOT_LINK)
        if not default_drive_link:
            default_drive_link = DEFAULT_GDRIVE_ROOT_LINK
        with st.popover("PDF & Excel direkt in Google Drive ablegen", use_container_width=True):
            st.markdown(
                f"**Ziel-Struktur in Google Drive:**  \n"
                f"`📁 {gender_subfolder}` ➔ `📁 {athlete_folder}` ➔  \n"
                f"• `📄 {pdf_filename}`  \n"
                f"• `📊 {excel_filename}`"
            )
            drive_link_in = st.text_input(
                "Google Drive Hauptordner-Link (oder lokaler Drive-Pfad):",
                value=default_drive_link,
                placeholder="https://drive.google.com/drive/folders/...",
                key=f"{key_prefix}_gdrive_link_input",
            )
            st.session_state["hycys_gdrive_root_link"] = drive_link_in

            app_dir = os.path.dirname(os.path.abspath(__file__))
            sa_path = os.path.join(app_dir, "gdrive_service_account.json")
            tok_path = os.path.join(app_dir, "gdrive_token.json")
            has_creds = os.path.isfile(sa_path) or os.path.isfile(tok_path)
            if not os.path.isdir(drive_link_in):
                with st.expander("🔑 Google-Drive-Verbindung einrichten / ändern", expanded=not has_creds):
                    st.caption(
                        "Für Ordner in **'Meine Ablage'** bitte eine **OAuth-Client JSON (Desktop-App)** hochladen (nutzt deinen eigenen Google-Speicherplatz). "
                        "Ein **Service-Account JSON** funktioniert bei Google Drive nur in **'Geteilten Ablagen' (Shared Drives)**."
                    )
                    cred_upload = st.file_uploader(
                        "Google JSON-Schlüsseldatei (.json) – OAuth-Client (Desktop) oder Service-Account",
                        type=["json"],
                        key=f"{key_prefix}_gdrive_cred_upload",
                    )
                    if cred_upload is not None:
                        with open(sa_path, "wb") as f_cred:
                            f_cred.write(cred_upload.getvalue())
                        if os.path.isfile(tok_path):
                            try:
                                os.remove(tok_path)
                            except Exception:
                                pass
                        st.success("Schlüsseldatei aktualisiert! Klicke jetzt unten auf Speichern.")

            if st.button(
                f"PDF + Excel in '{gender_subfolder}/{athlete_folder}' speichern",
                type="primary",
                use_container_width=True,
                key=f"{key_prefix}_gdrive_upload_btn",
            ):
                with st.spinner("Speichere PDF und HYCYS Excel-Auswertung in Google Drive..."):
                    ok, msg = upload_files_to_gdrive(
                        pdf_bytes=pdf_bytes,
                        pdf_filename=pdf_filename,
                        excel_bytes=excel_bytes,
                        excel_filename=excel_filename,
                        athlete_folder_name=athlete_folder,
                        gender_str=gender_str,
                        root_folder_link_or_id=drive_link_in,
                    )
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)


def render_centered_html_table(df: pd.DataFrame, highlight_first_row: bool = False):
    """
    Rendert einen Pandas DataFrame als übersichtliche Tabelle, in der alle Spaltenüberschriften
    und alle Zellenwerte garantiert mittig zentriert dargestellt werden.
    """
    if df is None or df.empty:
        return
    headers_html = "".join(
        f"<th style='text-align:center; padding:8px 10px; background-color:#2cb7b9; color:#ffffff; font-weight:700; font-size:13.5px; border:1px solid #e0e6e8;'>{col}</th>"
        for col in df.columns
    )
    rows_html = []
    for r_idx, (_, row) in enumerate(df.iterrows()):
        is_hl = highlight_first_row and (r_idx == 0)
        bg = "#eef8f8" if is_hl else ("#ffffff" if r_idx % 2 == 0 else "#f9fbfb")
        fw = "700" if is_hl else "400"
        cells_html = "".join(
            f"<td style='text-align:center; padding:7px 10px; border:1px solid #e6ecec; color:#1a1a1a; font-weight:{fw}; font-size:13.5px;'>{val if pd.notna(val) else ''}</td>"
            for val in row.values
        )
        rows_html.append(f"<tr style='background-color:{bg};'>{cells_html}</tr>")

    table_html = f"""
    <div style="overflow-x:auto; margin-bottom:10px;">
        <table style="width:100%; border-collapse:collapse; border-radius:6px; overflow:hidden; border:1px solid #d9e2e2;">
            <thead><tr>{headers_html}</tr></thead>
            <tbody>{''.join(rows_html)}</tbody>
        </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)


def render_hycys_top_basisdaten(
    weight: float,
    body_fat_pct: float,
    rel_vo2max: float,
    abs_vo2max: float,
    vlamax_str: str,
    ans_label: str,
    ans_main_str: str,
    ans_sub_str: str,
    ans_hr_str: str = "-",
    vo2_pct_str: str = "-",
    ltp1_label: str = "LTP1",
    ltp1_main_str: str = "-",
    ltp1_sub_str: str = "",
    ltp1_hr_str: str = "-",
    hf_ans_str: str = None,
    hf_ltp1_str: str = None,
    mode: str = "",
    athlete_name: str = "",
    test_date: str = "",
    is_presentation: bool = False,
    extra_header_info: str = "",
    spiro_note: str = "",
    bmi_val: float = None,
    height_cm: float = None,
):
    """
    Stellt den obersten Basisdaten-Abschnitt für Rad- und Laufauswertung einheitlich
    und analog zum HYCYS-PDF dar:
      1. Körperzusammensetzung: Gewicht, Körperfettanteil, Fettfreie Masse + HYCYS-Kreisdiagramm
      2. VO2max und VLamax gleichwertig gegenübergestellt
      3. Schwellenwerte der ANS bzw. LTP2 + HF
      4. Darunter LTP1 (bzw. LT1/Fatmax) + HF
    """
    if hf_ans_str is None:
        hf_ans_str = ans_hr_str
    if hf_ltp1_str is None:
        hf_ltp1_str = ltp1_hr_str

    if athlete_name:
        header_html = f"<b>Athlet:</b> {athlete_name} &nbsp;|&nbsp; <b>Test-Datum:</b> {test_date}"
        if not is_presentation and extra_header_info:
            header_html += f" &nbsp;|&nbsp; {extra_header_info}"

        st.markdown(
            f"""
            <div style="font-size: 16px; line-height: 1.5; margin-bottom: 16px; padding: 12px 16px; border-left: 5px solid #2cb7b9; background-color: #f8fafa; border-radius: 4px;">
                {header_html}
            </div>
            """,
            unsafe_allow_html=True,
        )

    vla_clean = str(vlamax_str).replace("mmol/l/s", "").replace("mmol/L/s", "").strip()
    vla_unit_html = '<span style="font-size:13px; color:#1a1a1a; font-weight:600;">mmol/l/s</span>' if vla_clean and vla_clean != "-" else ""

    fett_kg = weight * (body_fat_pct / 100.0)
    ffm_kg = weight - fett_kg
    ffm_pct = max(0.0, 100.0 - body_fat_pct)

    # Kuchendiagramm gut lesbar erzeugen; Schriftgröße im Kuchenstück wird über die exakte Bounding-Box
    # maximiert, sodass sie so groß wie möglich ist, aber niemals über die Ränder des Kuchenstücks hinausragt
    import base64, io
    fig_bf, ax_bf = plt.subplots(figsize=(3.15, 2.25), dpi=300)
    pct_fett = max(0.1, body_fat_pct)
    pct_frei = max(0.1, 100.0 - max(0.0, body_fat_pct))
    if body_fat_pct <= 0.0:
        lbl_fett = ""
    else:
        lbl_fett = f"{pct_fett:.1f}%".replace(".", ",") if pct_fett < 14.0 else f"{pct_fett:.1f} %".replace(".", ",")
    lbl_frei = f"{pct_frei:.1f} %".replace(".", ",")
    wedges, texts, autotexts = ax_bf.pie(
        [pct_fett, pct_frei],
        explode=(0.10, 0),
        colors=["#2cb7b9", "#cdb663"],
        autopct="",
        startangle=90,
        counterclock=False,
        textprops=dict(weight="bold", color="white"),
        wedgeprops=dict(edgecolor="white", linewidth=1.6),
    )
    autotexts[0].set_text(lbl_fett)
    autotexts[1].set_text(lbl_frei)
    ax_bf.axis("equal")
    fig_bf.patch.set_facecolor("white")
    fig_bf.canvas.draw()
    renderer_bf = fig_bf.canvas.get_renderer()
    inv_bf = ax_bf.transData.inverted()

    for i, (wedge, autotext) in enumerate(zip(wedges, autotexts)):
        slice_pct = pct_fett if i == 0 else pct_frei
        th1, th2 = float(wedge.theta1), float(wedge.theta2)
        th_mid = np.deg2rad((th1 + th2) / 2.0)
        cx, cy = wedge.center
        autotext.set_color("white")
        autotext.set_weight("bold")

        if slice_pct >= 50.0:
            d_large = 0.48
            autotext.set_position((cx + d_large * np.cos(th_mid), cy + d_large * np.sin(th_mid)))
            autotext.set_fontsize(12.5)
        else:
            best_fs = 5.0
            best_pos = (cx + 0.68 * np.cos(th_mid), cy + 0.68 * np.sin(th_mid))
            found_fit = False
            for fs_try in np.arange(12.5, 4.4, -0.5):
                autotext.set_fontsize(float(fs_try))
                for d_try in [0.55, 0.60, 0.65, 0.70, 0.74, 0.78, 0.81]:
                    for ang_off in [0.0, -1.2, -2.5, 1.2]:
                        if abs(ang_off) >= (th2 - th1) * 0.25:
                            continue
                        th_try = np.deg2rad((th1 + th2) / 2.0 + ang_off)
                        px = cx + d_try * np.cos(th_try)
                        py = cy + d_try * np.sin(th_try)
                        autotext.set_position((px, py))
                        bb = inv_bf.transform_bbox(autotext.get_window_extent(renderer_bf))
                        corners = [
                            (bb.x0, bb.y0),
                            (bb.x0, bb.y1),
                            (bb.x1, bb.y0),
                            (bb.x1, bb.y1),
                            ((bb.x0 + bb.x1) / 2.0, bb.y0),
                            ((bb.x0 + bb.x1) / 2.0, bb.y1),
                        ]
                        fits_wedge = True
                        for x_c, y_c in corners:
                            dx, dy = x_c - cx, y_c - cy
                            r_c = float(np.hypot(dx, dy))
                            ang_c = (float(np.rad2deg(np.arctan2(dy, dx))) + 360.0) % 360.0
                            if r_c > 0.90 or ang_c < (th1 + 1.8) or ang_c > (th2 - 1.8):
                                fits_wedge = False
                                break
                        if fits_wedge:
                            best_fs = float(fs_try)
                            best_pos = (px, py)
                            found_fit = True
                            break
                    if found_fit:
                        break
                if found_fit:
                    break
            autotext.set_fontsize(best_fs)
            autotext.set_position(best_pos)

    handles_bf = [Patch(color="#2cb7b9", label="Fett"), Patch(color="#cdb663", label="Fettfrei")]
    legend_bf = ax_bf.legend(
        handles=handles_bf,
        loc="center left",
        bbox_to_anchor=(0.94, 0.5),
        fontsize=10.0,
        frameon=False,
        handlelength=1.2,
        handletextpad=0.5,
    )
    for text in legend_bf.get_texts():
        text.set_weight("bold")
        text.set_color("#1a1a1a")
    plt.tight_layout(pad=0.04)
    buf_bf = io.BytesIO()
    fig_bf.savefig(buf_bf, format="png", dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="white")
    plt.close(fig_bf)
    b64_pie = base64.b64encode(buf_bf.getvalue()).decode("utf-8")

    # 1. ZEILE: KÖRPERZUSAMMENSETZUNG + GUT LESBARES KUCHENDIAGRAMM
    show_bmi_tile = (not is_presentation) and (bmi_val is not None and bmi_val > 0)
    anthro_cols_css = "1fr 1fr 1fr 1fr" if show_bmi_tile else "1fr 1fr 1fr"
    height_sub_html = f'<div style="font-size:11px; color:#1a1a1a;">(Größe: {int(round(height_cm))} cm)</div>' if (height_cm is not None and height_cm > 0) else ""
    bmi_tile_html = (
        f'<div style="background:#f8fafa; padding:10px 6px; border-radius:6px; border:1px solid #edf2f2;">'
        f'<div style="font-size:12px; color:#1a1a1a; font-weight:600;">BMI</div>'
        f'<div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{bmi_val:.1f} kg/m²</div>'
        f'{height_sub_html}'
        f'</div>'
    ) if show_bmi_tile else ""

    st.markdown(
        f'<div style="display:grid; grid-template-columns: 1.48fr 1.08fr; gap:14px; align-items:stretch;">'
        f'<div style="border:1px solid #e0e8e8; border-left:5px solid #2cb7b9; border-radius:6px; padding:14px 18px; background-color:#ffffff; display:flex; flex-direction:column; justify-content:center;">'
        f'<div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase; letter-spacing:0.4px; margin-bottom:10px; text-align:center;">'
        f'Anthropometrie'
        f'</div>'
        f'<div style="display:grid; grid-template-columns: {anthro_cols_css}; gap:10px; text-align:center;">'
        f'<div style="background:#f8fafa; padding:10px 6px; border-radius:6px; border:1px solid #edf2f2;">'
        f'<div style="font-size:12px; color:#1a1a1a; font-weight:600;">Gewicht</div>'
        f'<div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{weight:.1f} kg</div>'
        f'</div>'
        f'{bmi_tile_html}'
        f'<div style="background:#f8fafa; padding:10px 6px; border-radius:6px; border:1px solid #edf2f2;">'
        f'<div style="font-size:12px; color:#1a1a1a; font-weight:600;">Körperfettanteil</div>'
        f'<div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{body_fat_pct:.1f} %</div>'
        f'<div style="font-size:11px; color:#1a1a1a;">({fett_kg:.1f} kg Fettmasse)</div>'
        f'</div>'
        f'<div style="background:#f8fafa; padding:10px 6px; border-radius:6px; border:1px solid #edf2f2;">'
        f'<div style="font-size:12px; color:#1a1a1a; font-weight:600;">Fettfreie Masse (LBM)</div>'
        f'<div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{ffm_kg:.1f} kg</div>'
        f'<div style="font-size:11px; color:#1a1a1a;">({ffm_pct:.1f} % fettfrei)</div>'
        f'</div>'
        f'</div>'
        f'</div>'
        f'<div style="border:1px solid #e0e8e8; border-radius:6px; padding:6px 12px; background-color:#ffffff; display:flex; align-items:center; justify-content:center;">'
        f'<img src="data:image/png;base64,{b64_pie}" style="height:175px; width:auto; object-fit:contain; display:block;" />'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # 2. ZEILE: VO2MAX UND VLAMAX GLEICHWERTIG GEGENÜBERGESTELLT (BEIDE MIT BLAUER MARKIERUNG & SCHWARZER SCHRIFT)
    col_vo2, col_vla = st.columns(2)
    with col_vo2:
        st.markdown(
            f"""
            <div style="border:1px solid #e0e8e8; border-left:5px solid #2cb7b9; border-radius:6px; padding:12px 18px; background-color:#ffffff; margin-top:8px; text-align:center;">
                <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">Aerobe Kapazität · VO2max</div>
                <div style="display:flex; justify-content:center; align-items:baseline; gap:18px; margin-top:4px;">
                    <div>
                        <span style="font-size:24px; font-weight:800; color:#1a1a1a;">{rel_vo2max:.1f}</span>
                        <span style="font-size:13px; color:#1a1a1a; font-weight:600;"> ml/min/kg (rel.)</span>
                    </div>
                    <div style="color:#1a1a1a; font-size:18px;">|</div>
                    <div>
                        <span style="font-size:24px; font-weight:800; color:#1a1a1a;">{int(round(abs_vo2max))}</span>
                        <span style="font-size:13px; color:#1a1a1a; font-weight:600;"> ml/min (abs.)</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if spiro_note:
            st.caption(f"{spiro_note}")
    with col_vla:
        st.markdown(
            f"""
            <div style="border:1px solid #e0e8e8; border-left:5px solid #2cb7b9; border-radius:6px; padding:12px 18px; background-color:#ffffff; margin-top:8px; text-align:center;">
                <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">Glykolytische Rate · VLamax</div>
                <div style="display:flex; justify-content:center; align-items:baseline; gap:8px; margin-top:4px;">
                    <span style="font-size:24px; font-weight:800; color:#1a1a1a;">{vla_clean}</span>
                    {vla_unit_html}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 3. ZEILE: SCHWELLENWERTE DER ANS / LTP2 + HF + % VO2 (ALLES IN SCHWARZ, BLAUE MARKIERUNG LINKS)
    sub_ans_html = f"<span style='font-size:14px; color:#1a1a1a; font-weight:600; margin-left:8px;'>({ans_sub_str})</span>" if ans_sub_str else ""
    st.markdown(
        f"""
        <div style="border:1px solid #e0e8e8; border-left:5px solid #2cb7b9; border-radius:6px; padding:10px 18px; background-color:#ffffff; margin-top:8px;">
            <div style="display:grid; grid-template-columns: 1.2fr 1fr 1fr; gap:16px; text-align:center; align-items:center;">
                <div>
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">{ans_label}</div>
                    <div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{ans_main_str}{sub_ans_html}</div>
                </div>
                <div>
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">HF @ {ans_label}</div>
                    <div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{hf_ans_str}</div>
                </div>
                <div>
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">% VO2</div>
                    <div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{vo2_pct_str}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 4. ZEILE (DARUNTER): LTP1 UND DESSEN HF (IDENTISCHES FORMAT WIE ANS, ALLES IN SCHWARZ, BLAUE MARKIERUNG LINKS)
    sub_ltp1_html = f"<span style='font-size:14px; color:#1a1a1a; font-weight:600; margin-left:8px;'>({ltp1_sub_str})</span>" if ltp1_sub_str else ""
    st.markdown(
        f"""
        <div style="border:1px solid #e0e8e8; border-left:5px solid #2cb7b9; border-radius:6px; padding:10px 18px; background-color:#ffffff; margin-top:8px; margin-bottom:6px;">
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:16px; text-align:center; align-items:center;">
                <div>
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">{ltp1_label}</div>
                    <div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{ltp1_main_str}{sub_ltp1_html}</div>
                </div>
                <div>
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase;">HF @ {ltp1_label}</div>
                    <div style="font-size:21px; font-weight:800; color:#1a1a1a; margin-top:2px;">{hf_ltp1_str}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def calc_pot_weight_from_bf(
    weight_ist: float,
    bf_ist: float,
    bf_target: float,
    fm_loss_fraction: float = 0.85,
) -> float:
    """
    Berechnet das Ziel-Körpergewicht aus dem Ziel-Körperfettanteil unter der Annahme,
    dass bei einer Gewichtsänderung `fm_loss_fraction` (Standard: 85 %) aus Fettmasse (FM)
    und `1 - fm_loss_fraction` (Standard: 15 %) aus fettfreier Masse (FFM) stammen:
      FM_pot = FM_ist + fm_loss_fraction * (W_pot - W_ist)
      => W_pot * (bf_target / 100) = W_ist * (bf_ist / 100) + fm_loss_fraction * (W_pot - W_ist)
      => W_pot = W_ist * (fm_loss_fraction - bf_ist / 100) / (fm_loss_fraction - bf_target / 100)
    """
    if weight_ist <= 0 or bf_ist <= 0 or abs(bf_target - bf_ist) < 1e-6:
        return float(weight_ist)
    denom = fm_loss_fraction - (bf_target / 100.0)
    numer = fm_loss_fraction - (bf_ist / 100.0)
    if denom <= 0.05 or numer <= 0.0:
        return float(weight_ist)
    return float(weight_ist * (numer / denom))


def compute_galibier_v2_cycling(
    p_vals: np.ndarray,
    weight: float,
    rel_vo2max: float,
    vlamax: float,
    carb_intake_g_h: float = 90.0,
    km_val: float = 0.21,
    km_low: float = 0.15,
    km_high: float = 0.28,
    eta: float = 0.239,
    vo2_rest: float = 3.5,
) -> dict:
    """
    Version 2 (Galibier v4.2) für Radsport:
      - Ruhe-VO2 = 3,5 mL/kg/min (1 MET)
      - Konstanter Nettowirkungsgrad eta = 23,9 % (mechanische Leistung / Energieumsatz über Ruhe)
      - Substratabhängiges kalorisches Äquivalent AE = 19,6 * f_Fett + 21,1 * f_KH [kJ/L O2]
        => O2-Kosten variieren mit Substrat (~12,8 mL/min/W bei reinem Fett, ~11,9 bei reinem KH, ~12,0 am MLSS)
      - Michaelis-Menten PDH-Sättigung aus r = VLass / VLaoxmax:
        f_KH = r / (r + Km), f_Fett = Km / (r + Km) mit Km = 0,21 (Range: 0,15 - 0,28)
    """
    ks1, ks2, la_eq, vol_dist = 0.0631, 1.331, 0.02049, 0.4
    p_arr = np.asarray(p_vals, dtype=float)
    if weight <= 0 or rel_vo2max <= vo2_rest + 1.0:
        z = np.zeros_like(p_arr)
        return {
            "ee_tot": z, "ee_fat": z, "fat_g": z, "kh_kcal": z, "kh_g": z,
            "fat_pct": z, "kh_pct": z,
            "ee_fat_low": z, "ee_fat_high": z, "kh_g_low": z, "kh_g_high": z,
            "mlss_power": 0, "fatmax_power": 0, "carb_match_power": 0,
            "carb_match_kh_g": 0.0, "carb_match_fat_g": 0.0, "carb_match_ee_ges": 0.0,
        }

    vo2_grid = np.linspace(vo2_rest + 0.01, rel_vo2max - 0.02, 4000)
    adp = np.sqrt((ks1 * vo2_grid) / (rel_vo2max - vo2_grid))
    vlass = (60.0 * max(0.10, float(vlamax))) / (1.0 + ks2 / (adp ** 3))
    vlaox = (la_eq / vol_dist) * vo2_grid
    r_grid = np.maximum(0.0, vlass / np.maximum(1e-9, vlaox))

    def _eval_km(km_c: float):
        f_kh = r_grid / (r_grid + km_c)
        f_fat = km_c / (r_grid + km_c)
        ae = 19.6 * f_fat + 21.1 * f_kh
        e_net_kj_min = ((vo2_grid - vo2_rest) * weight / 1000.0) * ae
        p_grid = (e_net_kj_min * eta) / 0.06
        ee_tot_grid = ((vo2_grid * weight / 1000.0) * ae / 4.186) * 60.0
        ee_fat_grid = ee_tot_grid * f_fat
        ee_kh_grid = ee_tot_grid * f_kh
        kh_g_grid = ee_kh_grid / 4.063
        return p_grid, ee_tot_grid, ee_fat_grid, ee_kh_grid, kh_g_grid, f_fat * 100.0, f_kh * 100.0

    p_g, ee_tot_g, ee_fat_g, ee_kh_g, kh_g_g, fat_pct_g, kh_pct_g = _eval_km(km_val)
    p_g_lo, _, ee_fat_g_lo, _, kh_g_g_lo, _, _ = _eval_km(km_low)
    p_g_hi, _, ee_fat_g_hi, _, kh_g_g_hi, _, _ = _eval_km(km_high)

    ee_tot = np.interp(p_arr, p_g, ee_tot_g)
    ee_fat = np.interp(p_arr, p_g, ee_fat_g)
    kh_kcal = np.interp(p_arr, p_g, ee_kh_g)
    kh_g = np.interp(p_arr, p_g, kh_g_g)
    fat_g = ee_fat / 9.3
    fat_pct = np.interp(p_arr, p_g, fat_pct_g)
    kh_pct = np.interp(p_arr, p_g, kh_pct_g)

    # Km = 0.15 -> höherer KH-Anteil, niedrigerer Fett-Anteil; Km = 0.28 -> niedrigerer KH-Anteil, höherer Fett-Anteil
    ee_fat_low = np.interp(p_arr, p_g_lo, ee_fat_g_lo)
    ee_fat_high = np.interp(p_arr, p_g_hi, ee_fat_g_hi)
    kh_g_high = np.interp(p_arr, p_g_lo, kh_g_g_lo)
    kh_g_low = np.interp(p_arr, p_g_hi, kh_g_g_hi)

    idx_mlss = int(np.argmin(np.abs(r_grid - 1.0)))
    mlss_power = int(round(float(p_g[idx_mlss])))
    idx_fm = int(np.argmax(ee_fat_g))
    fatmax_power = int(round(float(p_g[idx_fm])))

    le_mask = kh_g_g <= carb_intake_g_h
    if np.any(le_mask):
        idx_khm = int(np.where(le_mask)[0][-1])
    else:
        idx_khm = int(np.argmin(np.abs(kh_g_g - carb_intake_g_h)))
    carb_match_power = int(round(float(p_g[idx_khm])))
    carb_match_kh_g = float(kh_g_g[idx_khm])
    carb_match_fat_g = float(ee_fat_g[idx_khm] / 9.3)
    carb_match_ee_ges = float(ee_tot_g[idx_khm])

    return {
        "ee_tot": ee_tot,
        "ee_fat": ee_fat,
        "fat_g": fat_g,
        "kh_kcal": kh_kcal,
        "kh_g": kh_g,
        "fat_pct": fat_pct,
        "kh_pct": kh_pct,
        "ee_fat_low": ee_fat_low,
        "ee_fat_high": ee_fat_high,
        "kh_g_low": kh_g_low,
        "kh_g_high": kh_g_high,
        "mlss_power": mlss_power,
        "fatmax_power": fatmax_power,
        "carb_match_power": carb_match_power,
        "carb_match_kh_g": carb_match_kh_g,
        "carb_match_fat_g": carb_match_fat_g,
        "carb_match_ee_ges": carb_match_ee_ges,
    }


def compute_galibier_v2_running(
    v_grid: np.ndarray,
    vo2_rel_dem_v1: np.ndarray,
    weight: float,
    rel_vo2max: float,
    vlamax: float,
    carb_intake_g_h: float = 90.0,
    km_val: float = 0.21,
    km_low: float = 0.15,
    km_high: float = 0.28,
) -> dict:
    """
    Version 2 (Galibier v4.2) übertragen auf das Laufen:
      - Konstanter Nettowirkungsgrad / konstanter Netto-Energiebedarf pro Meter bei der jeweiligen Geschwindigkeit.
      - Da Fett weniger Energie pro Liter O2 liefert (19,6 vs. 21,1 kJ/L O2; am MLSS ~20,84 kJ/L O2),
        variieren die O2-Kosten mit der Substratmischung: VO2_net_V2(v) = VO2_net_V1(v) * (20,84 / AE(r)).
      - Michaelis-Menten PDH-Sättigung aus r = VLass / VLaoxmax:
        f_KH = r / (r + Km), f_Fett = Km / (r + Km) mit Km = 0,21 (Range: 0,15 - 0,28).
    """
    ks1, ks2, la_eq, vol_dist = 0.0631, 1.331, 0.02049, 0.4
    v_arr = np.asarray(v_grid, dtype=float)
    vo2_v1 = np.asarray(vo2_rel_dem_v1, dtype=float)
    eff_vo2max = float(rel_vo2max if rel_vo2max > 10.0 else 60.0)
    vla_use = max(0.10, float(vlamax if vlamax > 0 else 0.45))

    def _eval_km_run(km_c: float):
        # Am MLSS (r=1) beträgt AE_mlss = 19.6 * (km_c/(1+km_c)) + 21.1 * (1/(1+km_c))
        ae_mlss = 19.6 + 1.5 * (1.0 / (1.0 + km_c))
        ae_cur = np.full_like(v_arr, ae_mlss, dtype=float)
        vo2_net_base = np.maximum(0.1, vo2_v1 - 3.5)
        for _ in range(5):
            vo2_rel_v2 = vo2_net_base * (ae_mlss / ae_cur) + 3.5
            d_vo2 = eff_vo2max - vo2_rel_v2
            valid = (d_vo2 > 0.02) & (vo2_rel_v2 > 0)
            r_arr = np.full_like(v_arr, 5.0, dtype=float)
            if np.any(valid):
                adp = np.sqrt((ks1 * vo2_rel_v2[valid]) / d_vo2[valid])
                vlass = (60.0 * vla_use) / (1.0 + ks2 / (adp ** 3))
                vlaox = (la_eq / vol_dist) * vo2_rel_v2[valid]
                r_arr[valid] = np.maximum(0.0, vlass / np.maximum(1e-9, vlaox))
            f_kh = r_arr / (r_arr + km_c)
            f_fat = km_c / (r_arr + km_c)
            ae_cur = 19.6 * f_fat + 21.1 * f_kh

        ee_tot = (((vo2_rel_v2 * weight / 1000.0) * ae_cur) * 60.0) / 4.186
        ee_fat = ee_tot * f_fat
        ee_kh = ee_tot * f_kh
        kh_g = ee_kh / 4.063
        fat_g = ee_fat / 9.3
        return vo2_rel_v2, r_arr, ee_tot, ee_fat, fat_g, ee_kh, kh_g, f_fat * 100.0, f_kh * 100.0

    vo2_v2, r_v2, ee_tot, ee_fat, fat_g, ee_kh, kh_g, fat_pct, kh_pct = _eval_km_run(km_val)
    _, _, _, ee_fat_lo, _, _, kh_g_hi, _, _ = _eval_km_run(km_low)
    _, _, _, ee_fat_hi, _, _, kh_g_lo, _, _ = _eval_km_run(km_high)

    sub_mask = (eff_vo2max - vo2_v2) > 0.05
    if np.any(sub_mask):
        v_sub = v_arr[sub_mask]
        r_sub = r_v2[sub_mask]
        fat_sub = ee_fat[sub_mask]
        kh_sub = kh_g[sub_mask]
        ee_sub = ee_tot[sub_mask]
        v_mlss_v2 = float(v_sub[int(np.argmin(np.abs(r_sub - 1.0)))])
        v_fatmax_v2 = float(v_sub[int(np.argmax(fat_sub))])
        le_m = kh_sub <= carb_intake_g_h
        idx_k = int(np.where(le_m)[0][-1]) if np.any(le_m) else int(np.argmin(np.abs(kh_sub - carb_intake_g_h)))
        v_khm_v2 = float(v_sub[idx_k])
        kh_at_m_v2 = float(kh_sub[idx_k])
        fat_g_at_m_v2 = float(fat_sub[idx_k] / 9.3)
        ee_at_m_v2 = float(ee_sub[idx_k])
    else:
        v_mlss_v2 = float(v_arr[-1])
        v_fatmax_v2 = float(v_arr[0])
        v_khm_v2 = float(v_arr[0])
        kh_at_m_v2 = float(carb_intake_g_h)
        fat_g_at_m_v2 = 0.0
        ee_at_m_v2 = 0.0

    return {
        "ee_tot": ee_tot,
        "ee_fat": ee_fat,
        "fat_g": fat_g,
        "kh_kcal": ee_kh,
        "kh_g": kh_g,
        "fat_pct": fat_pct,
        "kh_pct": kh_pct,
        "ee_fat_low": ee_fat_lo,
        "ee_fat_high": ee_fat_hi,
        "kh_g_low": kh_g_lo,
        "kh_g_high": kh_g_hi,
        "v_mlss": v_mlss_v2,
        "v_fatmax": v_fatmax_v2,
        "v_kh_match": v_khm_v2,
        "kh_at_match": kh_at_m_v2,
        "fat_g_at_match": fat_g_at_m_v2,
        "ee_at_match": ee_at_m_v2,
    }


def render_ans_coaching_diagrams(
    mode: str,
    weight_ist: float,
    bf_ist: float,
    abs_vo2max_ist: float,
    rel_vo2max_ist: float,
    vlamax_ist: float,
    ans_abs_ist: float,
    ans_rel_ist: float,
    solve_ans_fn,
    default_vo2_pct: float = 5.0,
    default_vla_pct: float = -10.0,
    default_bf_delta: float = -1.5,
    key_prefix: str = "lauf",
    is_presentation: bool = False,
):
    """
    Rendert die Live Simulation physiologisches Profil (VO2max, VLamax, Fettanteil -> ANS)
    mit:
      - Untereinander angeordneten Kennzahlen + Schiebereglern (ohne Nachkommastellen bei VO2max/VLamax, inkl. %)
      - Wasserfall-Diagramm (ANS Ist & Potenzial in Blau/Türkis, Potenzial-Hebel in Gold)
      - Spinnendiagramm (Pentagon/Quadrat, Potenzial in Grau gestrichelt)
    """
    st.markdown("---")
    st.subheader("Simulation physiologisches Profil")

    # Synchronisierte Session-State-Keys, damit Regler-Werte beim Wechsel zwischen Arbeitsansicht und Präsentations-Modus erhalten bleiben
    vo2_state_key = f"{key_prefix}_shared_pot_vo2_pct"
    vla_state_key = f"{key_prefix}_shared_pot_vla_pct"
    bf_state_key = f"{key_prefix}_shared_pot_bf_target"

    min_bf = max(4.0, round(bf_ist - 8.0, 1))
    max_bf = min(35.0, round(bf_ist + 5.0, 1))
    default_bf_target = max(min_bf, min(max_bf, round(bf_ist + default_bf_delta, 1)))

    if vo2_state_key not in st.session_state:
        st.session_state[vo2_state_key] = int(round(default_vo2_pct if default_vo2_pct != 0 else 5.0))
    if vla_state_key not in st.session_state:
        st.session_state[vla_state_key] = int(round(default_vla_pct if default_vla_pct != 0 else -10.0))
    if bf_state_key not in st.session_state or not (min_bf <= float(st.session_state[bf_state_key]) <= max_bf):
        st.session_state[bf_state_key] = float(default_bf_target)

    # Vorab aktuelle Regler-Werte aus Session State lesen, damit wir Wert und Schieberegler exakt untereinander in derselben Spalte darstellen können
    vo2_slider_key = f"{key_prefix}_{'pres' if is_presentation else 'work'}_vo2_pct"
    vla_slider_key = f"{key_prefix}_{'pres' if is_presentation else 'work'}_vla_pct"
    bf_slider_key = f"{key_prefix}_{'pres' if is_presentation else 'work'}_bf_target"

    cur_vo2_pct = int(round(float(st.session_state.get(vo2_slider_key, st.session_state[vo2_state_key]))))
    cur_vla_pct = int(round(float(st.session_state.get(vla_slider_key, st.session_state[vla_state_key]))))
    cur_bf_pot = float(st.session_state.get(bf_slider_key, st.session_state[bf_state_key]))
    cur_bf_pot = max(float(min_bf), min(float(max_bf), cur_bf_pot))

    # Vorab-Berechnung mit den aktuellen Slider-Werten (85 % Fettmasse / 15 % FFM bei Gewichtsänderung)
    weight_pot = calc_pot_weight_from_bf(weight_ist, bf_ist, cur_bf_pot, fm_loss_fraction=0.85)
    abs_vo2max_pot = abs_vo2max_ist * (1.0 + cur_vo2_pct / 100.0)
    rel_vo2max_pot = abs_vo2max_pot / weight_pot if weight_pot > 0 else rel_vo2max_ist
    vlamax_pot = max(0.15, vlamax_ist * (1.0 + cur_vla_pct / 100.0))

    ans_abs_0, ans_rel_0 = float(ans_abs_ist), float(ans_rel_ist)
    ans_abs_1, ans_rel_1 = solve_ans_fn(abs_vo2max_pot, vlamax_ist, weight_ist)
    ans_abs_2, ans_rel_2 = solve_ans_fn(abs_vo2max_pot, vlamax_pot, weight_ist)
    ans_abs_3, ans_rel_3 = solve_ans_fn(abs_vo2max_pot, vlamax_pot, weight_pot)

    # Beim Laufen ist die effektive ANS immer die reale Laufgeschwindigkeit (inkl. Gewichtseffekt auf rel. VO2max)
    if mode == "running":
        ans_abs_pot = float(ans_rel_3)
        ans_rel_pot = float(ans_rel_3)
    else:
        ans_abs_pot = float(ans_abs_3)
        ans_rel_pot = float(ans_rel_3)

    # 4 SPALTEN: JEWEILS WERT OBEN & ZUGEHÖRIGER SCHIEBEREGLER DIREKT DARUNTER
    c_col1, c_col2, c_col3, c_col4 = st.columns([1.05, 1.05, 1.05, 1.15])

    with c_col1:
        st.metric(
            "VO2max (Potenzial)",
            f"{rel_vo2max_pot:.1f} ml/min/kg",
            f"{rel_vo2max_pot - rel_vo2max_ist:+.1f} ml/min/kg ({int(round(abs_vo2max_pot))} ml)",
        )
        d_vo2_pct = st.slider(
            "Δ VO2max (absolut)",
            min_value=-10,
            max_value=20,
            value=int(cur_vo2_pct),
            step=1,
            format="%d %%",
            key=vo2_slider_key,
        )
        st.session_state[vo2_state_key] = int(d_vo2_pct)

    with c_col2:
        st.metric(
            "VLamax (Potenzial)",
            f"{vlamax_pot:.2f} mmol/l/s",
            f"{vlamax_pot - vlamax_ist:+.2f} mmol/l/s",
            delta_color="inverse",
        )
        d_vla_pct = st.slider(
            "Δ VLamax",
            min_value=-30,
            max_value=20,
            value=int(cur_vla_pct),
            step=1,
            format="%d %%",
            key=vla_slider_key,
        )
        st.session_state[vla_state_key] = int(d_vla_pct)

    with c_col3:
        st.metric(
            "Ziel-Gewicht & Körperfett",
            f"{weight_pot:.1f} kg ({cur_bf_pot:.1f} %)",
            f"{weight_pot - weight_ist:+.1f} kg ({cur_bf_pot - bf_ist:+.1f} % KF)",
            delta_color="inverse",
        )
        bf_pot = st.slider(
            "Ziel-Körperfettanteil",
            min_value=float(min_bf),
            max_value=float(max_bf),
            value=float(cur_bf_pot),
            step=0.5,
            format="%.1f %%",
            key=bf_slider_key,
        )
        st.session_state[bf_state_key] = float(bf_pot)

    with c_col4:
        if mode == "running":
            sec_i_rel = 1000.0 / ans_rel_0 if ans_rel_0 > 0 else 0.0
            sec_p_rel = 1000.0 / ans_rel_pot if ans_rel_pot > 0 else 0.0
            st.metric(
                "ANS (Potenzial)",
                f"{format_pace_mmss(ans_rel_pot)} min/km",
                f"{sec_p_rel - sec_i_rel:+.0f} s/km",
                delta_color="inverse",
            )
            wf_target = "ANS relativ"
        else:
            sub_c4_a, sub_c4_b = st.columns(2)
            with sub_c4_a:
                st.metric(
                    "ANS absolut",
                    f"{ans_abs_pot:.0f} W",
                    f"{ans_abs_pot - ans_abs_0:+.0f} W",
                )
            with sub_c4_b:
                st.metric(
                    "ANS relativ",
                    f"{ans_rel_pot:.2f} W/kg",
                    f"{ans_rel_pot - ans_rel_0:+.2f} W/kg",
                )
            wf_target = st.radio(
                "Ansicht Physiologischer Hebel:",
                options=["ANS relativ", "ANS absolut"],
                horizontal=True,
                key=f"{key_prefix}_{'pres' if is_presentation else 'work'}_wf_target",
            )

    # Normierungs-Funktion (0..100 Skala analog HYCYS Referenzringen)
    if mode == "running":
        ref_params = {
            "VO2max":     (60.0, 5.0,  False),
            "VLamax":     (0.55, 0.10, True),
            "Fettanteil": (13.0, 2.5,  True),
            "ANS":        (3.5,  0.4,  False),
        }
        fmt_ans_ist = f"{format_pace_mmss(ans_rel_0)}/km"
        fmt_ans_pot = f"{format_pace_mmss(ans_rel_pot)}/km"
    else:
        ref_params = {
            "VO2max":     (62.0,  6.0,  False),
            "VLamax":     (0.50,  0.10, True),
            "Fettanteil": (13.0,  2.5,  True),
            "ANS_abs":    (300.0, 35.0, False),
            "ANS_rel":    (4.1,   0.45, False),
        }
        fmt_ans_abs_ist = f"{ans_abs_0:.0f} W"
        fmt_ans_abs_pot = f"{ans_abs_pot:.0f} W"
        fmt_ans_rel_ist = f"{ans_rel_0:.2f} W/kg"
        fmt_ans_rel_pot = f"{ans_rel_pot:.2f} W/kg"

    def _norm_score(val, key):
        mean, sd, inv = ref_params[key]
        span = 6.0 * sd
        rel = (val - (mean - 3.0 * sd)) / span * 100.0 if span > 0 else 50.0
        if inv:
            rel = 100.0 - rel
        return float(max(5.0, min(98.0, rel)))

    col_wf, col_radar = st.columns([1.3, 1.0])

    # =========================================================================
    # DIAGRAMM 1 (LINKS): WASSERFALL-DIAGRAMM
    # Farben: ANS (Ist & Potenzial) in Blau/Türkis (#2cb7b9), Potenzial-Hebel in Gold (#cdb663)
    # =========================================================================
    with col_wf:
        st.markdown("#### Physiologischer Hebel")
        use_rel = (mode == "running") or (wf_target == "ANS relativ")
        if mode == "running":
            v0, v1, v2, v3 = ans_rel_0, ans_rel_1, ans_rel_2, ans_rel_3
            y_title = "ANS"
        elif use_rel:
            v0, v1, v2, v3 = ans_rel_0, ans_rel_1, ans_rel_2, ans_rel_3
            y_title = "ANS relativ [W/kg]"
        else:
            v0, v1, v2, v3 = ans_abs_0, ans_abs_1, ans_abs_2, ans_abs_3
            y_title = "ANS absolut [W]"

        step_vo2 = v1 - v0
        step_vla = v2 - v1
        step_fat = v3 - v2
        total_gain = v3 - v0

        def _fmt_delta(d_val, v_before, v_after):
            if mode == "running":
                sec_before = 1000.0 / v_before if v_before > 0 else 0.0
                sec_after = 1000.0 / v_after if v_after > 0 else 0.0
                d_sec = sec_after - sec_before
                return f"{d_sec:+.0f} s/km"
            else:
                if use_rel:
                    return f"{d_val:+.2f} W/kg"
                return f"{d_val:+.0f} W"

        def _fmt_total(v_val):
            if mode == "running":
                return f"{format_pace_mmss(v_val)} min/km"
            else:
                return f"{v_val:.2f} W/kg" if use_rel else f"{v_val:.0f} W"

        fig_wf = go.Figure(go.Waterfall(
            name="ANS Brücke",
            orientation="v",
            measure=["absolute", "relative", "relative", "relative", "total"],
            x=[
                "<b>1. ANS (Ist)</b>",
                f"<b>2. VO2max</b><br>({int(d_vo2_pct):+d} %)",
                f"<b>3. VLamax</b><br>({int(d_vla_pct):+d} %)",
                f"<b>4. Fettanteil</b><br>({bf_ist:.1f}% ➔ {bf_pot:.1f}%)",
                "<b>5. ANS (Potenzial)</b>",
            ],
            y=[v0, step_vo2, step_vla, step_fat, 0],
            text=[
                _fmt_total(v0),
                _fmt_delta(step_vo2, v0, v1),
                _fmt_delta(step_vla, v1, v2),
                _fmt_delta(step_fat, v2, v3),
                f"<b>{_fmt_total(v3)}</b><br>(Gesamt: {_fmt_delta(total_gain, v0, v3)})",
            ],
            textposition="outside",
            textinfo="text",
            hoverinfo="x+text",
            connector={"line": {"color": "#595a59", "width": 1.5, "dash": "dot"}},
            # Potenziale in Gold (#cdb663), ANS (Ist & Potenzial) in Blau/Türkis (#2cb7b9)
            increasing={"marker": {"color": "#cdb663", "line": {"color": "#9e8836", "width": 1.5}}},
            decreasing={"marker": {"color": "#595a59", "line": {"color": "#333333", "width": 1.5}}},
            totals={"marker": {"color": "#2cb7b9", "line": {"color": "#1a1a1a", "width": 2.0}}},
        ))

        min_v = min(v0, v1, v2, v3)
        max_v = max(v0, v1, v2, v3)
        span_v = max(abs(max_v - min_v), max_v * 0.12, 0.15)
        y_lo = max(0.0, min_v - span_v * 1.4)
        y_hi = max_v + span_v * 1.15

        fig_wf.update_layout(
            yaxis=dict(
                title=y_title,
                range=[y_lo, y_hi],
                gridcolor="#e6e6e6",
                showticklabels=(mode != "running"),
            ),
            xaxis=dict(title=""),
            height=430,
            margin=dict(l=45, r=30, t=30, b=40),
            plot_bgcolor="#f9f9f9",
            showlegend=False,
        )
        st.plotly_chart(fig_wf, use_container_width=True)

    # =========================================================================
    # DIAGRAMM 2 (RECHTS): SPINNENDIAGRAMM
    # Beim Laufen: 4 Variablen (VO2max, VLamax, ANS, Fettanteil - ohne Trennung ANS abs/rel)
    # Beim Radfahren: 5 Variablen (VO2max, VLamax, ANS abs. [W], ANS rel. [W/kg], Fettanteil)
    # =========================================================================
    with col_radar:
        st.markdown("#### Spinnendiagramm")
        fig_pent, ax_p = plt.subplots(figsize=(5.0, 4.8), dpi=250)
        ax_p.set_aspect("equal")

        if mode == "running":
            angles_deg = [90.0, 180.0, 270.0, 0.0]
            angles_rad = [np.deg2rad(a) for a in angles_deg]
            angles_closed = angles_rad + [angles_rad[0]]
            s_ist_ord = [
                _norm_score(rel_vo2max_ist, "VO2max"),
                _norm_score(vlamax_ist,     "VLamax"),
                _norm_score(ans_rel_0,      "ANS"),
                _norm_score(bf_ist,         "Fettanteil"),
            ]
            s_pot_ord = [
                _norm_score(rel_vo2max_pot, "VO2max"),
                _norm_score(vlamax_pot,     "VLamax"),
                _norm_score(ans_rel_pot,    "ANS"),
                _norm_score(bf_pot,         "Fettanteil"),
            ]
            s_ist_ord.append(s_ist_ord[0])
            s_pot_ord.append(s_pot_ord[0])
            labels_pent = [
                f"VO2max\n[{rel_vo2max_ist:.1f} ➔ {rel_vo2max_pot:.1f}]",
                f"VLamax\n[{vlamax_ist:.2f} ➔ {vlamax_pot:.2f}]",
                f"ANS\n[{format_pace_mmss(ans_rel_0)} ➔ {format_pace_mmss(ans_rel_pot)}/km]",
                f"Fettanteil\n[{bf_ist:.1f}% ➔ {bf_pot:.1f}%]",
            ]
        else:
            angles_deg = [90.0, 162.0, 234.0, 306.0, 18.0]
            order_idx = [0, 1, 3, 4, 2]
            angles_rad = [np.deg2rad(a) for a in angles_deg]
            angles_closed = angles_rad + [angles_rad[0]]
            scores_ist = [
                _norm_score(rel_vo2max_ist, "VO2max"),
                _norm_score(vlamax_ist,     "VLamax"),
                _norm_score(bf_ist,         "Fettanteil"),
                _norm_score(ans_abs_0,      "ANS_abs"),
                _norm_score(ans_rel_0,      "ANS_rel"),
            ]
            scores_pot = [
                _norm_score(rel_vo2max_pot, "VO2max"),
                _norm_score(vlamax_pot,     "VLamax"),
                _norm_score(bf_pot,         "Fettanteil"),
                _norm_score(ans_abs_pot,    "ANS_abs"),
                _norm_score(ans_rel_pot,    "ANS_rel"),
            ]
            s_ist_ord = [scores_ist[i] for i in order_idx] + [scores_ist[order_idx[0]]]
            s_pot_ord = [scores_pot[i] for i in order_idx] + [scores_pot[order_idx[0]]]
            labels_pent = [
                f"VO2max\n[{rel_vo2max_ist:.1f} ➔ {rel_vo2max_pot:.1f}]",
                f"VLamax\n[{vlamax_ist:.2f} ➔ {vlamax_pot:.2f}]",
                f"ANS abs.\n[{fmt_ans_abs_ist} ➔ {fmt_ans_abs_pot}]",
                f"ANS rel.\n[{fmt_ans_rel_ist} ➔ {fmt_ans_rel_pot}]",
                f"Fettanteil\n[{bf_ist:.1f}% ➔ {bf_pot:.1f}%]",
            ]

        for r_ring, col_ring, z_r in [
            (100, "#cdb663", 1),
            (80,  "#2cb7b9", 2),
            (60,  "#ffffff", 3),
            (35,  "#595a59", 4),
        ]:
            xr = [r_ring * np.cos(t) for t in angles_closed]
            yr = [r_ring * np.sin(t) for t in angles_closed]
            ax_p.fill(xr, yr, color=col_ring, edgecolor="none", zorder=z_r)

        for r_ring_line in [35, 60, 80]:
            ax_p.plot(
                [r_ring_line * np.cos(t) for t in angles_closed],
                [r_ring_line * np.sin(t) for t in angles_closed],
                color="#8c8d8c" if r_ring_line == 35 else "#b8b9b8",
                linewidth=0.6,
                zorder=5,
            )
        for t in angles_rad:
            ax_p.plot([0, 35 * np.cos(t)], [0, 35 * np.sin(t)], color="#9e9f9e", linewidth=0.8, zorder=5)
            ax_p.plot([35 * np.cos(t), 100 * np.cos(t)], [35 * np.sin(t), 100 * np.sin(t)], color="#595a59", linewidth=0.8, zorder=5)
        ax_p.plot([100 * np.cos(t) for t in angles_closed], [100 * np.sin(t) for t in angles_closed],
                  color="#595a59", linewidth=1.2, zorder=8)

        x_i = [r * np.cos(t) for r, t in zip(s_ist_ord, angles_closed)]
        y_i = [r * np.sin(t) for r, t in zip(s_ist_ord, angles_closed)]
        ax_p.plot(x_i, y_i, color="#1a1a1a", linewidth=2.6, zorder=6)
        ax_p.fill(x_i, y_i, color="#1a1a1a", alpha=0.10, zorder=5)

        # Potenzial-Polygon in Grau gestrichelt
        x_p = [r * np.cos(t) for r, t in zip(s_pot_ord, angles_closed)]
        y_p = [r * np.sin(t) for r, t in zip(s_pot_ord, angles_closed)]
        ax_p.plot(x_p, y_p, color="#595a59", linewidth=2.4, linestyle=(0, (4, 2)), zorder=7)

        for lbl, t in zip(labels_pent, angles_rad):
            r_lbl = 113.0
            lx, ly = r_lbl * np.cos(t), r_lbl * np.sin(t)
            ha = "center" if abs(np.cos(t)) < 0.25 else ("left" if np.cos(t) > 0 else "right")
            va = "bottom" if np.sin(t) > 0.3 else ("top" if np.sin(t) < -0.3 else "center")
            ax_p.text(
                lx, ly, lbl, size=7.2, fontweight="bold", color="#1a1a1a", ha=ha, va=va,
                zorder=10
            )

        legend_el = [
            Patch(facecolor="#595a59", label="Schwach"),
            Patch(facecolor="#ffffff", edgecolor="#cccccc", label="Zu verbessern"),
            Patch(facecolor="#2cb7b9", label="Gut"),
            Patch(facecolor="#cdb663", label="Sehr gut"),
            Line2D([0], [0], color="#1a1a1a", lw=2.5, label="Athlet (Ist)"),
            Line2D([0], [0], color="#595a59", lw=2.2, ls="--", label="Potenzial"),
        ]
        ax_p.legend(handles=legend_el, loc="upper center", bbox_to_anchor=(0.5, -0.04), ncol=3, frameon=False, fontsize=6.8)
        ax_p.set_xlim(-160, 160)
        ax_p.set_ylim(-140, 140)
        ax_p.axis("off")
        plt.tight_layout()
        st.pyplot(fig_pent)
        plt.close(fig_pent)

    return {
        "weight_pot": weight_pot,
        "bf_pot": bf_pot,
        "abs_vo2max_pot": abs_vo2max_pot,
        "rel_vo2max_pot": rel_vo2max_pot,
        "vlamax_pot": vlamax_pot,
        "ans_abs_pot": ans_abs_pot,
        "ans_rel_pot": ans_rel_pot,
    }


def sync_lauftest_input_df():
    anzahl = int(st.session_state.anzahl)
    start_v = float(st.session_state.start_v)
    v_increment = float(st.session_state.v_increment)
    
    stufe_names = ["Ruhe"] + [f"Stufe {i+1}" for i in range(anzahl)]
    speeds = [0.0] + [start_v + (i * v_increment) for i in range(anzahl)]
    paces = ["-"] + [format_pace_mmss(s) for s in speeds[1:]]
    
    default_ruhe_lac = 1.0
    default_ruhe_hr = 70
    default_lac = [1.03, 1.09, 1.19, 1.42, 2.36, 2.32, 6.05, 9.48]
    default_hr = [123, 137, 145, 154, 164, 170, 177, 183]
    
    target_len = anzahl + 1
    
    if "df_lauftest_input" not in st.session_state:
        st.session_state.df_lauftest_input = pd.DataFrame({
            "Stufe": stufe_names,
            "v (m/s)": np.round(speeds, 2),
            "v (min/km)": paces,
            "Laktat": [None] * target_len,
            "HF": [None] * target_len
        })
        st.session_state.last_sync_params = (anzahl, start_v, v_increment)
    else:
        df = st.session_state.df_lauftest_input
        current_len = len(df)
        
        expected_speeds = [0.0] + [round(start_v + (i * v_increment), 2) for i in range(anzahl)]
        current_speeds = [round(float(s), 2) if pd.notna(s) else 0.0 for s in df["v (m/s)"].values]
        
        if current_len != target_len or current_speeds != expected_speeds or "v (min/km)" not in df.columns or "v (km/h)" in df.columns:
            if "lauftest_lac_editor_key" in st.session_state:
                del st.session_state["lauftest_lac_editor_key"]
            if current_len != target_len:
                if target_len > current_len:
                    new_rows = []
                    for i in range(current_len, target_len):
                        new_rows.append({
                            "Stufe": stufe_names[i],
                            "v (m/s)": round(speeds[i], 2),
                            "v (min/km)": paces[i],
                            "Laktat": None,
                            "HF": None
                        })
                    df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
                else:
                    df = df.iloc[:target_len].copy()
            
            if "v (km/h)" in df.columns:
                df = df.drop(columns=["v (km/h)"])
            df["Stufe"] = stufe_names
            df["v (m/s)"] = np.round(speeds, 2)
            df["v (min/km)"] = paces
            df = df[["Stufe", "v (m/s)", "v (min/km)", "Laktat", "HF"]]
            st.session_state.df_lauftest_input = df
            st.session_state.last_sync_params = (anzahl, start_v, v_increment)

# Primary Sportart Selection (must be the very first sidebar widget)
_active_coach_user = st.session_state.get("logged_in_coach", st.session_state.get("coach", "Markus Hertlein"))
st.sidebar.markdown(
    f"<div style='font-size:12.5px; color:#595a59; background:#f4f8f8; padding:7px 10px; border-radius:5px; border-left:4px solid #2cb7b9; margin-bottom:8px;'>"
    f"👤 <b>Coach:</b> {_active_coach_user}"
    f"</div>",
    unsafe_allow_html=True,
)


def render_sidebar_account_and_remote_support(key_suffix: str = "main"):
    """Rendert ganz unten in der Sidebar (unter 'Auswertung starten') den Account-Bereich inkl. Session-ID & Remote-Support."""
    _cur_coach = st.session_state.get("logged_in_coach", st.session_state.get("coach", "Markus Hertlein"))
    _cur_sid = st.session_state.get("_hycys_sid", "-")
    with st.sidebar.expander("🔑 Account & Remote-Support", expanded=False):
        st.markdown(
            f"<div style='font-size:11.5px; color:#595a59; margin-bottom:8px;'>"
            f"<b>Eingeloggt als:</b> {_cur_coach}<br>"
            f"<b>Session-ID:</b> <code>{_cur_sid}</code>"
            f"</div>",
            unsafe_allow_html=True,
        )
        if st.button("Abmelden / Coach wechseln", use_container_width=True, key=f"btn_logout_coach_{key_suffix}"):
            st.session_state["hycys_password_authenticated"] = False
            try:
                if "sid" in st.query_params:
                    del st.query_params["sid"]
            except Exception:
                pass
            st.rerun()

        if _cur_coach in ADMIN_COACHES:
            st.markdown("---")
            st.markdown("**🛠️ Remote-Support (Aktive Coach-Sitzungen):**")
            _store_admin = _get_persistent_session_store()
            _active_sessions_list = []
            for _k_s, _v_s in sorted(_store_admin.items(), key=lambda x: x[1].get("last_active", 0), reverse=True):
                _c_u = _v_s.get("coach_user", "Unbekannt")
                _ath_u = _v_s.get("athlete_name", "-")
                _sp_u = _v_s.get("sportart", "-")
                _mins_ago = int(round((time.time() - _v_s.get("last_active", time.time())) / 60.0))
                _lbl_s = f"{_c_u} | {_ath_u} ({_sp_u}) – vor {_mins_ago} Min. [{_k_s[:6]}]"
                _active_sessions_list.append((_lbl_s, _k_s))

            if _active_sessions_list:
                _sel_lbl = st.selectbox(
                    "Aktive Sitzung auswählen",
                    options=[x[0] for x in _active_sessions_list],
                    key=f"remote_support_select_session_{key_suffix}",
                )
                _sel_sid = dict(_active_sessions_list).get(_sel_lbl)
                if st.button("In diese Coach-Sitzung schalten", type="primary", use_container_width=True, key=f"btn_remote_support_join_{key_suffix}"):
                    if _sel_sid and load_session_from_store(_sel_sid, keep_logged_in_coach=_cur_coach):
                        st.rerun()
            else:
                st.caption("Aktuell keine weiteren Coach-Sitzungen aktiv.")

        st.markdown("---")
        _manual_sid_in = st.text_input("Oder Session-ID direkt öffnen:", placeholder="z. B. a8f3c9...", key=f"remote_support_manual_sid_{key_suffix}")
        if st.button("Session-ID laden", use_container_width=True, key=f"btn_remote_support_load_sid_{key_suffix}"):
            _clean_sid = _manual_sid_in.strip().split("sid=")[-1].split("&")[0].strip()
            if _clean_sid and load_session_from_store(_clean_sid, keep_logged_in_coach=_cur_coach):
                st.rerun()
            else:
                st.error("Session-ID nicht gefunden oder abgelaufen.")


sportart_options = ["Bitte wählen...", "Laufen", "Radsport"]
sportart_val = st.session_state.sportart if st.session_state.sportart in sportart_options else "Bitte wählen..."
sportart = st.sidebar.selectbox("Sportart", options=sportart_options, index=sportart_options.index(sportart_val))
st.session_state.sportart = sportart

# Dynamically set title based on sportart
if sportart == "Radsport":
    title_text = "HYCYS Radsportauswertung"
elif sportart == "Laufen":
    title_text = "HYCYS Laufauswertung"
else:
    title_text = "HYCYS Diagnostik"

col_title, col_logo = st.columns([4, 1])
with col_title:
    st.title(title_text)
with col_logo:
    logo_candidates = ["image_822d59.png", "logo.png", "image_828f70.png"]
    logo_path = None
    for candidate in logo_candidates:
        if os.path.exists(candidate):
            st.image(candidate, width='stretch')
            logo_path = candidate
            break

st.markdown("---")

# 2. SEITENLEISTE
# Initialize widget keys in session state so they match general session state defaults
widget_keys_to_init = {
    "athlete_name": st.session_state.athlete_name,
    "birthdate": st.session_state.birthdate,
    "test_date": st.session_state.test_date,
    "height": int(st.session_state.height),
    "weight": float(st.session_state.weight),
    "body_fat_pct": float(st.session_state.body_fat_pct),
    "kategorie": st.session_state.kategorie,
    "coach": st.session_state.coach,
    
    "name_lauft_key": st.session_state.athlete_name,
    "birth_lauft_key": st.session_state.birthdate,
    "test_lauft_key": st.session_state.test_date,
    "height_lauft_key": int(st.session_state.height),
    "weight_lauft_key": float(st.session_state.weight),
    "bf_lauft_key": float(st.session_state.body_fat_pct),
    "kategorie_lauft_key": st.session_state.kategorie,
    "coach_lauft_key": st.session_state.coach,
    "gender": st.session_state.gender,
    "gender_lauft_key": st.session_state.gender,
}
for key, val in widget_keys_to_init.items():
    if key not in st.session_state:
        st.session_state[key] = val

def update_base_data(name, height, weight, birthdate, test_date, gender=None):
    if name:
        st.session_state.athlete_name = name
        st.session_state.name_lauft_key = name
        if st.session_state.get("sportart") == "Laufen":
            st.session_state.athlete_name_lauft_widget = name
    if height is not None:
        st.session_state.height = int(height)
        st.session_state.height_lauft_key = int(height)
        if st.session_state.get("sportart") == "Laufen":
            st.session_state.height_lauft_widget = int(height)
    if weight is not None:
        w_rounded = round(float(weight), 1)
        st.session_state.weight = w_rounded
        st.session_state.weight_lauft_key = w_rounded
        if st.session_state.get("sportart") == "Laufen":
            st.session_state.weight_lauft_widget = w_rounded
    if birthdate:
        st.session_state.birthdate = birthdate
        st.session_state.birth_lauft_key = birthdate
        if st.session_state.get("sportart") == "Laufen":
            st.session_state.birthdate_lauft_widget = birthdate
    if test_date:
        st.session_state.test_date = test_date
        st.session_state.test_lauft_key = test_date
        if st.session_state.get("sportart") == "Laufen":
            st.session_state.test_date_lauft_widget = test_date
    if gender:
        st.session_state.gender = gender
        st.session_state.gender_lauft_key = gender
        if st.session_state.get("sportart") == "Laufen":
            st.session_state.gender_lauft_widget = gender

if "df_sf" not in st.session_state:
    sf_names = ["Wange", "Kinn", "Achselfalte vorn", "10. Rippe", "Bauch (Nabel)", "Spina illiaca", "Oberschenkel", "Rücken", "Triceps", "Wade"]
    st.session_state.df_sf = pd.DataFrame({"Falte": sf_names, "M1": [0.0]*10, "M2": [0.0]*10, "M3": [0.0]*10})

# Sportart selection handled at startup
pass

# ------------------ CONDITIONAL SPORTART RENDERING ------------------
if sportart == "Bitte wählen...":
    render_sidebar_account_and_remote_support("init")
    st.info("Bitte wählen Sie zuerst eine Sportart in der Sidebar aus, um mit der Auswertung zu beginnen.")
    save_current_session_snapshot()
    st.stop()

elif sportart == "Radsport":
    # Pre-populate cycling constants and sprint/energetics parameters
    if "ppo_val" not in st.session_state: st.session_state.ppo_val = 436.0
    if "hfmax_val" not in st.session_state: st.session_state.hfmax_val = 179.0
    if "slope_val" not in st.session_state: st.session_state.slope_val = 0.11
    if "intercept_val" not in st.session_state: st.session_state.intercept_val = 138.0
    if "vo2max_spiro" not in st.session_state: st.session_state.vo2max_spiro = 4292.2
    if "t_bel" not in st.session_state: st.session_state.t_bel = 13.2
    if "t_alak" not in st.session_state: st.session_state.t_alak = 3.1
    if "t_glyc" not in st.session_state: st.session_state.t_glyc = 10.1
    if "t_bel_auto" not in st.session_state: st.session_state.t_bel_auto = 13.2
    if "t_alak_auto" not in st.session_state: st.session_state.t_alak_auto = 3.1
    if "t_bel_manual" not in st.session_state: st.session_state.t_bel_manual = 13.2
    if "t_alak_manual" not in st.session_state: st.session_state.t_alak_manual = 3.1
    if "rad_auswertung_gestartet" not in st.session_state: st.session_state.rad_auswertung_gestartet = False
    if "carb_intake_factor" not in st.session_state: st.session_state.carb_intake_factor = 1.0
    
    if "ks1" not in st.session_state: st.session_state.ks1 = 0.0631
    if "ks2" not in st.session_state: st.session_state.ks2 = 1.331
    if "ks4" not in st.session_state: st.session_state.ks4 = 11.7
    if "lac_eq" not in st.session_state: st.session_state.lac_eq = 0.02049
    if "lac_dist" not in st.session_state: st.session_state.lac_dist = 0.4
    if "vo2_rest" not in st.session_state: st.session_state.vo2_rest = 4.0

    coach_options = list(HYCYS_COACH_LIST)
    kategorie_options = ["Hobby", "Age-Grouper", "Amateur", "Profi"]

    # 0. File Uploaders placed at the very top of the sidebar
    uploaded_sprint = st.sidebar.file_uploader(
        "Sprintdatei hochladen (.xlsx, .srm, .fit)", 
        type=["xlsx", "srm", "fit"], 
        key="sprint_file_key"
    )
    uploaded_spiro = st.sidebar.file_uploader(
        "Spirodatei VO2max Rampe (10s) hochladen (.xlsx, .csv)", 
        type=["xlsx", "csv"], 
        key="spiro_rampe_key"
    )
    uploaded_fit = st.sidebar.file_uploader(
        "VO2max Rampe (.fit) hochladen", 
        type=["fit"], 
        key="fit_rampe_key"
    )

    st.sidebar.caption(
        "Original SRM-Ergometer Datei (.srm), Garmin/SRM Export (.fit) oder Sprint-Excel-Datei (.xlsx) hochladen. "
        "Power [W], Zeit und Trittfrequenz werden automatisch ausgelesen."
    )

    if uploaded_sprint is not None and uploaded_sprint.name != st.session_state.get("last_uploaded_sprint"):
        st.session_state.last_uploaded_sprint = uploaded_sprint.name
        _ext = os.path.splitext(uploaded_sprint.name)[1].lower()
        
        try:
            if _ext == ".xlsx":
                # Excel-Datei parsen
                sprint_powers, sprint_times, sprint_cadences = parse_sprint_excel(uploaded_sprint)
                _src_name = uploaded_sprint.name
                _format_info = "Excel"
                # Rohdaten für Download löschen
                if "sprint_records" in st.session_state:
                    del st.session_state["sprint_records"]
                if "last_uploaded_sprint_raw" in st.session_state:
                    del st.session_state["last_uploaded_sprint_raw"]
            else:
                # SRM/FIT-Datei parsen
                _suffix = _ext
                with tempfile.NamedTemporaryFile(delete=False, suffix=_suffix) as tmp_raw:
                    tmp_raw.write(uploaded_sprint.getvalue())
                    tmp_raw_path = tmp_raw.name
                
                try:
                    import importlib.util as _ilu
                    _spec = _ilu.spec_from_file_location(
                        "convert_sprint", 
                        os.path.join(os.path.dirname(os.path.abspath(__file__)), "convert_sprint_to_excel.py")
                    )
                    _conv = _ilu.module_from_spec(_spec)
                    _spec.loader.exec_module(_conv)
                    
                    _slope  = float(st.session_state.get("srm_slope", 1.0))
                    _offset = float(st.session_state.get("srm_offset", 0.0))
                    
                    if _ext == ".srm":
                        with open(tmp_raw_path, "rb") as _f:
                            _srm_data = _f.read()
                        _records, _hdr = _conv.parse_srm7_data(_srm_data, slope=_slope, zero_offset=_offset, trim_zeros=True)
                        _rec_int_label = f"{_hdr['rec_int']}s (SRM)"
                        _src_name = uploaded_sprint.name
                        _srm_ath = _hdr.get("athlete_name", "").strip()
                        _srm_dt = _hdr["record_date"].strftime("%d.%m.%Y") if _hdr.get("record_date") else None
                        if _srm_ath or _srm_dt:
                            update_base_data(_srm_ath or None, None, None, None, _srm_dt)
                    else:  # .fit
                        _records, _ath, _date = _conv.parse_fit_file(tmp_raw_path, trim_zeros=True)
                        _rec_int_label = "1s (FIT)"
                        _src_name = uploaded_sprint.name
                        if _ath or _date:
                            update_base_data(_ath or None, None, None, None, _date or None)
                    
                    if _records:
                        sprint_powers = [float(r.get("power_w", r.get("power", 0))) for r in _records]
                        sprint_times  = [float(r["elapsed_s"]) for r in _records]
                        sprint_cadences = [float(r.get("cadence", 0)) for r in _records]
                        st.session_state.sprint_records = _records
                        st.session_state.last_uploaded_sprint_raw = uploaded_sprint.name
                        st.session_state.sprint_rec_int_label = _rec_int_label
                        _format_info = "0.1s (SRM)" if _ext == ".srm" else "1s (FIT)"
                    else:
                        raise ValueError("Keine Leistungsdaten gefunden.")
                finally:
                    if os.path.exists(tmp_raw_path):
                        os.unlink(tmp_raw_path)
            
            if len(sprint_powers) > 0:
                p_max = max(sprint_powers)
                t_pmax_list = [t for p, t in zip(sprint_powers, sprint_times) if p == p_max]
                t_pmax_last = max(t_pmax_list) if t_pmax_list else 0.0
                t_alak = None
                p_threshold = p_max * 0.965
                for p, t in zip(sprint_powers, sprint_times):
                    if t > t_pmax_last and p < p_threshold:
                        t_alak = t
                        break
                if t_alak is None:
                    times_after_pmax = [t for t in sprint_times if t > t_pmax_last]
                    t_alak = times_after_pmax[-1] if times_after_pmax else (t_pmax_last + 0.1)
                t_bel, _plateau_start_idx = detect_sprint_end(sprint_powers, sprint_times)
                st.session_state.sprint_plateau_start_idx = _plateau_start_idx
                
                st.session_state.t_bel_auto = round(t_bel, 1)
                st.session_state.t_alak_auto = round(t_alak, 1)
                st.session_state.t_bel_manual = round(t_bel, 1)
                st.session_state.t_alak_manual = round(t_alak, 1)
                st.session_state.t_bel = round(t_bel, 1)
                st.session_state.t_alak = round(t_alak, 1)
                st.session_state.t_glyc = round(t_bel - t_alak, 1)
                st.session_state.sprint_powers = sprint_powers
                st.session_state.sprint_times = sprint_times
                st.session_state.sprint_cadences = sprint_cadences
                st.session_state.sprint_src_name = _src_name
                st.session_state.rad_auswertung_gestartet = False
                
                _plateau_info = (
                    f" Leistungsplateau am Ende erkannt — Sprintdauer bis {t_bel:.1f}s gewertet (Aufzeichnung {max(sprint_times):.1f}s)."
                    if _plateau_start_idx is not None else ""
                )
                st.success(f"✅ Geladen [{_format_info}]: {len(sprint_powers)} Punkte, Max {int(p_max)}W, Dauer {t_bel:.1f}s — Nullwerte am Anfang automatisch entfernt.{_plateau_info}")
                st.rerun()
            else:
                st.warning("Keine Leistungsdaten gefunden.")
        except Exception as e:
            st.error(f"Fehler beim Auslesen der Sprintdatei: {e}")

    # Converted Sprint Excel download button has been moved below the form.

    # Spiro uploader moved to the top of sidebar
    if uploaded_spiro is not None and uploaded_spiro.name != st.session_state.get("last_uploaded_spiro"):
        st.session_state.last_uploaded_spiro = uploaded_spiro.name
        try:
            df_spiro_excel = pd.read_excel(uploaded_spiro, header=None)
            lastname = df_spiro_excel.iloc[1, 1] if df_spiro_excel.shape[0] > 1 and df_spiro_excel.shape[1] > 1 else ""
            firstname = df_spiro_excel.iloc[2, 1] if df_spiro_excel.shape[0] > 2 and df_spiro_excel.shape[1] > 1 else ""
            name_val = f"{firstname} {lastname}".strip()
            
            h_val = None
            if df_spiro_excel.shape[0] > 5 and df_spiro_excel.shape[1] > 1:
                try: h_val = int(float(df_spiro_excel.iloc[5, 1]))
                except: pass
            
            w_val = None
            if df_spiro_excel.shape[0] > 6 and df_spiro_excel.shape[1] > 1:
                try: w_val = round(float(df_spiro_excel.iloc[6, 1]), 1)
                except: pass
                
            b_val = None
            if df_spiro_excel.shape[0] > 7 and df_spiro_excel.shape[1] > 1:
                try:
                    raw_birth = df_spiro_excel.iloc[7, 1]
                    birth_str = str(int(float(raw_birth))).zfill(8)
                    b_val = f"{birth_str[:2]}.{birth_str[2:4]}.{birth_str[4:]}"
                except: pass
                
            t_val = None
            if df_spiro_excel.shape[0] > 0 and df_spiro_excel.shape[1] > 4:
                try:
                    raw_test = df_spiro_excel.iloc[0, 4]
                    test_str = str(int(float(raw_test))).zfill(8)
                    t_val = f"{test_str[:2]}.{test_str[2:4]}.{test_str[4:]}"
                except: pass
                
            update_base_data(name_val, h_val, w_val, b_val, t_val)
                
            times_sec = df_spiro_excel.iloc[3:, 9].apply(parse_time_to_seconds)
            vo2_col = pd.to_numeric(df_spiro_excel.iloc[3:, 13], errors='coerce')
            vco2_col = pd.to_numeric(df_spiro_excel.iloc[3:, 14], errors='coerce') if df_spiro_excel.shape[1] > 14 else pd.Series(np.nan, index=df_spiro_excel.iloc[3:].index)
            eeh_col = pd.to_numeric(df_spiro_excel.iloc[3:, 36], errors='coerce') if df_spiro_excel.shape[1] > 36 else pd.Series(np.nan, index=df_spiro_excel.iloc[3:].index)
            fett_kcal_h_col = (pd.to_numeric(df_spiro_excel.iloc[3:, 41], errors='coerce') / 24.0) if df_spiro_excel.shape[1] > 41 else pd.Series(np.nan, index=df_spiro_excel.iloc[3:].index)
            kh_g_h_col = ((pd.to_numeric(df_spiro_excel.iloc[3:, 42], errors='coerce') / 24.0) / 4.063) if df_spiro_excel.shape[1] > 42 else pd.Series(np.nan, index=df_spiro_excel.iloc[3:].index)
            df_spiro_data = pd.DataFrame({
                'Time': times_sec,
                'VO2': vo2_col,
                'VCO2': vco2_col,
                'EEh': eeh_col,
                'FETT_kcal_h': fett_kcal_h_col,
                'KH_g_h': kh_g_h_col,
            }).dropna(subset=['Time', 'VO2'])
            st.session_state.rad_spiro_df = df_spiro_data
            if len(df_spiro_data) >= 2:
                t_diffs = df_spiro_data['Time'].diff().dropna()
                interval = int(round(t_diffs.median()))
            else:
                interval = 10
            window_pts = max(1, int(30 / interval))
            vo2_smooth = df_spiro_data['VO2'].rolling(window=window_pts, center=True).mean()
            st.session_state.vo2max_spiro = round(vo2_smooth.max(), 1)
            st.session_state.rad_auswertung_gestartet = False
            st.rerun()
        except Exception as e:
            st.error(f"Fehler beim Auslesen der Spirodatei: {e}")

    # FIT Rampen-Uploader moved to the top of sidebar
    if uploaded_fit is not None and uploaded_fit.name != st.session_state.get("last_uploaded_fit"):
        st.session_state.last_uploaded_fit = uploaded_fit.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".fit") as tmp:
            tmp.write(uploaded_fit.getvalue())
            tmp_path = tmp.name
        try:
            import fitparse
            fitfile = fitparse.FitFile(tmp_path)
            records_fit = []
            for record in fitfile.get_messages("record"):
                vals = record.get_values()
                ts = vals.get("timestamp")
                p = vals.get("power")
                hr = vals.get("heart_rate")
                if ts is not None:
                    records_fit.append({
                        "timestamp": ts,
                        "Power": float(p) if p is not None else 0.0,
                        "HR": float(hr) if hr is not None else np.nan
                    })
            
            if records_fit:
                df_raw = pd.DataFrame(records_fit)
                df_raw = df_raw.sort_values("timestamp")
                df_raw = df_raw.drop_duplicates(subset=["timestamp"], keep="last")
                df_raw.set_index("timestamp", inplace=True)
                df_fit = df_raw.resample("1s").ffill()
                df_fit = df_fit.reset_index()
            else:
                df_fit = pd.DataFrame(columns=["timestamp", "Power", "HR"])

            if not df_fit.empty:
                df_fit['Power_30s'] = df_fit['Power'].rolling(window=30, center=False).mean()
                max_p30 = df_fit['Power_30s'].max()
                if pd.isna(max_p30):
                    max_p30 = 0.0
                valid_hrs = df_fit['HR'].dropna().tolist()
                hf_max = max(valid_hrs) if valid_hrs else 180.0
            else:
                max_p30 = 0.0
                hf_max = 180.0
            
            # Only calculate regression once heart rate starts to rise significantly (excluding warmup and lag)
            # In the 3x3min warmup protocol, maximum power is ~150W; constant rise stabilizes around 215W.
            reg_df = df_fit[(df_fit['Power'] >= 215) & (~df_fit['HR'].isna())]
            if len(reg_df) >= 2:
                slope, intercept = np.polyfit(reg_df['Power'], reg_df['HR'], 1)
            else:
                slope, intercept = 0.15, 100.0
                
            st.session_state.ppo_val = float(int(round(max_p30)))
            st.session_state.hfmax_val = float(int(round(hf_max)))
            st.session_state.slope_val = round(slope, 2)
            st.session_state.intercept_val = round(intercept, 1)
            st.session_state.rad_auswertung_gestartet = False
            st.rerun()
        except Exception as e:
            st.error(f"Fehler beim Auslesen der Rampen-Datei: {e}")
        finally:
            if os.path.exists(tmp_path): os.unlink(tmp_path)

    # Pre-calculate body fat percentage from skinfolds to avoid rerun loops
    if "df_sf" in st.session_state:
        sf_means = st.session_state.df_sf[["M1", "M2", "M3"]].mean(axis=1)
        sum_sf = sf_means.sum()
        if sum_sf > 0:
            calculated_bf = (22.32 * np.log10(sum_sf)) - 29.2
            st.session_state.body_fat_pct = max(0.0, round(calculated_bf, 1))
            
    with st.sidebar.form("cycling_form", enter_to_submit=False):
        # 1. Basisdaten Expander (Default collapsed)
        with st.expander("Basisdaten", expanded=False):
            athlete_name = st.text_input("Name", value=st.session_state.athlete_name)
            st.session_state.athlete_name = athlete_name

            birthdate = st.text_input("Geburtsdatum", value=st.session_state.birthdate)
            st.session_state.birthdate = birthdate

            test_date = st.text_input("Testdatum", value=st.session_state.test_date)
            st.session_state.test_date = test_date

            height = st.number_input("Größe (cm)", value=int(st.session_state.height), step=1, format="%d")
            st.session_state.height = height

            weight = st.number_input("Gewicht (kg)", value=float(st.session_state.weight), step=0.1, format="%.1f")
            st.session_state.weight = round(weight, 1)

            coach_val = st.session_state.coach
            if coach_val not in coach_options:
                coach_options = coach_options + [coach_val]
            coach = st.selectbox("Coach", options=coach_options, index=coach_options.index(coach_val))
            st.session_state.coach = coach

            kategorie_val = st.session_state.kategorie
            if kategorie_val not in kategorie_options:
                kategorie_options = kategorie_options + [kategorie_val]
            kategorie = st.selectbox("Kategorie", options=kategorie_options, index=kategorie_options.index(kategorie_val))
            st.session_state.kategorie = kategorie

            gender_options = ["männlich", "weiblich"]
            gender_val = st.session_state.get("gender", "männlich")
            gender = st.selectbox("Geschlecht", options=gender_options, index=gender_options.index(gender_val))
            st.session_state.gender = gender

            diag_type_options = ["Cycling WHITE", "Cycling BLUE", "Cycling GOLD"]
            diag_type_val = st.session_state.get("diagnostik_type", "Cycling BLUE")
            if diag_type_val not in diag_type_options:
                diag_type_val = "Cycling BLUE"
            diagnostik_type = st.selectbox("Diagnostik Typ", options=diag_type_options, index=diag_type_options.index(diag_type_val))
            st.session_state.diagnostik_type = diagnostik_type

        # 2. Körperfettwerte Expander (eigenes Feld unter Basisdaten und oberhalb von Sprinttest Laktatwerte)
        with st.expander("Körperfettwerte", expanded=False):
            body_fat_pct = st.number_input("Körperfett (%)", value=float(st.session_state.body_fat_pct), step=0.1, format="%.1f")
            st.session_state.body_fat_pct = round(body_fat_pct, 1)

            st.caption("Körperfettmessung (Parizkova 10-Falten):")
            if "df_sf" not in st.session_state:
                sf_names = ["Wange", "Kinn", "Achselfalte vorn", "10. Rippe", "Bauch (Nabel)", "Spina illiaca", "Oberschenkel", "Rücken", "Triceps", "Wade"]
                st.session_state.df_sf = pd.DataFrame({"Falte": sf_names, "M1": [0.0]*10, "M2": [0.0]*10, "M3": [0.0]*10})
            edited_sf = st.data_editor(
                st.session_state.df_sf,
                hide_index=True,
                width='stretch',
                key="edited_sf_rad_key"
            )

        # 3. Sprinttest Laktatwerte Expander (Default collapsed)
        with st.expander("Sprinttest Laktatwerte", expanded=False):
            # Laktatwerte Sprint
            st.subheader("Laktatwerte Sprint")
            if "df_sprint_lac" not in st.session_state:
                times = ["Ruhelaktat 1", "Ruhelaktat 2", "Ruhelaktat 3", "Nachbelastung Min 0", "Nachbelastung Min 2", "Nachbelastung Min 3", "Nachbelastung Min 4", "Nachbelastung Min 5", "Nachbelastung Min 6", "Nachbelastung Min 7", "Nachbelastung Min 8", "Nachbelastung Min 9", "Nachbelastung Min 10"]
                default_lacs = [0.0] * len(times)
                st.session_state.df_sprint_lac = pd.DataFrame({
                    "Messpunkt": times,
                    "Laktat [mmol/L]": default_lacs
                })
                
            edited_sprint_lac = st.data_editor(
                st.session_state.df_sprint_lac,
                disabled=["Messpunkt"],
                hide_index=True,
                use_container_width=True,
                height=500,
                column_config={
                    "Messpunkt": st.column_config.TextColumn("Messpunkt", width="medium"),
                    "Laktat [mmol/L]": st.column_config.NumberColumn("Laktat", min_value=0.0, max_value=25.0, step=0.01, format="%.2f", width="small")
                },
                key="sprint_lac_editor_key"
            )
        
            st.subheader("Sprinttest Parameter")
            t_bel_auto = float(st.session_state.get("t_bel_auto", 13.2))
            t_alak_auto = float(st.session_state.get("t_alak_auto", 3.1))
            t_glyc_auto = round(t_bel_auto - t_alak_auto, 1)
            
            t_bel_manual = float(st.session_state.get("t_bel_manual", 13.2))
            t_alak_manual = float(st.session_state.get("t_alak_manual", 3.1))
            t_glyc_manual = round(t_bel_manual - t_alak_manual, 1)
            
            st.markdown(f"""
            | Parameter | Auto (Excel) | Manuell |
            | :--- | :--- | :--- |
            | **t_bel [s]** | {t_bel_auto:.1f} | {t_bel_manual:.1f} |
            | **t_alak [s]** | {t_alak_auto:.1f} | {t_alak_manual:.1f} |
            | **t_glyc [s]** | {t_glyc_auto:.1f} | {t_glyc_manual:.1f} |
            """)
            
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                t_bel_manual_input = st.number_input("Manuelle t_bel [s]", value=float(st.session_state.t_bel_manual), step=0.1, format="%.1f")
                st.session_state.t_bel_manual = round(t_bel_manual_input, 1)
            with col_t2:
                t_alak_manual_input = st.number_input("Manuelle t_alak [s]", value=float(st.session_state.t_alak_manual), step=0.1, format="%.1f")
                st.session_state.t_alak_manual = round(t_alak_manual_input, 1)
                    
            sprint_source = st.selectbox(
                "Aktive Quelle für Sprinttest",
                ["Automatisch aus Excel-Rohdaten", "Manuelle Anpassung"],
                key="sprint_source_key"
            )

        # 3. Rampentest Expander (Default collapsed)
        with st.expander("Rampentest", expanded=False):
            st.subheader("Rampentest Parameter")
            ppo_val_input = st.number_input("30s PPO [W]", value=int(round(float(st.session_state.ppo_val))), step=1, format="%d")
            st.session_state.ppo_val = float(ppo_val_input)
            
            hfmax_val_input = st.number_input("max. Herzfrequenz [bpm]", value=int(round(float(st.session_state.hfmax_val))), step=1, format="%d")
            st.session_state.hfmax_val = float(hfmax_val_input)
            
            slope_val_input = st.number_input("Steigung lin. Reg. [bpm/W]", value=round(float(st.session_state.slope_val), 2), step=0.01, format="%.2f")
            st.session_state.slope_val = round(float(slope_val_input), 2)
            
            intercept_val_input = st.number_input("Nullstelle lin. Reg. [bpm]", value=float(st.session_state.intercept_val), step=1.0, format="%.1f")
            st.session_state.intercept_val = round(intercept_val_input, 1)

            st.subheader("VO2max Quellenvergleich")
            weight = float(st.session_state.get("weight", 72.3))
            ppo_val = float(st.session_state.get("ppo_val", 436.0))
            spiro_abs = float(st.session_state.get("vo2max_spiro", 4292.2))
            spiro_rel = spiro_abs / weight if weight > 0 else 0.0
            m_abs = ppo_val * 9.44 + 592.0
            m_rel = m_abs / weight if weight > 0 else 0.0
            w_abs = ppo_val * 9.44 + 275.0
            w_rel = w_abs / weight if weight > 0 else 0.0
            override_abs = float(st.session_state.get("vo2_override", 4368.0))
            override_rel = override_abs / weight if weight > 0 else 0.0
            
            st.markdown(f"""
            | Quelle | Abs. [ml] | Rel. [ml/kg] |
            | :--- | :--- | :--- |
            | **Spiro (gemessen)** | {int(spiro_abs)} | {spiro_rel:.1f} |
            | **PPO (♂)** | {int(m_abs)} | {m_rel:.1f} |
            | **PPO (♀)** | {int(w_abs)} | {w_rel:.1f} |
            | **Manuell (override)** | {int(override_abs)} | {override_rel:.1f} |
            """)
            
            vo2_override_input = st.number_input("Manuelle VO2max [ml/min]", value=int(round(float(st.session_state.vo2_override))), step=50, format="%d")
            st.session_state.vo2_override = float(vo2_override_input)
            
            vo2_source = st.selectbox(
                "Aktive Quelle für Simulation",
                ["Gemessene Spirodatei", "Berechnet aus 30s PPO (männlich)", "Berechnet aus 30s PPO (weiblich)", "Manuelle Eingabe"],
                key="vo2_source_rad_key"
            )
            if vo2_source == "Manuelle Eingabe":
                vo2_override = float(st.session_state.vo2_override)
            else:
                vo2_override = 0.0

        st.markdown("---")
        start_button_rad = st.form_submit_button("Auswertung starten", type="primary")
        if start_button_rad:
            st.session_state.df_sprint_lac = edited_sprint_lac
            st.session_state.df_sf = edited_sf
            _sf_sum_rad = float(edited_sf[["M1", "M2", "M3"]].mean(axis=1).sum()) if edited_sf is not None and not edited_sf.empty else 0.0
            if _sf_sum_rad > 0:
                _calc_bf_rad = max(0.0, round((22.32 * np.log10(_sf_sum_rad)) - 29.2, 1))
                st.session_state.body_fat_pct = _calc_bf_rad
            else:
                _calc_bf_rad = float(st.session_state.get("body_fat_pct", 0.0))

            if _calc_bf_rad <= 0.0 and _sf_sum_rad <= 0.0:
                st.session_state.rad_auswertung_gestartet = False
                st.session_state.rad_kf_warnung_offen = True
                st.session_state.rad_kf_dialog_trigger = True
            else:
                st.session_state.rad_kf_warnung_offen = False
                st.session_state.rad_kf_dialog_trigger = False
                st.session_state.rad_auswertung_gestartet = True

    render_sidebar_account_and_remote_support("rad")

    if st.session_state.get("rad_kf_warnung_offen", False):
        if hasattr(st, "dialog"):
            @st.dialog("⚠️ Keine Körperfettwerte eingegeben")
            def _confirm_no_bf_rad_dialog():
                st.warning("Es wurden keine Körperfettwerte (bzw. Hautfalten) eingegeben. Möchtest du die Auswertung trotzdem ohne Körperfettwerte starten?")
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    if st.button("Ja, ohne KF auswerten", type="primary", use_container_width=True, key="btn_rad_kf_dlg_yes"):
                        st.session_state.rad_kf_warnung_offen = False
                        st.session_state.rad_auswertung_gestartet = True
                        st.rerun()
                with c_d2:
                    if st.button("Abbrechen", use_container_width=True, key="btn_rad_kf_dlg_no"):
                        st.session_state.rad_kf_warnung_offen = False
                        st.rerun()

            _confirm_no_bf_rad_dialog()

        st.sidebar.warning("⚠️ Es wurden keine Körperfettwerte eingegeben. Willst du trotzdem ohne Körperfett auswerten?")
        col_kf_r1, col_kf_r2 = st.sidebar.columns(2)
        with col_kf_r1:
            if st.button("Ja, weiter", type="primary", use_container_width=True, key="btn_rad_kf_yes_sb"):
                st.session_state.rad_kf_warnung_offen = False
                st.session_state.rad_auswertung_gestartet = True
                st.rerun()
        with col_kf_r2:
            if st.button("Abbrechen", use_container_width=True, key="btn_rad_kf_no_sb"):
                st.session_state.rad_kf_warnung_offen = False
                st.rerun()
        
    # 6. RENDER MAIN PAGE LAYOUT
    if st.session_state.rad_auswertung_gestartet:
        # Fetch base variables from key-bound session state
        weight = float(st.session_state.get("weight", 72.3))
        height = int(st.session_state.get("height", 178))
        body_fat_pct = float(st.session_state.get("body_fat_pct", 0.0))
        athlete_name = st.session_state.get("athlete_name", "Michael Wagner")
        birthdate = st.session_state.get("birthdate", "06.03.1996")
        test_date = st.session_state.get("test_date", "22.05.2026")
        kategorie = st.session_state.get("kategorie", "Amateur")
        coach = st.session_state.get("coach", "Markus Hertlein")
        
        ppo_val = float(st.session_state.get("ppo_val", 436.0))
        hfmax_val = float(st.session_state.get("hfmax_val", 179.0))
        slope_val = float(st.session_state.get("slope_val", 0.11))
        intercept_val = float(st.session_state.get("intercept_val", 138.0))
        carb_intake_factor = float(st.session_state.get("carb_intake_factor", 1.0))

        # A. Check and compute VO2max
        if vo2_source == "Gemessene Spirodatei":
            abs_vo2max = float(st.session_state.vo2max_spiro)
            vo2_source_lbl = f"Spirodatei ({abs_vo2max:.1f} ml/min)"
        elif vo2_source == "Berechnet aus 30s PPO (männlich)":
            abs_vo2max = float(st.session_state.ppo_val) * 9.44 + 592.0
            vo2_source_lbl = f"30s PPO männlich ({abs_vo2max:.1f} ml/min)"
        elif vo2_source == "Berechnet aus 30s PPO (weiblich)":
            abs_vo2max = float(st.session_state.ppo_val) * 9.44 + 275.0
            vo2_source_lbl = f"30s PPO weiblich ({abs_vo2max:.1f} ml/min)"
        else:
            abs_vo2max = float(vo2_override)
            vo2_source_lbl = f"Manuell überschrieben ({abs_vo2max:.1f} ml/min)"
            
        rel_vo2max = abs_vo2max / weight if weight > 0 else 0.0
        
        # B. Load Sprint parameters based on active source
        sprint_src = st.session_state.get("sprint_source_key", "Automatisch aus Excel-Rohdaten")
        if sprint_src == "Automatisch aus Excel-Rohdaten":
            st.session_state.t_bel = float(st.session_state.get("t_bel_auto", 13.2))
            st.session_state.t_alak = float(st.session_state.get("t_alak_auto", 3.1))
        else:
            st.session_state.t_bel = float(st.session_state.get("t_bel_manual", 13.2))
            st.session_state.t_alak = float(st.session_state.get("t_alak_manual", 3.1))
        st.session_state.t_glyc = round(st.session_state.t_bel - st.session_state.t_alak, 1)
        
        t_bel = float(st.session_state.t_bel)
        t_alak = float(st.session_state.t_alak)
        t_glyc = float(st.session_state.t_glyc)
        
        # Extract Lactates from table
        lac_vals = st.session_state.df_sprint_lac["Laktat [mmol/L]"].values
        ruhe_lacs = lac_vals[0:3]
        avg_rest_lac = np.mean(ruhe_lacs)
        post_lacs = lac_vals[3:]
        max_post_lac = np.max(post_lacs)
        
        # Calculate VLamax
        vlamax = (max_post_lac - avg_rest_lac) / t_glyc if t_glyc > 0 else 0.0
        
        # C. Run metabolic grid simulation (0 to 600 W) using session state constants
        sim_results = []
        ks1 = float(st.session_state.ks1)
        ks2 = float(st.session_state.ks2)
        ks4 = float(st.session_state.ks4)
        lac_eq = float(st.session_state.lac_eq)
        lac_dist = float(st.session_state.lac_dist)
        vo2_rest = float(st.session_state.vo2_rest)
        
        for p in range(601):
            vo2ss = (p * ks4 / weight) + vo2_rest
            
            # ADP calculation (using numpy to handle invalid operations gracefully)
            if rel_vo2max - vo2ss > 0:
                adp = np.sqrt((ks1 * vo2ss) / (rel_vo2max - vo2ss))
            else:
                adp = np.nan
                
            adp_3 = adp ** 3 if not np.isnan(adp) else np.nan
            
            # VLass (mmol/L/min)
            vlass = (vlamax * 60.0) / (1.0 + ks2 / adp_3) if not np.isnan(adp_3) and adp_3 > 0 else np.nan
            
            # VLaoxmax (mmol/L/min)
            vlaoxmax = (lac_eq / lac_dist) * vo2ss
            
            # Pyruvatdefizit (VLaoxmax - VLass)
            pyruvat_def = vlaoxmax - vlass if not np.isnan(vlass) else np.nan
            
            # Fuel splitting (capped for energetics and table display to match physiology)
            kh_pct = (vlass / vlaoxmax * 100.0) if vlaoxmax > 0 and not np.isnan(vlass) else np.nan
            fat_pct = 100.0 - kh_pct if not np.isnan(kh_pct) else np.nan
            
            if not np.isnan(kh_pct):
                kh_pct_capped = max(0.0, min(100.0, kh_pct))
                fat_pct_capped = max(0.0, min(100.0, 100.0 - kh_pct_capped))
            else:
                kh_pct_capped = np.nan
                fat_pct_capped = np.nan
            
            # RQ
            rq = 0.7 + 0.3 * (kh_pct_capped / 100.0) if not np.isnan(kh_pct_capped) else np.nan
            
            # AE (Energy equivalent, kJ/L O2)
            ae = (((rq - 0.7) * 0.05094) * 100.0) + 19.6 if not np.isnan(rq) else np.nan
            
            # EE gesamt [kcal/h]
            vo2_abs_ss = vo2ss * weight
            ee_gesamt = ((vo2_abs_ss / 1000.0) * ae / 4.186) * 60.0 if not np.isnan(ae) else np.nan
            
            # EE Fett [kcal/h]
            ee_fett = ee_gesamt * (fat_pct_capped / 100.0) if not np.isnan(ee_gesamt) and not np.isnan(fat_pct_capped) else np.nan
            
            # EE KH [g/h]
            ee_kh_g = (ee_gesamt * (kh_pct_capped / 100.0)) / 4.063 if not np.isnan(ee_gesamt) and not np.isnan(kh_pct_capped) else np.nan
            
            # HR
            hr_val = p * slope_val + intercept_val
            if hr_val > hfmax_val:
                hr_val = np.nan
                
            sim_results.append({
                'power': p,
                'vo2ss': vo2ss,
                'vlass': vlass,
                'vlaoxmax': vlaoxmax,
                'pyruvat_def': pyruvat_def,
                'kh_pct': kh_pct_capped,
                'fat_pct': fat_pct_capped,
                'rq': rq,
                'ee_gesamt': ee_gesamt,
                'ee_fett': ee_fett,
                'ee_kh_g': ee_kh_g,
                'hr': hr_val
            })
            
        df_sim = pd.DataFrame(sim_results)
        
        # D. Extract Derived Metrics
        # ANS crossover: power step where vlass > vlaoxmax and the positive difference is minimized
        ans_df = df_sim[df_sim['vlass'] > df_sim['vlaoxmax']].copy()
        if not ans_df.empty:
            diff = ans_df['vlass'] - ans_df['vlaoxmax']
            min_diff_idx = diff.idxmin()
            ans_power = int(df_sim.loc[min_diff_idx, 'power'])
        else:
            ans_power = 0
            
        ans_rel = ans_power / weight
        ans_hr = ans_power * slope_val + intercept_val
        
        # Note the Excel formula inconsistency in Report_Cycling!AA16 (adds vo2_rest before weight division)
        ans_vo2_rel = ((ans_power * ks4) + vo2_rest) / weight
        ans_vo2_abs = ans_vo2_rel * weight
        ans_ee_gesamt = df_sim.loc[df_sim['power'] == ans_power, 'ee_gesamt'].values[0]
        ans_ee_kh_g = df_sim.loc[df_sim['power'] == ans_power, 'ee_kh_g'].values[0]
        
        # Fatmax: power where pyruvat_def (vlaoxmax - vlass) is maximized
        fatmax_df = df_sim.dropna(subset=['pyruvat_def']).copy()
        fatmax_idx = fatmax_df['pyruvat_def'].idxmax()
        fatmax_row = df_sim.loc[fatmax_idx]
        fatmax_power = int(fatmax_row['power'])
        fatmax_rel = fatmax_power / weight
        fatmax_ee_gesamt = fatmax_row['ee_gesamt']
        fatmax_ee_fett = fatmax_row['ee_fett']
        
        # MAP (max. Leistung)
        map_val = (abs_vo2max - weight * vo2_rest) / ks4
        
        # Match-Schwelle (KH-Match): power where ee_kh_g matches carb intake [g/h]
        if "kh_match_g_h_rad" not in st.session_state:
            st.session_state["kh_match_g_h_rad"] = 90
        _rad_vm_cur = st.session_state.get("rad_view_mode", "🛠️ Diagnostik & Feineinstellung")
        _is_rad_pres_pre = (_rad_vm_cur == "🖥️ Präsentations-Modus (Athletengespräch)")
        _rad_kh_slider_key = "kh_match_slider_rad_pres" if _is_rad_pres_pre else "kh_match_slider_rad_diag"
        carb_intake = float(st.session_state.get(_rad_kh_slider_key, st.session_state["kh_match_g_h_rad"]))
        st.session_state["kh_match_g_h_rad"] = int(round(carb_intake))
        carb_intake_factor = (carb_intake / weight) if weight > 0 else 1.0
        st.session_state.carb_intake_factor = round(carb_intake_factor, 2)

        carb_match_df = df_sim[df_sim['ee_kh_g'] <= carb_intake]
        if not carb_match_df.empty:
            carb_match_power = int(carb_match_df['power'].max())
        else:
            carb_match_power = 0
        carb_match_rel = carb_match_power / weight if weight > 0 else 0.0
        carb_match_row_df = df_sim[df_sim['power'] == carb_match_power]
        if not carb_match_row_df.empty:
            carb_match_kh_g = float(carb_match_row_df.iloc[0]['ee_kh_g']) if pd.notna(carb_match_row_df.iloc[0]['ee_kh_g']) else carb_intake
            carb_match_fat_kcal = float(carb_match_row_df.iloc[0]['ee_fett']) if pd.notna(carb_match_row_df.iloc[0]['ee_fett']) else 0.0
            carb_match_ee_ges = float(carb_match_row_df.iloc[0]['ee_gesamt']) if pd.notna(carb_match_row_df.iloc[0]['ee_gesamt']) else 0.0
        else:
            carb_match_kh_g = carb_intake
            carb_match_fat_kcal = 0.0
            carb_match_ee_ges = 0.0
        carb_match_fat_g = carb_match_fat_kcal / 9.3
        carb_match_hr = int(round(carb_match_power * slope_val + intercept_val))
            
        # Body mass indices
        bmi = weight / ((height/100.0)**2) if height > 0 else 0.0
        fett_kg = weight * (body_fat_pct / 100.0)
        fettfrei_kg = weight - fett_kg
        vo2_ffm = abs_vo2max / fettfrei_kg if fettfrei_kg > 0 else 0.0
        
        # E. Potential Recalculation (aus Präsentations-Modus / Standard-Werten für PDF)
        vo2_change = float(st.session_state.get("rad_shared_pot_vo2_pct", 5)) / 100.0
        vla_change = float(st.session_state.get("rad_shared_pot_vla_pct", -10)) / 100.0
        pot_bf_target = float(st.session_state.get("rad_shared_pot_bf_target", max(4.0, round(body_fat_pct - 1.5, 1))))
        
        pot_weight = calc_pot_weight_from_bf(weight, body_fat_pct, pot_bf_target, fm_loss_fraction=0.85)
        pot_abs_vo2max = abs_vo2max * (1.0 + vo2_change)
        pot_rel_vo2max = pot_abs_vo2max / pot_weight if pot_weight > 0 else 0.0
        pot_vlamax = max(0.15, vlamax * (1.0 + vla_change))
        
        # Save values for PDF
        st.session_state.pot_weight_val = pot_weight
        st.session_state.pot_vlamax = pot_vlamax
        st.session_state.pot_vo2max = pot_rel_vo2max
        st.session_state.pot_fat = pot_bf_target
        
        # Run potential metabolic simulation
        pot_sim_results = []
        for p in range(601):
            vo2ss = (p * ks4 / pot_weight) + vo2_rest
            if pot_rel_vo2max - vo2ss > 0:
                adp = np.sqrt((ks1 * vo2ss) / (pot_rel_vo2max - vo2ss))
            else:
                adp = np.nan
            adp_3 = adp ** 3 if not np.isnan(adp) else np.nan
            vlass = (pot_vlamax * 60.0) / (1.0 + ks2 / adp_3) if not np.isnan(adp_3) and adp_3 > 0 else np.nan
            vlaoxmax = (lac_eq / lac_dist) * vo2ss
            pyruvat_def = vlaoxmax - vlass if not np.isnan(vlass) else np.nan
            kh_pct = (vlass / vlaoxmax * 100.0) if vlaoxmax > 0 and not np.isnan(vlass) else np.nan
            fat_pct = 100.0 - kh_pct if not np.isnan(kh_pct) else np.nan
            
            if not np.isnan(kh_pct):
                kh_pct_capped = max(0.0, min(100.0, kh_pct))
                fat_pct_capped = max(0.0, min(100.0, 100.0 - kh_pct_capped))
            else:
                kh_pct_capped = np.nan
                fat_pct_capped = np.nan
                
            rq = 0.7 + 0.3 * (kh_pct_capped / 100.0) if not np.isnan(kh_pct_capped) else np.nan
            ae = (((rq - 0.7) * 0.05094) * 100.0) + 19.6 if not np.isnan(rq) else np.nan
            vo2_abs_ss = vo2ss * pot_weight
            ee_gesamt = ((vo2_abs_ss / 1000.0) * ae / 4.186) * 60.0 if not np.isnan(ae) else np.nan
            ee_fett = ee_gesamt * (fat_pct_capped / 100.0) if not np.isnan(ee_gesamt) and not np.isnan(fat_pct_capped) else np.nan
            ee_kh_g = (ee_gesamt * (kh_pct_capped / 100.0)) / 4.063 if not np.isnan(ee_gesamt) and not np.isnan(kh_pct_capped) else np.nan
            
            pot_sim_results.append({
                'power': p,
                'vo2ss': vo2ss,
                'vlass': vlass,
                'vlaoxmax': vlaoxmax,
                'pyruvat_def': pyruvat_def,
                'kh_pct': kh_pct_capped,
                'fat_pct': fat_pct_capped,
                'rq': rq,
                'ee_gesamt': ee_gesamt,
                'ee_fett': ee_fett,
                'ee_kh_g': ee_kh_g
            })
            
        df_pot_sim = pd.DataFrame(pot_sim_results)
        
        # Calculate pot_ans_power
        pot_ans_df = df_pot_sim[df_pot_sim['vlass'] > df_pot_sim['vlaoxmax']].copy()
        if not pot_ans_df.empty:
            diff = pot_ans_df['vlass'] - pot_ans_df['vlaoxmax']
            min_diff_idx = diff.idxmin()
            pot_ans_power = int(df_pot_sim.loc[min_diff_idx, 'power'])
        else:
            pot_ans_power = 0
            
        pot_ans_rel = pot_ans_power / pot_weight if pot_weight > 0 else 0.0
        
        # Calculate pot_fatmax_power
        pot_fatmax_df = df_pot_sim.dropna(subset=['pyruvat_def']).copy()
        pot_fatmax_idx = pot_fatmax_df['pyruvat_def'].idxmax()
        pot_fatmax_row = df_pot_sim.loc[pot_fatmax_idx]
        pot_fatmax_power = int(pot_fatmax_row['power'])
        pot_fatmax_rel = pot_fatmax_power / pot_weight if pot_weight > 0 else 0.0
        
        # pot_match_power (bei eingestellter KH-Zufuhr in g/h)
        pot_carb_intake = carb_intake
        pot_carb_match_df = df_pot_sim[df_pot_sim['ee_kh_g'] <= pot_carb_intake]
        if not pot_carb_match_df.empty:
            pot_carb_match_power = int(pot_carb_match_df['power'].max())
        else:
            pot_carb_match_power = 0
        pot_match_rel = pot_carb_match_power / pot_weight if pot_weight > 0 else 0.0
        
        # pot_pmax_rel
        pot_sprint_powers = st.session_state.get("sprint_powers", [])
        if not pot_sprint_powers or len(pot_sprint_powers) < 2:
            pot_sprint_powers = [max(0.0, 1200.0 * (1.0 - ((t - 2.0)**2) / 64.0)) for t in [i * 0.1 for i in range(120)]]
        pot_max_sprint_power = np.max(pot_sprint_powers)
        pot_pmax_rel = pot_max_sprint_power / pot_weight if pot_weight > 0 else 0.0
        
        st.session_state.pot_ans_abs = float(pot_ans_power)
        st.session_state.pot_ans_rel = float(pot_ans_rel)
        st.session_state.pot_fatmax = float(pot_fatmax_rel)
        st.session_state.pot_match = float(pot_match_rel)
        st.session_state.pot_pmax = float(pot_pmax_rel)
        
        # Ansichts-Umschalter: Ansicht 2 (Diagnostik) <-> Ansicht 3 (Präsentations-Modus)
        if "rad_view_mode" not in st.session_state:
            st.session_state.rad_view_mode = "🛠️ Diagnostik & Feineinstellung"

        col_vm_r1, col_vm_r2 = st.columns([2.2, 1.0])
        with col_vm_r1:
            _rad_ansicht_lbl = "Diagnostik-Ansicht" if st.session_state.rad_view_mode == "🛠️ Diagnostik & Feineinstellung" else "Präsentationsansicht"
            st.markdown(
                f"<div style='font-size:12.5px; color:#595a59; font-weight:500; padding-top:8px;'>Aktive Ansicht: <span style='color:#1a1a1a; font-weight:700;'>{_rad_ansicht_lbl}</span></div>",
                unsafe_allow_html=True,
            )
        with col_vm_r2:
            if st.session_state.rad_view_mode == "🛠️ Diagnostik & Feineinstellung":
                if st.button("🖥️ Zur Präsentationsansicht wechseln", type="primary", use_container_width=True, key="btn_to_pres_rad_top"):
                    st.session_state.rad_view_mode = "🖥️ Präsentations-Modus (Athletengespräch)"
                    st.rerun()
            else:
                if st.button("🛠️ Zurück zur Diagnostik-Ansicht", use_container_width=True, key="btn_to_diag_rad_top"):
                    st.session_state.rad_view_mode = "🛠️ Diagnostik & Feineinstellung"
                    st.rerun()

        def _solve_ans_rad(vo2_abs_in, vla_in, w_in):
            if w_in <= 0:
                return float(ans_power), float(ans_rel)
            rel_vo2_in = vo2_abs_in / w_in
            best_p = float(ans_power)
            min_d = 1e9
            for p_cand in range(50, 601):
                vo2ss_c = (p_cand * ks4 / w_in) + vo2_rest
                if rel_vo2_in - vo2ss_c <= 0:
                    break
                adp_c = np.sqrt((ks1 * vo2ss_c) / (rel_vo2_in - vo2ss_c))
                vlass_c = (vla_in * 60.0) / (1.0 + ks2 / (adp_c ** 3))
                vlaox_c = (lac_eq / lac_dist) * vo2ss_c
                if vlass_c >= vlaox_c:
                    best_p = float(p_cand)
                    break
                d_c = abs(vlaox_c - vlass_c)
                if d_c < min_d:
                    min_d = d_c
                    best_p = float(p_cand)
            return best_p, best_p / w_in

        is_rad_pres = (st.session_state.rad_view_mode == "🖥️ Präsentations-Modus (Athletengespräch)")

        # Render Main Panels (einheitlich für Präsentations-Modus & Diagnostik & Feineinstellung analog zum Laufen)
        ans_vo2_pct_top = (ans_vo2_abs / abs_vo2max * 100.0) if abs_vo2max > 0 else 0.0
        ans_vo2_pct_top_str = f"{ans_vo2_pct_top:.1f} %".replace(".", ",") if ans_vo2_pct_top > 0 else "-"
        fatmax_hr_val = int(round(fatmax_row['hr'])) if 'hr' in fatmax_row and not np.isnan(fatmax_row['hr']) else int(round(fatmax_power * slope_val + intercept_val))
        if is_rad_pres:
            st.markdown(f"""
                <div style="font-size: 16px; line-height: 1.5; margin-bottom: 16px; padding: 12px 16px; border-left: 5px solid #2cb7b9; background-color: #f8fafa; border-radius: 4px;">
                    <b>Athlet:</b> {athlete_name} &nbsp;|&nbsp; <b>Test-Datum:</b> {test_date}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
                <div style="font-size: 16px; line-height: 1.5; margin-bottom: 16px; padding: 12px 16px; border-left: 5px solid #2cb7b9; background-color: #f8fafa; border-radius: 4px;">
                    <b>Athlet:</b> {athlete_name} &nbsp;|&nbsp; <b>Sportart:</b> Radsport &nbsp;|&nbsp; <b>Test-Datum:</b> {test_date} &nbsp;|&nbsp; <b>Kategorie:</b> {kategorie} &nbsp;|&nbsp; <b>Coach:</b> {coach}
                </div>
                """, unsafe_allow_html=True)

        render_hycys_top_basisdaten(
            weight=float(weight),
            body_fat_pct=float(body_fat_pct),
            rel_vo2max=float(rel_vo2max),
            abs_vo2max=float(abs_vo2max),
            vlamax_str=f"{vlamax:.2f}".replace(".", ",") + " mmol/l/s",
            ans_label="ANS / LTP2 (Schwelle)",
            ans_main_str=f"{int(ans_power)} W",
            ans_sub_str=f"Relativ: {ans_rel:.2f}".replace(".", ",") + " W/kg",
            ans_hr_str=f"{int(round(ans_hr))} bpm",
            vo2_pct_str=ans_vo2_pct_top_str,
            ltp1_label="LTP1 / Fatmax",
            ltp1_main_str=f"{int(fatmax_power)} W",
            ltp1_sub_str=f"Relativ: {fatmax_rel:.2f}".replace(".", ",") + " W/kg",
            ltp1_hr_str=f"{fatmax_hr_val} bpm",
            is_presentation=is_rad_pres,
            bmi_val=float(bmi),
            height_cm=float(height),
        )

        # Sprint-Verifizierung direkt unter den Basisdaten & Kennzahlen (in der Diagnostik-Ansicht)
        if not is_rad_pres:
            st.markdown("---")
            st.subheader("Sprint-Verifizierung & VLamax-Berechnung")
            col_sprint_1, col_sprint_2 = st.columns([2, 1])

            with col_sprint_1:
                if "sprint_powers" in st.session_state and len(st.session_state.sprint_powers) > 0:
                    fig_sprint = go.Figure()
                    _sp_times_hover = [f"{float(t):.2f}".replace(".", ",") for t in st.session_state.sprint_times]
                    fig_sprint.add_trace(go.Scatter(
                        x=st.session_state.sprint_times,
                        y=st.session_state.sprint_powers,
                        mode='lines',
                        line=dict(color='#00a1e0', width=3),
                        name='Leistung',
                        customdata=_sp_times_hover,
                        hovertemplate="<b>%{customdata} Sekunden und %{y:.0f} W</b><extra></extra>",
                    ))
                    sprint_pmax = max(st.session_state.sprint_powers)
                    sprint_palak = 0.965 * sprint_pmax
                    fig_sprint.add_hline(
                        y=sprint_palak,
                        line_dash="dot",
                        line_color="#cdb663",
                        annotation_text=f"96.5% Pmax ({sprint_palak:.0f} W)",
                        annotation_position="bottom left"
                    )
                    fig_sprint.add_vrect(
                        x0=0.0, x1=t_alak,
                        fillcolor="rgba(0, 161, 224, 0.1)",
                        layer="below", line_width=0,
                        annotation_text="Alaktazid (t_alak)",
                        annotation_position="top left"
                    )
                    fig_sprint.add_vrect(
                        x0=t_alak, x1=t_bel,
                        fillcolor="rgba(205, 182, 99, 0.1)",
                        layer="below", line_width=0,
                        annotation_text="Glykolytisch (t_glyc)",
                        annotation_position="top right"
                    )
                    _sp_t_max = float(max(st.session_state.sprint_times))
                    if _sp_t_max > t_bel + 0.05:
                        fig_sprint.add_vrect(
                            x0=t_bel, x1=_sp_t_max,
                            fillcolor="rgba(150, 150, 150, 0.18)",
                            layer="below", line_width=0,
                            annotation_text="nicht gewertet (Leistungsplateau)",
                            annotation_position="bottom right",
                            annotation_font=dict(size=10, color="#777"),
                        )
                    fig_sprint.add_vline(
                        x=t_bel,
                        line_dash="dash",
                        line_color="#d62728",
                        line_width=2,
                        annotation_text=f"Ende Sprint t_bel = {t_bel:.1f} s",
                        annotation_position="top left",
                        annotation_font=dict(size=11, color="#d62728"),
                    )
                    fig_sprint.update_layout(
                        xaxis_title="Zeit [Sekunden]",
                        yaxis_title="Leistung [Watt]",
                        height=460,
                        margin=dict(l=45, r=35, t=20, b=40),
                        plot_bgcolor='#f9f9f9',
                        legend=dict(x=0.02, y=0.98, bgcolor="rgba(255,255,255,0.7)")
                    )
                    st.plotly_chart(fig_sprint, use_container_width=True)
                else:
                    st.info("ℹ️ Keine Sprint-Rohdaten hochgeladen. Es wird die standardmäßig konfigurierte Belastungsdauer verwendet.")

            with col_sprint_2:
                st.markdown(f"""
                <div style="border: 1px solid #e6e6e6; border-radius: 8px; padding: 15px; background-color: #f9f9f9; color: #222;">
                    <h4 style="margin-top:0; color: #111;">Dauer der Phasen</h4>
                    <table style="width:100%; border-collapse: collapse;">
                        <tr style="border-bottom: 1px solid #e6e6e6;">
                            <td style="padding: 6px 0;"><b>Gesamtdauer (t_bel):</b></td>
                            <td style="text-align:right;">{t_bel:.1f} s</td>
                        </tr>
                        {(f'<tr style="border-bottom: 1px solid #e6e6e6; color:#777; font-size:0.85em;"><td style="padding: 4px 0;" colspan="2">Leistungsplateau am Ende (&gt; {SPRINT_PLATEAU_MIN_POINTS} gleiche Werte) ab {t_bel:.1f} s nicht gewertet (Aufzeichnung bis {float(max(st.session_state.sprint_times)):.1f} s)</td></tr>') if (st.session_state.get("sprint_times") and float(max(st.session_state.sprint_times)) > t_bel + 0.05) else ''}
                        <tr style="border-bottom: 1px solid #e6e6e6;">
                            <td style="padding: 6px 0;"><b>Alaktazide Phase (t_alak):</b></td>
                            <td style="text-align:right;">{t_alak:.1f} s</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e6e6e6; color: #cdb663;">
                            <td style="padding: 6px 0;"><b>Glykolytische Phase (t_glyc):</b></td>
                            <td style="text-align:right;"><b>{t_glyc:.1f} s</b></td>
                        </tr>
                    </table>
                    <h4 style="margin-top: 15px; color: #111;">Laktat-Differenz</h4>
                    <table style="width:100%; border-collapse: collapse;">
                        <tr style="border-bottom: 1px solid #e6e6e6;">
                            <td style="padding: 6px 0;">Ruhelaktat (Mittelwert):</td>
                            <td style="text-align:right;">{avg_rest_lac:.2f} mmol/L</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e6e6e6;">
                            <td style="padding: 6px 0;">Peak-Laktat (post-sprint):</td>
                            <td style="text-align:right;">{max_post_lac:.2f} mmol/L</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e6e6e6; color: #00a1e0;">
                            <td style="padding: 6px 0;"><b>Laktat-Anstieg (&Delta;La):</b></td>
                            <td style="text-align:right;"><b>{(max_post_lac - avg_rest_lac):.2f} mmol/L</b></td>
                        </tr>
                    </table>
                    <h4 style="margin-top:15px; margin-bottom:5px; color: #111;">VLamax Berechnung</h4>
                    <div style="font-family: monospace; font-size:13px; background-color:#ffffff; padding:10px; border-radius:4px; border: 1px solid #e6e6e6; text-align:center; color: #111;">
                        VLamax = &Delta;La / t_glyc <br>
                        VLamax = {max_post_lac - avg_rest_lac:.2f} / {t_glyc:.1f} <br>
                        <b>VLamax = {vlamax:.3f} mmol/l/s</b>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # Im Präsentations-Modus steht das interaktive Coaching-Potenzial direkt unter den Kennzahlen im Mittelpunkt
        if is_rad_pres:
            render_ans_coaching_diagrams(
                mode="cycling",
                weight_ist=float(weight),
                bf_ist=float(body_fat_pct),
                abs_vo2max_ist=float(abs_vo2max),
                rel_vo2max_ist=float(rel_vo2max),
                vlamax_ist=float(vlamax),
                ans_abs_ist=float(ans_power),
                ans_rel_ist=float(ans_rel),
                solve_ans_fn=_solve_ans_rad,
                default_vo2_pct=float(vo2_change * 100.0) if vo2_change != 0 else 5.0,
                default_vla_pct=float(vla_change * 100.0) if vla_change != 0 else -10.0,
                default_bf_delta=-1.5,
                key_prefix="rad",
                is_presentation=True,
            )

        # Fett & Kohlenhydrat Stoffwechsel (links, 2/3 Größe) + KH-Match (rechts daneben) – einheitlich wie beim Laufen
        st.markdown("---")
        max_x_plot = int(min(600, max(300, np.ceil((ans_power + 50) / 10.0) * 10.0)))
        df_plot_rad = df_sim[df_sim['power'] <= max_x_plot].dropna(subset=['ee_gesamt']).copy()
        p_vals_r = df_plot_rad['power'].values if not df_plot_rad.empty else np.arange(0, max_x_plot + 1, dtype=float)
        galibier_rad = compute_galibier_v2_cycling(
            p_vals=p_vals_r,
            weight=float(weight),
            rel_vo2max=float(rel_vo2max),
            vlamax=float(vlamax),
            carb_intake_g_h=float(carb_intake),
            km_val=0.21,
            km_low=0.15,
            km_high=0.28,
            eta=0.239,
            vo2_rest=3.5,
        )

        # Gemessene Spiro-Werte für Radsport aufbereiten (falls Spirodatei hochgeladen)
        rad_spiro_p_meas = []
        rad_spiro_fett_meas = []
        rad_spiro_kh_meas = []
        _df_sp_rad = st.session_state.get("rad_spiro_df", pd.DataFrame())
        if (_df_sp_rad is None or _df_sp_rad.empty or "FETT_kcal_h" not in _df_sp_rad.columns) and uploaded_spiro is not None:
            try:
                uploaded_spiro.seek(0)
                _df_sp_ex = pd.read_csv(uploaded_spiro, header=None, low_memory=False) if str(uploaded_spiro.name).lower().endswith(".csv") else pd.read_excel(uploaded_spiro, header=None)
                uploaded_spiro.seek(0)
                _df_sp_rad = pd.DataFrame({
                    "Time": _df_sp_ex.iloc[3:, 9].apply(parse_time_to_seconds),
                    "VO2": pd.to_numeric(_df_sp_ex.iloc[3:, 13], errors="coerce"),
                    "VCO2": pd.to_numeric(_df_sp_ex.iloc[3:, 14], errors="coerce") if _df_sp_ex.shape[1] > 14 else np.nan,
                    "FETT_kcal_h": (pd.to_numeric(_df_sp_ex.iloc[3:, 41], errors="coerce") / 24.0) if _df_sp_ex.shape[1] > 41 else np.nan,
                    "KH_g_h": ((pd.to_numeric(_df_sp_ex.iloc[3:, 42], errors="coerce") / 24.0) / 4.063) if _df_sp_ex.shape[1] > 42 else np.nan,
                }).dropna(subset=["Time", "VO2"])
                st.session_state.rad_spiro_df = _df_sp_rad
            except Exception:
                _df_sp_rad = pd.DataFrame()
        if _df_sp_rad is not None and not _df_sp_rad.empty and "VO2" in _df_sp_rad.columns:
            try:
                # Nur den ansteigenden Teil bis zum VO2-Peak verwenden
                _peak_pos = int(_df_sp_rad["VO2"].values.argmax())
                _df_sp_up = _df_sp_rad.iloc[: max(3, _peak_pos + 1)].copy()
                _df_sp_up["p_est"] = np.maximum(0.0, (_df_sp_up["VO2"] - float(weight) * float(vo2_rest)) / float(ks4))
                if "FETT_kcal_h" not in _df_sp_up.columns or _df_sp_up["FETT_kcal_h"].isna().all():
                    _rq_r = np.clip(_df_sp_up["VCO2"] / _df_sp_up["VO2"], 0.70, 1.00)
                    _fp_r = (1.0 - _rq_r) / 0.30
                    _ae_r = (_rq_r - 0.70) * 5.094 + 19.6
                    _ee_r = (_df_sp_up["VO2"] / 1000.0 * _ae_r * 60.0) / 4.186
                    _df_sp_up["FETT_kcal_h"] = _ee_r * _fp_r
                    _df_sp_up["KH_g_h"] = (_ee_r * (1.0 - _fp_r)) / 4.063
                _df_sp_up = _df_sp_up[(_df_sp_up["p_est"] >= 60.0) & (_df_sp_up["p_est"] <= float(max_x_plot))].dropna(subset=["FETT_kcal_h", "KH_g_h"])
                if len(_df_sp_up) >= 3:
                    _df_sp_up["p_bin"] = (np.round(_df_sp_up["p_est"] / 25.0) * 25.0).astype(int)
                    _grp_sp = _df_sp_up.groupby("p_bin", as_index=False)[["FETT_kcal_h", "KH_g_h"]].mean().sort_values("p_bin")
                    if len(_grp_sp) >= 2:
                        rad_spiro_p_meas = list(_grp_sp["p_bin"].astype(float).values)
                        rad_spiro_fett_meas = list(_grp_sp["FETT_kcal_h"].astype(float).values)
                        rad_spiro_kh_meas = list(_grp_sp["KH_g_h"].astype(float).values)
            except Exception:
                pass

        col_rad_chart_u, col_rad_khm_u = st.columns([1.95, 1.05])
        with col_rad_chart_u:
            st.markdown("<h3 style='text-align: center; margin-bottom: 6px;'>Fett & Kohlenhydrat Stoffwechsel</h3>", unsafe_allow_html=True)
            _c_v2_rad, _c_km_rad, _c_sp_rad = st.columns([1.35, 1.05, 1.25])
            with _c_v2_rad:
                show_v2_galibier_rad = st.checkbox(
                    "Version 2 - Galibier (AI Diagnostics) einblenden",
                    value=False,
                    key="show_v2_galibier_rad",
                )
            with _c_km_rad:
                show_v2_km_band_rad = st.checkbox(
                    "Km-Range (0,15–0,28) zeigen",
                    value=False,
                    disabled=not show_v2_galibier_rad,
                    key="show_v2_km_band_rad",
                )
            with _c_sp_rad:
                show_spiro_met_rad = False
                if len(rad_spiro_p_meas) >= 2:
                    show_spiro_met_rad = st.checkbox(
                        "Version 3 - Gemessen Spiro einblenden",
                        value=False,
                        key="show_spiro_met_rad",
                    )

        with col_rad_khm_u:
            st.markdown("<h3 style='text-align: center; margin-bottom: 6px;'>KH-Match</h3>", unsafe_allow_html=True)
            kh_sel_rad_u = st.slider(
                "KH-Zufuhr (Input) [g KH/h]",
                min_value=10,
                max_value=160,
                value=int(round(carb_intake)),
                step=1,
                format="%d g/h",
                key=_rad_kh_slider_key,
            )
            st.session_state["kh_match_g_h_rad"] = int(kh_sel_rad_u)
            st.caption(f"Entspricht **{(kh_sel_rad_u / weight):.2f} g/kg KG/h**".replace(".", ","))
            _khm_rel_str = f"{carb_match_rel:.2f}".replace(".", ",")
            _khm_fat_str = f"{carb_match_fat_g:.1f}".replace(".", ",")
            st.markdown(
                f"""
                <div style="border:1px solid #e0e8e8; border-left:5px solid #cdb663; border-radius:6px; padding:14px 16px; background-color:#ffffff; margin-top:6px;">
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase; margin-bottom:10px;">
                        ⚡ KH-Match bei {int(round(carb_intake))} g KH/h Zufuhr (Version 1 - HYCYS Berechnung (Mader))
                    </div>
                    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; text-align:center;">
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">KH-Match Leistung</div>
                            <div style="font-size:20px; font-weight:800; color:#1a1a1a;">{carb_match_power} W</div>
                            <div style="font-size:11px; color:#595a59;">({_khm_rel_str} W/kg)</div>
                        </div>
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">HF @ KH-Match</div>
                            <div style="font-size:20px; font-weight:800; color:#d9534f;">{carb_match_hr} bpm</div>
                        </div>
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">KH-Verbrauch</div>
                            <div style="font-size:20px; font-weight:800; color:#cdb663;">{int(round(carb_match_kh_g))} g/h</div>
                            <div style="font-size:11px; color:#595a59;">= Input ({int(round(carb_intake))} g/h)</div>
                        </div>
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">Fett & Gesamtenergie</div>
                            <div style="font-size:17px; font-weight:800; color:#2cb7b9;">{_khm_fat_str} g/h Fett</div>
                            <div style="font-size:11px; color:#595a59;">({int(round(carb_match_ee_ges))} kcal/h gesamt)</div>
                        </div>
                    </div>
                    <div style="margin-top:10px; font-size:13px; color:#1a1a1a; line-height:1.45;">
                        Bei <b>{int(round(carb_intake))} g KH/h</b> kannst du <b>{carb_match_power} W ({_khm_rel_str} W/kg)</b> fahren – dein Kohlenhydratverbrauch ist genauso groß wie dein Input.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if show_v2_galibier_rad:
                _v2_khm_p = int(galibier_rad["carb_match_power"])
                _v2_khm_rel = f"{(_v2_khm_p / weight if weight > 0 else 0.0):.2f}".replace(".", ",")
                _v2_khm_hr = int(round(_v2_khm_p * slope_val + intercept_val))
                _v2_fm_p = int(galibier_rad["fatmax_power"])
                _v2_mlss_p = int(galibier_rad["mlss_power"])
                _v2_fat_str = f"{galibier_rad['carb_match_fat_g']:.1f}".replace(".", ",")
                st.markdown(
                    f"""
                    <div style="border:1px solid #e0e8e8; border-left:5px solid #1f7a8c; border-radius:6px; padding:10px 14px; background-color:#f8fbfb; margin-top:8px; font-size:12.5px; color:#1a1a1a;">
                        <div style="font-weight:700; text-transform:uppercase; margin-bottom:6px; color:#1f7a8c;">
                            Version 2 - Galibier (AI Diagnostics)
                        </div>
                        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:6px; text-align:center;">
                            <div style="background:#ffffff; padding:6px; border-radius:4px; border:1px solid #e6ecec;">
                                <div style="font-size:10.5px; color:#595a59;">KH-Match (V2)</div>
                                <div style="font-size:15px; font-weight:800;">{_v2_khm_p} W ({_v2_khm_rel} W/kg)</div>
                                <div style="font-size:10.5px; color:#d9534f; font-weight:700;">HF: {_v2_khm_hr} bpm | {_v2_fat_str} g/h Fett</div>
                            </div>
                            <div style="background:#ffffff; padding:6px; border-radius:4px; border:1px solid #e6ecec;">
                                <div style="font-size:10.5px; color:#595a59;">Fatmax (V2 vs. V1)</div>
                                <div style="font-size:15px; font-weight:800; color:#1f7a8c;">{_v2_fm_p} W</div>
                                <div style="font-size:10.5px; color:#595a59;">V1: {fatmax_power} W</div>
                            </div>
                            <div style="background:#ffffff; padding:6px; border-radius:4px; border:1px solid #e6ecec;">
                                <div style="font-size:10.5px; color:#595a59;">MLSS / ANS (V2 vs. V1)</div>
                                <div style="font-size:15px; font-weight:800;">{_v2_mlss_p} W</div>
                                <div style="font-size:10.5px; color:#595a59;">V1: {ans_power} W</div>
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        with col_rad_chart_u:
            if not df_plot_rad.empty:
                ee_tot_r = df_plot_rad['ee_gesamt'].values
                ee_fat_r = df_plot_rad['ee_fett'].values
                fat_g_r = ee_fat_r / 9.3
                kh_g_r = df_plot_rad['ee_kh_g'].values
                kh_kcal_r = kh_g_r * 4.063
                custom_rad_hover = np.column_stack([
                    [f"{int(p)}" for p in p_vals_r],
                    [f"{p / weight:.2f}".replace(".", ",") for p in p_vals_r],
                    [f"{int(round(v))}" for v in ee_tot_r],
                    [f"{v:.1f}".replace(".", ",") for v in fat_g_r],
                    [f"{int(round(v))}" for v in ee_fat_r],
                    [f"{v:.1f}".replace(".", ",") for v in kh_g_r],
                    [f"{int(round(v))}" for v in kh_kcal_r],
                ])
                fig_met_u = go.Figure()
                fig_met_u.add_trace(go.Scatter(
                    x=p_vals_r, y=ee_tot_r,
                    mode='lines',
                    line=dict(color='#595a59', width=3),
                    name='Gesamtenergieverbrauch [kcal/h]',
                    customdata=custom_rad_hover,
                    hovertemplate=(
                        "<b>Leistung: %{customdata[0]} W (%{customdata[1]} W/kg)</b><br>"
                        "Gesamtenergie: <b>%{customdata[2]} kcal/h</b><br>"
                        "Fett (Version 1 - HYCYS Berechnung (Mader)): <b>%{customdata[3]} g/h Fett</b> (%{customdata[4]} kcal/h)<br>"
                        "KH (Version 1 - HYCYS Berechnung (Mader)): <b>%{customdata[5]} g/h KH</b> (%{customdata[6]} kcal/h)"
                        "<extra></extra>"
                    ),
                    yaxis='y'
                ))
                fig_met_u.add_trace(go.Scatter(
                    x=p_vals_r, y=ee_fat_r,
                    mode='lines',
                    line=dict(color='#2cb7b9', width=3),
                    name='Fett: Version 1 - HYCYS Berechnung (Mader) [kcal/h]',
                    hoverinfo='skip',
                    yaxis='y'
                ))
                fig_met_u.add_trace(go.Scatter(
                    x=p_vals_r, y=kh_g_r,
                    mode='lines',
                    line=dict(color='#cdb663', width=3),
                    name='KH: Version 1 - HYCYS Berechnung (Mader) [g/h]',
                    hoverinfo='skip',
                    yaxis='y2'
                ))

                # Version 2 - Galibier (AI Diagnostics) Kurven + optionale Km-Range (0.15 - 0.28)
                if show_v2_galibier_rad:
                    if show_v2_km_band_rad:
                        fig_met_u.add_trace(go.Scatter(
                            x=np.concatenate([p_vals_r, p_vals_r[::-1]]),
                            y=np.concatenate([galibier_rad["ee_fat_high"], galibier_rad["ee_fat_low"][::-1]]),
                            fill="toself",
                            fillcolor="rgba(31, 122, 140, 0.13)",
                            line=dict(color="rgba(255,255,255,0)"),
                            hoverinfo="skip",
                            name="Fett Spannweite: Version 2 - Galibier (AI Diagnostics)",
                            yaxis="y",
                        ))
                        fig_met_u.add_trace(go.Scatter(
                            x=np.concatenate([p_vals_r, p_vals_r[::-1]]),
                            y=np.concatenate([galibier_rad["kh_g_high"], galibier_rad["kh_g_low"][::-1]]),
                            fill="toself",
                            fillcolor="rgba(184, 134, 11, 0.13)",
                            line=dict(color="rgba(255,255,255,0)"),
                            hoverinfo="skip",
                            name="KH Spannweite: Version 2 - Galibier (AI Diagnostics)",
                            yaxis="y2",
                        ))

                    _v2_rad_hover = np.column_stack([
                        [f"{v:.1f}".replace(".", ",") for v in galibier_rad["fat_g"]],
                        [f"{int(round(v))}" for v in galibier_rad["ee_fat"]],
                        [f"{v:.0f}" for v in galibier_rad["fat_pct"]],
                        [f"{v:.1f}".replace(".", ",") for v in galibier_rad["kh_g"]],
                        [f"{int(round(v))}" for v in galibier_rad["kh_kcal"]],
                        [f"{v:.0f}" for v in galibier_rad["kh_pct"]],
                    ])
                    fig_met_u.add_trace(go.Scatter(
                        x=p_vals_r,
                        y=galibier_rad["ee_fat"],
                        mode="lines",
                        line=dict(color="#0d5c75", width=3, dash="dot"),
                        name="Fett: Version 2 - Galibier (AI Diagnostics) [kcal/h]",
                        customdata=_v2_rad_hover,
                        hovertemplate="Fett (Version 2 - Galibier (AI Diagnostics)): <b>%{customdata[0]} g/h Fett</b> (%{customdata[1]} kcal/h · %{customdata[2]} %)<extra></extra>",
                        yaxis="y",
                    ))
                    fig_met_u.add_trace(go.Scatter(
                        x=p_vals_r,
                        y=galibier_rad["kh_g"],
                        mode="lines",
                        line=dict(color="#b8860b", width=3, dash="dot"),
                        name="KH: Version 2 - Galibier (AI Diagnostics) [g/h]",
                        customdata=_v2_rad_hover,
                        hovertemplate="KH (Version 2 - Galibier (AI Diagnostics)): <b>%{customdata[3]} g/h KH</b> (%{customdata[4]} kcal/h · %{customdata[5]} %)<extra></extra>",
                        yaxis="y2",
                    ))

                # Version 3 - Gemessen Spiro Kurven für Radsport
                if show_spiro_met_rad and len(rad_spiro_p_meas) >= 2:
                    from scipy.interpolate import PchipInterpolator
                    _rp_arr = np.array(rad_spiro_p_meas, dtype=float)
                    _rf_arr = np.array(rad_spiro_fett_meas, dtype=float)
                    _rk_arr = np.array(rad_spiro_kh_meas, dtype=float)
                    _rp_u, _ru_idx = np.unique(_rp_arr, return_index=True)
                    _rf_u = _rf_arr[_ru_idx]
                    _rk_u = _rk_arr[_ru_idx]
                    if len(_rp_u) >= 2:
                        _rp_mask = (p_vals_r >= float(_rp_u[0])) & (p_vals_r <= float(_rp_u[-1]))
                        _rp_smooth = p_vals_r[_rp_mask] if np.sum(_rp_mask) >= 5 else np.linspace(float(_rp_u[0]), float(_rp_u[-1]), 100)
                        _rf_smooth = np.clip(PchipInterpolator(_rp_u, _rf_u)(_rp_smooth), 0.0, None)
                        _rk_smooth = np.clip(PchipInterpolator(_rp_u, _rk_u)(_rp_smooth), 0.0, None)
                        _rsp_hover = np.column_stack([
                            [f"{val / 9.3:.1f}".replace(".", ",") for val in _rf_smooth],
                            [f"{int(round(val))}" for val in _rf_smooth],
                            [f"{val:.1f}".replace(".", ",") for val in _rk_smooth],
                            [f"{int(round(val * 4.063))}" for val in _rk_smooth],
                        ])
                        fig_met_u.add_trace(go.Scatter(
                            x=_rp_smooth, y=_rf_smooth,
                            mode="lines",
                            line=dict(color="#11696b", width=3, dash="dash"),
                            name="Fett: Version 3 - Gemessen Spiro [kcal/h]",
                            customdata=_rsp_hover,
                            hovertemplate="Fett (Version 3 - Gemessen Spiro): <b>%{customdata[0]} g/h Fett</b> (%{customdata[1]} kcal/h)<extra></extra>",
                            yaxis="y",
                        ))
                        fig_met_u.add_trace(go.Scatter(
                            x=_rp_u, y=_rf_u,
                            mode="markers",
                            marker=dict(symbol="circle", size=7.5, color="#11696b", line=dict(width=1.5, color="#ffffff")),
                            showlegend=False,
                            hoverinfo="skip",
                            yaxis="y",
                        ))
                        fig_met_u.add_trace(go.Scatter(
                            x=_rp_smooth, y=_rk_smooth,
                            mode="lines",
                            line=dict(color="#8c7322", width=3, dash="dash"),
                            name="KH: Version 3 - Gemessen Spiro [g/h]",
                            customdata=_rsp_hover,
                            hovertemplate="KH (Version 3 - Gemessen Spiro): <b>%{customdata[2]} g/h KH</b> (%{customdata[3]} kcal/h)<extra></extra>",
                            yaxis="y2",
                        ))
                        fig_met_u.add_trace(go.Scatter(
                            x=_rp_u, y=_rk_u,
                            mode="markers",
                            marker=dict(symbol="circle", size=7.5, color="#8c7322", line=dict(width=1.5, color="#ffffff")),
                            showlegend=False,
                            hoverinfo="skip",
                            yaxis="y2",
                        ))
                if 0 < fatmax_power <= int(p_vals_r[-1]):
                    _fm_half_w_r = max(20.0, float(p_vals_r[-1]) * 0.055)
                    _fm_x0_r = max(0.0, float(fatmax_power) - _fm_half_w_r)
                    _fm_x1_r = min(float(p_vals_r[-1]), float(fatmax_power) + _fm_half_w_r)
                    fig_met_u.add_vrect(
                        x0=_fm_x0_r,
                        x1=_fm_x1_r,
                        fillcolor="rgba(44, 183, 185, 0.20)",
                        line_width=0,
                        layer="below",
                    )
                    fig_met_u.add_annotation(
                        x=0.5 * (_fm_x0_r + _fm_x1_r),
                        y=0.92,
                        xref="x",
                        yref="paper",
                        text="<b>Fatmax</b>",
                        showarrow=False,
                        xanchor="center",
                        yanchor="top",
                        font=dict(size=14, color="#1d8f91", family="Arial"),
                    )
                if 0 < carb_match_power <= int(p_vals_r[-1]):
                    fig_met_u.add_vline(
                        x=carb_match_power,
                        line_dash="dot",
                        line_color="#cdb663",
                        line_width=2.5,
                    )
                    fig_met_u.add_annotation(
                        x=carb_match_power,
                        y=0.10,
                        xref="x",
                        yref="paper",
                        text="<b>KH-Match</b>",
                        showarrow=False,
                        xanchor="left",
                        yanchor="bottom",
                        xshift=6,
                        font=dict(size=14, color="#9e8836", family="Arial"),
                    )
                if 0 < ans_power <= int(p_vals_r[-1]):
                    fig_met_u.add_vline(
                        x=ans_power,
                        line_dash="dash",
                        line_color="#595a59",
                    )
                    fig_met_u.add_annotation(
                        x=ans_power,
                        y=0.92,
                        xref="x",
                        yref="paper",
                        text="<b>ANS</b>",
                        showarrow=False,
                        xanchor="left",
                        yanchor="top",
                        xshift=6,
                        font=dict(size=13, color="#595a59", family="Arial"),
                    )
                fig_met_u.update_layout(
                    xaxis=dict(title="Leistung [Watt]", range=[0, int(p_vals_r[-1])], gridcolor="#e6e6e6"),
                    yaxis=dict(
                        title=dict(text="Energie [kcal/h]", font=dict(color="#595a59")),
                        tickfont=dict(color="#595a59"),
                        tickformat=".0f",
                        range=[0, float(np.max(ee_tot_r)) * 1.12],
                        gridcolor="#e6e6e6",
                    ),
                    yaxis2=dict(
                        title=dict(text="Kohlenhydratverbrauch [g/h]", font=dict(color="#cdb663")),
                        tickfont=dict(color="#cdb663"),
                        tickformat=".0f",
                        range=[0, float(np.max(kh_g_r)) * 1.12],
                        overlaying="y",
                        side="right",
                        showgrid=False,
                    ),
                    hovermode="x",
                    height=530,
                    legend=dict(orientation="h", x=0.5, xanchor="center", y=1.08, bgcolor="rgba(255,255,255,0.85)"),
                    margin=dict(l=55, r=55, t=55, b=50),
                    plot_bgcolor='#f9f9f9'
                )
                st.plotly_chart(fig_met_u, use_container_width=True)

        # TRAININGSBEREICHE RADSPORT (gleiche Bezeichnungen wie beim Laufen, radsportspezifische Berechnung)
        st.markdown("---")
        if is_rad_pres:
            st.subheader("Trainingsbereiche")
        else:
            st.subheader("Trainingsbereiche – basierend auf LTP1 / Fatmax & ANS")

        def _hr_at_p_rad(p_val):
            return int(round(min(hfmax_val, max(40.0, float(p_val) * slope_val + intercept_val))))

        def _rnd_rad_w(v_val):
            return int(round(float(v_val) / 5.0) * 5)

        p_lt2_r = float(ans_power) if ans_power > 0 else 250.0
        p_lt1_r = float(fatmax_power) if (fatmax_power > 0 and fatmax_power < p_lt2_r * 0.92) else (p_lt2_r * 0.75)
        p_z1_max_r = p_lt2_r * 0.50
        p_z2_mid_r = (p_z1_max_r + p_lt1_r) / 2.0
        p_z3_max_r = p_lt2_r * 0.90
        p_z4_min_r = p_z3_max_r
        p_z4_max_r = p_lt2_r * 1.05
        p_z5_min_r = p_z4_max_r
        p_z5_max_r = max(p_z5_min_r + 30.0, float(map_val) * 1.10)

        try:
            _g1_idx_r = (df_sim['ee_kh_g'] - 0.6 * weight).abs().idxmin()
            p_z2_hi_zl_r = _rnd_rad_w(float(df_sim.loc[_g1_idx_r, 'power']))
            if not (p_z2_mid_r <= p_z2_hi_zl_r <= p_lt1_r):
                p_z2_hi_zl_r = _rnd_rad_w((p_z2_mid_r + p_lt1_r) / 2.0)
        except Exception:
            p_z2_hi_zl_r = _rnd_rad_w((p_z2_mid_r + p_lt1_r) / 2.0)

        zones_rad_ui = [
            (
                "Zone 1 - Regeneration",
                "-",
                f"{_rnd_rad_w(p_z1_max_r)}",
                f"{_rnd_rad_w(p_z1_max_r / 2.0)}",
                "-",
                f"{_hr_at_p_rad(p_z1_max_r)}",
                f"{int(round((_hr_at_p_rad(0.0) + _hr_at_p_rad(p_z1_max_r)) / 2.0))}",
            ),
            (
                "Zone 2 - Grundlagenbereich unten",
                f"{_rnd_rad_w(p_z1_max_r)}",
                f"{_rnd_rad_w(p_z2_mid_r)}",
                f"{_rnd_rad_w((p_z1_max_r + p_z2_mid_r) / 2.0)}",
                f"{_hr_at_p_rad(p_z1_max_r)}",
                f"{_hr_at_p_rad(p_z2_mid_r)}",
                f"{int(round((_hr_at_p_rad(p_z1_max_r) + _hr_at_p_rad(p_z2_mid_r)) / 2.0))}",
            ),
            (
                "Zone 2 - Grundlagenbereich oben",
                f"{_rnd_rad_w(p_z2_mid_r)}",
                f"{_rnd_rad_w(p_lt1_r)}",
                f"{p_z2_hi_zl_r}",
                f"{_hr_at_p_rad(p_z2_mid_r)}",
                f"{_hr_at_p_rad(p_lt1_r)}",
                f"{int(round((_hr_at_p_rad(p_z2_mid_r) + _hr_at_p_rad(p_lt1_r)) / 2.0))}",
            ),
            (
                "Zone 3 - Übergang / Tempo",
                f"{_rnd_rad_w(p_lt1_r)}",
                f"{_rnd_rad_w(p_z3_max_r)}",
                f"{_rnd_rad_w((p_lt1_r + p_z3_max_r) / 2.0)}",
                f"{_hr_at_p_rad(p_lt1_r)}",
                f"{_hr_at_p_rad(p_z3_max_r)}",
                f"{int(round((_hr_at_p_rad(p_lt1_r) + _hr_at_p_rad(p_z3_max_r)) / 2.0))}",
            ),
            (
                "Zone 4 - Schwelle",
                f"{_rnd_rad_w(p_z4_min_r)}",
                f"{_rnd_rad_w(p_z4_max_r)}",
                f"{_rnd_rad_w(p_lt2_r)}",
                f"{_hr_at_p_rad(p_z4_min_r)}",
                f"{_hr_at_p_rad(p_z4_max_r)}",
                f"{_hr_at_p_rad(p_lt2_r)}",
            ),
            (
                "Zone 5 - VO2max / HIT",
                f"{_rnd_rad_w(p_z5_min_r)}",
                f"{_rnd_rad_w(p_z5_max_r)}",
                f"{_rnd_rad_w((p_z5_min_r + p_z5_max_r) / 2.0)}",
                f"{_hr_at_p_rad(p_z5_min_r)}",
                "max",
                f"{int(round((_hr_at_p_rad(p_z5_min_r) + hfmax_val) / 2.0))}",
            ),
        ]
        df_zones_rad_ui = pd.DataFrame([
            {
                "Trainingsbereich": zn,
                "Leistung min [W]": p_min,
                "Leistung max [W]": p_max,
                "Leistung Ziel [W]": p_zl,
                "HF min [bpm]": hr_mn,
                "HF max [bpm]": hr_mx,
                "HF Ziel [bpm]": hr_zl,
            }
            for (zn, p_min, p_max, p_zl, hr_mn, hr_mx, hr_zl) in zones_rad_ui
        ])
        render_centered_html_table(df_zones_rad_ui)

        if not is_rad_pres:
            st.markdown("---")
            
            # Detailed Grid Table (nur in Diagnostik-Ansicht)
            with st.expander("🔍 Detaillierter metabolischer Verlauf (Tabelle in 10W-Schritten)"):
                df_table_rows = []
                for p in range(0, max_x_plot + 1, 10):
                    rows = df_sim[df_sim['power'] == p]
                    if not rows.empty:
                        row_data = rows.iloc[0]
                        df_table_rows.append({
                            'Leistung [W]': p,
                            'VO2ss [ml/min/kg]': round(row_data['vo2ss'], 1) if not np.isnan(row_data['vo2ss']) else "-",
                            'VLass [mmol/L/min]': round(row_data['vlass'], 3) if not np.isnan(row_data['vlass']) else "-",
                            'VLaoxmax [mmol/L/min]': round(row_data['vlaoxmax'], 3) if not np.isnan(row_data['vlaoxmax']) else "-",
                            'Pyruvatdefizit [mmol/L/min]': round(row_data['pyruvat_def'], 3) if not np.isnan(row_data['pyruvat_def']) else "-",
                            'KH-Anteil [%]': f"{round(row_data['kh_pct'], 1)}%" if not np.isnan(row_data['kh_pct']) else "-",
                            'Fett-Anteil [%]': f"{round(row_data['fat_pct'], 1)}%" if not np.isnan(row_data['fat_pct']) else "-",
                            'kalk. RQ': round(row_data['rq'], 3) if not np.isnan(row_data['rq']) else "-",
                            'EE gesamt [kcal/h]': int(round(row_data['ee_gesamt'], 0)) if not np.isnan(row_data['ee_gesamt']) else "-",
                            'EE Fett [kcal/h]': int(round(row_data['ee_fett'], 0)) if not np.isnan(row_data['ee_fett']) else "-",
                            'KH-Verbrauch [g/h]': round(row_data['ee_kh_g'], 1) if not np.isnan(row_data['ee_kh_g']) else "-",
                            'HF [bpm]': f"{int(round(row_data['hr'], 0))}" if not np.isnan(row_data['hr']) else "-"
                        })
                render_centered_html_table(pd.DataFrame(df_table_rows))

        # --- RADSPORT PDF DOWNLOAD (nur in der Diagnostik-Ansicht) ---
        if not is_rad_pres:
            st.markdown("---")
            st.subheader("Report-Generierung")
            
            # 1. Sprint raw values extraction or dummy fallback
            sprint_times = st.session_state.get("sprint_times", [])
            sprint_powers = st.session_state.get("sprint_powers", [])
            sprint_cadences = st.session_state.get("sprint_cadences", [])
            
            if not sprint_powers or len(sprint_powers) < 2:
                # Fallback dummy sprint values if none uploaded
                sprint_times = [i * 0.1 for i in range(120)]
                # Peak power around 1200W
                sprint_powers = [max(0.0, 1200.0 * (1.0 - ((t - 2.0)**2) / 64.0)) for t in sprint_times]
                sprint_cadences = [max(40.0, 110.0 - 2.0 * t) for t in sprint_times]

            # 2. Calculate sprint force metrics using weight and 0.1725 crank length
            avg_sprint_power = np.mean(sprint_powers)
            max_sprint_power = np.max(sprint_powers)
            
            forces = []
            for p, c in zip(sprint_powers, sprint_cadences):
                if c > 0:
                    # Force F = Power / (crank_length * cadence * 2*pi / 60)
                    f = p / (0.1725 * c * 2.0 * np.pi / 60.0)
                else:
                    f = 0.0
                forces.append(f)
                
            max_force = np.max(forces) if forces else 0.0
            
            # Force at Pmax (find index of maximum power)
            pmax_idx = np.argmax(sprint_powers) if sprint_powers else 0
            force_at_pmax = forces[pmax_idx] if forces and pmax_idx < len(forces) else 0.0
            
            force_pct_max = (force_at_pmax / max_force * 100.0) if max_force > 0 else 0.0
            
            # 3. Dynamic percentage of VO2max at Anaerobic Threshold
            ans_vo2_pct = (ans_vo2_abs / abs_vo2max * 100.0) if abs_vo2max > 0 else 0.0
            rel_vo2max = (abs_vo2max / weight) if weight > 0 else 0.0
            
            # 4. Coaching potential overrides
            pot_vo2max = float(st.session_state.get("pot_vo2max", 72.0))
            pot_fat = float(st.session_state.get("pot_fat", 10.5))
            pot_ans_abs = float(st.session_state.get("pot_ans_abs", 295.0))
            pot_ans_rel = float(st.session_state.get("pot_ans_rel", 4.8))
            pot_fatmax = float(st.session_state.get("pot_fatmax", 3.4))
            pot_pmax = float(st.session_state.get("pot_pmax", 14.5))
            pot_match = float(st.session_state.get("pot_match", 3.6))
            pot_vlamax = float(st.session_state.get("pot_vlamax", 0.45))
            
            # Create PDF Bytes
            pdf_bytes = None
            pdf_error = None
            with st.spinner("PDF wird erstellt..."):
                try:
                    pdf_bytes = create_pdf_rad(
                        athlete_name=athlete_name,
                        birthdate=birthdate,
                        test_date=test_date,
                        test_type=st.session_state.get("diagnostik_type", "Cycling BLUE"),
                        weight=weight,
                        body_fat_pct=body_fat_pct,
                        rel_vo2max=rel_vo2max,
                        ans_power=ans_power,
                        ans_rel=ans_rel,
                        ans_ee_gesamt=ans_ee_gesamt,
                        ans_hr=ans_hr,
                        ans_vo2_abs=ans_vo2_abs,
                        ans_vo2_pct=ans_vo2_pct,
                        ans_ee_kh_g=ans_ee_kh_g,
                        fatmax_power=fatmax_power,
                        fatmax_rel=fatmax_rel,
                        fatmax_ee_gesamt=fatmax_ee_gesamt,
                        fatmax_ee_fett=fatmax_ee_fett,
                        map_val=map_val,
                        carb_match_power=carb_match_power,
                        carb_intake=carb_intake,
                        vlamax=vlamax,
                        avg_rest_lac=avg_rest_lac,
                        max_post_lac=max_post_lac,
                        t_glyc=t_glyc,
                        t_bel=t_bel,
                        t_alak=t_alak,
                        coach=coach,
                        sportart="Radsport",
                        kategorie=kategorie,
                        height=height,
                        slope_val=slope_val,
                        intercept_val=intercept_val,
                        hfmax_val=hfmax_val,
                        df_sim=df_sim,
                        sprint_times=sprint_times,
                        sprint_powers=sprint_powers,
                        sprint_cadences=sprint_cadences,
                        avg_sprint_power=avg_sprint_power,
                        max_sprint_power=max_sprint_power,
                        max_force=max_force,
                        force_at_pmax=force_at_pmax,
                        force_pct_max=force_pct_max,
                        pot_vo2max=pot_vo2max,
                        pot_fat=pot_fat,
                        pot_ans_abs=pot_ans_abs,
                        pot_ans_rel=pot_ans_rel,
                        pot_fatmax=pot_fatmax,
                        pot_pmax=pot_pmax,
                        pot_match=pot_match,
                        pot_vlamax=pot_vlamax,
                        gender=st.session_state.get("gender", "männlich")
                    )
                except Exception as e:
                    import traceback
                    pdf_error = traceback.format_exc()

            if pdf_bytes is not None:
                excel_bytes_rad = None
                try:
                    excel_bytes_rad = create_hycys_excel_rad(
                        athlete_name=athlete_name,
                        birthdate=birthdate,
                        test_date=test_date,
                        diagnostik_type=st.session_state.get("diagnostik_type", "Cycling BLUE"),
                        coach=coach,
                        kategorie=kategorie,
                        height=height,
                        weight=weight,
                        body_fat_pct=body_fat_pct,
                        df_sf=st.session_state.get("df_sf"),
                        df_sprint_lac=st.session_state.get("df_sprint_lac"),
                        ppo_val=ppo_val,
                        abs_vo2max=abs_vo2max,
                        t_alak=t_alak,
                        t_bel=t_bel,
                        slope_val=slope_val,
                        intercept_val=intercept_val,
                        hfmax_val=hfmax_val,
                        sprint_times=sprint_times,
                        sprint_powers=sprint_powers,
                        sprint_cadences=sprint_cadences,
                        uploaded_spiro=uploaded_spiro,
                    )
                except Exception as e_xls_rad:
                    st.warning(f"Hinweis: HYCYS-Excel konnte nicht erzeugt werden ({e_xls_rad}).")

                render_gdrive_upload_section(
                    pdf_bytes=pdf_bytes,
                    diagnostik_type=st.session_state.get("diagnostik_type", "Cycling BLUE"),
                    athlete_name=athlete_name,
                    test_date=test_date,
                    gender_str=st.session_state.get("gender", "männlich"),
                    key_prefix="radsport",
                    excel_bytes=excel_bytes_rad,
                )
            else:
                st.error(f"Fehler bei der PDF-Erstellung: Bitte Eingaben prüfen.")
                if pdf_error:
                    with st.expander("Fehlerdetails anzeigen"):
                        st.code(pdf_error)

    else:
        # Default view before starting the calculation
        st.markdown(f"""
            <div style="font-size: 16px; line-height: 1.5; margin-bottom: 20px;">
                <b>Athlet:</b> {athlete_name}<br>
                <b>Geschlecht:</b> {st.session_state.get("gender", "männlich")}<br>
                <b>Sportart:</b> Radsport<br>
                <b>Test-Datum:</b> {test_date}
            </div>
            """, unsafe_allow_html=True)
            
        if st.session_state.get("rad_kf_warnung_offen", False):
            st.warning("⚠️ **Hinweis:** Es wurden keine Körperfettwerte eingegeben. Willst du trotzdem ohne Körperfett auswerten?")
            col_kf_m1, col_kf_m2, _ = st.columns([1.2, 1.0, 2.5])
            with col_kf_m1:
                if st.button("Ja, trotzdem weitermachen", type="primary", use_container_width=True, key="btn_rad_kf_yes_main"):
                    st.session_state.rad_kf_warnung_offen = False
                    st.session_state.rad_auswertung_gestartet = True
                    st.rerun()
            with col_kf_m2:
                if st.button("Abbrechen & KF eingeben", use_container_width=True, key="btn_rad_kf_no_main"):
                    st.session_state.rad_kf_warnung_offen = False
                    st.rerun()

        st.info("💡 Radsport-Diagnostik: Bitte laden Sie die gewünschten Dateien hoch und klicken Sie in der Seitenleiste auf **'Auswertung starten'**, um die Stoffwechselprofile zu generieren.")
        
        st.subheader("Schritt 1: Dateien hochladen")
        c1, c2, c3 = st.columns(3)
        with c1:
            if uploaded_spiro is not None:
                st.success(f"Spirodatei geladen:\n{uploaded_spiro.name}")
                st.caption(f"Berechnete VO2max: {st.session_state.get('vo2max_spiro', 4292.2):.0f} ml/min")
            else:
                st.warning("Keine Spirodatei hochgeladen")
        with c2:
            if uploaded_fit is not None:
                st.success(f".fit-Datei geladen:\n{uploaded_fit.name}")
                st.caption(f"Steigung: {st.session_state.get('slope_val', 0.11):.2f} | Nullstelle: {st.session_state.get('intercept_val', 138):.1f}")
            else:
                st.warning("Keine .fit-Datei hochgeladen")
        with c3:
            if uploaded_sprint is not None:
                st.success(f"Sprintdatei geladen:\n{uploaded_sprint.name}")
            elif st.session_state.get("sprint_src_name"):
                st.success(f"Sprintdatei geladen:\n{st.session_state.sprint_src_name}")
            else:
                st.warning("Keine Sprintdatei hochgeladen")
                
        st.subheader("Schritt 2: Laktatwerte eingeben")
        st.subheader("Schritt 3: Auswertung starten")
        
    save_current_session_snapshot()
    st.stop()


# Testprotokoll selection placed at the very top of the sidebar for Laufauswertung
test_types = [
    "HYCYS Standart (5min / +0,4m/s)",
    "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)"
]

# 1. Upload (Spiro file)
uploaded_file = st.sidebar.file_uploader("Spiro-Datei hochladen (.xlsx oder .csv)", type=["xlsx", "csv"], key="uploaded_file_key")

separate_ramp = st.sidebar.checkbox("Stufe und Rampe getrennt", key="separate_ramp_checkbox")
uploaded_ramp = None
if separate_ramp:
    uploaded_ramp = st.sidebar.file_uploader("Rampendatei (Spiro, 10s Format) hochladen (.xlsx oder .csv)", type=["xlsx", "csv"], key="uploaded_ramp_key")

# 2. Testprotokoll (eingeklappt)
if "selected_threshold_model" not in st.session_state:
    st.session_state.selected_threshold_model = "OBLA 4.0 (ANS)"

# Falls ein vorheriger Rerun (z. B. Datei-Upload) Widget-Werte vor der Instanziierung synchronisieren möchte:
if "_pending_lauft_widget_sync" in st.session_state:
    for _k_sync, _v_sync in st.session_state["_pending_lauft_widget_sync"].items():
        st.session_state[_k_sync] = _v_sync
    del st.session_state["_pending_lauft_widget_sync"]

def _on_sidebar_thresh_change():
    st.session_state.selected_threshold_model = st.session_state.selected_threshold_model_widget

def _on_test_type_change():
    new_tt = st.session_state.get("test_type_lauft_widget", st.session_state.get("last_test_type"))
    st.session_state.last_test_type = new_tt
    if new_tt == "HYCYS Standart (5min / +0,4m/s)":
        st.session_state.start_v = 2.5
        st.session_state.v_increment = 0.400
        st.session_state.anzahl = 5
        st.session_state.stufendauer = 5.0
    elif new_tt == "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)":
        st.session_state.start_v = 2.5
        st.session_state.v_increment = 0.278
        st.session_state.anzahl = 8
        st.session_state.stufendauer = 3.0
    st.session_state.vorlauf = 60
    st.session_state.pausendauer = 30

    # In on_change (vor Widget-Instanziierung) dürfen die Widget-Keys direkt gesetzt werden:
    st.session_state.start_v_lauft_widget = st.session_state.start_v
    st.session_state.anzahl_lauft_widget = st.session_state.anzahl
    st.session_state.stufendauer_lauft_widget = st.session_state.stufendauer
    st.session_state.vorlauf_lauft_widget = st.session_state.vorlauf
    st.session_state.pausendauer_lauft_widget = st.session_state.pausendauer

    if "df_lauftest_input" in st.session_state:
        del st.session_state["df_lauftest_input"]
    if "lauftest_lac_editor_key" in st.session_state:
        del st.session_state["lauftest_lac_editor_key"]
    st.session_state.lauf_auswertung_gestartet = False
    for key in list(st.session_state.keys()):
        if key.startswith("lauftest_end_t_") or key.startswith("lauftest_offset_"):
            del st.session_state[key]

with st.sidebar.expander("Testprotokoll", expanded=False):
    test_type = st.selectbox(
        "Testprotokoll",
        options=test_types,
        key="test_type_lauft_widget",
        on_change=_on_test_type_change,
    )

    st.session_state.selected_threshold_model_widget = st.session_state.selected_threshold_model
    selected_model_name = st.selectbox(
        "Auswertungs-Schwelle (LT2 / ANS)",
        options=["OBLA 4.0 (ANS)", "LTP2"],
        key="selected_threshold_model_widget",
        on_change=_on_sidebar_thresh_change
    )
    
    # Step parameters (inside the expander)
    start_v = st.number_input("Startgeschwindigkeit (m/s)", value=float(st.session_state.start_v), step=0.1, key="start_v_lauft_widget")
    st.session_state.start_v = start_v

    # Geschwindigkeitssteigerung fest über das gewählte Testprotokoll vorgegeben (nicht manuell auswählbar)
    v_increment = 0.278 if test_type == "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)" else 0.400
    st.session_state.v_increment = v_increment

    anzahl = st.number_input("Anzahl Stufen", value=int(st.session_state.anzahl), min_value=1, max_value=20, step=1, key="anzahl_lauft_widget")
    st.session_state.anzahl = anzahl

    # Collapsible Protokoll-Setup expander (nested, collapsed)
    with st.expander("Protokoll-Setup", expanded=False):
        vorlauf = st.number_input("Ruhemessung (Sekunden)", value=int(st.session_state.vorlauf), step=10, key="vorlauf_lauft_widget")
        st.session_state.vorlauf = vorlauf
        
        stufendauer = st.number_input("Stufendauer (Minuten)", value=float(st.session_state.stufendauer), step=0.5, key="stufendauer_lauft_widget")
        st.session_state.stufendauer = stufendauer
        
        pausendauer = st.number_input("Pausendauer (Sekunden)", value=int(st.session_state.pausendauer), step=5, key="pausendauer_lauft_widget")
        st.session_state.pausendauer = pausendauer
        
        ausbelastung = st.checkbox("Test bis zur Ausbelastung", value=True, key="ausbelastung_lauft_widget")

def detect_spiro_stage_ends(df_spiro, anzahl):
    n_rows = len(df_spiro)
    if n_rows < 10:
        return []
    t_diffs = df_spiro['Time'].diff().dropna()
    dt = t_diffs.median() if len(t_diffs) > 0 else 10.0
    step_30s = max(1, int(round(30.0 / dt)))
    step_60s = max(step_30s, int(round(60.0 / dt)))
    drops = []
    for i in range(n_rows - step_30s):
        t1 = float(df_spiro['Time'].iloc[i])
        vo2_1 = float(df_spiro['VO2'].iloc[i])
        if vo2_1 <= 0 or pd.isna(vo2_1):
            continue
        vo2_2 = float(df_spiro['VO2'].iloc[i + step_30s])
        t2 = float(df_spiro['Time'].iloc[i + step_30s])
        pct_drop = (vo2_1 - vo2_2) / vo2_1 if (20 <= (t2 - t1) <= 45 and pd.notna(vo2_2)) else 0.0
        # Falls der Abfall erst im zweiten 30s-Intervall der Pause voll sichtbar wird (z. B. bei 30s-Mischkammer):
        if i + step_60s < n_rows:
            vo2_3 = float(df_spiro['VO2'].iloc[i + step_60s])
            t3 = float(df_spiro['Time'].iloc[i + step_60s])
            if 45 <= (t3 - t1) <= 75 and pd.notna(vo2_3) and vo2_2 < vo2_1 * 0.97:
                pct_drop_60 = (vo2_1 - vo2_3) / vo2_1
                if pct_drop_60 > pct_drop:
                    pct_drop = pct_drop_60
        drops.append((t1, pct_drop, vo2_1))
    candidate_drops = [d for d in drops if d[1] > 0.15 and d[0] > 120]
    candidate_drops.sort(key=lambda x: x[0])
    grouped_ends = []
    for d in candidate_drops:
        t, pct, val = d
        if not grouped_ends or t - grouped_ends[-1][0] > 90:
            grouped_ends.append(d)
        else:
            if val > grouped_ends[-1][2]:
                grouped_ends[-1] = d
    detected_times = [g[0] for g in grouped_ends]
    detected_times.sort()
    return [float(x) for x in detected_times[:anzahl]]


if uploaded_file is not None and uploaded_file.name != st.session_state.last_uploaded_file:
    st.session_state.last_uploaded_file = uploaded_file.name
    try:
        df_excel = pd.read_csv(uploaded_file, header=None, low_memory=False) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
        
        # Name
        name_val = None
        if df_excel.shape[0] > 2 and df_excel.shape[1] > 1:
            name_val = f"{df_excel.iloc[2, 1]} {df_excel.iloc[1, 1]}"
        
        # Height
        h_val = None
        try:
            if df_excel.shape[0] > 5 and df_excel.shape[1] > 1:
                h_val = int(float(df_excel.iloc[5, 1]))
        except:
            pass
            
        # Weight
        w_val = None
        try:
            if df_excel.shape[0] > 6 and df_excel.shape[1] > 1:
                w_val = round(float(df_excel.iloc[6, 1]), 1)
        except:
            pass
            
        # Birthdate
        b_val = None
        if df_excel.shape[0] > 7 and df_excel.shape[1] > 1:
            raw_date = df_excel.iloc[7, 1] 
            if pd.notna(raw_date):
                date_str = str(int(float(raw_date))).zfill(8)
                b_val = f"{date_str[:2]}.{date_str[2:4]}.{date_str[4:]}"
            
        # Testdatum
        t_val = None
        if df_excel.shape[0] > 0 and df_excel.shape[1] > 4:
            raw_test_date = df_excel.iloc[0, 4]
            if pd.notna(raw_test_date):
                try:
                    t_date_str = str(int(float(raw_test_date))).zfill(8)
                    t_val = f"{t_date_str[:2]}.{t_date_str[2:4]}.{t_date_str[4:]}"
                except:
                    t_val = str(raw_test_date)

        # Geschlecht
        g_val = None
        if df_excel.shape[0] > 3 and df_excel.shape[1] > 1:
            raw_gender = df_excel.iloc[3, 1]
            if pd.notna(raw_gender):
                raw_gender_str = str(raw_gender).strip().lower()
                if "w" in raw_gender_str or "fem" in raw_gender_str or "weib" in raw_gender_str:
                    g_val = "weiblich"
                else:
                    g_val = "männlich"
                    
        update_base_data(name_val, h_val, w_val, b_val, t_val, g_val)
        
        # Parse and save spiro data to session state to prevent slow Excel reloading
        try:
            times_sec = df_excel.iloc[3:, 9].apply(parse_time_to_seconds)
            vo2_col = pd.to_numeric(df_excel.iloc[3:, 13], errors='coerce')
            vco2_col = pd.to_numeric(df_excel.iloc[3:, 14], errors='coerce')
            eeh_col = pd.to_numeric(df_excel.iloc[3:, 36], errors='coerce') if df_excel.shape[1] > 36 else pd.Series(np.nan, index=df_excel.iloc[3:].index)
            fett_kcal_h_col = (pd.to_numeric(df_excel.iloc[3:, 41], errors='coerce') / 24.0) if df_excel.shape[1] > 41 else pd.Series(np.nan, index=df_excel.iloc[3:].index)
            kh_g_h_col = ((pd.to_numeric(df_excel.iloc[3:, 42], errors='coerce') / 24.0) / 4.063) if df_excel.shape[1] > 42 else pd.Series(np.nan, index=df_excel.iloc[3:].index)
            spiro_df = pd.DataFrame({
                'Time': times_sec,
                'VO2': vo2_col,
                'VCO2': vco2_col,
                'EEh': eeh_col,
                'FETT_kcal_h': fett_kcal_h_col,
                'KH_g_h': kh_g_h_col,
            }).dropna(subset=['Time'])
            st.session_state.lauftest_spiro_df = spiro_df
            
            # Detect stage ends and store in session state
            detected = detect_spiro_stage_ends(spiro_df, 20)
            if detected:
                st.session_state.detected_stage_ends = detected
                if len(detected) >= 2:
                    med_interval = float(np.median(np.diff(detected)))
                    if 280.0 <= med_interval <= 380.0:
                        st.session_state.stufendauer = 5.0
                        st.session_state["_pending_lauft_widget_sync"] = {"stufendauer_lauft_widget": 5.0}
                    elif 170.0 <= med_interval <= 250.0:
                        st.session_state.stufendauer = 3.0
                        st.session_state["_pending_lauft_widget_sync"] = {"stufendauer_lauft_widget": 3.0}
            else:
                if "detected_stage_ends" in st.session_state:
                    del st.session_state.detected_stage_ends
        except Exception as detection_err:
            pass

        # Reset manual end times and offsets
        for key in list(st.session_state.keys()):
            if key.startswith("lauftest_end_t_") or key.startswith("lauftest_offset_"):
                del st.session_state[key]

        # Reset started flag
        st.session_state.lauf_auswertung_gestartet = False

        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Fehler beim Extrahieren der Stammdaten: {e}")

# Kategorie, Coach, and Basisdaten reorganized
# Pre-calculate body fat percentage if skinfolds table has entries (Laufen)
if "df_sf" in st.session_state:
    sf_means = st.session_state.df_sf[["M1", "M2", "M3"]].mean(axis=1)
    sum_sf = sf_means.sum()
    if sum_sf > 0:
        calculated_bf = (22.32 * np.log10(sum_sf)) - 29.2
        st.session_state.body_fat_pct = max(0.0, round(calculated_bf, 1))

# 3. Basisdaten (collapsed expander)
with st.sidebar.expander("Basisdaten", expanded=False):
    athlete_name = st.text_input("Name", value=st.session_state.athlete_name, key="athlete_name_lauft_widget")
    st.session_state.athlete_name = athlete_name

    birthdate = st.text_input("Geburtsdatum", value=st.session_state.birthdate, key="birthdate_lauft_widget")
    st.session_state.birthdate = birthdate

    test_date = st.text_input("Testdatum", value=st.session_state.test_date, key="test_date_lauft_widget")
    st.session_state.test_date = test_date

    # Größe formatted as 123cm without decimals
    height = st.number_input("Größe (cm)", value=int(st.session_state.height), step=1, format="%d", key="height_lauft_widget")
    st.session_state.height = height

    # Gewicht (float, 1 decimal)
    weight = st.number_input("Gewicht (kg)", value=float(st.session_state.weight), step=0.1, format="%.1f", key="weight_lauft_widget")
    st.session_state.weight = round(weight, 1)

    # Kategorie selectbox
    kategorie_options = ["Hobby", "Age-Grouper", "Amateur", "Profi"]
    kategorie_val = st.session_state.kategorie
    if kategorie_val not in kategorie_options:
        kategorie_options = kategorie_options + [kategorie_val]
    kategorie = st.selectbox("Kategorie", options=kategorie_options, index=kategorie_options.index(kategorie_val), key="kategorie_lauft_widget")
    st.session_state.kategorie = kategorie

    # Geschlecht selectbox
    gender_options = ["männlich", "weiblich"]
    gender_val = st.session_state.get("gender", "männlich")
    gender = st.selectbox("Geschlecht", options=gender_options, index=gender_options.index(gender_val), key="gender_lauft_widget")
    st.session_state.gender = gender

    # Coach selectbox
    coach_options = list(HYCYS_COACH_LIST)
    coach_val = st.session_state.coach
    if coach_val not in coach_options:
        coach_options = coach_options + [coach_val]
    coach = st.selectbox("Coach", options=coach_options, index=coach_options.index(coach_val), key="coach_lauft_widget")
    st.session_state.coach = coach

    diag_type_options = ["Running WHITE", "Running BLUE", "Running GOLD"]
    diag_type_val = st.session_state.get("diagnostik_type", "Running BLUE")
    if diag_type_val not in diag_type_options:
        diag_type_val = "Running BLUE"
    diagnostik_type = st.selectbox("Diagnostik Typ", options=diag_type_options, index=diag_type_options.index(diag_type_val), key="diagnostik_type_lauft_widget")
    st.session_state.diagnostik_type = diagnostik_type

# 3b. Körperfettwerte (eigenes Feld unter Basisdaten und oberhalb von Stufentest Laktat)
with st.sidebar.expander("Körperfettwerte", expanded=False):
    body_fat_pct = st.number_input("Körperfett (%)", value=float(st.session_state.body_fat_pct), step=0.1, format="%.1f", key="bf_lauft_widget")
    st.session_state.body_fat_pct = round(body_fat_pct, 1)

    st.caption("Körperfettmessung (Parizkova 10-Falten):")
    edited_sf = st.data_editor(
        st.session_state.df_sf,
        hide_index=True,
        width='stretch',
        key="edited_sf_key"
    )

# 4. Stufentest Laktat (collapsed expander)
sync_lauftest_input_df()

with st.sidebar.expander("Stufentest Laktat", expanded=False):
    edited_df = st.data_editor(
        st.session_state.df_lauftest_input,
        disabled=["Stufe", "v (m/s)", "v (min/km)"],
        hide_index=True,
        use_container_width=True,
        height=320,
        column_config={
            "Stufe": st.column_config.TextColumn("Stufe", width="small"),
            "v (m/s)": st.column_config.NumberColumn("v (m/s)", format="%.2f", width="small"),
            "v (min/km)": st.column_config.TextColumn("v (min/km)", width="small"),
            "Laktat": st.column_config.NumberColumn("Laktat", min_value=0.0, max_value=25.0, step=0.01, format="%.2f", width="small"),
            "HF": st.column_config.NumberColumn("HF", min_value=30, max_value=240, step=1, format="%d", width="small")
        },
        key="lauftest_lac_editor_key"
    )

# 5. Spiro-Fenster Feineinstellung (collapsed expander, outside any form so it reruns immediately)
with st.sidebar.expander("Spiro-Fenster Feineinstellung", expanded=False):
    vo2max_override = st.number_input(
        "VO2max überschreiben (ml/min)", 
        value=0, 
        step=50, 
        help="Ermöglicht das manuelle Überschreiben des berechneten VO2max-Wertes (z.B. 3652 für Barbara Bauer)"
    )
    
    if uploaded_file is not None:
        st.caption("Endzeitpunkt (in Sekunden) eingeben:")
        unregel = st.checkbox("Es gab während des Tests Unregelmäßigkeiten in der Stufendauer", value=False, key="unregel_stufen_widget")
        
        anz = int(st.session_state.anzahl)
        stage_sec = int(st.session_state.stufendauer * 60)
        pause_sec = int(st.session_state.pausendauer)
        vorlauf = int(st.session_state.vorlauf)
        interval = stage_sec + pause_sec

        if "detected_stage_ends" not in st.session_state and "lauftest_spiro_df" in st.session_state:
            _det = detect_spiro_stage_ends(st.session_state.lauftest_spiro_df, 20)
            if _det:
                st.session_state.detected_stage_ends = _det

        det_list_raw = st.session_state.get("detected_stage_ends", [])
        det_list = []
        if len(det_list_raw) >= 2:
            det_interval = int(round(float(np.median(np.diff(det_list_raw)))))
            # Nur übernehmen, wenn das erkannte Stufenintervall zur gewählten Stufendauer passt (±45s)
            if abs(det_interval - (stage_sec + pause_sec)) <= 45:
                interval = det_interval
                det_list = det_list_raw
        
        for i in range(anz):
            if i < len(det_list):
                end_t_base = int(det_list[i])
            elif len(det_list) >= 2:
                end_t_base = int(det_list[-1]) + (i - len(det_list) + 1) * interval
            else:
                end_t_base = int(vorlauf + (i + 1) * stage_sec + i * pause_sec)
            
            end_t_key = f"lauftest_end_t_{i}"
            
            if i == 0:
                if end_t_key not in st.session_state:
                    st.session_state[end_t_key] = end_t_base
                new_end_t = st.number_input(
                    f"Stufe {i+1} Endzeit (s)",
                    min_value=0,
                    max_value=10000,
                    value=int(st.session_state[end_t_key]),
                    step=1,
                    key=f"lauftest_end_t_widget_{i}"
                )
                st.session_state[end_t_key] = new_end_t
            else:
                if len(det_list) >= 2 and int(st.session_state["lauftest_end_t_0"]) == int(det_list[0]):
                    if i < len(det_list):
                        calculated_end_t = int(det_list[i])
                    else:
                        calculated_end_t = int(det_list[-1]) + (i - len(det_list) + 1) * interval
                else:
                    calculated_end_t = int(st.session_state["lauftest_end_t_0"]) + i * interval
                
                if unregel:
                    if end_t_key not in st.session_state:
                        st.session_state[end_t_key] = calculated_end_t
                    new_end_t = st.number_input(
                        f"Stufe {i+1} Endzeit (s)",
                        min_value=0,
                        max_value=10000,
                        value=int(st.session_state[end_t_key]),
                        step=1,
                        key=f"lauftest_end_t_widget_{i}"
                    )
                    st.session_state[end_t_key] = new_end_t
                else:
                    st.session_state[end_t_key] = calculated_end_t

st.sidebar.markdown("---")
start_button = st.sidebar.button("Auswertung starten", type="primary")
render_sidebar_account_and_remote_support("lauf")

def check_lauftest_lactate(df):
    lac_series = pd.to_numeric(df["Laktat"], errors="coerce")
    if len(df) > 1:
        return not lac_series.iloc[1:].isna().any()
    return not lac_series.isna().any()

def check_lauftest_hf_complete(df):
    hf_series = pd.to_numeric(df["HF"], errors="coerce")
    if len(df) > 1:
        return not (hf_series.iloc[1:].isna() | (hf_series.iloc[1:] <= 0)).any()
    return not (hf_series.isna() | (hf_series <= 0)).any()

if start_button:
    _sf_sum_lauf = float(edited_sf[["M1", "M2", "M3"]].mean(axis=1).sum()) if edited_sf is not None and not edited_sf.empty else 0.0
    if _sf_sum_lauf > 0:
        _calc_bf_lauf = max(0.0, round((22.32 * np.log10(_sf_sum_lauf)) - 29.2, 1))
        st.session_state.body_fat_pct = _calc_bf_lauf
    else:
        _calc_bf_lauf = float(st.session_state.get("body_fat_pct", 0.0))

    if not check_lauftest_lactate(edited_df):
        st.session_state.lauf_auswertung_gestartet = False
        st.session_state.lauf_hf_warnung_offen = False
        st.session_state.lauf_kf_warnung_offen = False
        st.sidebar.error("Bitte alle Laktat-Werte ausfüllen, bevor die Auswertung gestartet wird.")
    elif not check_lauftest_hf_complete(edited_df):
        st.session_state.df_lauftest_input = edited_df
        st.session_state.df_sf = edited_sf
        st.session_state.lauf_auswertung_gestartet = False
        st.session_state.lauf_hf_warnung_offen = True
        st.session_state.lauf_kf_warnung_offen = (_calc_bf_lauf <= 0.0 and _sf_sum_lauf <= 0.0)
    elif _calc_bf_lauf <= 0.0 and _sf_sum_lauf <= 0.0:
        st.session_state.df_lauftest_input = edited_df
        st.session_state.df_sf = edited_sf
        st.session_state.lauf_auswertung_gestartet = False
        st.session_state.lauf_hf_warnung_offen = False
        st.session_state.lauf_kf_warnung_offen = True
    else:
        st.session_state.df_lauftest_input = edited_df
        st.session_state.df_sf = edited_sf
        st.session_state.lauf_hf_warnung_offen = False
        st.session_state.lauf_kf_warnung_offen = False
        st.session_state.lauf_auswertung_gestartet = True
elif st.session_state.get("lauf_auswertung_gestartet", False):
    if check_lauftest_lactate(edited_df):
        st.session_state.df_lauftest_input = edited_df
        st.session_state.df_sf = edited_sf
    else:
        st.session_state.lauf_auswertung_gestartet = False
        st.sidebar.error("Auswertung gestoppt: Es fehlen Laktat-Werte.")

if st.session_state.get("lauf_hf_warnung_offen", False):
    st.sidebar.warning("⚠️ Die HF-Werte sind nicht vollständig. Willst du trotzdem weitermachen?")
    col_hf_w1, col_hf_w2 = st.sidebar.columns(2)
    with col_hf_w1:
        if st.button("Ja, weiter", type="primary", use_container_width=True, key="btn_hf_warn_yes_sb"):
            st.session_state.df_lauftest_input = edited_df
            st.session_state.df_sf = edited_sf
            st.session_state.lauf_hf_warnung_offen = False
            if not st.session_state.get("lauf_kf_warnung_offen", False):
                st.session_state.lauf_auswertung_gestartet = True
            st.rerun()
    with col_hf_w2:
        if st.button("Abbrechen", use_container_width=True, key="btn_hf_warn_no_sb"):
            st.session_state.lauf_hf_warnung_offen = False
            st.session_state.lauf_kf_warnung_offen = False
            st.rerun()

    st.warning("⚠️ **Hinweis:** Die HF-Werte sind nicht vollständig. Willst du trotzdem weitermachen?")
    col_hf_m1, col_hf_m2, _ = st.columns([1.2, 1.0, 2.5])
    with col_hf_m1:
        if st.button("Ja, trotzdem weitermachen", type="primary", use_container_width=True, key="btn_hf_warn_yes_main"):
            st.session_state.df_lauftest_input = edited_df
            st.session_state.df_sf = edited_sf
            st.session_state.lauf_hf_warnung_offen = False
            if not st.session_state.get("lauf_kf_warnung_offen", False):
                st.session_state.lauf_auswertung_gestartet = True
            st.rerun()
    with col_hf_m2:
        if st.button("Abbrechen & HF eingeben", use_container_width=True, key="btn_hf_warn_no_main"):
            st.session_state.lauf_hf_warnung_offen = False
            st.session_state.lauf_kf_warnung_offen = False
            st.rerun()

if st.session_state.get("lauf_kf_warnung_offen", False) and not st.session_state.get("lauf_hf_warnung_offen", False):
    if hasattr(st, "dialog"):
        @st.dialog("⚠️ Keine Körperfettwerte eingegeben")
        def _confirm_no_bf_lauf_dialog():
            st.warning("Es wurden keine Körperfettwerte (bzw. Hautfalten) eingegeben. Möchtest du die Auswertung trotzdem ohne Körperfettwerte starten?")
            c_dl1, c_dl2 = st.columns(2)
            with c_dl1:
                if st.button("Ja, ohne KF auswerten", type="primary", use_container_width=True, key="btn_lauf_kf_dlg_yes"):
                    st.session_state.df_lauftest_input = edited_df
                    st.session_state.df_sf = edited_sf
                    st.session_state.lauf_kf_warnung_offen = False
                    st.session_state.lauf_auswertung_gestartet = True
                    st.rerun()
            with c_dl2:
                if st.button("Abbrechen", use_container_width=True, key="btn_lauf_kf_dlg_no"):
                    st.session_state.lauf_kf_warnung_offen = False
                    st.rerun()

        _confirm_no_bf_lauf_dialog()

    st.sidebar.warning("⚠️ Es wurden keine Körperfettwerte eingegeben. Willst du trotzdem ohne Körperfett auswerten?")
    col_kf_l1, col_kf_l2 = st.sidebar.columns(2)
    with col_kf_l1:
        if st.button("Ja, weiter", type="primary", use_container_width=True, key="btn_lauf_kf_yes_sb"):
            st.session_state.df_lauftest_input = edited_df
            st.session_state.df_sf = edited_sf
            st.session_state.lauf_kf_warnung_offen = False
            st.session_state.lauf_auswertung_gestartet = True
            st.rerun()
    with col_kf_l2:
        if st.button("Abbrechen", use_container_width=True, key="btn_lauf_kf_no_sb"):
            st.session_state.lauf_kf_warnung_offen = False
            st.rerun()

    st.warning("⚠️ **Hinweis:** Es wurden keine Körperfettwerte eingegeben. Willst du trotzdem ohne Körperfett auswerten?")
    col_kf_lm1, col_kf_lm2, _ = st.columns([1.2, 1.0, 2.5])
    with col_kf_lm1:
        if st.button("Ja, trotzdem weitermachen", type="primary", use_container_width=True, key="btn_lauf_kf_yes_main"):
            st.session_state.df_lauftest_input = edited_df
            st.session_state.df_sf = edited_sf
            st.session_state.lauf_kf_warnung_offen = False
            st.session_state.lauf_auswertung_gestartet = True
            st.rerun()
    with col_kf_lm2:
        if st.button("Abbrechen & KF eingeben", use_container_width=True, key="btn_lauf_kf_no_main"):
            st.session_state.lauf_kf_warnung_offen = False
            st.rerun()

# --- HILFSFUNKTIONEN ---

def calculate_vlamax_for_v(v_thresh, rel_vo2max, speed_arr, vo2_steady_abs, vco2_steady_abs, weight):
    if not (rel_vo2max > 0 and v_thresh > 0 and len(vo2_steady_abs) > 0 and weight > 0): 
        return None, None, None
        
    vo2_rest = 3.5 * weight
    re_abs_m_s = []
    ex_speeds = []
    
    # Calculate VO2_CHO100 running economy for exercise stages with valid Spiro data
    # (inkl. der ersten Stufe unmittelbar über v_thresh für die lineare Interpolation wie in HYCYS Excel,
    # weitere überschwellige Stufen darüber hinaus werden ignoriert)
    passed_thresh = False
    for v, vo2, vco2 in zip(speed_arr, vo2_steady_abs, vco2_steady_abs):
        if v <= 0 or vo2 <= 0:
            continue
        if passed_thresh:
            break
        if v > v_thresh + 1e-6:
            passed_thresh = True
        vo2_working = max(0.0, vo2 - vo2_rest)
        if vo2_working <= 0:
            continue
        if vo2 > 0 and vco2 > 0:
            rq = vco2 / vo2
        else:
            rq = 0.85
        cal_eq = (rq - 0.7) * 5.094 + 19.6
        vo2_cho100 = vo2_working * cal_eq / 21.1
        re_abs_m_s.append(vo2_cho100 / v)
        ex_speeds.append(v)
        
    if len(re_abs_m_s) == 0 or len(ex_speeds) == 0:
        return None, None, None
        
    # Discretize speed to 0.01 step using floor (with floating-point tolerance)
    # to match Excel's approximate match behavior in J25:K425 table
    v_discretized = np.floor(v_thresh * 100 + 1e-9) / 100
    re_abs_at_thresh = np.interp(v_discretized, ex_speeds, re_abs_m_s)
    
    # Calculate relative VO2 demand at threshold (adding B28 = 1.0)
    rel_vo2_demand = (v_thresh * re_abs_at_thresh / weight) + 1.0
    
    delta_vo2 = rel_vo2max - rel_vo2_demand
    if delta_vo2 <= 0 or rel_vo2_demand <= 0:
        return None, re_abs_at_thresh / weight, rel_vo2_demand
        
    ks1, ks2, la_eq, vol_dist = 0.0631, 1.3310, 0.02049, 0.4
    vla_ox_max = (rel_vo2_demand * la_eq) / vol_dist
    adp = np.sqrt((ks1 * rel_vo2_demand) / delta_vo2)
    adp_3 = adp ** 3
    term2 = 1 + (ks2 / adp_3)
    vla_sec = (vla_ox_max * term2) / 60
    return vla_sec, re_abs_at_thresh / weight, rel_vo2_demand

def calculate_laufokonomie_hycys(vo2_abs, vco2_abs, weight, speed_ms):
    """
    Einzelwert-Hilfsfunktion (wird nicht mehr direkt für die HYCYS-Methodik verwendet).
    Berechnet vereinfachte LÖ = VO2ss_rel / v [m/s].
    Für die exakte HYCYS-Methodik bitte calculate_laufokonomie_hycys_stages() verwenden.
    """
    if weight <= 0 or speed_ms <= 0 or vo2_abs <= 0:
        return None
    vo2_rel = vo2_abs / weight
    return vo2_rel / speed_ms


def calculate_laufokonomie_hycys_stages(speeds_arr, vo2_abs_arr, vco2_abs_arr, weight, v_thresh=None):
    """
    Laufökonomie nach exakter HYCYS Excel-Methodik für alle Stufen.

    Entspricht den Formeln in 'Calculation Physiologie', Spalten AC/AF/AG:
      1. AC = VO2_CHO100 / speed [dm/s] / 10
         mit VO2_CHO100 = (VO2 - 3.5*kg) * kal_equiv / 21.1
         und kal_equiv = (RQ - 0.7) * 5.094 + 19.6
      2. Lineare Regression (SLOPE/INTERCEPT) zwischen je 2 aufeinanderfolgenden Stufen
      3. LÖ[i] = (SLOPE * speed_dm_s[i] + INTERCEPT) * 100 / weight

    Gibt eine Liste von LÖ-Werten zurück [ml/min/kg/(m/s)],
    überschwellige Stufen (v > v_thresh) geben None zurück.
    """
    n = len(speeds_arr)
    if n < 2 or weight <= 0:
        return [None] * n

    vo2_rest = 3.5 * weight  # Ruhe-VO2 [ml/min]
    b19 = 21.1               # kcal/L O2 bei RQ=1

    # AC-Werte berechnen (inkl. der ersten Stufe unmittelbar über v_thresh für die SLOPE-Formel wie in HYCYS Excel)
    ac_vals = []
    passed_thresh = False
    for vo2, vco2, v in zip(vo2_abs_arr, vco2_abs_arr, speeds_arr):
        if vo2 <= 0 or v <= 0 or passed_thresh:
            ac_vals.append(None)
            continue
        if v_thresh is not None and v_thresh > 0 and v > v_thresh + 1e-6:
            passed_thresh = True
        vo2_working = max(0.0, vo2 - vo2_rest)
        if vo2_working <= 0:
            ac_vals.append(None)
            continue
        rq = (vco2 / vo2) if vco2 > 0 else 0.85
        cal_eq = (rq - 0.7) * 5.094 + 19.6
        vo2_cho100 = vo2_working * cal_eq / b19
        # AC = VO2_CHO100 / speed_dm_s / 10  (entspricht Excel-Spalte AC)
        speed_dm_s = v * 10
        ac_vals.append(vo2_cho100 / speed_dm_s / 10)

    # Speeds in dm/s
    speeds_dm_s = [v * 10 for v in speeds_arr]

    # LÖ via SLOPE zwischen je zwei aufeinanderfolgenden Stufen;
    # dargestellt und berechnet werden alle unterschwelligen Stufen + die erste überschwellige Stufe (ac_vals[i] is not None)
    lo_vals = []
    for i in range(n):
        if ac_vals[i] is None:
            lo_vals.append(None)
            continue
        if i < n - 1 and ac_vals[i + 1] is not None:
            x0, x1 = speeds_dm_s[i], speeds_dm_s[i + 1]
            y0, y1 = ac_vals[i],      ac_vals[i + 1]
            slope     = (y1 - y0) / (x1 - x0) if (x1 - x0) != 0 else 0
            intercept = y0 - slope * x0
            # F = (slope * v_dm_s + intercept) * 100
            f_val = (slope * speeds_dm_s[i] + intercept) * 100
            lo_vals.append(f_val / weight)
        else:
            # Letzte berücksichtigte Stufe (inkl. erster überschwelliger Stufe): SLOPE mit vorheriger Stufe rückwärts
            if i > 0 and ac_vals[i - 1] is not None:
                x0, x1 = speeds_dm_s[i - 1], speeds_dm_s[i]
                y0, y1 = ac_vals[i - 1],      ac_vals[i]
                slope     = (y1 - y0) / (x1 - x0) if (x1 - x0) != 0 else 0
                intercept = y0 - slope * x0
                f_val = (slope * speeds_dm_s[i] + intercept) * 100
                lo_vals.append(f_val / weight)
            else:
                lo_vals.append(None)

    return lo_vals


def build_training_zones_variante_a(v_lt1, hf_lt1, v_lt2, hf_lt2):
    """
    Erstellt die 6 Trainingsbereiche nach Variante A (Zone 1-5 mit aufgeteiltem Grundlagenbereich Zone 2 unten/oben):
      - Zone 1 - Regeneration:             < 85 % von LTP1
      - Zone 2 - Grundlagenbereich unten:  85 % - 93 % von LTP1
      - Zone 2 - Grundlagenbereich oben:   93 % - 100 % von LTP1
      - Zone 3 - Übergang / Tempo:         100 % von LTP1 bis 97 % von LT2 (OBLA 4.0 bzw. LTP2)
      - Zone 4 - Schwelle:                 97 % - 103 % von LT2 (OBLA 4.0 bzw. LTP2)
      - Zone 5 - VO2max / HIT:             > 103 % von LT2 (OBLA 4.0 bzw. LTP2)
    """
    def _fmt_pace(v):
        if not isinstance(v, (int, float, np.number)) or v <= 0:
            return "-"
        seconds_per_km = 1000.0 / float(v)
        minutes = int(seconds_per_km // 60)
        seconds = int(round(seconds_per_km % 60))
        if seconds == 60:
            minutes += 1
            seconds = 0
        return f"{minutes:02d}:{seconds:02d}"

    def _fmt_hr(hr_val):
        if not isinstance(hr_val, (int, float, np.number)) or hr_val <= 0:
            return "-"
        return str(int(round(float(hr_val))))

    if not (isinstance(v_lt2, (int, float, np.number)) and pd.notna(v_lt2) and v_lt2 > 0):
        return [
            ("Zone 1 - Regeneration",            "-", "-", "-", "-", "-", "-"),
            ("Zone 2 - Grundlagenbereich unten", "-", "-", "-", "-", "-", "-"),
            ("Zone 2 - Grundlagenbereich oben",  "-", "-", "-", "-", "-", "-"),
            ("Zone 3 - Übergang / Tempo",        "-", "-", "-", "-", "-", "-"),
            ("Zone 4 - Schwelle",                "-", "-", "-", "-", "-", "-"),
            ("Zone 5 - VO2max / HIT",            "-", "-", "-", "-", "-", "-"),
        ]

    v2 = float(v_lt2)
    v1 = float(v_lt1) if (isinstance(v_lt1, (int, float, np.number)) and pd.notna(v_lt1) and v_lt1 > 0) else (v2 * 0.79)
    has_hf = isinstance(hf_lt2, (int, float, np.number)) and pd.notna(hf_lt2) and float(hf_lt2) > 0
    h2 = float(hf_lt2) if has_hf else 0.0
    h1 = float(hf_lt1) if (has_hf and isinstance(hf_lt1, (int, float, np.number)) and pd.notna(hf_lt1) and float(hf_lt1) > 0) else (h2 * 0.85 if has_hf else 0.0)

    return [
        ("Zone 1 - Regeneration",
         "-", _fmt_pace(v1 * 0.85), _fmt_pace(v1 * 0.80),
         "-", _fmt_hr(h1 * 0.90), _fmt_hr(h1 * 0.85)),

        ("Zone 2 - Grundlagenbereich unten",
         _fmt_pace(v1 * 0.85), _fmt_pace(v1 * 0.93), _fmt_pace(v1 * 0.89),
         _fmt_hr(h1 * 0.90), _fmt_hr(h1 * 0.95), _fmt_hr(h1 * 0.93)),

        ("Zone 2 - Grundlagenbereich oben",
         _fmt_pace(v1 * 0.93), _fmt_pace(v1), _fmt_pace(v1 * 0.97),
         _fmt_hr(h1 * 0.95), _fmt_hr(h1), _fmt_hr(h1 * 0.98)),

        ("Zone 3 - Übergang / Tempo",
         _fmt_pace(v1), _fmt_pace(v2 * 0.97), _fmt_pace((v1 + v2 * 0.97) / 2.0),
         _fmt_hr(h1), _fmt_hr(h2 * 0.97), _fmt_hr((h1 + h2 * 0.97) / 2.0)),

        ("Zone 4 - Schwelle",
         _fmt_pace(v2 * 0.97), _fmt_pace(v2 * 1.03), _fmt_pace(v2),
         _fmt_hr(h2 * 0.97), _fmt_hr(h2 * 1.03), _fmt_hr(h2)),

        ("Zone 5 - VO2max / HIT",
         _fmt_pace(v2 * 1.03), "max", _fmt_pace(v2 * 1.08),
         _fmt_hr(h2 * 1.03), "max" if has_hf else "-", _fmt_hr(h2 * 1.05)),
    ]


# --- PDF GENERATOR FUNKTION ---
def create_pdf(athlete_name, birthdate, test_date, test_type, weight, body_fat_pct, rel_vo2max, df_res, fig_laktat,
               coach, sportart, kategorie, height, speed, lactate, hr, vo2_steady_values, vo2_steady_abs, vco2_steady_abs, uploaded_file,
               v_ltp1=None, hf_ltp1=None):
    pdf = FPDF(unit='pt')
    pdf.set_auto_page_break(auto=False)

    # Standard colors from the template
    c_teal = (44, 183, 185)       # #2cb7b9
    c_gold = (205, 182, 99)       # #cdb663
    c_dark_grey = (89, 90, 89)    # #595a59
    c_black = (0, 0, 0)
    c_white = (255, 255, 255)

    # Helper function to set fonts & colors and draw text
    def draw_text(x, y, text, size=8.3, font_style='', color=c_black, font_name='Arial'):
        pdf.set_font(font_name, font_style, size)
        pdf.set_text_color(*color)
        pdf.text(x, y + size * 0.85, text.encode('latin-1', 'replace').decode('latin-1'))

    def draw_text_centered(cx, y, text, size=8.3, font_style='', color=c_black):
        pdf.set_font('Arial', font_style, size)
        w = pdf.get_string_width(str(text).encode('latin-1', 'replace').decode('latin-1'))
        draw_text(cx - w / 2.0, y, text, size=size, font_style=font_style, color=color)

    # Helper function to draw colored rectangles
    def draw_rect(x, y, w, h, fill_color=None, stroke_color=None, stroke_width=0.5):
        if fill_color:
            pdf.set_fill_color(*fill_color)
        if stroke_color:
            pdf.set_draw_color(*stroke_color)
            pdf.set_line_width(stroke_width)
        style = 'F' if fill_color and not stroke_color else 'D' if stroke_color and not fill_color else 'FD' if fill_color and stroke_color else ''
        pdf.rect(x, y, w, h, style)

    def page_logo(pg_num):
        """Draw small logo top-right (pages 2-5)."""
        asset = f'pdf_assets/page{pg_num}_img1.png'
        if not os.path.exists(asset):
            asset = 'pdf_assets/page2_img1.png'
        if os.path.exists(asset):
            coords = {
                2: (396.23, 68.88, 109.29, 20.47),
                3: (399.91, 69.42, 108.63, 20.09),
                4: (408.48, 68.88, 109.23, 20.58),
                5: (399.14, 68.88, 108.80, 20.58)
            }
            x, y, w, h = coords.get(pg_num, (396.23, 68.88, 109.29, 20.47))
            pdf.image(asset, x=x, y=y, w=w, h=h)

    def format_pace_helper(v):
        if not isinstance(v, (int, float)) or v <= 0:
            return "-"
        sec = 1000.0 / v
        m = int(sec // 60)
        s = int(round(sec % 60))
        if s == 60:
            m += 1
            s = 0
        return f"{m:02d}:{s:02d} min/km"

    def format_time_helper(t_min):
        h = int(t_min // 60)
        m = int(t_min % 60)
        s = int(round((t_min * 60) % 60))
        if s == 60:
            m += 1
            s = 0
        if m == 60:
            h += 1
            m = 0
        return f"{h:02d}:{m:02d}:{s:02d}"

    temp_files = []

    # --- METABOLIC GRID CALCULATIONS FOR REPORT ---
    vla_val = 0.6
    ans_row = pd.DataFrame()
    try:
        if not df_res.empty:
            ans_row = df_res.iloc[[0]]
            vla_raw = ans_row.iloc[0].get('VLamax [mmol/l/s]', 0.6)
            if pd.notna(vla_raw) and vla_raw != 'N/A' and 'bitte' not in str(vla_raw).lower():
                vla_val = float(vla_raw)
    except:
        pass

    v_ans = 3.5
    if not ans_row.empty:
        try:
            v_ans = float(ans_row.iloc[0].get('v', 3.5))
        except:
            pass
    v_ans_val = v_ans

    has_spiro = (uploaded_file is not None) and (len(vo2_steady_abs) > 0)
    
    # Defaults/Fallbacks
    fatmax_speed = None
    v_grid = None
    lac_clear_grid = None
    lac_prod_grid = None
    fatmax_pace_str = "-"
    fatmax_total_energy = "-"
    fatmax_fat_energy = "-"
    ans_total_energy = "-"
    ans_carb_consumption = "-"
    marathon_paces = {0: "-", 5: "-", 10: "-", 15: "-"}
    marathon_times = {0: "-", 5: "-", 10: "-", 15: "-"}
    marathon_deltas = {5: "-", 10: "-", 15: "-"}

    if has_spiro:
        try:
            glycogen_store = 500.0  # g (HYCYS Standard Eingabedaten!E16)
            carb_intake_rate = 80.0  # g/h (HYCYS Standard Eingabedaten!E15)
            vo2_rest = 3.5 * weight

            # 1. Compute AC for each stage up to first stage above v_ans_val (for linear interpolation at ANS)
            ex_speeds = []
            ac_stages = []
            passed_thresh_pdf = False
            for v, vo2, vco2 in zip(speed, vo2_steady_abs, vco2_steady_abs):
                if v <= 0 or vo2 <= 0:
                    continue
                if passed_thresh_pdf:
                    break
                if v_ans_val > 0 and v > v_ans_val + 1e-6:
                    passed_thresh_pdf = True
                net_vo2 = max(0.0, vo2 - vo2_rest)
                rq = vco2 / vo2 if vco2 > 0 else 0.85
                kal_equiv = (rq - 0.7) * 5.094 + 19.6
                ee_kJ_min = (net_vo2 / 1000.0) * kal_equiv
                vo2_100kh = ee_kJ_min / 21.1 * 1000.0
                ac = vo2_100kh / (v * 100.0)
                ex_speeds.append(float(v))
                ac_stages.append(ac)
            
            if len(ex_speeds) >= 1:
                v_grid_start = min(2.6, float(ex_speeds[0]))
                v_grid = np.arange(v_grid_start, 8.61, 0.01)
                ac_grid = []
                for v in v_grid:
                    if len(ex_speeds) == 1:
                        ac = ac_stages[0]
                    elif v <= ex_speeds[0]:
                        m = (ac_stages[1] - ac_stages[0]) / (ex_speeds[1] - ex_speeds[0])
                        c_val = ac_stages[0] - m * ex_speeds[0]
                        ac = m * v + c_val
                    elif v >= ex_speeds[-1]:
                        # Entspricht HYCYS Excel Calculation Spirometrie J25:K425:
                        # ab der letzten gemessenen Spiro-Stufe wird der letzte AC-Wert gehalten (LOOKUP)
                        ac = ac_stages[-1]
                    else:
                        ac = ac_stages[-1]
                        for i in range(len(ex_speeds) - 1):
                            if ex_speeds[i] <= v < ex_speeds[i+1]:
                                m = (ac_stages[i+1] - ac_stages[i]) / (ex_speeds[i+1] - ex_speeds[i])
                                c_val = ac_stages[i] - m * ex_speeds[i]
                                ac = m * v + c_val
                                break
                    ac_grid.append(ac)

                ac_grid = np.array(ac_grid)
                vo2_rel_demand_grid = (ac_grid * v_grid * 100.0) / weight + 1.0
                delta_vo2_grid = rel_vo2max - vo2_rel_demand_grid

                lac_prod_grid = np.zeros_like(v_grid)
                lac_clear_grid = (0.02049 * vo2_rel_demand_grid) / 0.4
                fat_pct_grid = np.zeros_like(v_grid)
                carb_pct_grid = np.zeros_like(v_grid)
                net_clearance = np.full_like(v_grid, -9999.0)

                ks1, ks2, la_eq, vol_dist = 0.0631, 1.331, 0.02049, 0.4

                for idx, (v, vo2_rel, d_vo2) in enumerate(zip(v_grid, vo2_rel_demand_grid, delta_vo2_grid)):
                    if d_vo2 > 0 and vo2_rel > 0:
                        adp = np.sqrt((ks1 * vo2_rel) / d_vo2)
                        adp_3 = adp ** 3
                        pfk = 1 + (ks2 / adp_3)
                        vla_min = vla_val * 60 / pfk  # mmol/l/min
                        lac_prod_grid[idx] = vla_min
                        
                        fat_pct = 100.0 - (vla_min / lac_clear_grid[idx] * 100.0)
                        fat_pct = max(0.0, min(100.0, fat_pct))
                        net_clearance[idx] = lac_clear_grid[idx] - vla_min
                    else:
                        fat_pct = 0.0
                        net_clearance[idx] = -9999.0
                    
                    fat_pct_grid[idx] = fat_pct
                    carb_pct_grid[idx] = 100.0 - fat_pct

                max_clearance_idx = np.argmax(net_clearance)
                fatmax_speed = v_grid[max_clearance_idx]
                fatmax_pace_str = format_pace_helper(fatmax_speed)

                rq_grid = 0.7 + 0.3 * (carb_pct_grid / 100.0)
                kal_equiv_grid = (rq_grid - 0.7) * 5.094 + 19.6
                ee_kJ_min_grid = (vo2_rel_demand_grid * weight / 1000.0) * kal_equiv_grid
                ee_kJ_h_grid = ee_kJ_min_grid * 60
                ee_kcal_h_grid = ee_kJ_h_grid / 4.186
                ee_fat_kcal_h_grid = ee_kcal_h_grid * (fat_pct_grid / 100.0)
                ee_kh_kcal_h_grid = ee_kcal_h_grid * (carb_pct_grid / 100.0)
                carb_g_h_grid = ee_kh_kcal_h_grid / 4.063

                fatmax_idx = max_clearance_idx
                fatmax_total_energy = f"{int(round(ee_kcal_h_grid[fatmax_idx]))}"
                fatmax_fat_energy = f"{int(round(ee_fat_kcal_h_grid[fatmax_idx]))}"

                ans_speed_lookup = np.floor(v_ans * 100 + 1e-9) / 100.0
                ans_idx = np.argmin(np.abs(v_grid - ans_speed_lookup))
                ans_total_energy = f"{int(round(ee_kcal_h_grid[ans_idx]))}"
                ans_carb_consumption = f"{int(round(carb_g_h_grid[ans_idx]))}"

                def calc_marathon_speed_0():
                    for idx, v in enumerate(v_grid):
                        t_mara_min = 42195.0 / (v * 60.0)
                        carb_req_g = carb_g_h_grid[idx] / 60.0 * t_mara_min
                        carb_avail_g = (carb_intake_rate / 60.0 * t_mara_min) + glycogen_store
                        deficit = carb_req_g - carb_avail_g
                        if deficit > 0:
                            return float(v_grid[max(0, idx - 1)] - 0.01)
                    return float(v_grid[-1])

                v_mara_0 = calc_marathon_speed_0()
                t_min_0 = 42195.0 / (v_mara_0 * 60.0)
                t_min_5 = t_min_0 * 0.95
                t_min_10 = t_min_0 * 0.90
                t_min_15 = t_min_0 * 0.85
                v_mara_5 = 42195.0 / (t_min_5 * 60.0)
                v_mara_10 = 42195.0 / (t_min_10 * 60.0)
                v_mara_15 = 42195.0 / (t_min_15 * 60.0)

                for pct_key, v_m, t_min in [(0, v_mara_0, t_min_0), (5, v_mara_5, t_min_5), (10, v_mara_10, t_min_10), (15, v_mara_15, t_min_15)]:
                    h = int(t_min // 60)
                    m = int(t_min % 60)
                    s = int(round((t_min * 60) % 60))
                    if s == 60:
                        m += 1
                        s = 0
                    if m == 60:
                        h += 1
                        m = 0
                    marathon_paces[pct_key] = format_pace_helper(v_m)
                    marathon_times[pct_key] = f"{h:02d}:{m:02d}:{s:02d} h"

                for pct_key, t_min_p in [(5, t_min_5), (10, t_min_10), (15, t_min_15)]:
                    delta_min = t_min_0 - t_min_p
                    dm = int(delta_min)
                    ds = int(round((delta_min - dm) * 60.0))
                    if ds == 60:
                        dm += 1
                        ds = 0
                    marathon_deltas[pct_key] = f"- {dm:02d}:{ds:02d} min"
        except Exception as e:
            pass

    # --- SEITE 1: DECKBLATT ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    
    if os.path.exists('pdf_assets/page1_img3.png'):
        pdf.image('pdf_assets/page1_img3.png', x=86.8, y=194.4, w=400, h=124)
        
    draw_rect(50.9, 460.4, 479.1, 24.5, fill_color=c_gold)
    draw_text(111.7, 464.7, f"HYCYS {test_type}", size=13.8, font_style='B', color=c_white)

    labels_s1 = ["Name", "Geburtsdatum", "Coach", "Testdatum", "Sportart", "Kategorie"]
    vals_s1 = [athlete_name, birthdate, coach, test_date, sportart, kategorie]
    y_coords_s1 = [490.7, 502.7, 542.6, 555.9, 582.6, 595.9]
    for lbl, val, y_coord in zip(labels_s1, vals_s1, y_coords_s1):
        draw_text(111.0, y_coord, lbl, size=8.3, color=c_black)
        draw_text(214.5, y_coord, val, size=8.3, font_style='B', color=c_black)

    if os.path.exists('pdf_assets/page1_img1.png'):
        pdf.image('pdf_assets/page1_img1.png', x=114.4, y=681.4, w=15.6, h=14.4)
    draw_text(162.6, 685.8, "www.hycys.de", size=7.6, font_style='B', color=c_teal)
    
    if os.path.exists('pdf_assets/page1_img2.png'):
        pdf.image('pdf_assets/page1_img2.png', x=137.3, y=680.6, w=15.6, h=16.1)
    draw_text(266.1, 685.8, "contact@hycys.de", size=7.6, font_style='B', color=c_teal)

    # --- SEITE 2: ANTHROPOMETRIE & LAKTATKINETIK ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 471.6, 663.6, fill_color=c_white)
    
    page_logo(2)

    draw_rect(57.6, 133.9, 171.9, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 171.9 / 2.0, 135.0, "Anthropologie", size=8.2, font_style='B', color=c_white)

    draw_text(59.3, 173.2, "Gewicht", size=8.3, font_style='B', color=c_black)
    draw_text(199.9, 173.2, f"{weight:.1f} kg".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    draw_text(59.3, 183.8, "Größe", size=8.3, font_style='B', color=c_black)
    draw_text(199.9, 183.8, f"{int(height)}cm", size=8.3, font_style='B', color=c_dark_grey)
    
    bmi = weight / ((height/100.0)**2) if height > 0 else 0.0
    draw_text(59.3, 209.7, "Body Mass Index", size=8.3, font_style='B', color=c_black)
    draw_text(187.5, 209.7, f"{bmi:.1f} kg/m²".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    
    draw_rect(113.18, 245.58, 112.94, 10.44, fill_color=c_gold)
    draw_text_centered(141.42, 246.8, "Masse", size=8.2, font_style='B', color=c_white)
    draw_text_centered(197.89, 246.8, "%",     size=8.2, font_style='B', color=c_white)
    
    fett_kg = weight * (body_fat_pct / 100.0)
    fettfrei_kg = weight - fett_kg
    draw_text(59.3, 259.6, "Fett", size=8.3, font_style='B', color=c_black)
    draw_text_centered(141.42, 259.6, f"{fett_kg:.1f} kg".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(197.89, 259.6, f"{body_fat_pct:.1f} %".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    
    draw_text(59.3, 270.1, "Fettfrei", size=8.3, font_style='B', color=c_black)
    draw_text_centered(141.42, 270.1, f"{fettfrei_kg:.1f} kg".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(197.89, 270.1, f"{100.0 - body_fat_pct:.1f} %".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)

    # Exploded Pie Chart matching Cycling Report
    fig, ax = plt.subplots(figsize=(3.856, 2.0), dpi=300)
    pct_fett = max(0.1, body_fat_pct)
    pct_frei = max(0.1, 100.0 - body_fat_pct)
    wedges, texts, autotexts = ax.pie(
        [pct_fett, pct_frei],
        explode=(0.12, 0),
        colors=['#2cb7b9', '#cdb663'],
        autopct=lambda p: f"{p:.1f} %".replace('.', ','),
        startangle=90,
        counterclock=False,
        pctdistance=0.50,
        textprops=dict(size=8, weight='bold', color='white'),
        wedgeprops=dict(edgecolor='white', linewidth=2),
    )
    
    for i, (wedge, autotext) in enumerate(zip(wedges, autotexts)):
        theta = np.deg2rad(wedge.theta1 + (wedge.theta2 - wedge.theta1) / 2.0)
        e = 0.12 if i == 0 else 0.0
        d = 0.65 if i == 0 else 0.50
        x = (e + d) * np.cos(theta)
        y = (e + d) * np.sin(theta)
        autotext.set_position((x, y))
        autotext.set_color('white')

    ax.axis('equal')
    from matplotlib.patches import Patch
    handles = [Patch(color='#2cb7b9', label='Fett'), Patch(color='#cdb663', label='Fettfrei')]
    legend = ax.legend(handles=handles, loc='center left', bbox_to_anchor=(1.05, 0.5), fontsize=8, frameon=False)
    for text in legend.get_texts():
        text.set_weight('bold')
        text.set_color('#595a59')
    fig.patch.set_facecolor('white')
    plt.tight_layout(pad=0.3)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_pie:
        fig.savefig(tmp_pie.name, format="png", bbox_inches="tight", facecolor='white')
        plt.close(fig)
        pdf.image(tmp_pie.name, x=268.05, y=134.92, w=227.56, h=118.01)
        temp_files.append(tmp_pie.name)

    # VLamax Section
    draw_rect(57.6, 293.2, 225.8, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 225.8 / 2.0, 294.3, "Anaerober Stoffwechsel - max. Laktatbildungsrate", size=8.2, font_style='B', color=c_white)
    
    draw_rect(114.02, 314.9, 112.94, 10.44, fill_color=c_gold)
    draw_text_centered(114.02 + 112.94 / 2.0, 316.1, "VLamax", size=8.2, font_style='B', color=c_white)
    
    vla_ans_str = f"{vla_val:.2f} mmol/l/s".replace('.', ',') if vla_val > 0 else "-"
    draw_text_centered(114.02 + 112.94 / 2.0, 314.9 + 10.44 + 0.8, vla_ans_str, size=8.2, font_style='B', color=c_dark_grey)

    # VO2max Section
    draw_rect(57.6, 355.0, 343.6, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 343.6 / 2.0, 356.1, "Aerober Stoffwechsel - max. Sauerstoffaufnahme", size=8.2, font_style='B', color=c_white)
    
    vo2_abs = rel_vo2max * weight
    vo2_abs_str = f"{int(vo2_abs):,}".replace(",", ".") + " ml/min" if vo2_abs > 0 else "-"
    vo2_rel_str = f"{rel_vo2max:.1f} ml/min/kg".replace('.', ',') if rel_vo2max > 0 else "-"
    
    ffm_pct = 100.0 - body_fat_pct
    ffm_kg = weight * (ffm_pct / 100.0)
    vo2_ffm = (rel_vo2max * weight) / ffm_kg if ffm_kg > 0 else 0.0
    vo2_ffm_str = f"{vo2_ffm:.1f} ml/min/kg".replace('.', ',') if vo2_ffm > 0 else "-"

    draw_rect(57.6, 375.0, 143.6, 10.44, fill_color=c_gold)
    draw_text_centered(57.6 + 143.6 / 2.0, 376.2, "VO2max abs.", size=8.2, font_style='B', color=c_white)
    draw_text_centered(57.6 + 143.6 / 2.0, 375.0 + 10.44 + 0.8, vo2_abs_str, size=8.2, font_style='B', color=c_dark_grey)
    
    draw_rect(201.2, 375.0, 143.6, 10.44, fill_color=c_gold)
    draw_text_centered(201.2 + 143.6 / 2.0, 376.2, "VO2max rel.", size=8.2, font_style='B', color=c_white)
    draw_text_centered(201.2 + 143.6 / 2.0, 375.0 + 10.44 + 0.8, vo2_rel_str, size=8.2, font_style='B', color=c_dark_grey)
    
    draw_rect(344.8, 375.0, 143.4, 10.44, fill_color=c_gold)
    draw_text_centered(344.8 + 143.4 / 2.0, 376.2, "VO2max rel. FFM", size=8.2, font_style='B', color=c_white)
    draw_text_centered(344.8 + 143.4 / 2.0, 375.0 + 10.44 + 0.8, vo2_ffm_str, size=8.2, font_style='B', color=c_dark_grey)

    # Lactate kinetics plot
    draw_rect(57.6, 415.0, 458.1, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 465.1 / 2.0, 416.1, "Interaktion Laktatauf- & abbau", size=8.2, font_style='B', color=c_white)
    
    import math
    import matplotlib.ticker as ticker
    fig, ax = plt.subplots(figsize=(6.0, 3.2), dpi=300)
    stage_speeds_pdf = [float(v) for v in speed if pd.notna(v) and float(v) > 0]
    if len(stage_speeds_pdf) >= 2:
        _v_step_pdf = round(stage_speeds_pdf[-1] - stage_speeds_pdf[-2], 3)
        if _v_step_pdf <= 0:
            _v_step_pdf = 0.4
        stage_speeds_pdf = stage_speeds_pdf + [round(stage_speeds_pdf[-1] + _v_step_pdf, 2)]
    elif len(stage_speeds_pdf) == 1:
        stage_speeds_pdf = stage_speeds_pdf + [round(stage_speeds_pdf[0] + 0.4, 2)]
    v_start = float(stage_speeds_pdf[0]) if len(stage_speeds_pdf) > 0 else 2.5
    v_end = float(stage_speeds_pdf[-1]) if len(stage_speeds_pdf) > 0 else 4.5
    v_plot_end = v_end
    v_arr = np.linspace(v_start, v_plot_end, 100)

    ax2 = ax.twinx()
    bbox_props = dict(boxstyle='square,pad=0.22', facecolor='white', edgecolor='none', alpha=0.92)

    if v_grid is not None and lac_clear_grid is not None and lac_prod_grid is not None and 'vo2_rel_demand_grid' in locals():
        mask_lac = (v_grid >= v_start - 1e-6) & (v_grid <= v_plot_end + 1e-6)
        v_lac_plot = v_grid[mask_lac]
        vo2_abs_plot = vo2_rel_demand_grid[mask_lac] * weight
        lac_clear_plot = lac_clear_grid[mask_lac]
        lac_prod_raw = lac_prod_grid[mask_lac]
        # Wo delta_vo2 <= 0 ist, Laktatproduktion auf > y2_max setzen, damit die goldene Kurve wie in Excel nach oben ausläuft
        lac_prod_plot = np.where(delta_vo2_grid[mask_lac] > 0, lac_prod_raw, 15.0)

        ax.plot(v_lac_plot, vo2_abs_plot, color='#595a59', linewidth=2.5)
        ax2.plot(v_lac_plot, lac_clear_plot, color='#2cb7b9', linewidth=2.5)
        ax2.plot(v_lac_plot, lac_prod_plot, color='#cdb663', linewidth=2.5, linestyle='-')

        max_vo2_plot = float(np.max(vo2_abs_plot)) if len(vo2_abs_plot) > 0 else 4500.0
        y1_lac_max = max(5500.0, math.ceil(max_vo2_plot * 1.20 / 500.0) * 500.0)
        ax.set_ylim(0, y1_lac_max)
        ax.set_yticks(np.arange(0, y1_lac_max + 1, 500))

        # HYCYS Excel Report Running skaliert die rechte Achse (Laktatkinetik) von 0,0 bis 5,0 mmol/L/min (Schrittweite 0,5)
        y2_lac_max = 5.0
        ax2.set_ylim(0, y2_lac_max)
        ax2.set_yticks(np.arange(0.0, y2_lac_max + 0.01, 0.5))

        if len(v_lac_plot) >= 2:
            x_span = max(v_lac_plot[-1] - v_lac_plot[0], 1e-6)
            # 1. Label 'Sauerstoffaufnahme - VO2' wie im Original-HYCYS-PDF links über der grauen VO2-Kurve
            v_lbl_vo2 = v_lac_plot[0] + 0.25 * x_span
            y_lbl_vo2 = float(np.interp(v_lbl_vo2, v_lac_plot, vo2_abs_plot)) + 0.12 * y1_lac_max
            ax.text(v_lbl_vo2, min(y_lbl_vo2, 0.88 * y1_lac_max), 'Sauerstoffaufnahme - VO2',
                    color='#595a59', fontsize=7.5, fontweight='bold', ha='center', va='center', bbox=bbox_props, zorder=10)

            # 2. Label 'Laktatabbau' über der türkisen Laktatabbau-Kurve bei ~35 % der X-Achse
            v_lbl_clr = v_lac_plot[0] + 0.35 * x_span
            y_clr_at_v = float(np.interp(v_lbl_clr, v_lac_plot, lac_clear_plot))
            ax2.text(v_lbl_clr, min(0.85 * y2_lac_max, y_clr_at_v + 0.11 * y2_lac_max), 'Laktatabbau',
                     color='#2cb7b9', fontsize=7.5, fontweight='bold', ha='center', va='center', bbox=bbox_props, zorder=10)

            # 3. Label 'Laktatproduktion' links vom steilen Ast der goldenen Kurve über der Kurve (wie im Original-HYCYS-PDF)
            valid_prd_mask = lac_prod_plot <= y2_lac_max
            if np.any(valid_prd_mask):
                v_prd_sub = v_lac_plot[valid_prd_mask]
                y_prd_sub = lac_prod_plot[valid_prd_mask]
                idx_prd = int(np.argmin(np.abs(y_prd_sub - 1.50)))
                v_lbl_prd = min(v_lac_plot[-1] - 0.14 * x_span, float(v_prd_sub[idx_prd]) + 0.08 * x_span)
            else:
                v_lbl_prd = v_lac_plot[0] + 0.70 * x_span
            ax2.text(v_lbl_prd, 1.15, 'Laktatproduktion',
                     color='#cdb663', fontsize=7.5, fontweight='bold', ha='center', va='center', bbox=bbox_props, zorder=10)
    else:
        if uploaded_file is not None and len(vo2_steady_abs) > 0:
            ax.plot(speed[:len(vo2_steady_abs)], vo2_steady_abs, color='#595a59', linewidth=2.5)
        ax.set_ylim(0, 5500)
        ax.set_yticks(np.arange(0, 5501, 500))
        poly_coeffs = np.polyfit(speed, lactate, 3)
        poly_func = np.poly1d(poly_coeffs)
        ax2.plot(v_arr, poly_func(v_arr), color='#2cb7b9', linewidth=2.5)
        lac_prod = lactate[0] + (lactate[-1] - lactate[0]) * ((v_arr - v_start) / max(v_end - v_start, 1e-6)) ** 3.5
        ax2.plot(v_arr, lac_prod, color='#cdb663', linewidth=2.5, linestyle='-')
        ax2.set_ylim(0, 5.0)
        ax2.set_yticks(np.arange(0.0, 5.01, 0.5))
        ax.text(0.26, 0.56, 'Sauerstoffaufnahme - VO2', color='#595a59', fontsize=7.5, fontweight='bold', ha='center', va='center', transform=ax.transAxes, bbox=bbox_props, zorder=10)
        ax2.text(0.36, 0.40, 'Laktatabbau', color='#2cb7b9', fontsize=7.5, fontweight='bold', ha='center', va='center', transform=ax2.transAxes, bbox=bbox_props, zorder=10)
        ax2.text(0.75, 0.29, 'Laktatproduktion', color='#cdb663', fontsize=7.5, fontweight='bold', ha='center', va='center', transform=ax2.transAxes, bbox=bbox_props, zorder=10)

    ax.set_ylabel(r'$\mathbf{VO_2}$ [ml/min]', color='#595a59', fontsize=7, fontweight='bold')
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f"{int(round(x)):,}".replace(",", ".")))
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6, width=1.2)
    ax.grid(True, axis='y', linestyle='-', linewidth=0.8, color='#d8d8d8')
    ax.set_axisbelow(True)

    ax2.set_ylabel('Laktatkinetik [mmol/L/min]', color='#2cb7b9', fontsize=7, fontweight='bold')
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f"{x:.1f}".replace(".", ",")))
    ax2.tick_params(axis='y', labelcolor='#2cb7b9', labelsize=6, width=1.2, colors='#595a59')
    for lbl_y2 in ax2.get_yticklabels():
        lbl_y2.set_color('#2cb7b9')
    
    ax.set_xlabel('Geschwindigkeit [m/s]', color='#595a59', fontsize=7, fontweight='bold')
    ax.set_xlim(v_start, v_plot_end)
    if len(stage_speeds_pdf) > 0:
        ax.set_xticks(stage_speeds_pdf)
        _all_1dec = all(abs(round(v, 1) - round(v, 2)) < 0.015 for v in stage_speeds_pdf)
        ax.set_xticklabels([f"{v:.1f}".replace(".", ",") if _all_1dec else f"{v:.2f}".replace(".", ",") for v in stage_speeds_pdf])
    ax.tick_params(axis='x', labelcolor='#595a59', labelsize=6, width=1.2)
    
    for s in ax.spines.values():
        s.set_color('#595a59')
        s.set_linewidth(1.2)
    for s in ax2.spines.values():
        s.set_color('#595a59')
        s.set_linewidth(1.2)

    for label in ax.get_xticklabels(): label.set_weight('bold')
    for label in ax.get_yticklabels(): label.set_weight('bold')
    for label in ax2.get_yticklabels(): label.set_weight('bold')
    
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_lactate:
        fig.savefig(tmp_lactate.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_lactate.name, x=55.2, y=432.0, w=455, h=248)
        temp_files.append(tmp_lactate.name)

    # --- SEITE 3: LAUFÖKONOMIE ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    
    page_logo(3)
        
    draw_rect(57.6, 133.9, 207.0, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 207.0 / 2.0, 135.0, "Laufökonomie", size=8.2, font_style='B', color=c_white)

    # Table Header
    draw_rect(57.6, 159.1, 207.0, 10.44, fill_color=c_gold)
    draw_text_centered(57.6 + 103.5 / 2.0, 160.3, "Geschwindigkeit", size=8.2, font_style='B', color=c_white)
    draw_text_centered(57.6 + 103.5 + 103.5 / 2.0, 160.3, "Laufökonomie", size=8.2, font_style='B', color=c_white)

    # Populate running economy values in original HYCYS PDF format (ml/min/kg*m/s, nur unterschwellige Stufen)
    speeds_s3 = []
    eco_vals_s3 = []
    valid_lo_s3 = []
    valid_speeds_s3 = []
    y_coords_s3_default = [173.2, 185.8, 198.4, 211.0, 223.6]
    
    n_pdf = len(speed)
    lo_pdf_list = calculate_laufokonomie_hycys_stages(
        speeds_arr   = list(speed[:n_pdf]),
        vo2_abs_arr  = list(vo2_steady_abs[:n_pdf]) if len(vo2_steady_abs) >= n_pdf else [0.0]*n_pdf,
        vco2_abs_arr = list(vco2_steady_abs[:n_pdf]) if len(vco2_steady_abs) >= n_pdf else [0.0]*n_pdf,
        weight       = weight,
        v_thresh     = v_ans_val if v_ans_val > 0 else None,
    ) if len(vo2_steady_abs) > 0 else [None]*n_pdf

    _s3_all_1dec = all(abs(round(float(v), 1) - round(float(v), 2)) < 0.015 for v in speed if pd.notna(v))
    for i in range(n_pdf):
        v_ms = float(speed[i])
        lo_v = lo_pdf_list[i] if i < len(lo_pdf_list) else None
        v_fmt = f"{v_ms:.1f} m/s".replace('.', ',') if _s3_all_1dec else f"{v_ms:.2f} m/s".replace('.', ',')
        if lo_v is not None:
            eco_str = f"{lo_v:.1f} ml/min/kg*m/s".replace('.', ',')
            speeds_s3.append(v_fmt)
            eco_vals_s3.append(eco_str)
            valid_lo_s3.append(float(lo_v))
            valid_speeds_s3.append(f"{v_ms:.1f}".replace('.', ',') if _s3_all_1dec else f"{v_ms:.2f}".replace('.', ','))
        elif len(vo2_steady_abs) == 0:
            speeds_s3.append(v_fmt)
            eco_vals_s3.append("-")
        
    y_coords_s3 = y_coords_s3_default if len(speeds_s3) <= 5 else list(np.linspace(172.0, 226.5, len(speeds_s3)))
    font_s3 = 8.3 if len(speeds_s3) <= 5 else 7.2
    for i in range(len(speeds_s3)):
        y_coord = y_coords_s3[i]
        draw_text_centered(57.6 + 103.5 / 2.0, y_coord, speeds_s3[i], size=font_s3, font_style='B', color=c_dark_grey)
        draw_text_centered(57.6 + 103.5 + 103.5 / 2.0, y_coord, eco_vals_s3[i], size=font_s3, font_style='B', color=c_dark_grey)

    # Chart 3a: Running Economy bar chart in original HYCYS PDF format (Y-axis: Laufökonomie 5..15)
    fig, ax = plt.subplots(figsize=(3.2, 1.85), dpi=300)
    if len(valid_lo_s3) > 0:
        x_pos_s3 = np.arange(len(valid_lo_s3))
        ax.bar(x_pos_s3, valid_lo_s3, width=0.68, color='#2cb7b9', zorder=3)
        ax.set_xticks(x_pos_s3)
        ax.set_xticklabels(valid_speeds_s3)
        ax.set_xlim(-0.55, max(len(valid_lo_s3) - 0.45, 0.55))
    ax.set_ylim(5, 15)
    ax.set_yticks([5, 7, 9, 11, 13, 15])
    ax.set_ylabel('Laufökonomie', color='#595a59', fontsize=6.5, fontweight='bold')
    ax.set_xlabel('Geschwindigkeit [m/s]', color='#595a59', fontsize=6.5, fontweight='bold')
    ax.tick_params(axis='both', labelcolor='#595a59', labelsize=5.5, width=1.2)
    for label in ax.get_xticklabels(): label.set_weight('bold')
    for label in ax.get_yticklabels(): label.set_weight('bold')
    ax.grid(True, axis='y', linestyle='-', linewidth=0.8, color='#d8d8d8')
    ax.set_axisbelow(True)
    for s in ax.spines.values():
        s.set_color('#595a59')
        s.set_linewidth(1.2)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_eco_chart:
        fig.savefig(tmp_eco_chart.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_eco_chart.name, x=266.8, y=156.2, w=257, h=135.6)
        temp_files.append(tmp_eco_chart.name)

    # Marathon predictions table populating calculated values (0%, 5%, 10%, 15% wie im Original-HYCYS-PDF)
    draw_rect(57.6, 334.1, 414.0, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 414.0 / 2.0, 335.2, "Einfluss Laufökonomie auf Marathon-Zeit", size=8.2, font_style='B', color=c_white)

    draw_rect(57.6, 355.1, 414.0, 10.44, fill_color=c_gold)
    draw_text_centered(109.35, 356.2, "Verbesserung", size=8.2, font_style='B', color=c_white)
    draw_text_centered(212.85, 356.2, "Marathon-Pace", size=8.2, font_style='B', color=c_white)
    draw_text_centered(316.35, 356.2, "Zielzeit", size=8.2, font_style='B', color=c_white)
    draw_text_centered(419.85, 356.2, "Delta zu 0%", size=8.2, font_style='B', color=c_white)

    y_coords_s3_m = [373.0, 385.6, 398.2, 410.8]
    improves = ["0%", "5%", "10%", "15%"]
    for imp, y_coord in zip(improves, y_coords_s3_m):
        pct_key = int(imp.replace('%', ''))
        pace_val = marathon_paces[pct_key]
        time_val = marathon_times[pct_key]
        delta_val = marathon_deltas[pct_key] if pct_key != 0 else "-"
        
        draw_text_centered(109.35, y_coord, imp, size=8.3, font_style='B', color=c_dark_grey)
        draw_text_centered(212.85, y_coord, pace_val, size=8.3, font_style='B', color=c_dark_grey)
        draw_text_centered(316.35, y_coord, time_val, size=8.3, font_style='B', color=c_dark_grey)
        if pct_key != 0:
            draw_text_centered(419.85, y_coord, delta_val, size=8.3, font_style='B', color=c_dark_grey)

    # Marathon prediction horizontal bar chart (1..4 wie im Original-HYCYS-PDF)
    fig, ax = plt.subplots(figsize=(5.2, 2.3), dpi=300)
    
    if has_spiro and fatmax_speed is not None:
        y_vals = [1, 2, 3, 4]
        t_0_h = 42195.0 / (v_mara_0 * 3600.0)
        t_5_h = 42195.0 / (v_mara_5 * 3600.0)
        t_10_h = 42195.0 / (v_mara_10 * 3600.0)
        t_15_h = 42195.0 / (v_mara_15 * 3600.0)
        x_vals = [t_0_h, t_5_h, t_10_h, t_15_h]
        
        step_10m = 10.0 / 60.0
        x_min = math.floor((min(x_vals) - 0.15) / step_10m) * step_10m
        if 2.0 <= min(x_vals) <= 3.15:
            x_min = 2.0
        x_max = math.ceil((max(x_vals) + 0.08) / step_10m) * step_10m
        
        widths = [max(0.02, x - x_min) for x in x_vals]
        ax.barh(y_vals, widths, height=0.66, left=x_min, color='#2cb7b9', zorder=3)
        
        ax.set_ylim(0.45, 4.55)
        ax.set_yticks([1, 2, 3, 4])
        ax.set_yticklabels(['1', '2', '3', '4'])
        
        ax.set_xlim(x_min, x_max)
        x_ticks_mara = np.arange(x_min, x_max + 0.01, step_10m)
        ax.set_xticks(x_ticks_mara)
        
        def format_hours_hm(x, pos):
            total_min = int(round(x * 60.0))
            h = total_min // 60
            m = total_min % 60
            return f"{h}:{m:02d}"
        ax.xaxis.set_major_formatter(ticker.FuncFormatter(format_hours_hm))
    else:
        ax.set_xlim(2.0, 3.1667)
        ax.set_ylim(0.4, 4.6)
        ax.set_yticks([1, 2, 3, 4])
        ax.set_yticklabels(['1', '2', '3', '4'])

    for label in ax.get_xticklabels(): label.set_weight('bold'); label.set_size(5.5)
    for label in ax.get_yticklabels(): label.set_weight('bold'); label.set_size(5.5)
    ax.set_xlabel('Marathon-Zielzeit [h:mm]', color='#595a59', fontsize=6.5, fontweight='bold')
    ax.set_ylabel('Verbesserung Laufökonomie [%]', color='#595a59', fontsize=6.5, fontweight='bold')
    ax.tick_params(axis='both', labelcolor='#595a59', width=1.2)
    for s in ax.spines.values(): s.set_color('#595a59'); s.set_linewidth(1.2)
    ax.grid(True, axis='x', linestyle='-', linewidth=0.8, color='#d8d8d8')
    ax.set_axisbelow(True)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_mara_chart:
        fig.savefig(tmp_mara_chart.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_mara_chart.name, x=66.1, y=475.1, w=449.2, h=209.4)
        temp_files.append(tmp_mara_chart.name)

    # --- SEITE 4: ANAEROBE SCHWELLE & ENERGETIK ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    
    page_logo(4)
        
    draw_rect(57.6, 133.9, 465.7, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 465.7 / 2.0, 135.0, "Anaerobe Schwelle (ANS) - Geschwindigkeit & Herzfrequenz", size=8.2, font_style='B', color=c_white)

    # Box 1
    draw_rect(57.6, 159.1, 207.0, 10.44, fill_color=c_gold)
    draw_text_centered(57.6 + 207.0 / 2.0, 160.3, "Laufgeschwindigkeit", size=8.2, font_style='B', color=c_white)
    
    draw_rect(368.0, 159.1, 103.6, 10.44, fill_color=c_gold)
    draw_text_centered(368.0 + 103.6 / 2.0, 160.3, "Herzfrequenz", size=8.2, font_style='B', color=c_white)

    ans_ms_str, ans_pace_str, ans_hf_str = "-", "-", "-"
    v_ans_val = 0.0
    if not ans_row.empty:
        raw_v = ans_row.iloc[0].get('v', 0.0)
        try:
            v_ans_val = float(raw_v)
        except (ValueError, TypeError):
            v_ans_val = 0.0
        if v_ans_val > 0:
            ans_ms_str = f"{v_ans_val:.2f} m/s".replace('.', ',')
            ans_pace_str = format_pace_helper(v_ans_val)
            try:
                raw_hf = ans_row.iloc[0].get('HF', 0)
                ans_hf_str = f"{int(float(raw_hf))} bpm"
            except (ValueError, TypeError):
                ans_hf_str = "-"

    draw_text_centered(109.35, 170.34, ans_ms_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(212.85, 170.34, ans_pace_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(368.0 + 103.6 / 2.0, 170.34, ans_hf_str, size=8.3, font_style='B', color=c_dark_grey)
 
    # Box 2: VO2
    draw_rect(57.6, 206.9, 465.7, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 465.7 / 2.0, 208.0, "Anaerobe Schwelle (ANS) - VO2", size=8.2, font_style='B', color=c_white)
 
    draw_rect(57.6, 234.9, 103.6, 10.44, fill_color=c_gold)
    draw_text_centered(57.6 + 103.6 / 2.0, 236.1, "VO2 abs.", size=8.2, font_style='B', color=c_white)
 
    draw_rect(212.8, 234.9, 103.6, 10.44, fill_color=c_gold)
    draw_text_centered(212.8 + 103.6 / 2.0, 236.1, "VO2 rel.", size=8.2, font_style='B', color=c_white)
 
    draw_rect(368.0, 234.9, 103.6, 10.44, fill_color=c_gold)
    draw_text_centered(368.0 + 103.6 / 2.0, 236.1, "% VO2max", size=8.2, font_style='B', color=c_white)
 
    ans_vo2_abs_str, ans_vo2_rel_str, ans_vo2_pct_str = "-", "-", "-"
    if uploaded_file is not None and v_ans_val > 0 and len(vo2_steady_values) > 0:
        if v_grid is not None and 'vo2_rel_demand_grid' in locals():
            ans_v_lkp = np.floor(v_ans_val * 100.0 + 1e-9) / 100.0
            ans_vo2_rel = float(vo2_rel_demand_grid[int(np.argmin(np.abs(v_grid - ans_v_lkp)))])
        else:
            ans_vo2_rel = float(np.interp(v_ans_val, speed, vo2_steady_values))
        ans_vo2_abs = ans_vo2_rel * weight
        ans_vo2_pct = (ans_vo2_rel / rel_vo2max * 100.0) if rel_vo2max > 0 else 0.0
        
        ans_vo2_abs_str = f"{int(round(ans_vo2_abs)):,}".replace(",", ".") + " ml/min"
        ans_vo2_rel_str = f"{ans_vo2_rel:.1f} ml/min/kg".replace('.', ',')
        ans_vo2_pct_str = f"{int(round(ans_vo2_pct))} %"
 
    draw_text_centered(57.6 + 103.6 / 2.0, 246.14, ans_vo2_abs_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(212.8 + 103.6 / 2.0, 246.14, ans_vo2_rel_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(368.0 + 103.6 / 2.0, 246.14, ans_vo2_pct_str, size=8.3, font_style='B', color=c_dark_grey)
 
    # Energetik Box with calculated total energy & KH-Verbrauch
    draw_rect(57.6, 279.9, 465.7, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 465.7 / 2.0, 281.0, "Anaerobe Schwelle (ANS) - Energetik", size=8.2, font_style='B', color=c_white)
 
    draw_rect(57.6, 307.8, 103.6, 10.44, fill_color=c_gold)
    draw_text_centered(57.6 + 103.6 / 2.0, 309.0, "Energieverbrauch", size=8.2, font_style='B', color=c_white)
 
    draw_rect(212.8, 307.8, 103.6, 10.44, fill_color=c_gold)
    draw_text_centered(212.8 + 103.6 / 2.0, 309.0, "KH-Verbrauch", size=8.2, font_style='B', color=c_white)
 
    ans_total_energy_str = f"{ans_total_energy} kcal/h" if ans_total_energy != "-" else "-"
    ans_carb_consumption_str = f"{ans_carb_consumption} g/h" if ans_carb_consumption != "-" else "-"
    draw_text_centered(57.6 + 103.6 / 2.0, 319.04, ans_total_energy_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(212.8 + 103.6 / 2.0, 319.04, ans_carb_consumption_str, size=8.3, font_style='B', color=c_dark_grey)
 
    # Fatmax Box with calculated values
    draw_rect(57.6, 355.1, 465.7, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 465.7 / 2.0, 356.2, "Fettstoffwechsel - Fatmax", size=8.2, font_style='B', color=c_white)
 
    draw_rect(57.6, 383.1, 207.0, 10.44, fill_color=c_gold)
    draw_text_centered(57.6 + 207.0 / 2.0, 384.3, "Laufgeschwindigkeit", size=8.2, font_style='B', color=c_white)
 
    draw_rect(316.3, 383.1, 207.0, 10.44, fill_color=c_gold)
    draw_text_centered(316.3 + 207.0 / 4.0, 384.3, "Energieverbrauch", size=8.2, font_style='B', color=c_white)
    draw_text_centered(316.3 + 3 * 207.0 / 4.0, 384.3, "Fettverbrauch", size=8.2, font_style='B', color=c_white)
 
    fatmax_speed_str = f"{fatmax_speed:.2f} m/s".replace('.', ',') if fatmax_speed is not None else "-"
    fatmax_total_energy_str = f"{fatmax_total_energy} kcal/h" if fatmax_total_energy != "-" else "-"
    fatmax_fat_energy_str = f"{fatmax_fat_energy} kcal/h" if fatmax_fat_energy != "-" else "-"
    
    draw_text_centered(109.35, 394.34, fatmax_speed_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(212.85, 394.34, fatmax_pace_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(316.3 + 207.0 / 4.0, 394.34, fatmax_total_energy_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text_centered(316.3 + 3 * 207.0 / 4.0, 394.34, fatmax_fat_energy_str, size=8.3, font_style='B', color=c_dark_grey)

    # Türkiser Balken für das zweite Diagramm: "Fett- & Kohlenhydrat-Stoffwechsel"
    draw_rect(57.6, 431.0, 465.7, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 465.7 / 2.0, 432.1, "Fett- & Kohlenhydrat-Stoffwechsel", size=8.2, font_style='B', color=c_white)

    # Energetik chart matching original HYCYS PDF Page 4 1:1
    fig, ax = plt.subplots(figsize=(6.0, 3.2), dpi=300)
    ax.set_xlabel('Geschwindigkeit [m/s]', color='#595a59', fontsize=7, fontweight='bold')
    ax.set_ylabel('Energie [kcal/h]', color='#595a59', fontsize=7, fontweight='bold')
    
    v_start = float(stage_speeds_pdf[0]) if len(stage_speeds_pdf) > 0 else 2.5
    v_end = float(stage_speeds_pdf[-1]) if len(stage_speeds_pdf) > 0 else 4.5
    x_lo = v_start
    x_hi = v_end
    ax.set_xlim(x_lo, x_hi)
    if len(stage_speeds_pdf) > 0:
        ax.set_xticks(stage_speeds_pdf)
        _all_1dec_e = all(abs(round(v, 1) - round(v, 2)) < 0.015 for v in stage_speeds_pdf)
        ax.set_xticklabels([f"{v:.1f}".replace(".", ",") if _all_1dec_e else f"{v:.2f}".replace(".", ",") for v in stage_speeds_pdf])
    
    ax2 = ax.twinx()
    ax2.set_ylabel('Kohlenhydratverbrauch [g/h]', color='#cdb663', fontsize=7, fontweight='bold')
    
    if has_spiro and fatmax_speed is not None:
        vis_mask = (v_grid >= x_lo - 1e-6) & (v_grid <= x_hi + 1e-6)
        v_vis = v_grid[vis_mask]
        vis_ee = ee_kcal_h_grid[vis_mask]
        vis_fat = ee_fat_kcal_h_grid[vis_mask]
        vis_carb = carb_g_h_grid[vis_mask]

        # Compute Y-axis limits matching original HYCYS Report Running Page 4 (0..1.500 kcal/h and 0..450 g/h)
        import math
        from matplotlib.patches import Rectangle
        max_ee_vis = float(np.max(vis_ee)) if len(vis_ee) > 0 else 1350.0
        y1_max = 1500.0 if max_ee_vis <= 1480.0 else math.ceil(max_ee_vis * 1.05 / 250.0) * 250.0
        max_carb_vis = float(np.max(vis_carb)) if len(vis_carb) > 0 else 330.0
        y2_max = 450.0 if max_carb_vis <= 430.0 else math.ceil(max_carb_vis * 1.08 / 50.0) * 50.0
        
        ax.set_ylim(0, y1_max)
        ax.set_yticks(np.arange(0, y1_max + 1, 250))
        ax2.set_ylim(0, min(y2_max, 420.0) if y2_max == 450.0 else y2_max)
        ax2.set_yticks(np.arange(0, 401, 50) if y2_max == 450.0 else np.arange(0, y2_max + 1, 50))
        if y2_max == 450.0:
            ax2.set_ylim(0, 425.0)
            y2_max = 425.0

        # Fatmax bar in background (height ~72% of y1_max, turquoise fill with grey outline and white bold "Fatmax" inside top).
        # Falls Fatmax nahe am oder unter dem linken Diagrammrand liegt, den Balken vollständig ab dem linken Rand (x_lo) darstellen.
        if fatmax_speed > 0 and fatmax_speed <= x_hi:
            x_span_e = max(x_hi - x_lo, 1e-6)
            bar_w_fm = 0.095 * x_span_e
            bar_h_fm = 0.72 * y1_max
            bar_x0_fm = max(x_lo, min(x_hi - bar_w_fm, fatmax_speed - bar_w_fm / 2.0))
            bar_cx_fm = bar_x0_fm + bar_w_fm / 2.0
            rect_fm = Rectangle(
                (bar_x0_fm, 0.0),
                bar_w_fm,
                bar_h_fm,
                facecolor='#2cb7b9',
                edgecolor='#595a59',
                linewidth=1.1,
                alpha=0.65,
                zorder=2,
            )
            ax.add_patch(rect_fm)
            ax.text(
                bar_cx_fm,
                bar_h_fm - 0.045 * y1_max,
                "Fatmax",
                color='white',
                fontsize=7.2,
                fontweight='bold',
                ha='center',
                va='top',
                zorder=5,
            )

        # Plot fat curve (left axis) and carbohydrate curve (right axis) matching original HYCYS PDF Page 4
        ax.plot(v_vis, vis_fat, color='#2cb7b9', linewidth=2.5, zorder=4)
        ax2.plot(v_vis, vis_carb, color='#cdb663', linewidth=2.5, linestyle='-', zorder=4)
        
        bbox_props_e = dict(boxstyle='square,pad=0.22', facecolor='white', edgecolor='none', alpha=0.92)
        x_span_e = max(x_hi - x_lo, 1e-6)

        # 1. 'Verbrauch von Fetten' Rechts vom Fatmax-Balken direkt über der türkisen Fettkurve (wie im Original-PDF)
        if fatmax_speed > 0 and fatmax_speed <= x_hi:
            _fm_right_edge = max(x_lo, min(x_hi - 0.095 * x_span_e, fatmax_speed - 0.0475 * x_span_e)) + 0.095 * x_span_e
            fat_lbl_v = min(x_hi - 0.18 * x_span_e, _fm_right_edge + 0.13 * x_span_e)
        else:
            fat_lbl_v = x_lo + 0.45 * x_span_e
        fat_at_lbl = float(np.interp(fat_lbl_v, v_vis, vis_fat))
        fat_lbl_y = min(0.85 * y1_max, fat_at_lbl + 0.115 * y1_max)
        ax.text(
            fat_lbl_v,
            fat_lbl_y,
            'Verbrauch von Fetten',
            color='#2cb7b9',
            fontsize=7.5,
            fontweight='bold',
            ha='center',
            va='center',
            bbox=bbox_props_e,
            zorder=10,
        )

        # 2. 'Verbrauch von\nKohlenhydraten' über dem oberen Ast der goldenen Kohlenhydratkurve (wie im Original-PDF)
        carb_lbl_v = x_lo + 0.71 * x_span_e
        carb_at_lbl = float(np.interp(carb_lbl_v, v_vis, vis_carb))
        carb_lbl_y = min(0.91 * y2_max, max(0.55 * y2_max, carb_at_lbl + 0.11 * y2_max))
        ax2.text(
            carb_lbl_v,
            carb_lbl_y,
            'Verbrauch von\nKohlenhydraten',
            color='#cdb663',
            fontsize=7.5,
            fontweight='bold',
            ha='center',
            va='center',
            bbox=bbox_props_e,
            zorder=10,
        )
    else:
        ax.set_ylim(0, 1500)
        ax.set_yticks(np.arange(0, 1501, 250))
        ax2.set_ylim(0, 425)
        ax2.set_yticks(np.arange(0, 401, 50))
        
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f"{int(round(x)):,}".replace(",", ".")))
    ax.tick_params(axis='both', labelcolor='#595a59', labelsize=6, width=1.2)
    ax2.set_ylabel('Kohlenhydratverbrauch [g/h]', color='#595a59', fontsize=7, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#595a59', labelsize=6, width=1.2, colors='#595a59')
    for lbl_y2_e in ax2.get_yticklabels():
        lbl_y2_e.set_color('#595a59')
        lbl_y2_e.set_weight('bold')
    for lbl_x_e in ax.get_xticklabels():
        lbl_x_e.set_weight('bold')
    for lbl_y_e in ax.get_yticklabels():
        lbl_y_e.set_weight('bold')
    for s in ax.spines.values(): s.set_color('#595a59'); s.set_linewidth(1.2)
    for s in ax2.spines.values(): s.set_color('#595a59'); s.set_linewidth(1.2)
    ax.grid(True, axis='y', linestyle='-', linewidth=0.8, color='#d8d8d8')
    ax.set_axisbelow(True)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_ener_chart:
        fig.savefig(tmp_ener_chart.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_ener_chart.name, x=55.9, y=449.5, w=472, h=254)
        temp_files.append(tmp_ener_chart.name)

    # --- SEITE 5: TRAININGSBEREICHE & REFERENZDATEN ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    
    page_logo(5)
        
    draw_rect(57.6, 133.9, 155.3, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 155.3 / 2.0, 135.0, "Referenzdaten", size=8.2, font_style='B', color=c_white)

    # Cartesian Radar chart for Reference data matching original HYCYS PDF Page 5 1:1
    fig, ax = plt.subplots(figsize=(5.0, 5.0), dpi=300)
    ax.set_aspect('equal')
    
    radar_cats = ['VO2max', 'Fettanteil', 'Laufökonomie', 'ANS', 'Fatmax', 'KH-Match', 'VLamax']
    radar_labels = [
        'VO2max [ml/min/kg]',
        'Fettanteil [%]',
        'Laufökonomie [ml/min/kg*m/s]',
        'ANS [m/s]',
        'Fatmax [m/s]',
        'KH-Match [m/s]',
        'VLamax [mmol/l/s]'
    ]
    N = len(radar_cats)
    # Excel Radar Chart: Kategorie 0 (VO2max) startet oben bei 90° und läuft im Uhrzeigersinn!
    angles = [np.pi / 2.0 - n / float(N) * 2.0 * np.pi for n in range(N)]
    angles += angles[:1]
    
    # Background rings (Sehr gut > Gut > Zu verbessern > Schwach: 100 / 80 / 60 / 35)
    x_100 = [100 * np.cos(t) for t in angles]
    y_100 = [100 * np.sin(t) for t in angles]
    ax.fill(x_100, y_100, color='#cdb663', edgecolor='none', zorder=1, label='Sehr gut')
    
    x_80 = [80 * np.cos(t) for t in angles]
    y_80 = [80 * np.sin(t) for t in angles]
    ax.fill(x_80, y_80, color='#2cb7b9', edgecolor='none', zorder=2, label='Gut')
    
    x_60 = [60 * np.cos(t) for t in angles]
    y_60 = [60 * np.sin(t) for t in angles]
    ax.fill(x_60, y_60, color='#ffffff', edgecolor='none', zorder=3, label='Zu verbessern')
    
    x_35 = [35 * np.cos(t) for t in angles]
    y_35 = [35 * np.sin(t) for t in angles]
    ax.fill(x_35, y_35, color='#595a59', edgecolor='none', zorder=4, label='Schwach')
    
    # Normalize running reference values smoothly matching HYCYS Reference data
    def normalize_ref_running(val, cat):
        ref_params = {
            'VO2max':       (52.0,  5.0,   False),
            'Fettanteil':   (14.0,  2.5,   True),
            'Laufökonomie': (10.5,  1.2,   True),
            'ANS':          (3.6,   0.35,  False),
            'Fatmax':       (2.8,   0.35,  False),
            'KH-Match':     (2.6,   0.35,  False),
            'VLamax':       (0.58,  0.10,  True),
        }
        if cat not in ref_params:
            return 65.0
        mean, sd, inverted = ref_params[cat]
        if sd <= 0:
            return 65.0
        z = (mean - val) / sd if inverted else (val - mean) / sd
        raw_score = 62.0 + 13.5 * np.tanh(z / 1.65) + 4.5 * np.clip(z, -2.2, 2.2)
        return float(np.clip(raw_score, 44.0, 94.0))

    lo_mean_pdf = float(np.mean(valid_lo_s3)) if len(valid_lo_s3) > 0 else 10.5
    v_fatmax = fatmax_speed if (fatmax_speed is not None and fatmax_speed > 0) else (v_ans * 0.78)
    # In HYCYS Reference data für Laufen wird KH-Match standardmäßig mit 0,6 g/kg KG/h (bzw. eingestelltem Slider) berechnet
    kh_target_pdf = 0.6 * float(weight)
    v_kh_match = v_ans * 0.78
    if has_spiro and v_grid is not None and 'carb_g_h_grid' in locals():
        try:
            sub_m_khm = (delta_vo2_grid > 0)
            if np.any(sub_m_khm):
                v_sub_k = v_grid[sub_m_khm]
                c_sub_k = carb_g_h_grid[sub_m_khm]
                le_k = c_sub_k <= kh_target_pdf
                if np.any(le_k):
                    v_kh_match = float(v_sub_k[int(np.where(le_k)[0][-1])])
                else:
                    v_kh_match = float(v_sub_k[int(np.argmin(np.abs(c_sub_k - kh_target_pdf)))])
        except Exception:
            pass

    athlete_scores = [
        normalize_ref_running(rel_vo2max,   'VO2max'),
        normalize_ref_running(body_fat_pct, 'Fettanteil'),
        normalize_ref_running(lo_mean_pdf,  'Laufökonomie'),
        normalize_ref_running(v_ans,        'ANS'),
        normalize_ref_running(v_fatmax,     'Fatmax'),
        normalize_ref_running(v_kh_match,   'KH-Match'),
        normalize_ref_running(vla_val,      'VLamax'),
    ]
    athlete_scores += athlete_scores[:1]

    # Draw inner ring helper boundaries and radial spokes so helper lines extend visibly into the dark-grey center
    for r_ring_line in [35, 60, 80]:
        ax.plot(
            [r_ring_line * np.cos(t) for t in angles],
            [r_ring_line * np.sin(t) for t in angles],
            color='#8c8d8c' if r_ring_line == 35 else '#b8b9b8',
            linestyle='-',
            linewidth=0.6,
            zorder=5,
        )
    for t in angles[:-1]:
        ax.plot([0, 35 * np.cos(t)], [0, 35 * np.sin(t)], color='#9e9f9e', linestyle='-', linewidth=0.8, zorder=5)
        ax.plot([35 * np.cos(t), 100 * np.cos(t)], [35 * np.sin(t), 100 * np.sin(t)], color='#595a59', linestyle='-', linewidth=0.8, zorder=5)

    # Draw outer heptagon spine boundary
    ax.plot(x_100, y_100, color='#595a59', linewidth=1.2, zorder=8)

    # Draw athlete line (markante schwarze Linie wie im Original-HYCYS-PDF)
    x_ath = [r * np.cos(t) for r, t in zip(athlete_scores, angles)]
    y_ath = [r * np.sin(t) for r, t in zip(athlete_scores, angles)]
    ax.plot(x_ath, y_ath, color='#000000', linewidth=2.8, zorder=7)

    # Add text labels manually
    for i, label in enumerate(radar_labels):
        t = angles[i]
        r_label = 107
        x = r_label * np.cos(t)
        y = r_label * np.sin(t)
        
        ha = 'center'
        if np.cos(t) > 0.15:
            ha = 'left'
        elif np.cos(t) < -0.15:
            ha = 'right'
            
        va = 'center'
        if np.sin(t) > 0.25:
            va = 'bottom'
        elif np.sin(t) < -0.25:
            va = 'top'
            
        if abs(np.cos(t)) < 0.15:
            if np.sin(t) > 0:
                y += 1.5
            else:
                y -= 1.5
                
        ax.text(x, y, label, fontsize=7.0, fontweight='bold', color='#595a59', ha=ha, va=va)

    ax.axis('off')
    ax.set_xlim(-145, 145)
    ax.set_ylim(-130, 130)
    
    from matplotlib.lines import Line2D
    legend_handles = [
        Patch(facecolor='#cdb663', alpha=1.0,  label='Sehr gut'),
        Patch(facecolor='#2cb7b9', alpha=1.0,  label='Gut'),
        Patch(facecolor='#ffffff', edgecolor='#cccccc', alpha=1.0,  label='Zu verbessern'),
        Patch(facecolor='#595a59', alpha=1.0,  label='Schwach'),
        Line2D([0], [0], color='#000000', lw=2.5, label='Daten Sportler'),
    ]
    legend = ax.legend(
        handles=legend_handles,
        loc='lower center',
        bbox_to_anchor=(0.5, -0.06),
        ncol=5,
        fontsize=6.4,
        frameon=False,
        handlelength=1.1,
        handletextpad=0.4,
        columnspacing=1.0,
    )
    for text in legend.get_texts():
        text.set_color('#1a1a1a')
        text.set_weight('bold')
        
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_radar:
        fig.savefig(tmp_radar.name, format="png", bbox_inches="tight", facecolor='white')
        plt.close(fig)
        
        img = Image.open(tmp_radar.name)
        w_img, h_img = img.size
        _w_spider_l = 320.0
        _h_spider_l = _w_spider_l * (h_img / float(w_img))
        pdf.image(tmp_radar.name, x=130.0, y=145.0, w=_w_spider_l, h=_h_spider_l)
        temp_files.append(tmp_radar.name)

    # Trainingsbereiche Table
    draw_rect(57.6, 471.4, 207.0, 10.44, fill_color=c_teal)
    draw_text_centered(57.6 + 207.0 / 2.0, 472.5, "Trainingsbereiche", size=8.2, font_style='B', color=c_white)

    draw_text(99.3, 497.0, "Trainingsbereich", size=9.0, font_style='B', color=c_dark_grey)
    draw_text(254.8, 497.0, "Geschwindigkeit", size=9.0, font_style='B', color=c_dark_grey)
    draw_text(417.4, 497.0, "Herzfrequenz", size=9.0, font_style='B', color=c_dark_grey)
    draw_text(275.3, 510.7, "[min/km]", size=8.3, color=c_dark_grey)
    draw_text(435.8, 510.7, "[bpm]", size=8.3, color=c_dark_grey)
    
    draw_text(231.8, 523.9, "min", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(282.6, 523.9, "max", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(335.5, 523.9, "Ziel", size=8.3, color=c_dark_grey)
    draw_text(387.0, 523.9, "min", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(437.9, 523.9, "max", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(490.7, 523.9, "Ziel", size=8.3, color=c_dark_grey)

    def format_pace_sec(sec_per_km):
        if not isinstance(sec_per_km, (int, float, np.number)) or sec_per_km <= 0:
            return "-"
        minutes = int(sec_per_km // 60)
        seconds = int(round(sec_per_km % 60))
        if seconds == 60:
            minutes += 1
            seconds = 0
        return f"{minutes:02d}:{seconds:02d}"

    def format_hr(hr_val):
        if not isinstance(hr_val, (int, float, np.number)) or hr_val <= 0:
            return "-"
        return str(int(round(float(hr_val))))

    # Original HYCYS Excel Report Running Trainingsbereiche (DLreg, DLext, DLint, DLTempo, TText, TTint)
    valid_hr_mask = (np.asarray(speed, dtype=float) > 0) & (np.asarray(hr, dtype=float) > 0) & (~np.isnan(np.asarray(speed, dtype=float))) & (~np.isnan(np.asarray(hr, dtype=float)))
    valid_sp_hr = np.asarray(speed, dtype=float)[valid_hr_mask]
    valid_hr_arr = np.asarray(hr, dtype=float)[valid_hr_mask]

    def interp_hr_at_v(v_target):
        if len(valid_sp_hr) == 0 or v_target <= 0:
            return 0.0
        if len(valid_sp_hr) == 1:
            return float(valid_hr_arr[0])
        if v_target <= valid_sp_hr[0]:
            m_h = (valid_hr_arr[1] - valid_hr_arr[0]) / max(valid_sp_hr[1] - valid_sp_hr[0], 1e-6)
            return float(valid_hr_arr[0] + m_h * (v_target - valid_sp_hr[0]))
        if v_target >= valid_sp_hr[-1]:
            m_h = (valid_hr_arr[-1] - valid_hr_arr[-2]) / max(valid_sp_hr[-1] - valid_sp_hr[-2], 1e-6)
            return float(valid_hr_arr[-1] + m_h * (v_target - valid_sp_hr[-1]))
        return float(np.interp(v_target, valid_sp_hr, valid_hr_arr))

    if isinstance(v_ans, (int, float, np.number)) and v_ans > 0:
        sec_68  = 1000.0 / (v_ans * 0.68)
        sec_79  = 1000.0 / (v_ans * 0.79)
        sec_93  = 1000.0 / (v_ans * 0.93)
        sec_105 = 1000.0 / (v_ans * 1.05)
        sec_123 = 1000.0 / (v_ans * 1.23)
        sec_125 = 1000.0 / (v_ans * 1.25)

        hr_76  = interp_hr_at_v(v_ans * 0.76)
        hr_80  = interp_hr_at_v(v_ans * 0.80)
        hr_86  = interp_hr_at_v(v_ans * 0.86)
        hr_93  = interp_hr_at_v(v_ans * 0.93)
        hr_96  = interp_hr_at_v(v_ans * 0.96)
        hr_105 = interp_hr_at_v(v_ans * 1.05)
        hr_110 = interp_hr_at_v(v_ans * 1.10)
        hr_120 = interp_hr_at_v(v_ans * 1.20)
        hr_123 = interp_hr_at_v(v_ans * 1.23)

        zones_data = [
            ("regenerativer Dauerlauf (DLreg)",
             "-", format_pace_sec(sec_68), format_pace_sec(sec_68 * 1.05),
             "-", format_hr(hr_76), format_hr(hr_76 * 0.90)),
            ("extensiver Dauerlauf (DLext)",
             format_pace_sec(sec_68), format_pace_sec(sec_79), format_pace_sec(((sec_68 + sec_79) / 2.0) * 0.95),
             format_hr(hr_76), format_hr(hr_86), format_hr(hr_80)),
            ("intensiver  Dauerlauf (DLint)",
             format_pace_sec(sec_79), format_pace_sec(sec_93), format_pace_sec((sec_79 + sec_93) / 2.0),
             format_hr(hr_86), format_hr(hr_96), format_hr(hr_93)),
            ("Tempodauerlauf (DLTempo)",
             format_pace_sec(sec_93), format_pace_sec(sec_105), format_pace_sec((sec_93 + sec_105) / 2.0),
             format_hr(hr_96), format_hr(hr_105), format_hr((hr_96 + hr_105) / 2.0)),
            ("extensives Tempotraining (TText)",
             format_pace_sec(sec_105), format_pace_sec(sec_123), format_pace_sec((sec_105 + sec_123) / 2.0),
             format_hr(hr_105), format_hr(hr_120), format_hr((hr_105 + hr_120) / 2.0)),
            ("intensives Tempotraining (TTint)",
             format_pace_sec(sec_123), "max", format_pace_sec(sec_125),
             format_hr(hr_110), "max" if len(valid_sp_hr) > 0 else "-", format_hr((hr_110 + hr_123) / 2.0)),
        ]
    else:
        zones_data = [
            ("regenerativer Dauerlauf (DLreg)",  "-", "-", "-", "-", "-", "-"),
            ("extensiver Dauerlauf (DLext)",     "-", "-", "-", "-", "-", "-"),
            ("intensiver  Dauerlauf (DLint)",    "-", "-", "-", "-", "-", "-"),
            ("Tempodauerlauf (DLTempo)",         "-", "-", "-", "-", "-", "-"),
            ("extensives Tempotraining (TText)", "-", "-", "-", "-", "-", "-"),
            ("intensives Tempotraining (TTint)", "-", "-", "-", "-", "-", "-"),
        ]

    y_coords_s5 = [536.1, 549.4, 562.7, 576.0, 589.3, 602.6]
    for (zn, p_min, p_max, p_zl, hr_mn, hr_mx, hr_zl), y_coord in zip(zones_data, y_coords_s5):
        draw_text(59.3, y_coord, zn, size=8.3, font_style='B', color=c_dark_grey)
        
        if p_min == "-":
            draw_text_centered(231.8 + 6.0, y_coord, "-", size=8.3, color=c_dark_grey)
        else:
            draw_text(231.8, y_coord, p_min, size=8.3, color=c_dark_grey)
            
        draw_text(282.6, y_coord, p_max, size=8.3, color=c_dark_grey)
        draw_text(335.5, y_coord, p_zl, size=8.3, color=c_dark_grey)
        
        if hr_mn == "-":
            draw_text_centered(387.0 + 6.0, y_coord, "-", size=8.3, color=c_dark_grey)
        else:
            draw_text(387.0, y_coord, hr_mn, size=8.3, color=c_dark_grey)
            
        draw_text(437.9, y_coord, hr_mx, size=8.3, color=c_dark_grey)
        if hr_zl:
            draw_text(490.7, y_coord, hr_zl, size=8.3, color=c_dark_grey)

    for x in [212.9, 368.8, 524.7]:
        draw_rect(x, 522.0, 0.6, 93.6, fill_color=c_dark_grey)
    draw_rect(212.9, 615.0, 311.8, 0.6, fill_color=c_dark_grey)

    pdf_output = pdf.output(dest='S')
    
    for f in temp_files:
        try:
            os.unlink(f)
        except Exception:
            pass
            
    return pdf_output.encode('latin-1', 'replace')


# 3. BERECHNUNG & ANZEIGE
if st.session_state.get("lauf_auswertung_gestartet", False):
    df_valid = st.session_state.df_lauftest_input.copy()
    df_valid["v (m/s)"] = pd.to_numeric(df_valid["v (m/s)"], errors='coerce')
    df_valid["Laktat"] = pd.to_numeric(df_valid["Laktat"], errors='coerce')
    df_valid["HF"] = pd.to_numeric(df_valid["HF"], errors='coerce')
    
    ruhe_lac = float(df_valid.iloc[0]["Laktat"]) if len(df_valid) > 0 and pd.notna(df_valid.iloc[0]["Laktat"]) else 1.0
    ruhe_hr = float(df_valid.iloc[0]["HF"]) if len(df_valid) > 0 and pd.notna(df_valid.iloc[0]["HF"]) else np.nan
    
    df_stages = df_valid.iloc[1:].dropna(subset=["v (m/s)", "Laktat"]).reset_index(drop=True)

    # Exclude Ruhe (first row) for exercise steps calculations
    speed = df_stages["v (m/s)"].values
    lactate = df_stages["Laktat"].values
    hr = df_stages["HF"].values
    
    raw_speeds = np.concatenate([[0.0], speed])
    raw_lactate = np.concatenate([[ruhe_lac], lactate])
    raw_hr = np.concatenate([[ruhe_hr], hr])
    abs_vo2max, rel_vo2max = 0.0, 0.0
    _cur_stufendauer = float(st.session_state.get("stufendauer", 5.0))
    _is_3min_proto = (test_type == "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)") or (_cur_stufendauer <= 3.5)
    steady_fenster = 90 if _is_3min_proto else 120      # 90s für 3min-Stufen, 120s für 5min-Stufen
    vo2_steady_values = []
    vo2_steady_abs = []      # absoluter VO2-Steady-State [ml/min] je Stufe
    vco2_steady_abs = []     # absoluter VCO2-Steady-State [ml/min] je Stufe
    spiro_fett_steady_kcal_h = []  # gemessener Fettverbrauch [kcal/h] je Stufe aus Spiro
    spiro_kh_steady_g_h = []       # gemessener KH-Verbrauch [g/h] je Stufe aus Spiro
    window_coords = []
    spiro_df = pd.DataFrame()

    sf_means = st.session_state.df_sf[["M1", "M2", "M3"]].mean(axis=1)
    sum_sf = sf_means.sum()
    if sum_sf > 0:
        body_fat_pct = (22.32 * np.log10(sum_sf)) - 29.2

    spiro_interval_sec = None   # wird beim Einlesen ermittelt (10 oder 30)
    spiro_interval_note = ""    # Hinweis bei 30s für VO2max-Anzeige

    if uploaded_file is not None or ("lauftest_spiro_df" in st.session_state and not st.session_state.lauftest_spiro_df.empty):
        try:
            if "lauftest_spiro_df" in st.session_state and "FETT_kcal_h" in st.session_state.lauftest_spiro_df.columns:
                spiro_df = st.session_state.lauftest_spiro_df
            else:
                uploaded_file.seek(0)
                df_excel = pd.read_csv(uploaded_file, header=None, low_memory=False) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
                uploaded_file.seek(0)
                times_sec = df_excel.iloc[3:, 9].apply(parse_time_to_seconds)
                vo2_col = pd.to_numeric(df_excel.iloc[3:, 13], errors='coerce')
                vco2_col = pd.to_numeric(df_excel.iloc[3:, 14], errors='coerce')
                eeh_col = pd.to_numeric(df_excel.iloc[3:, 36], errors='coerce') if df_excel.shape[1] > 36 else pd.Series(np.nan, index=df_excel.iloc[3:].index)
                fett_kcal_h_col = (pd.to_numeric(df_excel.iloc[3:, 41], errors='coerce') / 24.0) if df_excel.shape[1] > 41 else pd.Series(np.nan, index=df_excel.iloc[3:].index)
                kh_g_h_col = ((pd.to_numeric(df_excel.iloc[3:, 42], errors='coerce') / 24.0) / 4.063) if df_excel.shape[1] > 42 else pd.Series(np.nan, index=df_excel.iloc[3:].index)
                spiro_df = pd.DataFrame({
                    'Time': times_sec,
                    'VO2': vo2_col,
                    'VCO2': vco2_col,
                    'EEh': eeh_col,
                    'FETT_kcal_h': fett_kcal_h_col,
                    'KH_g_h': kh_g_h_col,
                }).dropna(subset=['Time'])
                st.session_state.lauftest_spiro_df = spiro_df

            # --- Zeitintervall automatisch erkennen ---
            if len(spiro_df) >= 2:
                t_diffs = spiro_df['Time'].diff().dropna()
                spiro_interval_sec = int(round(t_diffs.median()))
            else:
                spiro_interval_sec = 10  # default

            if spiro_interval_sec == 30:
                spiro_interval_note = "(aus 30s-Auflösung, Genauigkeit eingeschränkt)"
            elif spiro_interval_sec == 10:
                spiro_interval_note = ""
            else:
                spiro_interval_note = f"({spiro_interval_sec}s-Auflösung)"

            if not spiro_df.empty:
                # VO2max Rolling-Window: immer der höchste über 30s gemittelte Wert
                # (3 Punkte bei 10s-Auflösung = 30s, 1 Punkt bei 30s-Auflösung = 30s)
                if not separate_ramp:
                    window_pts = max(1, int(round(30.0 / spiro_interval_sec)))
                    abs_vo2max = float(spiro_df['VO2'].rolling(window=window_pts, center=True).mean().max())
                else:
                    abs_vo2max = 0.0
                
                # Wenn Rampe getrennt hochgeladen wurde, VO2max ausschließlich aus der Rampe
                # als höchsten über 30s gemittelten Wert berechnen (Stufentest nur für Laufökonomie)
                if separate_ramp and uploaded_ramp is not None:
                    try:
                        uploaded_ramp.seek(0)
                        df_ramp_excel = pd.read_csv(uploaded_ramp, header=None, low_memory=False) if uploaded_ramp.name.endswith('.csv') else pd.read_excel(uploaded_ramp, header=None)
                        ramp_times_sec = df_ramp_excel.iloc[3:, 9].apply(parse_time_to_seconds)
                        ramp_vo2_col = pd.to_numeric(df_ramp_excel.iloc[3:, 13], errors='coerce')
                        ramp_df = pd.DataFrame({'Time': ramp_times_sec, 'VO2': ramp_vo2_col}).dropna(subset=['Time']).reset_index(drop=True)
                        if len(ramp_df) >= 2:
                            r_t_diffs = ramp_df['Time'].diff().dropna()
                            ramp_interval_sec = int(round(r_t_diffs.median()))
                        else:
                            ramp_interval_sec = 10
                        if ramp_interval_sec == 10:
                            spiro_interval_note = ""
                        if not ramp_df.empty and ramp_df['VO2'].notna().any():
                            ramp_window_pts = max(1, int(round(30.0 / ramp_interval_sec)))
                            abs_vo2max = float(ramp_df['VO2'].rolling(window=ramp_window_pts, center=True).mean().max())
                    except Exception as e:
                        st.error(f"Fehler beim Auslesen der Rampendatei: {e}")

                rel_vo2max = abs_vo2max / weight if weight > 0 else 0

                stufendauer = float(st.session_state.get("stufendauer", 5.0))
                pausendauer = st.session_state.get("pausendauer", 30)
                vorlauf = st.session_state.get("vorlauf", 60)
                stage_sec = int(stufendauer * 60)
                pause_sec = int(pausendauer)
                steady_fenster = 90 if ((test_type == "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)") or (stufendauer <= 3.5)) else 120
                max_spiro_t = float(spiro_df['Time'].max()) if not spiro_df.empty else 0.0
                det_ends_valid = []
                if "detected_stage_ends" in st.session_state and len(st.session_state.detected_stage_ends) >= 2:
                    det_med_int = float(np.median(np.diff(st.session_state.detected_stage_ends)))
                    if abs(det_med_int - float(stage_sec + pause_sec)) <= 45.0:
                        det_ends_valid = st.session_state.detected_stage_ends
                for i in range(anzahl):
                    if i < len(det_ends_valid):
                        end_t_base = det_ends_valid[i]
                    else:
                        end_t_base = vorlauf + (i + 1) * stage_sec + i * pause_sec
                    
                    end_t_val = st.session_state.get(f"lauftest_end_t_{i}", end_t_base)
                    end_t = float(end_t_val)
                    start_t = end_t - steady_fenster
                    
                    # Prüfen, ob für diese Stufe im Stufentest noch valide Spiro-Daten vorliegen
                    fett_st_val = np.nan
                    kh_st_val = np.nan
                    if start_t >= max_spiro_t or (end_t - max_spiro_t) > max(30.0, float(spiro_interval_sec)):
                        vo2_val = 0.0
                        vco2_val = 0.0
                    else:
                        mask = (spiro_df['Time'] > start_t) & (spiro_df['Time'] <= end_t)
                        stage_data = spiro_df.loc[mask]
                        vo2_val  = float(stage_data['VO2'].mean())  if not stage_data.empty and stage_data['VO2'].notna().any() else 0.0
                        vco2_val = float(stage_data['VCO2'].mean()) if not stage_data.empty and stage_data['VCO2'].notna().any() else 0.0
                        # Falls am Ende der Datei nur noch Abbruch/Cooldown-Punkte liegen (VO2 fällt stark ab gegenüber Vorstufe)
                        if i > 0 and vo2_steady_abs[i - 1] > 0 and vo2_val < 0.80 * vo2_steady_abs[i - 1]:
                            vo2_val = 0.0
                            vco2_val = 0.0
                        else:
                            window_coords.append((start_t, end_t))
                            if 'FETT_kcal_h' in stage_data.columns and stage_data['FETT_kcal_h'].notna().any():
                                fett_st_val = float(stage_data['FETT_kcal_h'].mean())
                            elif vo2_val > 0 and vco2_val > 0:
                                _rq_st = float(np.clip(vco2_val / vo2_val, 0.70, 1.00))
                                _fat_pct_st = (1.0 - _rq_st) / 0.30
                                _kal_eq_st = (_rq_st - 0.70) * 5.094 + 19.6
                                _ee_st = (vo2_val / 1000.0 * _kal_eq_st * 60.0) / 4.186
                                fett_st_val = float(_ee_st * _fat_pct_st)
                            if 'KH_g_h' in stage_data.columns and stage_data['KH_g_h'].notna().any():
                                kh_st_val = float(stage_data['KH_g_h'].mean())
                            elif vo2_val > 0 and vco2_val > 0:
                                _rq_st = float(np.clip(vco2_val / vo2_val, 0.70, 1.00))
                                _kh_pct_st = (_rq_st - 0.70) / 0.30
                                _kal_eq_st = (_rq_st - 0.70) * 5.094 + 19.6
                                _ee_st = (vo2_val / 1000.0 * _kal_eq_st * 60.0) / 4.186
                                kh_st_val = float((_ee_st * _kh_pct_st) / 4.063)
                    vo2_steady_values.append(vo2_val / weight if weight > 0 else 0.0)
                    vo2_steady_abs.append(vo2_val)
                    vco2_steady_abs.append(vco2_val)
                    spiro_fett_steady_kcal_h.append(fett_st_val)
                    spiro_kh_steady_g_h.append(kh_st_val)
        except Exception as e:
            st.error(f"Fehler beim Einlesen: {e}")

    # Override VO2max if manual value is provided
    if vo2max_override > 0:
        abs_vo2max = float(vo2max_override)
        rel_vo2max = abs_vo2max / weight if weight > 0 else 0.0


    # --- SCHWELLENBERECHNUNG ---
    v_start, v_end = speed[0], speed[-1]
    results = []

    # 1. Fits
    # Fit 3rd degree polynomial on exercise stages (excluding baseline, matching ExPhysLab / lactater):
    poly_coeffs_no = np.polyfit(speed, lactate, 3)
    poly_func_no = np.poly1d(poly_coeffs_no)

    # 2. OBLA 4.0 (ANS)
    if 4.0 < lactate[0] or 4.0 > lactate[-1]:
        v_exact_40 = None
    else:
        v_exact_40 = np.interp(4.0, lactate, speed)
        
    if v_exact_40 is not None:
        results.append({'Modell': 'OBLA 4.0 (ANS)', 'v': v_exact_40, 'Laktat': 4.0})
    else:
        results.append({'Modell': 'OBLA 4.0 (ANS)', 'v': 'Laktatwerte zu niedrig um OBLA zu kalkulieren', 'Laktat': None})

    # 3. LTP1 & LTP2
    try:
        # Interpolate raw speed and lactate onto 0.1 m/s grid (excluding Ruhe)
        grid_speed = np.arange(speed[0], speed[-1] + 1e-9, 0.1)
        interpolated_lactate = np.interp(grid_speed, speed, lactate)
        
        def fit_3seg_fixed_bps(bp1, bp2, x, y):
            if bp1 <= x[0] or bp2 >= x[-1] or bp1 >= bp2:
                return 1e10
            A = np.column_stack([
                np.ones_like(x),
                x,
                np.maximum(0.0, x - bp1),
                np.maximum(0.0, x - bp2)
            ])
            try:
                coefs, residuals, rank, s_vals = np.linalg.lstsq(A, y, rcond=None)
                if len(residuals) > 0:
                    return residuals[0]
                else:
                    pred = A @ coefs
                    return np.sum((y - pred)**2)
            except:
                return 1e10

        # Coarse grid search (30 x 30)
        best_err = 1e10
        best_bps = None
        bp1_vals = np.linspace(speed[0] + 0.05, speed[-1] - 0.05, 30)
        bp2_vals = np.linspace(speed[0] + 0.05, speed[-1] - 0.05, 30)
        for bp1 in bp1_vals:
            for bp2 in bp2_vals:
                if bp1 < bp2:
                    err = fit_3seg_fixed_bps(bp1, bp2, grid_speed, interpolated_lactate)
                    if err < best_err:
                        best_err = err
                        best_bps = (bp1, bp2)

        # Refine search with Nelder-Mead
        import scipy.optimize as opt
        def loss_func(bps):
            return fit_3seg_fixed_bps(bps[0], bps[1], grid_speed, interpolated_lactate)
            
        res_ltp = opt.minimize(loss_func, x0=best_bps, method='Nelder-Mead')
        v_exact_ltp1, v_exact_ltp2 = res_ltp.x[0], res_ltp.x[1]
        
        # Associated lactates from polynomial on exercise stages (matching ExPhysLab retrieve_lactate: round to 1 decimal)
        lac_ltp1 = round(float(poly_func_no(v_exact_ltp1)), 1)
        lac_ltp2 = round(float(poly_func_no(v_exact_ltp2)), 1)
        
        v_ltp1_val = float(v_exact_ltp1)
        lac_ltp1_val = float(lac_ltp1)
        hr_mask_ltp1 = (speed > 0) & (hr > 0) & (~np.isnan(speed)) & (~np.isnan(hr))
        if hr_mask_ltp1.sum() >= 2:
            hr_slope_ltp1, hr_intercept_ltp1 = np.polyfit(speed[hr_mask_ltp1], hr[hr_mask_ltp1], 1)
            hf_ltp1_val = int(round(hr_slope_ltp1 * v_ltp1_val + hr_intercept_ltp1, 0))
        elif hr_mask_ltp1.sum() == 1:
            hf_ltp1_val = int(round(float(hr[hr_mask_ltp1][0]), 0))
        else:
            hf_ltp1_val = None
        
        # LTP1 wird immer im Hintergrund kalkuliert (für Trainingsbereiche & Kurve),
        # in die Schwellen & VLamax Matrix kommen jedoch nur OBLA 4.0 (ANS) und LTP2
        results.append({'Modell': 'LTP2', 'v': v_exact_ltp2, 'Laktat': lac_ltp2})
    except Exception as e:
        v_ltp1_val = None
        lac_ltp1_val = None
        hf_ltp1_val = None


    # Ergebnisanzeige Matrix aufbauen mit Error-Handling (nur OBLA 4.0 und LTP2)
    df_res_list = []
    for res in results:
        row = {'Modell': res['Modell'], 'v': res['v']}
        
        if isinstance(res['v'], str):
            row['m/s'] = res['v']
            row['min/km'] = ""
            row['Laktat'] = ""
            row['HF'] = ""
            row['VLamax [mmol/l/s]'] = ""
        else:
            row['m/s'] = round(res['v'], 2)
            row['min/km'] = format_pace_mmss(res['v'])
            row['Laktat'] = round(res['Laktat'], 1)
            
            # Linear model for HF instead of interp
            hr_mask = (speed > 0) & (hr > 0) & (~np.isnan(speed)) & (~np.isnan(hr))
            if hr_mask.sum() >= 2:
                hr_slope, hr_intercept = np.polyfit(speed[hr_mask], hr[hr_mask], 1)
                row['HF'] = int(round(hr_slope * res['v'] + hr_intercept, 0))
            elif hr_mask.sum() == 1:
                row['HF'] = int(round(float(hr[hr_mask][0]), 0))
            else:
                row['HF'] = "-"
            
            if not check_lauftest_lactate(edited_df):
                row['VLamax [mmol/l/s]'] = "Bitte Laktatwerte prüfen"
            elif uploaded_file is None and "lauftest_spiro_df" not in st.session_state and vo2max_override == 0:
                row['VLamax [mmol/l/s]'] = "Bitte lade eine Spirodatei hoch"
            else:
                vla, _, _ = calculate_vlamax_for_v(res['v'], rel_vo2max, speed, vo2_steady_abs, vco2_steady_abs, weight)
                row['VLamax [mmol/l/s]'] = f"{vla:.2f}" if vla is not None else "N/A"
                
        df_res_list.append(row)

    df_res = pd.DataFrame(df_res_list)
    
    # 1. Determine active threshold model name (always selectable between OBLA 4.0 (ANS) and LTP2)
    selected_model_name = st.session_state.get("selected_threshold_model", "OBLA 4.0 (ANS)")
        
    # 2. Extract v_ans and vla_val for downstream calculations
    v_ans = 3.5
    vla_val = 0.6
    hf_ans_val = "-"
    ans_pace_str = "-"
    vla_display = "N/A"
    
    if not df_res.empty:
        selected_row = df_res[df_res['Modell'] == selected_model_name]
        if selected_row.empty:
            selected_row = df_res[df_res['Modell'].str.contains('OBLA 4.0', regex=False, na=False)]
        if not selected_row.empty:
            v_val = selected_row.iloc[0]['v']
            if not isinstance(v_val, str) and pd.notna(v_val):
                v_ans = float(v_val)
                p_s = 1000.0 / v_ans if v_ans > 0 else 0
                ans_pace_str = f"{int(p_s // 60)}:{int(p_s % 60):02d}"
                
            hf_ans_val = selected_row.iloc[0]['HF']
            vla_raw = selected_row.iloc[0]['VLamax [mmol/l/s]']
            vla_display = str(vla_raw).replace('.', ',') if isinstance(vla_raw, (int, float)) else str(vla_raw)
            if pd.notna(vla_raw) and vla_raw != 'N/A' and 'bitte' not in str(vla_raw).lower():
                try:
                    vla_exact, _, _ = calculate_vlamax_for_v(v_ans, rel_vo2max, speed, vo2_steady_abs, vco2_steady_abs, weight)
                    vla_val = float(vla_exact) if vla_exact is not None else float(vla_raw)
                except ValueError:
                    pass

    # 3. Reorder df_res to put the active model row at the very top (index 0)
    if not df_res.empty:
        selected_idx = df_res[df_res['Modell'] == selected_model_name].index
        if not selected_idx.empty:
            idx = selected_idx[0]
            df_res = pd.concat([df_res.iloc[[idx]], df_res.drop(idx)]).reset_index(drop=True)

    # Ansichts-Umschalter: Ansicht 2 (Diagnostik) <-> Ansicht 3 (Präsentations-Modus)
    if "lauf_view_mode" not in st.session_state:
        st.session_state.lauf_view_mode = "🛠️ Diagnostik & Feineinstellung"

    col_vm_l1, col_vm_l2 = st.columns([2.2, 1.0])
    with col_vm_l1:
        _lauf_ansicht_lbl = "Diagnostik-Ansicht" if st.session_state.lauf_view_mode == "🛠️ Diagnostik & Feineinstellung" else "Präsentationsansicht"
        st.markdown(
            f"<div style='font-size:12.5px; color:#595a59; font-weight:500; padding-top:8px;'>Aktive Ansicht: <span style='color:#1a1a1a; font-weight:700;'>{_lauf_ansicht_lbl}</span></div>",
            unsafe_allow_html=True,
        )
    with col_vm_l2:
        if st.session_state.lauf_view_mode == "🛠️ Diagnostik & Feineinstellung":
            if st.button("🖥️ Zur Präsentationsansicht wechseln", type="primary", use_container_width=True, key="btn_to_pres_lauf_top"):
                st.session_state.lauf_view_mode = "🖥️ Präsentations-Modus (Athletengespräch)"
                st.rerun()
        else:
            if st.button("🛠️ Zurück zur Diagnostik-Ansicht", use_container_width=True, key="btn_to_diag_lauf_top"):
                st.session_state.lauf_view_mode = "🛠️ Diagnostik & Feineinstellung"
                st.rerun()

    is_lauf_pres = (st.session_state.lauf_view_mode == "🖥️ Präsentations-Modus (Athletengespräch)")

    # Header-Metriken & Körperfett-Tacho
    if is_lauf_pres:
        st.markdown(f"""
            <div style="font-size: 16px; line-height: 1.5; margin-bottom: 16px; padding: 12px 16px; border-left: 5px solid #2cb7b9; background-color: #f8fafa; border-radius: 4px;">
                <b>Athlet:</b> {athlete_name} &nbsp;|&nbsp; <b>Test-Datum:</b> {test_date}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
            <div style="font-size: 16px; line-height: 1.5; margin-bottom: 16px; padding: 12px 16px; border-left: 5px solid #2cb7b9; background-color: #f8fafa; border-radius: 4px;">
                <b>Athlet:</b> {athlete_name} &nbsp;|&nbsp; <b>Protokoll:</b> {test_type} &nbsp;|&nbsp; <b>Test-Datum:</b> {test_date} &nbsp;|&nbsp; <b>Kategorie:</b> {kategorie} &nbsp;|&nbsp; <b>Coach:</b> {coach}
            </div>
            """, unsafe_allow_html=True)
        
    v_fatmax = v_ans * 0.78
    has_spiro_for_fatmax = ((uploaded_file is not None) or ("lauftest_spiro_df" in st.session_state and not st.session_state.lauftest_spiro_df.empty)) and (len(vo2_steady_abs) > 0)
    if has_spiro_for_fatmax:
        try:
            vo2_rest_fm = 3.5 * weight
            ex_speeds_fm = []
            ac_stages_fm = []
            passed_thresh_fm = False
            for v_s, vo2_s, vco2_s in zip(speed, vo2_steady_abs, vco2_steady_abs):
                if v_s <= 0 or vo2_s <= 0:
                    continue
                if passed_thresh_fm:
                    break
                if v_ans > 0 and v_s > v_ans + 1e-6:
                    passed_thresh_fm = True
                net_vo2_s = max(0.0, vo2_s - vo2_rest_fm)
                rq_s = vco2_s / vo2_s if vco2_s > 0 else 0.85
                kal_equiv_s = (rq_s - 0.7) * 5.094 + 19.6
                ee_kJ_min_s = (net_vo2_s / 1000.0) * kal_equiv_s
                vo2_100kh_s = ee_kJ_min_s / 21.1 * 1000.0
                ex_speeds_fm.append(float(v_s))
                ac_stages_fm.append(vo2_100kh_s / (v_s * 100.0))
            if len(ex_speeds_fm) >= 1:
                v_grid_fm = np.arange(min(2.6, float(ex_speeds_fm[0])), 8.61, 0.01)
                ac_grid_fm = []
                for v_g in v_grid_fm:
                    if len(ex_speeds_fm) == 1:
                        ac_g = ac_stages_fm[0]
                    elif v_g <= ex_speeds_fm[0]:
                        m_fm = (ac_stages_fm[1] - ac_stages_fm[0]) / (ex_speeds_fm[1] - ex_speeds_fm[0])
                        c_fm = ac_stages_fm[0] - m_fm * ex_speeds_fm[0]
                        ac_g = m_fm * v_g + c_fm
                    elif v_g >= ex_speeds_fm[-1]:
                        ac_g = ac_stages_fm[-1]
                    else:
                        ac_g = ac_stages_fm[-1]
                        for idx_fm in range(len(ex_speeds_fm) - 1):
                            if ex_speeds_fm[idx_fm] <= v_g < ex_speeds_fm[idx_fm + 1]:
                                m_fm = (ac_stages_fm[idx_fm + 1] - ac_stages_fm[idx_fm]) / (ex_speeds_fm[idx_fm + 1] - ex_speeds_fm[idx_fm])
                                c_fm = ac_stages_fm[idx_fm] - m_fm * ex_speeds_fm[idx_fm]
                                ac_g = m_fm * v_g + c_fm
                                break
                    ac_grid_fm.append(ac_g)
                ac_grid_fm = np.array(ac_grid_fm)
                vo2_rel_demand_fm = (ac_grid_fm * v_grid_fm * 100.0) / weight + 1.0
                delta_vo2_fm = rel_vo2max - vo2_rel_demand_fm
                lac_clear_fm = (0.02049 * vo2_rel_demand_fm) / 0.4
                net_clear_fm = np.full_like(v_grid_fm, -9999.0)
                for idx_fm, (vo2_r_g, d_vo2_g) in enumerate(zip(vo2_rel_demand_fm, delta_vo2_fm)):
                    if d_vo2_g > 0 and vo2_r_g > 0:
                        adp_g = np.sqrt((0.0631 * vo2_r_g) / d_vo2_g)
                        pfk_g = 1.0 + (1.331 / (adp_g ** 3))
                        vla_min_g = vla_val * 60.0 / pfk_g
                        net_clear_fm[idx_fm] = lac_clear_fm[idx_fm] - vla_min_g
                v_fatmax = float(v_grid_fm[int(np.argmax(net_clear_fm))])
        except Exception:
            pass
    fatmax_pace_str = format_pace_mmss(v_fatmax) if v_fatmax > 0 else "-"
    ltp1_pace_str = format_pace_mmss(v_ltp1_val) if (v_ltp1_val is not None and v_ltp1_val > 0) else "-"
    hf_ltp1_str = f"{hf_ltp1_val} bpm" if hf_ltp1_val is not None else "-"

    lac_ans_val = None
    if not df_res.empty and "Laktat" in df_res.columns:
        raw_lac_ans = df_res.iloc[0]["Laktat"]
        if isinstance(raw_lac_ans, (int, float, np.number)) and pd.notna(raw_lac_ans):
            lac_ans_val = float(raw_lac_ans)

    ans_sub_lac_str = f"{lac_ans_val:.1f}".replace(".", ",") + " mmol/l Laktat" if lac_ans_val is not None else ""
    ltp1_sub_lac_str = f"{lac_ltp1_val:.1f}".replace(".", ",") + " mmol/l Laktat" if lac_ltp1_val is not None else ""
    ans_hr_str = f"{int(round(float(hf_ans_val)))} bpm" if (isinstance(hf_ans_val, (int, float, np.number)) and pd.notna(hf_ans_val) and float(hf_ans_val) > 0) else "-"

    ans_vo2_pct_lauf_str = "-"
    _has_spiro_loaded = (uploaded_file is not None) or ("lauftest_spiro_df" in st.session_state and not st.session_state.lauftest_spiro_df.empty)
    if _has_spiro_loaded and v_ans > 0 and len(vo2_steady_values) > 0 and any(v > 0 for v in vo2_steady_values) and rel_vo2max > 0:
        ans_vo2_rel_top = float(np.interp(v_ans, speed, vo2_steady_values))
        ans_vo2_pct_lauf = (ans_vo2_rel_top / rel_vo2max * 100.0)
        if ans_vo2_pct_lauf > 0:
            ans_vo2_pct_lauf_str = f"{ans_vo2_pct_lauf:.1f} %".replace(".", ",")
    elif rel_vo2max > 0 and v_ans > 0:
        ans_vo2_rel_est = v_ans * 12.6 + 3.5
        ans_vo2_pct_lauf = float(np.clip(ans_vo2_rel_est / rel_vo2max * 100.0, 0.0, 100.0))
        if ans_vo2_pct_lauf > 0:
            ans_vo2_pct_lauf_str = f"{ans_vo2_pct_lauf:.1f} %".replace(".", ",")

    bmi_lauf = (float(weight) / ((float(height) / 100.0) ** 2)) if float(height) > 0 else 0.0
    render_hycys_top_basisdaten(
        weight=float(weight),
        body_fat_pct=float(body_fat_pct),
        rel_vo2max=float(rel_vo2max),
        abs_vo2max=float(abs_vo2max),
        vlamax_str=str(vla_display).replace(".", ","),
        ans_label=f"ANS ({selected_model_name})",
        ans_main_str=f"{ans_pace_str} min/km",
        ans_sub_str=ans_sub_lac_str,
        ans_hr_str=ans_hr_str,
        vo2_pct_str=ans_vo2_pct_lauf_str,
        ltp1_label="LTP1",
        ltp1_main_str=f"{ltp1_pace_str} min/km",
        ltp1_sub_str=ltp1_sub_lac_str,
        ltp1_hr_str=hf_ltp1_str,
        is_presentation=is_lauf_pres,
        bmi_val=bmi_lauf,
        height_cm=float(height),
    )
    if not is_lauf_pres:
        st.caption(f"Kalkulierte Pace @ Fatmax: **{fatmax_pace_str} min/km**")
        if spiro_interval_note:
            st.caption(f"{spiro_interval_note}")

    # 4. Potential Simulation & KH-Match Parameter
    if "kh_match_g_h_lauf" not in st.session_state:
        st.session_state["kh_match_g_h_lauf"] = 90
    _lauf_kh_slider_key = "kh_match_slider_lauf_pres" if is_lauf_pres else "kh_match_slider_lauf_diag"
    kh_intake_lauf = float(st.session_state.get(_lauf_kh_slider_key, st.session_state["kh_match_g_h_lauf"]))
    st.session_state["kh_match_g_h_lauf"] = int(round(kh_intake_lauf))

    vo2_change = float(st.session_state.get("lauf_shared_pot_vo2_pct", 5)) / 100.0
    vla_change = float(st.session_state.get("lauf_shared_pot_vla_pct", -10)) / 100.0
    ans_change = 0.0
    pot_bf_target_lauf = float(st.session_state.get("lauf_shared_pot_bf_target", max(4.0, round(body_fat_pct - 1.5, 1))))
    pot_weight_lauf = calc_pot_weight_from_bf(weight, body_fat_pct, pot_bf_target_lauf, fm_loss_fraction=0.85)

    pot_abs_vo2max_lauf = abs_vo2max * (1.0 + vo2_change)
    pot_vo2max = (pot_abs_vo2max_lauf / pot_weight_lauf) if pot_weight_lauf > 0 else (rel_vo2max * (1.0 + vo2_change))
    pot_vlamax = max(0.15, vla_val * (1.0 + vla_change))

    has_spiro = _has_spiro_loaded and (len(vo2_steady_abs) > 0)
    if has_spiro:
        try:
            vo2_rest = 3.5 * weight
            ex_speeds = []
            ac_stages = []
            passed_thresh_main = False
            for v, vo2, vco2 in zip(speed, vo2_steady_abs, vco2_steady_abs):
                if v <= 0 or vo2 <= 0:
                    continue
                if passed_thresh_main:
                    break
                if v_ans > 0 and v > v_ans + 1e-6:
                    passed_thresh_main = True
                net_vo2 = max(0.0, vo2 - vo2_rest)
                rq = vco2 / vo2 if vco2 > 0 else 0.85
                kal_equiv = (rq - 0.7) * 5.094 + 19.6
                ee_kJ_min = (net_vo2 / 1000.0) * kal_equiv
                vo2_100kh = ee_kJ_min / 21.1 * 1000.0
                ac = vo2_100kh / (v * 100.0)
                ex_speeds.append(float(v))
                ac_stages.append(ac)

            if len(ex_speeds) >= 1:
                v_grid_start = min(2.6, float(ex_speeds[0]))
                v_grid = np.arange(v_grid_start, 8.61, 0.01)
                ac_grid = []
                for v in v_grid:
                    if len(ex_speeds) == 1:
                        ac = ac_stages[0]
                    elif v <= ex_speeds[0]:
                        m = (ac_stages[1] - ac_stages[0]) / (ex_speeds[1] - ex_speeds[0])
                        c_val = ac_stages[0] - m * ex_speeds[0]
                        ac = m * v + c_val
                    elif v >= ex_speeds[-1]:
                        ac = ac_stages[-1]
                    else:
                        ac = ac_stages[-1]
                        for i in range(len(ex_speeds) - 1):
                            if ex_speeds[i] <= v < ex_speeds[i+1]:
                                m = (ac_stages[i+1] - ac_stages[i]) / (ex_speeds[i+1] - ex_speeds[i])
                                c_val = ac_stages[i] - m * ex_speeds[i]
                                ac = m * v + c_val
                                break
                    ac_grid.append(ac)
                ac_grid = np.array(ac_grid)

                # Improve running economy if ANS is overridden
                eco_factor = 1.0 / (1.0 + ans_change)
                pot_ac_grid = ac_grid * eco_factor

                pot_vo2_rel_demand_grid = (pot_ac_grid * v_grid * 100.0) / weight + 1.0
                pot_delta_vo2_grid = pot_vo2max - pot_vo2_rel_demand_grid

                pot_lac_clear_grid = (0.02049 * pot_vo2_rel_demand_grid) / 0.4
                pot_net_clearance = np.full_like(v_grid, -9999.0)
                pot_carb_pct_grid = np.zeros_like(v_grid)

                ks1, ks2 = 0.0631, 1.331
                for idx, (v, vo2_rel, d_vo2) in enumerate(zip(v_grid, pot_vo2_rel_demand_grid, pot_delta_vo2_grid)):
                    if d_vo2 > 0 and vo2_rel > 0:
                        adp = np.sqrt((ks1 * vo2_rel) / d_vo2)
                        adp_3 = adp ** 3
                        pfk = 1 + (ks2 / adp_3)
                        vla_min = pot_vlamax * 60 / pfk
                        pot_net_clearance[idx] = pot_lac_clear_grid[idx] - vla_min
                        fat_pct = max(0.0, min(100.0, 100.0 - (vla_min / pot_lac_clear_grid[idx] * 100.0)))
                        pot_carb_pct_grid[idx] = 100.0 - fat_pct
                    else:
                        pot_net_clearance[idx] = -9999.0
                        pot_carb_pct_grid[idx] = 100.0

                pot_ans_idx = None
                for idx in range(len(v_grid)):
                    if pot_net_clearance[idx] < 0 and idx > 0 and pot_net_clearance[idx-1] >= 0:
                        pot_ans_idx = idx - 1
                        break
                if pot_ans_idx is None:
                    pot_ans_idx = np.argmin(np.abs(pot_net_clearance))

                pot_v_ans = v_grid[pot_ans_idx]

                pot_fatmax_idx = np.argmax(pot_net_clearance)
                pot_v_fatmax = v_grid[pot_fatmax_idx]

                pot_rq_grid = 0.7 + 0.3 * (pot_carb_pct_grid / 100.0)
                pot_kal_equiv_grid = (pot_rq_grid - 0.7) * 5.094 + 19.6
                pot_ee_kcal_h_grid = (((pot_vo2_rel_demand_grid * weight / 1000.0) * pot_kal_equiv_grid) * 60.0) / 4.186
                pot_carb_g_h_grid = (pot_ee_kcal_h_grid * (pot_carb_pct_grid / 100.0)) / 4.063

                pot_match_idx = np.argmin(np.abs(pot_carb_g_h_grid - kh_intake_lauf))
                pot_v_match = v_grid[pot_match_idx]
            else:
                pot_v_ans = v_ans * (1.0 + ans_change)
                pot_v_fatmax = pot_v_ans * 0.78
                pot_v_match = pot_v_ans * 0.85
        except:
            pot_v_ans = v_ans * (1.0 + ans_change)
            pot_v_fatmax = pot_v_ans * 0.78
            pot_v_match = pot_v_ans * 0.85
    else:
        pot_v_ans = v_ans * (1.0 + ans_change)
        pot_v_fatmax = pot_v_ans * 0.78
        pot_v_match = pot_v_ans * 0.85

    st.session_state.pot_weight_val = pot_weight_lauf
    st.session_state.pot_vo2max = pot_vo2max
    st.session_state.pot_vlamax = pot_vlamax
    st.session_state.pot_ans_rel = pot_v_ans
    st.session_state.pot_fatmax = pot_v_fatmax
    st.session_state.pot_match = pot_v_match
    st.session_state.pot_fat = pot_bf_target_lauf

    poly_func = poly_func_no
    fig_laktat = None

    def _solve_mader_v_lauf(rel_vo2_in, vla_in):
        if has_spiro and 'ac_grid' in locals() and 'v_grid' in locals() and len(v_grid) > 10:
            vo2_rel_dem = (ac_grid * v_grid * 100.0) / weight + 1.0
            d_vo2 = rel_vo2_in - vo2_rel_dem
            lac_cl = (0.02049 * vo2_rel_dem) / 0.4
            net_cl = np.full_like(v_grid, -9999.0)
            for idx_m, (vr_m, dv_m) in enumerate(zip(vo2_rel_dem, d_vo2)):
                if dv_m > 0 and vr_m > 0:
                    adp_m = np.sqrt((0.0631 * vr_m) / dv_m)
                    pfk_m = 1.0 + (1.331 / (adp_m ** 3))
                    vla_m = (vla_in * 60.0) / pfk_m
                    net_cl[idx_m] = lac_cl[idx_m] - vla_m
            ans_i = None
            for idx_m in range(len(v_grid)):
                if net_cl[idx_m] < 0 and idx_m > 0 and net_cl[idx_m - 1] >= 0:
                    ans_i = idx_m - 1
                    break
            if ans_i is None:
                ans_i = int(np.argmin(np.abs(net_cl)))
            return float(v_grid[ans_i])
        else:
            vo2_ratio = (rel_vo2_in / rel_vo2max) if rel_vo2max > 0 else 1.0
            vla_diff = (vla_val - vla_in) if vla_val > 0 else 0.0
            return float(v_ans * (vo2_ratio ** 0.85) * (1.0 + 0.28 * vla_diff))

    _base_mader_v = _solve_mader_v_lauf(float(rel_vo2max), float(vla_val))
    _scale_to_vans = (float(v_ans) / _base_mader_v) if _base_mader_v > 0 else 1.0

    def _solve_ans_lauf(vo2_abs_in, vla_in, w_in):
        rel_vo2_motor_only = (vo2_abs_in / weight) if weight > 0 else rel_vo2max
        v_abs_res = _solve_mader_v_lauf(rel_vo2_motor_only, vla_in) * _scale_to_vans
        rel_vo2_with_weight = (vo2_abs_in / w_in) if w_in > 0 else rel_vo2_motor_only
        v_rel_res = _solve_mader_v_lauf(rel_vo2_with_weight, vla_in) * _scale_to_vans
        return float(v_abs_res), float(v_rel_res)

    # SCHWELLEN (direkt unter den Basisdaten, erst danach kommt die Live-Simulation mit dem Wasserfall-Diagramm)
    if not df_res.empty:
        st.markdown("---")
        st.markdown("<h3 style='text-align: center; margin-bottom: 8px;'>Schwellen</h3>", unsafe_allow_html=True)

        _, col_schw_tbl, _ = st.columns([0.15, 0.70, 0.15])
        with col_schw_tbl:
            def _on_main_thresh_change():
                st.session_state.selected_threshold_model = st.session_state.selected_threshold_model_main

            if not is_lauf_pres:
                st.session_state.selected_threshold_model_main = st.session_state.get("selected_threshold_model", "OBLA 4.0 (ANS)")
                st.radio(
                    "Aktive Schwelle (LT2 / ANS) für Auswertung & Trainingsbereiche:",
                    options=["OBLA 4.0 (ANS)", "LTP2"],
                    horizontal=True,
                    key="selected_threshold_model_main",
                    on_change=_on_main_thresh_change
                )
            
            def fmt_2_decimals(x):
                if pd.isna(x) or x == "": return "-"
                if isinstance(x, (int, float, np.number)): return f"{x:.2f}".replace('.', ',')
                try: return f"{float(x):.2f}".replace('.', ',')
                except ValueError: return str(x)

            def fmt_1_decimal(x):
                if pd.isna(x) or x == "": return "-"
                if isinstance(x, (int, float, np.number)): return f"{x:.1f}".replace('.', ',')
                try: return f"{float(x):.1f}".replace('.', ',')
                except ValueError: return str(x)

            def fmt_int(x):
                if pd.isna(x) or x == "": return "-"
                if isinstance(x, (int, float, np.number)): return f"{int(round(float(x)))}"
                try: return f"{int(round(float(x)))}"
                except ValueError: return str(x)

            def fmt_kmh(ms_val):
                if pd.isna(ms_val) or ms_val == "":
                    return "-"
                try:
                    return f"{float(ms_val) * 3.6:.1f}".replace(".", ",")
                except (ValueError, TypeError):
                    return str(ms_val)

            st.markdown("<div style='margin-top: 4px; margin-bottom: 4px; font-weight: 700; font-size: 13px; color: #1a1a1a; text-align: center;'>Anaerobe Schwelle</div>", unsafe_allow_html=True)
            df_for_schwellen_ui = df_res.iloc[[0]] if is_lauf_pres else df_res
            schwellen_rows_fmt = []
            for _, r_s in df_for_schwellen_ui.iterrows():
                row_s_dict = {
                    "Schwelle": r_s["Modell"],
                    "km/h": fmt_kmh(r_s["v"]),
                    "min/km": str(r_s["min/km"]),
                    "Laktat": fmt_1_decimal(r_s["Laktat"]),
                    "HF": fmt_int(r_s["HF"]),
                }
                if not is_lauf_pres:
                    row_s_dict["VLamax [mmol/l/s]"] = fmt_2_decimals(r_s["VLamax [mmol/l/s]"])
                schwellen_rows_fmt.append(row_s_dict)
            df_schwellen_fmt = pd.DataFrame(schwellen_rows_fmt)
            render_centered_html_table(df_schwellen_fmt, highlight_first_row=(not is_lauf_pres))

            if v_ltp1_val is not None:
                st.markdown("<div style='margin-top: 12px; margin-bottom: 4px; font-weight: 700; font-size: 13px; color: #1a1a1a; text-align: center;'>Aerobe Schwelle</div>", unsafe_allow_html=True)
                df_ltp1_only = pd.DataFrame([{
                    "Schwelle": "LTP1",
                    "km/h": fmt_kmh(v_ltp1_val),
                    "min/km": format_pace_mmss(v_ltp1_val),
                    "Laktat": fmt_1_decimal(lac_ltp1_val),
                    "HF": fmt_int(hf_ltp1_val),
                }])
                render_centered_html_table(df_ltp1_only, highlight_first_row=False)

        st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)
        _, col_schw_chart, _ = st.columns([0.08, 0.84, 0.08])
        with col_schw_chart:
            hr_valid_mask = (speed > 0) & pd.notna(hr) & (hr > 0)
            speed_hr_valid = speed[hr_valid_mask]
            hr_valid_vals = hr[hr_valid_mask]
            stage_tickvals = [float(v) for v in speed if float(v) > 0]
            stage_ticktext = [format_pace_mmss(float(v)) for v in stage_tickvals]

            # 1. Im Hintergrund Matplotlib-Graph für den PDF-Export erzeugen (immer mit LTP1, LTP2 und OBLA 4.0)
            fig, ax = plt.subplots(figsize=(8, 4))
            v_smooth = np.linspace(v_start, v_end, 200)
            ax.plot(v_smooth, poly_func(v_smooth), color='#00a1e0', linewidth=2.0, label='Laktatkurve', zorder=2)
            ax.plot(
                speed,
                lactate,
                linestyle='None',
                marker='o',
                color='#1a1a1a',
                markersize=6.0,
                zorder=5,
                label='_nolegend_',
            )
            for xv_pt, ylac_pt in zip(speed, lactate):
                ax.annotate(
                    f"{ylac_pt:.2f}".replace(".", ","),
                    (xv_pt, ylac_pt),
                    textcoords="offset points",
                    xytext=(0, 6),
                    ha="center",
                    va="bottom",
                    fontsize=7.5,
                    color="#1a1a1a",
                    fontweight="bold",
                    bbox=dict(boxstyle="round,pad=0.15", facecolor="white", edgecolor="none", alpha=0.75),
                    zorder=6,
                )
            if v_ltp1_val is not None and lac_ltp1_val is not None:
                ax.plot(
                    v_ltp1_val,
                    lac_ltp1_val,
                    'X',
                    color='#2cb7b9',
                    markeredgecolor='#1d8f91',
                    markeredgewidth=1.4,
                    markersize=9.5,
                    zorder=6,
                    label='LTP1',
                )
            # Reihenfolge in der Legende immer: LTP1 -> LTP2 -> OBLA 4.0 (ohne "(ANS)")
            df_res_legend_order = pd.concat([
                df_res[~df_res['Modell'].astype(str).str.contains('OBLA', na=False)],
                df_res[df_res['Modell'].astype(str).str.contains('OBLA', na=False)],
            ]).reset_index(drop=True)
            for _, r in df_res_legend_order.iterrows():
                if isinstance(r['v'], (int, float, np.number)) and pd.notna(r['v']) and pd.notna(r['Laktat']):
                    is_obla_pdf = 'OBLA' in str(r['Modell'])
                    m_col = '#cdb663' if is_obla_pdf else '#2cb7b9'
                    m_edge = '#9e8836' if is_obla_pdf else '#1d8f91'
                    leg_lbl_pdf = 'OBLA 4.0' if is_obla_pdf else str(r['Modell'])
                    ax.plot(
                        float(r['v']),
                        float(r['Laktat']),
                        'X',
                        color=m_col,
                        markeredgecolor=m_edge,
                        markeredgewidth=1.4,
                        markersize=9.5,
                        zorder=6,
                        label=leg_lbl_pdf,
                    )

            if len(hr_valid_vals) > 0:
                from matplotlib.ticker import MaxNLocator
                ax_hr = ax.twinx()
                ax_hr.plot(
                    speed_hr_valid,
                    hr_valid_vals,
                    linestyle='--',
                    linewidth=1.8,
                    marker='o',
                    color='#d9534f',
                    markerfacecolor='#1a1a1a',
                    markeredgecolor='#1a1a1a',
                    markersize=5.8,
                    zorder=4,
                    label='_nolegend_',
                )
                for xv_pt, yhr_pt in zip(speed_hr_valid, hr_valid_vals):
                    ax_hr.annotate(
                        f"{int(round(float(yhr_pt)))}",
                        (xv_pt, yhr_pt),
                        textcoords="offset points",
                        xytext=(0, 6),
                        ha="center",
                        va="bottom",
                        fontsize=7.5,
                        color="#1a1a1a",
                        fontweight="bold",
                        bbox=dict(boxstyle="round,pad=0.15", facecolor="white", edgecolor="none", alpha=0.75),
                        zorder=6,
                    )
                min_hr_plot = float(np.min(hr_valid_vals))
                max_hr_plot = float(np.max(hr_valid_vals))
                ax_hr.set_ylim(bottom=max(40.0, min_hr_plot - 18.0), top=max_hr_plot + 20.0)
                ax_hr.yaxis.set_major_locator(MaxNLocator(integer=True))
                ax_hr.set_ylabel("Herzfrequenz (S/min)", color="#d9534f")
                ax_hr.tick_params(axis='y', labelcolor="#d9534f")

            max_lac_plot = max(list(lactate) + [float(poly_func(v_end))]) if len(lactate) > 0 else 10.0
            ax.set_ylim(bottom=0.0, top=max_lac_plot * 1.16)
            if len(stage_tickvals) > 0:
                ax.set_xticks(stage_tickvals)
                ax.set_xticklabels(stage_ticktext)
            ax.set_xlabel("Geschwindigkeit (min/km)")
            ax.set_ylabel("Laktat (mmol/L)")
            handles_l, labels_l = ax.get_legend_handles_labels()
            ax.legend(handles_l, labels_l, loc='upper left', fontsize=7.5, framealpha=0.9)
            fig_laktat = fig
            plt.close(fig)

            # 2. In der App-Ansicht (Präsentation & Diagnostik) im modernen Plotly-Format (immer mit LTP1, LTP2 und OBLA 4.0)
            fig_lak_ui = go.Figure()
            fig_lak_ui.add_trace(
                go.Scatter(
                    x=v_smooth,
                    y=poly_func(v_smooth),
                    mode="lines",
                    line=dict(color="#00a1e0", width=2.8),
                    name="Laktatkurve",
                    hovertemplate="Pace: %{x:.2f}<br>Laktat: %{y:.2f} mmol/l<extra></extra>",
                )
            )
            fig_lak_ui.add_trace(
                go.Scatter(
                    x=speed,
                    y=lactate,
                    mode="markers+text",
                    marker=dict(symbol="circle", size=8.5, color="#1a1a1a", line=dict(width=1.2, color="#1a1a1a")),
                    text=[f"{yl:.2f}".replace(".", ",") for yl in lactate],
                    textposition="top center",
                    textfont=dict(size=10, color="#1a1a1a", family="Arial"),
                    customdata=[format_pace_mmss(float(v)) for v in speed],
                    hovertemplate="Stufe: %{customdata} min/km<br>Laktat: %{y:.2f} mmol/l<extra></extra>",
                    name="Messwerte (Laktat)",
                    showlegend=False,
                )
            )
            if len(hr_valid_vals) > 0:
                fig_lak_ui.add_trace(
                    go.Scatter(
                        x=speed_hr_valid,
                        y=hr_valid_vals,
                        mode="lines+markers+text",
                        line=dict(color="#d9534f", width=2.2, dash="dot"),
                        marker=dict(symbol="circle", size=8.5, color="#1a1a1a", line=dict(width=1.2, color="#1a1a1a")),
                        text=[f"{int(round(float(yh)))}" for yh in hr_valid_vals],
                        textposition="top center",
                        textfont=dict(size=10, color="#1a1a1a", family="Arial"),
                        customdata=[format_pace_mmss(float(v)) for v in speed_hr_valid],
                        hovertemplate="Stufe: %{customdata} min/km<br>HF: %{y:.0f} S/min<extra></extra>",
                        name="Herzfrequenz [S/min]",
                        yaxis="y2",
                        showlegend=False,
                    )
                )
            if v_ltp1_val is not None and lac_ltp1_val is not None:
                fig_lak_ui.add_trace(
                    go.Scatter(
                        x=[v_ltp1_val],
                        y=[lac_ltp1_val],
                        mode="markers",
                        marker=dict(symbol="x", size=13, color="#2cb7b9", line=dict(width=2.5, color="#1d8f91")),
                        name="LTP1",
                        customdata=[format_pace_mmss(v_ltp1_val)],
                        hovertemplate="LTP1: %{customdata} min/km<br>Laktat: %{y:.2f} mmol/l<extra></extra>",
                    )
                )
            for _, r in df_res_legend_order.iterrows():
                if isinstance(r["v"], (int, float, np.number)) and pd.notna(r["v"]) and pd.notna(r["Laktat"]):
                    is_obla = "OBLA" in str(r["Modell"])
                    m_color = "#cdb663" if is_obla else "#2cb7b9"
                    m_line_color = "#9e8836" if is_obla else "#1d8f91"
                    leg_lbl_ui = "OBLA 4.0" if is_obla else f"{r['Modell']}"
                    fig_lak_ui.add_trace(
                        go.Scatter(
                            x=[float(r["v"])],
                            y=[float(r["Laktat"])],
                            mode="markers",
                            marker=dict(symbol="x", size=13, color=m_color, line=dict(width=2.5, color=m_line_color)),
                            name=leg_lbl_ui,
                            customdata=[format_pace_mmss(float(r["v"]))],
                            hovertemplate=f"{leg_lbl_ui}: %{{customdata}} min/km<br>Laktat: %{{y:.2f}} mmol/l<extra></extra>",
                        )
                    )
            layout_lak_kwargs = dict(
                xaxis=dict(
                    title="Geschwindigkeit [min/km]",
                    tickmode="array" if len(stage_tickvals) > 0 else "auto",
                    tickvals=stage_tickvals if len(stage_tickvals) > 0 else None,
                    ticktext=stage_ticktext if len(stage_tickvals) > 0 else None,
                    gridcolor="#e6e6e6",
                ),
                yaxis=dict(
                    title="Laktat [mmol/l]",
                    range=[0.0, max_lac_plot * 1.18],
                    tickformat=".0f",
                    gridcolor="#e6e6e6",
                ),
                legend=dict(
                    orientation="v",
                    x=0.015,
                    xanchor="left",
                    y=0.98,
                    yanchor="top",
                    bgcolor="rgba(255,255,255,0.88)",
                    bordercolor="#e0e0e0",
                    borderwidth=1,
                    font=dict(size=9.5),
                ),
                height=460,
                margin=dict(l=55, r=55 if len(hr_valid_vals) > 0 else 30, t=25, b=45),
                plot_bgcolor="#f9f9f9",
            )
            if len(hr_valid_vals) > 0:
                min_hr_ui = float(np.min(hr_valid_vals))
                max_hr_ui = float(np.max(hr_valid_vals))
                layout_lak_kwargs["yaxis2"] = dict(
                    title=dict(text="Herzfrequenz [S/min]", font=dict(color="#d9534f")),
                    tickfont=dict(color="#d9534f"),
                    tickformat=".0f",
                    overlaying="y",
                    side="right",
                    range=[max(40.0, min_hr_ui - 18.0), max_hr_ui + 20.0],
                    showgrid=False,
                )
            fig_lak_ui.update_layout(**layout_lak_kwargs)
            st.plotly_chart(fig_lak_ui, use_container_width=True)

    # Im Präsentations-Modus kommt die Live-Simulation (inkl. Wasserfall-Diagramm) nach den Schwellen
    pot_sim_dict_lauf = None
    if is_lauf_pres:
        pot_sim_dict_lauf = render_ans_coaching_diagrams(
            mode="running",
            weight_ist=float(weight),
            bf_ist=float(body_fat_pct),
            abs_vo2max_ist=float(abs_vo2max),
            rel_vo2max_ist=float(rel_vo2max),
            vlamax_ist=float(vla_val if vla_val > 0 else 0.45),
            ans_abs_ist=float(v_ans),
            ans_rel_ist=float(v_ans),
            solve_ans_fn=_solve_ans_lauf,
            default_vo2_pct=float(vo2_change * 100.0) if vo2_change != 0 else 5.0,
            default_vla_pct=float(vla_change * 100.0) if vla_change != 0 else -10.0,
            default_bf_delta=-1.5,
            key_prefix="lauf",
            is_presentation=True,
        )

    # FETT & KOHLENHYDRAT STOFFWECHSEL & KH-MATCH (im rechten Fenster einstellbar)
    try:
        st.markdown("---")

        v_lo_ener = float(speed[0]) if len(speed) > 0 else 2.6
        v_hi_ener = float(speed[-1]) if len(speed) > 0 else 4.5
        v_grid_ener = np.arange(min(1.8, v_lo_ener), max(6.5, v_hi_ener + 0.5), 0.01)

        if has_spiro and 'ac_stages' in locals() and 'ex_speeds' in locals() and len(ac_stages) >= 1:
            ex_sp_ener = ex_speeds
            ac_grid_ener = []
            for vg in v_grid_ener:
                if len(ex_sp_ener) == 1:
                    ac_grid_ener.append(ac_stages[0])
                elif vg <= ex_sp_ener[0]:
                    ac_grid_ener.append(ac_stages[0])
                elif vg >= ex_sp_ener[-1]:
                    ac_grid_ener.append(ac_stages[-1])
                else:
                    ac_val_e = ac_stages[-1]
                    for idx_e in range(len(ex_sp_ener) - 1):
                        if ex_sp_ener[idx_e] <= vg < ex_sp_ener[idx_e + 1]:
                            m_e = (ac_stages[idx_e + 1] - ac_stages[idx_e]) / (ex_sp_ener[idx_e + 1] - ex_sp_ener[idx_e])
                            c_e = ac_stages[idx_e] - m_e * ex_sp_ener[idx_e]
                            ac_val_e = m_e * vg + c_e
                            break
                    ac_grid_ener.append(ac_val_e)
            ac_grid_ener = np.array(ac_grid_ener)
            vo2_rel_dem_ener = (ac_grid_ener * v_grid_ener * 100.0) / weight + 1.0
        else:
            # Fallback über Standard-Laufökonomie (ca. 210 ml/kg/km = 12.6 ml/min/kg pro m/s)
            vo2_rel_dem_ener = (v_grid_ener * 12.6) + 3.5

        d_vo2_ener = rel_vo2max - vo2_rel_dem_ener if rel_vo2max > 0 else (60.0 - vo2_rel_dem_ener)
        lac_cl_ener = (0.02049 * vo2_rel_dem_ener) / 0.4
        fat_pct_ener = np.zeros_like(v_grid_ener)
        carb_pct_ener = np.zeros_like(v_grid_ener)
        net_cl_ener = np.full_like(v_grid_ener, -9999.0)
        vla_use_ener = float(vla_val if vla_val > 0 else 0.45)

        for idx_e, (vr_e, dv_e) in enumerate(zip(vo2_rel_dem_ener, d_vo2_ener)):
            if dv_e > 0 and vr_e > 0 and lac_cl_ener[idx_e] > 0:
                adp_e = np.sqrt((0.0631 * vr_e) / dv_e)
                pfk_e = 1.0 + (1.331 / (adp_e ** 3))
                vla_min_e = (vla_use_ener * 60.0) / pfk_e
                f_pct = max(0.0, min(100.0, 100.0 - (vla_min_e / lac_cl_ener[idx_e] * 100.0)))
                fat_pct_ener[idx_e] = f_pct
                carb_pct_ener[idx_e] = 100.0 - f_pct
                net_cl_ener[idx_e] = lac_cl_ener[idx_e] - vla_min_e
            else:
                fat_pct_ener[idx_e] = 0.0
                carb_pct_ener[idx_e] = 100.0

        rq_ener = 0.7 + 0.3 * (carb_pct_ener / 100.0)
        kal_eq_ener = (rq_ener - 0.7) * 5.094 + 19.6
        ee_kcal_h_ener = (((vo2_rel_dem_ener * weight / 1000.0) * kal_eq_ener) * 60.0) / 4.186
        ee_fat_kcal_h_ener = ee_kcal_h_ener * (fat_pct_ener / 100.0)
        ee_kh_kcal_h_ener = ee_kcal_h_ener * (carb_pct_ener / 100.0)
        carb_g_h_ener = ee_kh_kcal_h_ener / 4.063
        fat_g_h_ener = ee_fat_kcal_h_ener / 9.3

        # KH-Match für Laufen berechnen (Geschwindigkeit, bei der KH-Verbrauch = KH-Zufuhr)
        submax_mask_khm = d_vo2_ener > 0
        if np.any(submax_mask_khm):
            v_sub_khm = v_grid_ener[submax_mask_khm]
            kh_sub_khm = carb_g_h_ener[submax_mask_khm]
            fat_g_sub_khm = fat_g_h_ener[submax_mask_khm]
            ee_sub_khm = ee_kcal_h_ener[submax_mask_khm]
            le_mask_khm = kh_sub_khm <= kh_intake_lauf
            if np.any(le_mask_khm):
                khm_idx = int(np.where(le_mask_khm)[0][-1])
            else:
                khm_idx = int(np.argmin(np.abs(kh_sub_khm - kh_intake_lauf)))
            v_kh_match_lauf = float(v_sub_khm[khm_idx])
            kh_at_match_lauf = float(kh_sub_khm[khm_idx])
            fat_g_at_match_lauf = float(fat_g_sub_khm[khm_idx])
            ee_at_match_lauf = float(ee_sub_khm[khm_idx])
        else:
            v_kh_match_lauf = float(v_ans * 0.85)
            kh_at_match_lauf = float(kh_intake_lauf)
            fat_g_at_match_lauf = 0.0
            ee_at_match_lauf = 0.0

        st.session_state["v_kh_match_lauf"] = v_kh_match_lauf
        pace_khm_lauf_str = format_pace_mmss(v_kh_match_lauf)
        kmh_khm_lauf_str = f"{v_kh_match_lauf * 3.6:.1f}".replace(".", ",")
        ms_khm_lauf_str = f"{v_kh_match_lauf:.2f}".replace(".", ",")

        hr_mask_khm = (speed > 0) & (hr > 0) & (~np.isnan(speed)) & (~np.isnan(hr))
        if hr_mask_khm.sum() >= 2:
            hr_slope_khm, hr_intercept_khm = np.polyfit(speed[hr_mask_khm], hr[hr_mask_khm], 1)
            hf_khm_lauf_str = f"{int(round(hr_slope_khm * v_kh_match_lauf + hr_intercept_khm, 0))} bpm"
        elif hr_mask_khm.sum() == 1:
            hf_khm_lauf_str = f"{int(round(float(hr[hr_mask_khm][0]), 0))} bpm"
        else:
            hf_khm_lauf_str = "-"

        galibier_lauf = compute_galibier_v2_running(
            v_grid=v_grid_ener,
            vo2_rel_dem_v1=vo2_rel_dem_ener,
            weight=float(weight),
            rel_vo2max=float(rel_vo2max),
            vlamax=float(vla_use_ener),
            carb_intake_g_h=float(kh_intake_lauf),
            km_val=0.21,
            km_low=0.15,
            km_high=0.28,
        )

        # Gemessene Stufenwerte aus den Spiro-Rohdaten des Stufentests für die jeweiligen Geschwindigkeiten sammeln
        spiro_v_meas = []
        spiro_fett_kcal_meas = []
        spiro_kh_g_meas = []
        if has_spiro and len(spiro_fett_steady_kcal_h) > 0:
            for _idx_sm in range(min(len(speed), len(spiro_fett_steady_kcal_h), len(spiro_kh_steady_g_h))):
                _v_sm = float(speed[_idx_sm])
                _f_sm = spiro_fett_steady_kcal_h[_idx_sm]
                _k_sm = spiro_kh_steady_g_h[_idx_sm]
                if _v_sm > 0 and pd.notna(_f_sm) and pd.notna(_k_sm):
                    spiro_v_meas.append(_v_sm)
                    spiro_fett_kcal_meas.append(float(_f_sm))
                    spiro_kh_g_meas.append(float(_k_sm))

        col_stoff_chart, col_stoff_khm = st.columns([1.95, 1.05])
        with col_stoff_chart:
            st.markdown("<h3 style='text-align: center; margin-bottom: 6px;'>Fett & Kohlenhydrat Stoffwechsel</h3>", unsafe_allow_html=True)
            _c_v2_lauf, _c_km_lauf, _c_sp_lauf = st.columns([1.35, 1.05, 1.25])
            with _c_v2_lauf:
                show_v2_galibier_lauf = st.checkbox(
                    "Version 2 - Galibier (AI Diagnostics) einblenden",
                    value=False,
                    key="show_v2_galibier_lauf",
                )
            with _c_km_lauf:
                show_v2_km_band_lauf = st.checkbox(
                    "Km-Range (0,15–0,28) zeigen",
                    value=False,
                    disabled=not show_v2_galibier_lauf,
                    key="show_v2_km_band_lauf",
                )
            with _c_sp_lauf:
                show_spiro_met_curves = False
                if len(spiro_v_meas) >= 2:
                    show_spiro_met_curves = st.checkbox(
                        "Version 3 - Gemessen Spiro einblenden",
                        value=False,
                        key="show_spiro_met_curves_lauf",
                    )

        with col_stoff_khm:
            st.markdown("<h3 style='text-align: center; margin-bottom: 6px;'>KH-Match</h3>", unsafe_allow_html=True)
            kh_sel_lauf = st.slider(
                "KH-Zufuhr (Input) [g KH/h]",
                min_value=10,
                max_value=160,
                value=int(round(kh_intake_lauf)),
                step=1,
                format="%d g/h",
                key=_lauf_kh_slider_key,
            )
            st.session_state["kh_match_g_h_lauf"] = int(kh_sel_lauf)
            st.caption(f"Entspricht **{(kh_sel_lauf / weight):.2f} g/kg KG/h**".replace(".", ","))
            _fat_l_str = f"{fat_g_at_match_lauf:.1f}".replace(".", ",")
            st.markdown(
                f"""
                <div style="border:1px solid #e0e8e8; border-left:5px solid #cdb663; border-radius:6px; padding:14px 16px; background-color:#ffffff; margin-top:6px;">
                    <div style="font-size:12.5px; font-weight:700; color:#1a1a1a; text-transform:uppercase; margin-bottom:10px;">
                        ⚡ KH-Match bei {int(round(kh_intake_lauf))} g KH/h Zufuhr (Version 1 - HYCYS Berechnung (Mader))
                    </div>
                    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; text-align:center;">
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">KH-Match Pace</div>
                            <div style="font-size:20px; font-weight:800; color:#1a1a1a;">{pace_khm_lauf_str} min/km</div>
                            <div style="font-size:11px; color:#595a59;">({kmh_khm_lauf_str} km/h | {ms_khm_lauf_str} m/s)</div>
                        </div>
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">HF @ KH-Match</div>
                            <div style="font-size:20px; font-weight:800; color:#d9534f;">{hf_khm_lauf_str}</div>
                        </div>
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">KH-Verbrauch</div>
                            <div style="font-size:20px; font-weight:800; color:#cdb663;">{int(round(kh_at_match_lauf))} g/h</div>
                            <div style="font-size:11px; color:#595a59;">= Input ({int(round(kh_intake_lauf))} g/h)</div>
                        </div>
                        <div style="background:#f8fafa; padding:10px 8px; border-radius:5px; border:1px solid #edf2f2;">
                            <div style="font-size:11px; color:#595a59; font-weight:600;">Fett & Gesamtenergie</div>
                            <div style="font-size:17px; font-weight:800; color:#2cb7b9;">{_fat_l_str} g/h Fett</div>
                            <div style="font-size:11px; color:#595a59;">({int(round(ee_at_match_lauf))} kcal/h gesamt)</div>
                        </div>
                    </div>
                    <div style="margin-top:10px; font-size:13px; color:#1a1a1a; line-height:1.45;">
                        Bei <b>{int(round(kh_intake_lauf))} g KH/h</b> kannst du <b>{pace_khm_lauf_str} min/km ({kmh_khm_lauf_str} km/h)</b> laufen – dein Kohlenhydratverbrauch ist genauso groß wie dein Input.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if show_v2_galibier_lauf:
                _v2_v_khm = float(galibier_lauf["v_kh_match"])
                _v2_pace_khm = format_pace_mmss(_v2_v_khm)
                _v2_kmh_khm = f"{_v2_v_khm * 3.6:.1f}".replace(".", ",")
                if hr_mask_khm.sum() >= 2:
                    _v2_hf_khm = f"{int(round(hr_slope_khm * _v2_v_khm + hr_intercept_khm, 0))} bpm"
                else:
                    _v2_hf_khm = "-"
                _v2_pace_fm = format_pace_mmss(float(galibier_lauf["v_fatmax"]))
                _v1_pace_fm = format_pace_mmss(float(v_grid_ener[int(np.argmax(net_cl_ener))]))
                _v2_pace_mlss = format_pace_mmss(float(galibier_lauf["v_mlss"]))
                _v2_fat_l_str = f"{galibier_lauf['fat_g_at_match']:.1f}".replace(".", ",")
                st.markdown(
                    f"""
                    <div style="border:1px solid #e0e8e8; border-left:5px solid #1f7a8c; border-radius:6px; padding:10px 14px; background-color:#f8fbfb; margin-top:8px; font-size:12.5px; color:#1a1a1a;">
                        <div style="font-weight:700; text-transform:uppercase; margin-bottom:6px; color:#1f7a8c;">
                            Version 2 - Galibier (AI Diagnostics)
                        </div>
                        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:6px; text-align:center;">
                            <div style="background:#ffffff; padding:6px; border-radius:4px; border:1px solid #e6ecec;">
                                <div style="font-size:10.5px; color:#595a59;">KH-Match (V2)</div>
                                <div style="font-size:15px; font-weight:800;">{_v2_pace_khm} min/km ({_v2_kmh_khm} km/h)</div>
                                <div style="font-size:10.5px; color:#d9534f; font-weight:700;">HF: {_v2_hf_khm} | {_v2_fat_l_str} g/h Fett</div>
                            </div>
                            <div style="background:#ffffff; padding:6px; border-radius:4px; border:1px solid #e6ecec;">
                                <div style="font-size:10.5px; color:#595a59;">Fatmax (V2 vs. V1)</div>
                                <div style="font-size:15px; font-weight:800; color:#1f7a8c;">{_v2_pace_fm} min/km</div>
                                <div style="font-size:10.5px; color:#595a59;">V1: {_v1_pace_fm} min/km</div>
                            </div>
                            <div style="background:#ffffff; padding:6px; border-radius:4px; border:1px solid #e6ecec;">
                                <div style="font-size:10.5px; color:#595a59;">MLSS / ANS (V2 vs. V1)</div>
                                <div style="font-size:15px; font-weight:800;">{_v2_pace_mlss} min/km</div>
                                <div style="font-size:10.5px; color:#595a59;">V1: {ans_pace_str} min/km</div>
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        with col_stoff_chart:
            v_plot_lo = min(v_lo_ener, max(1.8, v_kh_match_lauf - 0.1))
            valid_ener_mask = (d_vo2_ener > 0) & (v_grid_ener >= v_plot_lo - 1e-6) & (v_grid_ener <= v_hi_ener + 1e-6)
            if np.sum(valid_ener_mask) < 5:
                valid_ener_mask = (v_grid_ener >= v_plot_lo - 1e-6) & (v_grid_ener <= v_hi_ener + 1e-6)

            v_vis_e = v_grid_ener[valid_ener_mask]
            ee_vis_e = ee_kcal_h_ener[valid_ener_mask]
            fat_kcal_vis_e = ee_fat_kcal_h_ener[valid_ener_mask]
            fat_g_vis_e = fat_g_h_ener[valid_ener_mask]
            kh_kcal_vis_e = ee_kh_kcal_h_ener[valid_ener_mask]
            kh_g_vis_e = carb_g_h_ener[valid_ener_mask]
            paces_vis_e = [format_pace_mmss(vv) for vv in v_vis_e]
            kmh_vis_e = [vv * 3.6 for vv in v_vis_e]

            fatmax_idx_e = int(np.argmax(net_cl_ener))
            v_fatmax_e = float(v_grid_ener[fatmax_idx_e])

            custom_hover_e = np.column_stack([
                paces_vis_e,
                [f"{k:.1f}".replace(".", ",") for k in kmh_vis_e],
                [f"{int(round(val))}" for val in ee_vis_e],
                [f"{val:.1f}".replace(".", ",") for val in fat_g_vis_e],
                [f"{int(round(val))}" for val in fat_kcal_vis_e],
                [f"{val:.1f}".replace(".", ",") for val in kh_g_vis_e],
                [f"{int(round(val))}" for val in kh_kcal_vis_e],
            ])

            fig_stoff = go.Figure()

            # Kurve 1: Gesamtenergieverbrauch [kcal/h] (Linke Y-Achse)
            fig_stoff.add_trace(
                go.Scatter(
                    x=v_vis_e,
                    y=ee_vis_e,
                    mode="lines",
                    line=dict(color="#595a59", width=3),
                    name="Gesamtenergieverbrauch [kcal/h]",
                    customdata=custom_hover_e,
                    hovertemplate=(
                        "<b>Pace: %{customdata[0]} min/km (%{customdata[1]} km/h)</b><br>"
                        "Gesamtenergie: <b>%{customdata[2]} kcal/h</b><br>"
                        "Fett (Version 1 - HYCYS Berechnung (Mader)): <b>%{customdata[3]} g/h Fett</b> (%{customdata[4]} kcal/h)<br>"
                        "KH (Version 1 - HYCYS Berechnung (Mader)): <b>%{customdata[5]} g/h KH</b> (%{customdata[6]} kcal/h)"
                        "<extra></extra>"
                    ),
                    yaxis="y",
                )
            )

            # Kurve 2: Fettverbrennung [kcal/h] (Linke Y-Achse)
            fig_stoff.add_trace(
                go.Scatter(
                    x=v_vis_e,
                    y=fat_kcal_vis_e,
                    mode="lines",
                    line=dict(color="#2cb7b9", width=3),
                    name="Fett: Version 1 - HYCYS Berechnung (Mader) [kcal/h]",
                    hoverinfo="skip",
                    yaxis="y",
                )
            )

            # Kurve 3: Kohlenhydratverbrauch [g/h] (Rechte Y-Achse)
            fig_stoff.add_trace(
                go.Scatter(
                    x=v_vis_e,
                    y=kh_g_vis_e,
                    mode="lines",
                    line=dict(color="#cdb663", width=3),
                    name="KH: Version 1 - HYCYS Berechnung (Mader) [g/h]",
                    hoverinfo="skip",
                    yaxis="y2",
                )
            )

            # Optionale Kurven: Version 2 - Galibier (AI Diagnostics)
            if show_v2_galibier_lauf:
                _v2_fat_vis = galibier_lauf["ee_fat"][valid_ener_mask]
                _v2_fat_g_vis = galibier_lauf["fat_g"][valid_ener_mask]
                _v2_fat_pct_vis = galibier_lauf["fat_pct"][valid_ener_mask]
                _v2_kh_g_vis = galibier_lauf["kh_g"][valid_ener_mask]
                _v2_kh_kcal_vis = galibier_lauf["kh_kcal"][valid_ener_mask]
                _v2_kh_pct_vis = galibier_lauf["kh_pct"][valid_ener_mask]

                if show_v2_km_band_lauf:
                    _v2_fat_lo_vis = galibier_lauf["ee_fat_low"][valid_ener_mask]
                    _v2_fat_hi_vis = galibier_lauf["ee_fat_high"][valid_ener_mask]
                    _v2_kh_lo_vis = galibier_lauf["kh_g_low"][valid_ener_mask]
                    _v2_kh_hi_vis = galibier_lauf["kh_g_high"][valid_ener_mask]
                    fig_stoff.add_trace(
                        go.Scatter(
                            x=np.concatenate([v_vis_e, v_vis_e[::-1]]),
                            y=np.concatenate([_v2_fat_hi_vis, _v2_fat_lo_vis[::-1]]),
                            fill="toself",
                            fillcolor="rgba(31, 122, 140, 0.13)",
                            line=dict(color="rgba(255,255,255,0)"),
                            hoverinfo="skip",
                            name="Fett Spannweite: Version 2 - Galibier (AI Diagnostics)",
                            yaxis="y",
                        )
                    )
                    fig_stoff.add_trace(
                        go.Scatter(
                            x=np.concatenate([v_vis_e, v_vis_e[::-1]]),
                            y=np.concatenate([_v2_kh_hi_vis, _v2_kh_lo_vis[::-1]]),
                            fill="toself",
                            fillcolor="rgba(184, 134, 11, 0.13)",
                            line=dict(color="rgba(255,255,255,0)"),
                            hoverinfo="skip",
                            name="KH Spannweite: Version 2 - Galibier (AI Diagnostics)",
                            yaxis="y2",
                        )
                    )

                _v2_lauf_hover = np.column_stack([
                    [f"{v:.1f}".replace(".", ",") for v in _v2_fat_g_vis],
                    [f"{int(round(v))}" for v in _v2_fat_vis],
                    [f"{v:.0f}" for v in _v2_fat_pct_vis],
                    [f"{v:.1f}".replace(".", ",") for v in _v2_kh_g_vis],
                    [f"{int(round(v))}" for v in _v2_kh_kcal_vis],
                    [f"{v:.0f}" for v in _v2_kh_pct_vis],
                ])
                fig_stoff.add_trace(
                    go.Scatter(
                        x=v_vis_e,
                        y=_v2_fat_vis,
                        mode="lines",
                        line=dict(color="#0d5c75", width=3, dash="dot"),
                        name="Fett: Version 2 - Galibier (AI Diagnostics) [kcal/h]",
                        customdata=_v2_lauf_hover,
                        hovertemplate="Fett (Version 2 - Galibier (AI Diagnostics)): <b>%{customdata[0]} g/h Fett</b> (%{customdata[1]} kcal/h · %{customdata[2]} %)<extra></extra>",
                        yaxis="y",
                    )
                )
                fig_stoff.add_trace(
                    go.Scatter(
                        x=v_vis_e,
                        y=_v2_kh_g_vis,
                        mode="lines",
                        line=dict(color="#b8860b", width=3, dash="dot"),
                        name="KH: Version 2 - Galibier (AI Diagnostics) [g/h]",
                        customdata=_v2_lauf_hover,
                        hovertemplate="KH (Version 2 - Galibier (AI Diagnostics)): <b>%{customdata[3]} g/h KH</b> (%{customdata[4]} kcal/h · %{customdata[5]} %)<extra></extra>",
                        yaxis="y2",
                    )
                )

            # Optionale Vergleichskurven: Version 3 - Gemessen Spiro
            if show_spiro_met_curves and len(spiro_v_meas) >= 2:
                from scipy.interpolate import PchipInterpolator
                _sp_order = np.argsort(spiro_v_meas)
                _sp_v_arr = np.array(spiro_v_meas, dtype=float)[_sp_order]
                _sp_f_arr = np.array(spiro_fett_kcal_meas, dtype=float)[_sp_order]
                _sp_k_arr = np.array(spiro_kh_g_meas, dtype=float)[_sp_order]
                _sp_v_unique, _sp_u_idx = np.unique(_sp_v_arr, return_index=True)
                _sp_f_unique = _sp_f_arr[_sp_u_idx]
                _sp_k_unique = _sp_k_arr[_sp_u_idx]

                if len(_sp_v_unique) >= 2:
                    _sp_grid_mask = (v_vis_e >= float(_sp_v_unique[0]) - 1e-6) & (v_vis_e <= float(_sp_v_unique[-1]) + 1e-6)
                    _sp_v_smooth = v_vis_e[_sp_grid_mask] if np.sum(_sp_grid_mask) >= 5 else np.linspace(float(_sp_v_unique[0]), float(_sp_v_unique[-1]), 100)
                    _pchip_fett = PchipInterpolator(_sp_v_unique, _sp_f_unique)
                    _pchip_kh = PchipInterpolator(_sp_v_unique, _sp_k_unique)
                    _sp_f_smooth = np.clip(_pchip_fett(_sp_v_smooth), 0.0, None)
                    _sp_k_smooth = np.clip(_pchip_kh(_sp_v_smooth), 0.0, None)

                    _sp_custom_smooth = np.column_stack([
                        [f"{val / 9.3:.1f}".replace(".", ",") for val in _sp_f_smooth],
                        [f"{int(round(val))}" for val in _sp_f_smooth],
                        [f"{val:.1f}".replace(".", ",") for val in _sp_k_smooth],
                        [f"{int(round(val * 4.063))}" for val in _sp_k_smooth],
                    ])
                    _sp_custom_pts = np.column_stack([
                        [format_pace_mmss(vv) for vv in _sp_v_unique],
                        [f"{vv * 3.6:.1f}".replace(".", ",") for vv in _sp_v_unique],
                        [f"{val / 9.3:.1f}".replace(".", ",") for val in _sp_f_unique],
                        [f"{int(round(val))}" for val in _sp_f_unique],
                        [f"{val:.1f}".replace(".", ",") for val in _sp_k_unique],
                        [f"{int(round(val * 4.063))}" for val in _sp_k_unique],
                    ])

                    # Kurve 4: Gemessener Fettstoffwechsel aus Spiro [kcal/h] (Linke Y-Achse)
                    fig_stoff.add_trace(
                        go.Scatter(
                            x=_sp_v_smooth,
                            y=_sp_f_smooth,
                            mode="lines",
                            line=dict(color="#11696b", width=3, dash="dash"),
                            name="Fett: Version 3 - Gemessen Spiro [kcal/h]",
                            customdata=_sp_custom_smooth,
                            hovertemplate="Fett (Version 3 - Gemessen Spiro): <b>%{customdata[0]} g/h Fett</b> (%{customdata[1]} kcal/h)<extra></extra>",
                            yaxis="y",
                        )
                    )
                    fig_stoff.add_trace(
                        go.Scatter(
                            x=_sp_v_unique,
                            y=_sp_f_unique,
                            mode="markers",
                            marker=dict(symbol="circle", size=8, color="#11696b", line=dict(width=1.5, color="#ffffff")),
                            name="Fett Stufenwert (Version 3 - Gemessen Spiro)",
                            customdata=_sp_custom_pts,
                            hovertemplate="Stufe %{customdata[0]} min/km (%{customdata[1]} km/h)<br>Fett (Version 3 - Gemessen Spiro): <b>%{customdata[2]} g/h</b> (%{customdata[3]} kcal/h)<extra></extra>",
                            showlegend=False,
                            yaxis="y",
                        )
                    )

                    # Kurve 5: Gemessener Kohlenhydratverbrauch aus Spiro [g/h] (Rechte Y-Achse)
                    fig_stoff.add_trace(
                        go.Scatter(
                            x=_sp_v_smooth,
                            y=_sp_k_smooth,
                            mode="lines",
                            line=dict(color="#8c7322", width=3, dash="dash"),
                            name="KH: Version 3 - Gemessen Spiro [g/h]",
                            customdata=_sp_custom_smooth,
                            hovertemplate="KH (Version 3 - Gemessen Spiro): <b>%{customdata[2]} g/h KH</b> (%{customdata[3]} kcal/h)<extra></extra>",
                            yaxis="y2",
                        )
                    )
                    fig_stoff.add_trace(
                        go.Scatter(
                            x=_sp_v_unique,
                            y=_sp_k_unique,
                            mode="markers",
                            marker=dict(symbol="circle", size=8, color="#8c7322", line=dict(width=1.5, color="#ffffff")),
                            name="KH Stufenwert (Version 3 - Gemessen Spiro)",
                            customdata=_sp_custom_pts,
                            hovertemplate="Stufe %{customdata[0]} min/km (%{customdata[1]} km/h)<br>KH (Version 3 - Gemessen Spiro): <b>%{customdata[4]} g/h</b> (%{customdata[5]} kcal/h)<extra></extra>",
                            showlegend=False,
                            yaxis="y2",
                        )
                    )

            if float(v_vis_e[0]) <= v_fatmax_e <= float(v_vis_e[-1]):
                _v_span_e = max(1.0, float(v_vis_e[-1]) - float(v_vis_e[0]))
                _fm_half_v = max(0.115, _v_span_e * 0.055)
                _fm_x0_l = max(float(v_vis_e[0]), v_fatmax_e - _fm_half_v)
                _fm_x1_l = min(float(v_vis_e[-1]), v_fatmax_e + _fm_half_v)
                fig_stoff.add_vrect(
                    x0=_fm_x0_l,
                    x1=_fm_x1_l,
                    fillcolor="rgba(44, 183, 185, 0.20)",
                    line_width=0,
                    layer="below",
                )
                fig_stoff.add_annotation(
                    x=0.5 * (_fm_x0_l + _fm_x1_l),
                    y=0.92,
                    xref="x",
                    yref="paper",
                    text="<b>Fatmax</b>",
                    showarrow=False,
                    xanchor="center",
                    yanchor="top",
                    font=dict(size=14, color="#1d8f91", family="Arial"),
                )

            if float(v_vis_e[0]) <= v_kh_match_lauf <= float(v_vis_e[-1]):
                fig_stoff.add_vline(
                    x=v_kh_match_lauf,
                    line_dash="dot",
                    line_color="#cdb663",
                    line_width=2.5,
                )
                fig_stoff.add_annotation(
                    x=v_kh_match_lauf,
                    y=0.10,
                    xref="x",
                    yref="paper",
                    text="<b>KH-Match</b>",
                    showarrow=False,
                    xanchor="left",
                    yanchor="bottom",
                    xshift=6,
                    font=dict(size=14, color="#9e8836", family="Arial"),
                )

            if float(v_vis_e[0]) <= v_ans <= float(v_vis_e[-1]):
                fig_stoff.add_vline(
                    x=v_ans,
                    line_dash="dash",
                    line_color="#595a59",
                )
                fig_stoff.add_annotation(
                    x=v_ans,
                    y=0.92,
                    xref="x",
                    yref="paper",
                    text="<b>ANS</b>",
                    showarrow=False,
                    xanchor="left",
                    yanchor="top",
                    xshift=6,
                    font=dict(size=13, color="#595a59", family="Arial"),
                )

            # X-Achse mit Pace (min/km) Ticks beschriften, ohne km/h
            tick_v_vals = np.linspace(float(v_vis_e[0]), float(v_vis_e[-1]), 6)
            tick_v_texts = [f"{format_pace_mmss(tv)} min/km" for tv in tick_v_vals]

            _max_ee_base = float(np.max(ee_vis_e)) if len(ee_vis_e) > 0 else 1000.0
            _max_kh_base = float(np.max(kh_g_vis_e)) if len(kh_g_vis_e) > 0 else 250.0
            if show_spiro_met_curves and len(spiro_fett_kcal_meas) > 0 and len(spiro_kh_g_meas) > 0:
                _max_ee_base = max(_max_ee_base, float(np.max(spiro_fett_kcal_meas)))
                _max_kh_base = max(_max_kh_base, float(np.max(spiro_kh_g_meas)))
            max_ee_y = _max_ee_base * 1.12
            max_kh_y = _max_kh_base * 1.12

            fig_stoff.update_layout(
                xaxis=dict(
                    title="Pace [min/km]",
                    range=[float(v_vis_e[0]), float(v_vis_e[-1])],
                    tickvals=list(tick_v_vals),
                    ticktext=tick_v_texts,
                    gridcolor="#e6e6e6",
                ),
                yaxis=dict(
                    title=dict(text="Energie [kcal/h]", font=dict(color="#595a59")),
                    tickfont=dict(color="#595a59"),
                    tickformat=".0f",
                    range=[0.0, max_ee_y],
                    gridcolor="#e6e6e6",
                ),
                yaxis2=dict(
                    title=dict(text="Kohlenhydratverbrauch [g/h]", font=dict(color="#cdb663")),
                    tickfont=dict(color="#cdb663"),
                    tickformat=".0f",
                    range=[0.0, max_kh_y],
                    overlaying="y",
                    side="right",
                    showgrid=False,
                ),
                hovermode="x",
                legend=dict(
                    orientation="h",
                    x=0.5,
                    xanchor="center",
                    y=1.08,
                    bgcolor="rgba(255,255,255,0.85)",
                ),
                height=530,
                margin=dict(l=55, r=55, t=55, b=50),
                plot_bgcolor="#f9f9f9",
            )
            st.plotly_chart(fig_stoff, use_container_width=True)
    except Exception:
        pass

    # LAUFÖKONOMIE SECTION (über den Trainingsbereichen; standardmäßig nur RE [l/km], ohne m/s)
    if not spiro_df.empty and len(vo2_steady_abs) > 0 and any(v > 0 for v in vo2_steady_abs):
        st.markdown("---")
        st.subheader("Laufökonomie")

        n_stages = min(anzahl, len(vo2_steady_abs))
        lo_hycys_list = calculate_laufokonomie_hycys_stages(
            speeds_arr   = list(speed[:n_stages]),
            vo2_abs_arr  = list(vo2_steady_abs[:n_stages]),
            vco2_abs_arr = list(vco2_steady_abs[:n_stages]) if len(vco2_steady_abs) >= n_stages else [0.0]*n_stages,
            weight       = weight,
            v_thresh     = v_ans if v_ans > 0 else None,
        )

        ffm_kg_re = weight * (1.0 - float(body_fat_pct) / 100.0) if weight > 0 else 0.0
        lauf_rows = []
        for i in range(anzahl):
            v_ms   = speed[i]
            v_kmh  = round(v_ms * 3.6, 1)
            vo2_a  = vo2_steady_abs[i]  if i < len(vo2_steady_abs)  else 0.0
            vco2_a = vco2_steady_abs[i] if i < len(vco2_steady_abs) else 0.0
            vo2_r  = vo2_steady_values[i]   # ml/min/kg

            re_ml_min_kg_ms = lo_hycys_list[i] if i < len(lo_hycys_list) else None
            re_ml_kg_km = (re_ml_min_kg_ms * 1000.0 / 60.0) if re_ml_min_kg_ms is not None else None
            re_lit_ml_kg_km = (vo2_r / (v_ms * 3.6) * 60.0) if (vo2_r > 0 and v_ms > 0) else None
            re_l_km = (re_ml_kg_km * weight / 1000.0) if (re_ml_kg_km is not None and weight > 0) else None
            re_ml_km = (re_ml_kg_km * weight) if (re_ml_kg_km is not None and weight > 0) else None
            re_ml_ffm_km = (re_ml_km / ffm_kg_re) if (re_ml_km is not None and ffm_kg_re > 0) else None
            re_kcal_km = (re_l_km * 21.1 / 4.186) if re_l_km is not None else None
            rq_val = (vco2_a / vo2_a) if (vco2_a > 0 and vo2_a > 0) else None

            vla_stage = None
            if rel_vo2max > 0 and vo2_r > 0 and len(vo2_steady_values) > 1:
                vla_stage, _, _ = calculate_vlamax_for_v(
                    v_ms, rel_vo2max, speed, vo2_steady_abs, vco2_steady_abs, weight
                )

            pace_sec = (1000.0 / v_ms) if v_ms > 0 else 0
            pace_str = f"{int(pace_sec // 60)}:{int(pace_sec % 60):02d}"

            lauf_rows.append({
                'stage_idx':       i + 1,
                'v_ms':            v_ms,
                'v_kmh':           v_kmh,
                'pace':            pace_str,
                'vo2_abs':         vo2_a,
                'vo2_rel':         vo2_r,
                'rq':              rq_val,
                're_ml_kg_km':     re_ml_kg_km,
                're_lit_ml_kg_km': re_lit_ml_kg_km,
                're_l_km':         re_l_km,
                're_ml_km':        re_ml_km,
                're_ml_ffm_km':    re_ml_ffm_km,
                're_kcal_km':      re_kcal_km,
                're_ml_min_kg_ms': re_ml_min_kg_ms,
                'vla_stage':       vla_stage
            })

        valid_rows = [
            r for r in lauf_rows
            if r['re_l_km'] is not None
        ]

        if len(valid_rows) >= 1:
            re_pot_pct = int(st.session_state.get("lauf_re_pot_slider", -5))
            re_pot_factor = 1.0 + (re_pot_pct / 100.0)
            col_lt, col_lc = st.columns([1.15, 1.25])

            with col_lt:
                col_ms_narrow, _ = st.columns([1.0, 1.0])
                with col_ms_narrow:
                    optional_re_units = st.multiselect(
                        "Weitere Einheiten einblenden",
                        options=[
                            "ml/min/kg/(m/s)",
                            "L/km",
                        ],
                        default=[],
                        key="lauf_re_optional_units_v4",
                    )

                st.markdown("**Messwerte je Stufe**")
                display_records = []
                for r in valid_rows:
                    re_ffm_v = r['re_ml_ffm_km']
                    row_dict = {
                        'Pace [min/km]': r['pace'],
                        'ml/FFM/km':     f"{re_ffm_v:.1f}".replace('.', ',') if re_ffm_v is not None else '-',
                    }
                    if "ml/min/kg/(m/s)" in optional_re_units:
                        row_dict['ml/min/kg/(m/s)'] = f"{r['re_ml_min_kg_ms']:.2f}".replace('.', ',') if r['re_ml_min_kg_ms'] is not None else '-'
                    if "L/km" in optional_re_units:
                        row_dict['L/km'] = f"{r['re_l_km']:.2f}".replace('.', ',') if r['re_l_km'] is not None else '-'
                    display_records.append(row_dict)

                df_lauf_display = pd.DataFrame(display_records)
                render_centered_html_table(df_lauf_display)
                if not is_lauf_pres:
                    st.caption("ml/FFM/km = Sauerstoffverbrauch pro Kilometer bezogen auf die fettfreie Masse (HYCYS netto, RQ-normiert). Je niedriger der Wert, desto ökonomischer der Laufstil.")

                # Praxis-Übersetzung des Laufökonomie-Potenzials (3 anschauliche Coaching-Karten)
                avg_re_ist = float(np.mean([r['re_l_km'] for r in valid_rows]))
                avg_re_pot = avg_re_ist * re_pot_factor
                avg_o2_saved_l_km = max(0.0, avg_re_ist - avg_re_pot)
                kcal_saved_marathon = avg_o2_saved_l_km * (21.1 / 4.186) * 42.195
                kh_saved_marathon_g = kcal_saved_marathon / 4.063
                v_ans_re_pot = (v_ans / re_pot_factor) if re_pot_factor > 0 else v_ans
                pace_ans_ist_s = (1000.0 / v_ans) if v_ans > 0 else 0.0
                pace_ans_pot_s = (1000.0 / v_ans_re_pot) if v_ans_re_pot > 0 else 0.0
                sec_gained_per_km = max(0.0, pace_ans_ist_s - pace_ans_pot_s)
                sec_gain_str = f"{sec_gained_per_km:.0f}"

                # Marathon-Berechnung analog zu HYCYS Excel (Calculation Energetik & Calculation Physiologie)
                v_mara_ist = float(v_ans * 0.88) if v_ans > 0 else 3.33
                try:
                    if 'ex_speeds' in locals() and 'ac_stages' in locals() and len(ex_speeds) >= 1 and rel_vo2max > 0:
                        v_grid_m = np.arange(min(2.6, float(ex_speeds[0])), 8.61, 0.01)
                        ac_grid_m = []
                        for vm_pt in v_grid_m:
                            if len(ex_speeds) == 1:
                                ac_m = ac_stages[0]
                            elif vm_pt <= ex_speeds[0]:
                                m_sl = (ac_stages[1] - ac_stages[0]) / (ex_speeds[1] - ex_speeds[0])
                                ac_m = m_sl * vm_pt + (ac_stages[0] - m_sl * ex_speeds[0])
                            elif vm_pt >= ex_speeds[-1]:
                                ac_m = ac_stages[-1]
                            else:
                                ac_m = ac_stages[-1]
                                for idx_s in range(len(ex_speeds) - 1):
                                    if ex_speeds[idx_s] <= vm_pt < ex_speeds[idx_s + 1]:
                                        m_sl = (ac_stages[idx_s + 1] - ac_stages[idx_s]) / (ex_speeds[idx_s + 1] - ex_speeds[idx_s])
                                        ac_m = m_sl * vm_pt + (ac_stages[idx_s] - m_sl * ex_speeds[idx_s])
                                        break
                            ac_grid_m.append(ac_m)
                        ac_grid_m = np.array(ac_grid_m)
                        vo2_rel_m = (ac_grid_m * v_grid_m * 100.0) / weight + 1.0
                        d_vo2_m = rel_vo2max - vo2_rel_m
                        lac_clear_m = (0.02049 * vo2_rel_m) / 0.4
                        carb_pct_m = np.zeros_like(v_grid_m)
                        vla_use_m = float(vla_val if vla_val > 0 else 0.45)
                        for idx_m, (vr_m, dv_m) in enumerate(zip(vo2_rel_m, d_vo2_m)):
                            if dv_m > 0 and vr_m > 0 and lac_clear_m[idx_m] > 0:
                                adp_m = np.sqrt((0.0631 * vr_m) / dv_m)
                                pfk_m = 1.0 + (1.331 / (adp_m ** 3))
                                vla_min_m = (vla_use_m * 60.0) / pfk_m
                                fat_p_m = max(0.0, min(100.0, 100.0 - (vla_min_m / lac_clear_m[idx_m] * 100.0)))
                                carb_pct_m[idx_m] = 100.0 - fat_p_m
                            else:
                                carb_pct_m[idx_m] = 100.0
                        rq_m = 0.7 + 0.3 * (carb_pct_m / 100.0)
                        kal_eq_m = (rq_m - 0.7) * 5.094 + 19.6
                        ee_kcal_h_m = (((vo2_rel_m * weight / 1000.0) * kal_eq_m) * 60.0) / 4.186
                        ee_kh_kcal_h_m = ee_kcal_h_m * (carb_pct_m / 100.0)

                        # HYCYS Standard: 45 g/h KH-Zufuhr & 250 g Glykogenmenge für Marathon-Defizitpunkt
                        carb_intake_mara = 45.0
                        glycogen_mara = 250.0
                        v_mara_found = float(v_grid_m[-1])
                        for idx_m, vm_pt in enumerate(v_grid_m):
                            t_mara_min = 42195.0 / (vm_pt * 60.0)
                            carb_req_m = (ee_kh_kcal_h_m[idx_m] / 60.0) * t_mara_min
                            carb_avail_m = (carb_intake_mara / 60.0 * t_mara_min * 4.063) + (glycogen_mara * 4.063)
                            if (carb_req_m - carb_avail_m) > 0 or d_vo2_m[idx_m] <= 0:
                                v_mara_found = float(v_grid_m[max(0, idx_m - 1)])
                                break
                        v_mara_ist = v_mara_found
                except Exception:
                    pass

                v_mara_pot = (v_mara_ist / re_pot_factor) if re_pot_factor > 0 else v_mara_ist

                def _fmt_mara_hms_ui(v_ms_val):
                    if v_ms_val is None or v_ms_val <= 0:
                        return "-"
                    t_min_val = 42195.0 / (v_ms_val * 60.0)
                    h_v = int(t_min_val // 60)
                    m_v = int(t_min_val % 60)
                    s_v = int(round((t_min_val * 60) % 60))
                    if s_v == 60:
                        m_v += 1
                        s_v = 0
                    if m_v == 60:
                        h_v += 1
                        m_v = 0
                    return f"{h_v:02d}:{m_v:02d}:{s_v:02d}"

                mara_pace_ist_str = format_pace_mmss(v_mara_ist)
                mara_pace_pot_str = format_pace_mmss(v_mara_pot)
                mara_time_ist_str = _fmt_mara_hms_ui(v_mara_ist)
                mara_time_pot_str = _fmt_mara_hms_ui(v_mara_pot)

                t_mara_ist_min = (42195.0 / (v_mara_ist * 60.0)) if v_mara_ist > 0 else 0.0
                t_mara_pot_min = (42195.0 / (v_mara_pot * 60.0)) if v_mara_pot > 0 else 0.0
                delta_mara_min = max(0.0, t_mara_ist_min - t_mara_pot_min)
                dm_ui = int(delta_mara_min)
                ds_ui = int(round((delta_mara_min - dm_ui) * 60.0))
                if ds_ui == 60:
                    dm_ui += 1
                    ds_ui = 0
                mara_delta_ui_str = f"-{dm_ui:02d}:{ds_ui:02d} min"
                if re_pot_pct != 0:
                    mara_sub_str = f"statt {mara_pace_ist_str} min/km ({mara_time_ist_str} h | {mara_delta_ui_str})"
                else:
                    mara_sub_str = f"Pace: {mara_pace_ist_str} min/km ({v_mara_ist * 3.6:.1f} km/h)".replace(".", ",")

                st.markdown(
                    f"""
                    <div style="margin-top: 10px; background: #f8fafa; border: 1px solid #d5ecec; border-left: 4px solid #cdb663; border-radius: 6px; padding: 10px 14px;">
                        <div style="font-size: 12px; font-weight: 700; color: #1a1a1a; text-transform: uppercase; margin-bottom: 6px;">
                            Praxis-Effekt bei {abs(re_pot_pct)} % besserer Laufökonomie
                        </div>
                        <div style="display: grid; grid-template-columns: 1.2fr 1fr 1fr; gap: 8px; text-align: center;">
                            <div style="background: #ffffff; border: 1px solid #e6e6e6; border-radius: 5px; padding: 6px;">
                                <div style="font-size: 11px; color: #595a59;">Marathon-Prognose</div>
                                <div style="font-size: 15px; font-weight: 800; color: #2cb7b9;">{mara_pace_pot_str} min/km ({mara_time_pot_str} h)</div>
                                <div style="font-size: 10.5px; color: #777;">{mara_sub_str}</div>
                            </div>
                            <div style="background: #ffffff; border: 1px solid #e6e6e6; border-radius: 5px; padding: 6px;">
                                <div style="font-size: 11px; color: #595a59;">Ersparnis Marathon</div>
                                <div style="font-size: 15px; font-weight: 800; color: #cdb663;">-{int(round(kcal_saved_marathon))} kcal</div>
                                <div style="font-size: 10.5px; color: #777;">≈ {int(round(kh_saved_marathon_g))} g Kohlenhydrate</div>
                            </div>
                            <div style="background: #ffffff; border: 1px solid #e6e6e6; border-radius: 5px; padding: 6px;">
                                <div style="font-size: 11px; color: #595a59;">ANS-Pace Gewinn</div>
                                <div style="font-size: 15px; font-weight: 800; color: #1b8a5a;">-{sec_gain_str} Sek./km</div>
                                <div style="font-size: 10.5px; color: #777;">{format_pace_mmss(v_ans_re_pot)} statt {ans_pace_str} min/km</div>
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with col_lc:
                st.markdown("<div style='margin-top: 76px;'></div>", unsafe_allow_html=True)
                re_pot_pct = int(
                    st.slider(
                        "Potenzial Laufökonomie-Verbesserung",
                        min_value=-15,
                        max_value=0,
                        value=int(re_pot_pct),
                        step=1,
                        format="%d %%",
                        key="lauf_re_pot_slider",
                    )
                )
                re_pot_factor = 1.0 + (re_pot_pct / 100.0)
                x_labels = [
                    f"<b>{r['pace']} min/km</b>"
                    for r in valid_rows
                ]
                re_ist_vals = [float(r['re_l_km']) for r in valid_rows]
                re_pot_vals = [float(r['re_l_km']) * re_pot_factor for r in valid_rows]
                re_diff_vals = [pot_v - ist_v for ist_v, pot_v in zip(re_ist_vals, re_pot_vals)]

                fig_lauf = go.Figure()
                fig_lauf.add_trace(
                    go.Bar(
                        x=x_labels,
                        y=re_ist_vals,
                        name="Ist-Verbrauch [L/km]",
                        marker=dict(color="#2cb7b9", line=dict(color="#1d8f91", width=1)),
                        text=[f"{v:.2f} L/km".replace(".", ",") for v in re_ist_vals],
                        textposition="outside",
                        textfont=dict(size=11, color="#1a1a1a", family="Arial"),
                    )
                )
                if re_pot_pct < 0:
                    fig_lauf.add_trace(
                        go.Bar(
                            x=x_labels,
                            y=re_pot_vals,
                            name=f"Potenzial ({re_pot_pct} %) [L/km]",
                            marker=dict(color="#cdb663", line=dict(color="#a69043", width=1)),
                            text=[
                                f"<b>{pv:.2f}</b><br><span style='font-size:10px;'>({dv:+.2f})</span>".replace(".", ",")
                                for pv, dv in zip(re_pot_vals, re_diff_vals)
                            ],
                            textposition="outside",
                            textfont=dict(size=11, color="#6b5c22", family="Arial"),
                        )
                    )

                max_bar_y = max(re_ist_vals) if len(re_ist_vals) > 0 else 15.0
                min_bar_y = max(0.0, min(re_pot_vals) * 0.72) if len(re_pot_vals) > 0 else 0.0
                fig_lauf.update_layout(
                    title=dict(
                        text="<b>Sauerstoffverbrauch pro Kilometer: Ist vs. Potenzial (RE [L/km])</b>",
                        font=dict(size=13.5, color="#1a1a1a"),
                    ),
                    barmode="group",
                    bargap=0.25,
                    bargroupgap=0.10,
                    xaxis=dict(title="Pace [min/km]"),
                    yaxis=dict(
                        title=dict(text="RE [L O₂ / km]", font=dict(color="#1a1a1a")),
                        range=[min_bar_y, max_bar_y * 1.18],
                        gridcolor="rgba(0,0,0,0.08)",
                    ),
                    legend=dict(
                        orientation="h",
                        x=0.5,
                        xanchor="center",
                        y=1.11,
                        bgcolor="rgba(255,255,255,0.8)",
                    ),
                    height=395,
                    margin=dict(l=55, r=25, t=65, b=45),
                    plot_bgcolor="#f9f9f9",
                )
                st.plotly_chart(fig_lauf, use_container_width=True)

    if not is_lauf_pres:
        # EXPERTEN TOOLS (nur in der Diagnostik-Ansicht)
        st.markdown("---")
        with st.expander("Experten Tools zum Vergleich mit einer bestehenden HYCYS Diagnostik"):
            if not spiro_df.empty:
                st.subheader("Spiro-Synchronisation & Rohdaten")
                import plotly.graph_objects as go
                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(
                    x=spiro_df['Time'],
                    y=spiro_df['VO2'],
                    mode='lines+markers',
                    name='VO2',
                    line=dict(color='gray', width=1.5),
                    marker=dict(size=4),
                    hovertemplate='Zeit: %{x}s<br>VO2: %{y:.1f} ml/min<extra></extra>'
                ))
                
                shapes = []
                for i, (s_t, e_t) in enumerate(window_coords):
                    shapes.append(dict(
                        type="rect",
                        xref="x",
                        yref="paper",
                        x0=s_t,
                        y0=0,
                        x1=e_t,
                        y1=1,
                        fillcolor="#00a1e0",
                        opacity=0.3,
                        layer="below",
                        line_width=0,
                    ))
                
                fig2.update_layout(
                    xaxis=dict(title=dict(text="Zeit (s)")),
                    yaxis=dict(title=dict(text="VO2 (ml/min)")),
                    shapes=shapes,
                    height=350,
                    margin=dict(l=60, r=60, t=20, b=50),
                    plot_bgcolor='#f9f9f9',
                    showlegend=False
                )
                st.plotly_chart(fig2, use_container_width=True)
                
                col_diag1, col_diag2 = st.columns(2)
                with col_diag1:
                    st.write("**Metabolische Vergleichswerte:**")
                    st.write(f"- Berechnete VO2max: {rel_vo2max:.2f} ml/min/kg")
                with col_diag2:
                    st.write("**Stufen-VO2 (relativ):**")
                    render_centered_html_table(pd.DataFrame({"Stufe": [f"S{i+1}" for i in range(anzahl)], "VO2 rel": [f"{v:.2f}" for v in vo2_steady_values]}))
            
            st.subheader("Stufenauswertung")
            if not spiro_df.empty:
                step_data = []
                for i in range(anzahl):
                    v_ms   = speed[i]
                    pace_s = (1000.0 / v_ms) if v_ms > 0 else 0
                    pace_str = f"{int(pace_s // 60)}:{int(pace_s % 60):02d}"
                    vo2_a  = vo2_steady_abs[i]  if i < len(vo2_steady_abs)  else 0.0
                    vco2_a = vco2_steady_abs[i] if i < len(vco2_steady_abs) else 0.0
                    vo2_r  = vo2_steady_values[i]
                    rq_val   = (vco2_a / vo2_a) if (vco2_a > 0 and vo2_a > 0) else None
                    rq_str   = f"{rq_val:.3f}" if rq_val else '-'
                    # LO via exakter HYCYS-Methodik (stages-Funktion, nur unterschwellige Stufen)
                    lo_exp_list = calculate_laufokonomie_hycys_stages(
                        speeds_arr   = list(speed[:anzahl]),
                        vo2_abs_arr  = list(vo2_steady_abs[:anzahl]) if len(vo2_steady_abs) >= anzahl else [0.0]*anzahl,
                        vco2_abs_arr = list(vco2_steady_abs[:anzahl]) if len(vco2_steady_abs) >= anzahl else [0.0]*anzahl,
                        weight       = weight,
                        v_thresh     = v_ans if v_ans > 0 else None,
                    ) if len(vo2_steady_abs) > 0 else [None]*anzahl
                    lo_hycys = lo_exp_list[i] if i < len(lo_exp_list) else None
                    re_kg_km = (lo_hycys * 1000.0 / 60.0) if lo_hycys is not None else None
                    re_lit_kg_km = (vo2_r / (v_ms * 3.6) * 60.0) if (vo2_r > 0 and v_ms > 0) else None
                    re_l_km_exp = (re_kg_km * weight / 1000.0) if (re_kg_km is not None and weight > 0) else None
                    re_kcal_km_exp = (re_l_km_exp * 21.1 / 4.186) if re_l_km_exp is not None else None

                    step_data.append({
                        "Stufe":                     f"{i+1}",
                        "v [m/s]":                   v_ms,
                        "Pace [min/km]":             pace_str,
                        "VO2ss [ml/min]":            round(vo2_a, 1)  if vo2_a > 0 else '-',
                        "VO2ss [ml/min/kg]":         round(vo2_r, 2) if vo2_r > 0 else '-',
                        "RQ":                        rq_str,
                        "LÖ [ml/min/kg/(m/s)]":     round(lo_hycys, 2) if lo_hycys else '-',
                        "RE [ml/kg/km]":             round(re_kg_km, 1) if re_kg_km else '-',
                        "RE Lit. [ml/kg/km]":        round(re_lit_kg_km, 1) if re_lit_kg_km else '-',
                        "RE [l/km]":                 round(re_l_km_exp, 2) if re_l_km_exp else '-',
                        "Energie [kcal/km]":         round(re_kcal_km_exp, 1) if re_kcal_km_exp else '-'
                    })
                render_centered_html_table(pd.DataFrame(step_data))
                st.caption("LÖ [ml/min/kg/(m/s)] & RE [ml/kg/km] = HYCYS Methodik (netto, RQ-normiert)  |  RE Lit. [ml/kg/km] = Literaturstandard (brutto)  |  RE [l/km] & Energie [kcal/km] = gewichtsneutral (HYCYS netto)")

    # TRAININGSBEREICHE AN DIE UNTERSTE STELLE DER SEITE
    if not df_res.empty:
        st.markdown("---")
        if is_lauf_pres:
            st.subheader("Trainingsbereiche")
        else:
            st.subheader(f"Trainingsbereiche – basierend auf LTP1 & {selected_model_name}")

        ans_hf_num = float(hf_ans_val) if isinstance(hf_ans_val, (int, float, np.number)) else 0.0
        zones_ui = build_training_zones_variante_a(v_ltp1_val, hf_ltp1_val, v_ans, ans_hf_num)
        df_zones_ui = pd.DataFrame([
            {
                "Trainingsbereich": zn,
                "Pace min [min/km]": p_min,
                "Pace max [min/km]": p_max,
                "Pace Ziel [min/km]": p_zl,
                "HF min [bpm]": hr_mn,
                "HF max [bpm]": hr_mx,
                "HF Ziel [bpm]": hr_zl,
            }
            for (zn, p_min, p_max, p_zl, hr_mn, hr_mx, hr_zl) in zones_ui
        ])
        render_centered_html_table(df_zones_ui)


    # --- PDF DOWNLOAD & GOOGLE DRIVE ABLAGE (nur in der Diagnostik-Ansicht) ---
    if not is_lauf_pres and fig_laktat is not None:
        st.markdown("---")
        pdf_bytes = create_pdf(
            athlete_name, birthdate, test_date, st.session_state.get("diagnostik_type", "Running BLUE"), weight, body_fat_pct, rel_vo2max, df_res, fig_laktat,
            coach, sportart, kategorie, height, speed, lactate, hr, vo2_steady_values, vo2_steady_abs, vco2_steady_abs, _has_spiro_loaded,
            v_ltp1=v_ltp1_val, hf_ltp1=hf_ltp1_val
        )

        excel_bytes_lauf = None
        try:
            excel_bytes_lauf = create_hycys_excel_lauf(
                athlete_name=athlete_name,
                birthdate=birthdate,
                test_date=test_date,
                gender_str=st.session_state.get("gender", "männlich"),
                diagnostik_type=st.session_state.get("diagnostik_type", "Running BLUE"),
                coach=coach,
                sportart=sportart,
                kategorie=kategorie,
                height=height,
                weight=weight,
                body_fat_pct=body_fat_pct,
                df_sf=st.session_state.get("df_sf"),
                ruhe_lac=ruhe_lac,
                ruhe_hr=ruhe_hr,
                speed=list(speed),
                lactate=list(lactate),
                hr=list(hr),
                abs_vo2max=abs_vo2max,
                vo2_steady_abs=list(vo2_steady_abs),
                vco2_steady_abs=list(vco2_steady_abs),
                uploaded_file=uploaded_file,
                uploaded_ramp=uploaded_ramp if separate_ramp else None,
            )
        except Exception as e_xls_lauf:
            st.warning(f"Hinweis: HYCYS-Excel konnte nicht erzeugt werden ({e_xls_lauf}).")

        render_gdrive_upload_section(
            pdf_bytes=pdf_bytes,
            diagnostik_type=st.session_state.get("diagnostik_type", "Running BLUE"),
            athlete_name=athlete_name,
            test_date=test_date,
            gender_str=st.session_state.get("gender", "männlich"),
            key_prefix="laufen",
            excel_bytes=excel_bytes_lauf,
        )

save_current_session_snapshot()

