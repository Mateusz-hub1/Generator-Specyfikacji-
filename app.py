import os
import re
import datetime
import streamlit as st
from googleapiclient.errors import HttpError

from sheets_engine import generate_spreadsheet, cascade_nr, parse_nr_prac

# ── Konfiguracja strony ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Generator Specyfikacji | Korcz",
    page_icon="🚂",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Nagłówek */
.korcz-header {
    background: linear-gradient(135deg, #0D1E31 0%, #1C3557 60%, #2E5D8E 100%);
    border-radius: 16px;
    padding: 28px 32px 24px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 20px;
    position: relative;
    overflow: hidden;
}
.korcz-header::before {
    content: '';
    position: absolute;
    inset: 0;
    background: repeating-linear-gradient(
        90deg, transparent 0px, transparent 28px,
        rgba(255,255,255,0.04) 28px, rgba(255,255,255,0.04) 30px
    );
}
.korcz-header-icon {
    width: 56px; height: 56px;
    background: rgba(255,255,255,0.12);
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-size: 28px;
    flex-shrink: 0;
    border: 1px solid rgba(255,255,255,0.18);
    position: relative;
}
.korcz-header-text { position: relative; }
.korcz-header-eyebrow {
    font-size: 10px; font-weight: 700;
    letter-spacing: 0.12em; text-transform: uppercase;
    color: #7EC8F4; margin-bottom: 4px;
}
.korcz-header-title {
    font-size: 22px; font-weight: 700;
    color: #fff; line-height: 1.2; margin: 0;
}
.korcz-header-sub {
    font-size: 13px; color: rgba(255,255,255,0.55);
    margin-top: 3px;
}

/* Sekcje */
.section-label {
    font-size: 10px; font-weight: 700;
    letter-spacing: 0.12em; text-transform: uppercase;
    color: #94A3B8;
    display: flex; align-items: center; gap: 10px;
    margin: 24px 0 14px;
}
.section-label::after {
    content: ''; flex: 1; height: 1px; background: #E2E8F0;
}

/* Preview kaskadowy */
.cascade-box {
    background: #EEF5FF;
    border: 1px solid #C3D9F5;
    border-radius: 10px;
    padding: 12px 16px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    color: #1C3557;
    line-height: 2;
    margin-top: 8px;
}
.cascade-label {
    font-size: 9px; font-weight: 700;
    letter-spacing: 0.1em; text-transform: uppercase;
    color: #94A3B8; margin-bottom: 4px;
}

/* Sukces */
.success-box {
    background: #F0FBF4;
    border: 1.5px solid #68D391;
    border-radius: 12px;
    padding: 20px 24px;
    display: flex; gap: 14px;
    align-items: flex-start;
    margin-top: 8px;
}
.success-icon { font-size: 22px; flex-shrink: 0; margin-top: 1px; }
.success-title { font-weight: 700; color: #276749; font-size: 15px; margin-bottom: 4px; }
.success-sub { color: #2F855A; font-size: 13px; }
.success-link {
    display: inline-block; margin-top: 10px;
    background: #276749; color: #fff !important;
    padding: 8px 18px; border-radius: 8px;
    font-size: 13px; font-weight: 600;
    text-decoration: none !important;
}
.success-link:hover { background: #22543D; }

/* Ukryj footer Streamlit */
footer { visibility: hidden; }
#MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Nagłówek ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="korcz-header">
  <div class="korcz-header-icon">🚂</div>
  <div class="korcz-header-text">
    <div class="korcz-header-eyebrow">Serwis Pojazdów Kolejowych A. Korcz Sp. J.</div>
    <div class="korcz-header-title">Generator Specyfikacji</div>
    <div class="korcz-header-sub">Automatyczne tworzenie arkuszy w Google Sheets</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Odczyt konfiguracji (Streamlit Secrets lub env) ──────────────────────────
def get_config():
    """Zwraca (target_folder_id, credentials_json_str, credentials_file)."""
    try:
        # Streamlit Cloud: st.secrets["TARGET_FOLDER_ID"]
        folder_id   = st.secrets["TARGET_FOLDER_ID"]
        creds_json  = st.secrets.get("GOOGLE_CREDENTIALS_JSON", None)
        creds_file  = st.secrets.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
    except Exception:
        # Lokalnie: zmienne środowiskowe
        folder_id   = os.environ.get("TARGET_FOLDER_ID", "")
        creds_json  = os.environ.get("GOOGLE_CREDENTIALS_JSON", None)
        creds_file  = os.environ.get("GOOGLE_CREDENTIALS", "credentials.json")
    return folder_id, creds_json, creds_file


TARGET_FOLDER_ID, CREDS_JSON, CREDS_FILE = get_config()

# Ostrzeżenie jeśli brak konfiguracji
if not TARGET_FOLDER_ID or TARGET_FOLDER_ID == "TWOJ_FOLDER_ID_TUTAJ":
    st.warning(
        "⚠️ Brak konfiguracji `TARGET_FOLDER_ID`. "
        "Dodaj go w `.streamlit/secrets.toml` lub jako zmienną środowiskową.",
        icon="⚙️",
    )


# ── Formularz ─────────────────────────────────────────────────────────────────

st.markdown('<div class="section-label">Dane zlecenia</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    nazwa_pliku = st.text_input(
        "Nazwa pliku docelowego",
        placeholder="np. Specyfikacja_DB_Cargo_06_2026",
        help="Tak będzie nazwany arkusz na Google Drive",
    )
with col2:
    firma = st.text_input(
        "Firma (odbiorca usługi)",
        placeholder="np. DB Cargo Polska S.A.",
    )

st.markdown('<div class="section-label">Parametry usługi</div>', unsafe_allow_html=True)

col3, col4, col5 = st.columns([1, 1.4, 1.6])
with col3:
    wagon_count = st.number_input(
        "Ilość wagonów",
        min_value=1, max_value=200, value=1, step=1,
    )
with col4:
    data_usl = st.date_input(
        "Data wykonania usługi",
        value=datetime.date.today(),
    )
with col5:
    miejsce = st.text_input(
        "Miejsce naprawy",
        placeholder="np. Tarnowskie Góry",
    )

st.markdown('<div class="section-label">Numeracja prac</div>', unsafe_allow_html=True)

nr_prac = st.text_input(
    "Nr prac",
    placeholder="14/06/2026",
    help="Format: numer/miesiąc/rok, np. 14/06/2026",
)

# ── Podgląd kaskadowy ─────────────────────────────────────────────────────────
if nr_prac and re.match(r"^\d+/.+$", nr_prac.strip()):
    try:
        parse_nr_prac(nr_prac)
        preview_count = min(int(wagon_count), 5)
        lines = [f"a{i+1}  →  {cascade_nr(nr_prac, i+1)}" for i in range(preview_count)]
        if wagon_count > 5:
            lines.append(f"... ({wagon_count - 5} więcej)")
        preview_html = "<br>".join(lines)
        st.markdown(f"""
        <div class="cascade-label">Podgląd numeracji kaskadowej</div>
        <div class="cascade-box">{preview_html}</div>
        """, unsafe_allow_html=True)
    except ValueError:
        pass

st.write("")  # odstęp

# ── Przycisk generowania ──────────────────────────────────────────────────────
generate_btn = st.button(
    "🗂️  Generuj arkusz",
    type="primary",
    use_container_width=True,
    disabled=not TARGET_FOLDER_ID or TARGET_FOLDER_ID == "TWOJ_FOLDER_ID_TUTAJ",
)

if generate_btn:
    # Walidacja
    errors = []
    if not nazwa_pliku.strip():   errors.append("Podaj nazwę pliku.")
    if not firma.strip():         errors.append("Podaj nazwę firmy.")
    if not miejsce.strip():       errors.append("Podaj miejsce naprawy.")
    if not nr_prac.strip():       errors.append("Podaj nr prac.")
    else:
        try:
            parse_nr_prac(nr_prac)
        except ValueError as e:
            errors.append(str(e))

    if errors:
        for e in errors:
            st.error(e, icon="❌")
    else:
        # Pasek postępu
        progress_bar = st.progress(0, text="Łączenie z Google API…")

        def update_progress(val: float):
            labels = {
                0.1: "Tworzenie pliku na Drive…",
                0.2: "Inicjalizacja zakładek…",
                0.3: "Dodawanie zakładek a1…aX…",
                0.4: "Formatowanie specyfikacji…",
                0.6: "Wstawianie danych…",
            }
            text = labels.get(round(val, 1), "Formatowanie zakładek podrzędnych…")
            progress_bar.progress(val, text=text)

        try:
            result = generate_spreadsheet(
                target_folder_id=TARGET_FOLDER_ID,
                credentials_json=CREDS_JSON,
                credentials_file=CREDS_FILE,
                nazwa_pliku=nazwa_pliku.strip(),
                firma=firma.strip(),
                wagon_count=int(wagon_count),
                data_usl=str(data_usl),
                miejsce=miejsce.strip(),
                nr_prac=nr_prac.strip(),
                progress_cb=update_progress,
            )
            progress_bar.progress(1.0, text="Gotowe!")

            sheets_count = result["sheets"]
            st.markdown(f"""
            <div class="success-box">
              <div class="success-icon">✅</div>
              <div>
                <div class="success-title">Arkusz utworzony pomyślnie!</div>
                <div class="success-sub">
                  Plik <strong>{result['filename']}</strong> z {sheets_count} zakładkami
                  ({sheets_count - 1} wagonów) został zapisany w docelowym folderze.
                </div>
                <a class="success-link" href="{result['url']}" target="_blank">
                  Otwórz w Google Sheets →
                </a>
              </div>
            </div>
            """, unsafe_allow_html=True)

        except ValueError as e:
            progress_bar.empty()
            st.error(f"Błąd danych: {e}", icon="⚠️")
        except HttpError as e:
            progress_bar.empty()
            st.error(f"Błąd Google API: {e.reason}", icon="🔴")
        except Exception as e:
            progress_bar.empty()
            st.error(f"Nieoczekiwany błąd: {e}", icon="🔴")

# ── Stopka ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Serwis Pojazdów Kolejowych A. Korcz Sp. J. · Generator dokumentacji technicznej")
