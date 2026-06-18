# Generator Specyfikacji — wersja Streamlit

Aplikacja do automatycznego generowania sformatowanych arkuszy Google Sheets.
Działa lokalnie i na **Streamlit Community Cloud** (darmowy hosting).

---

## Struktura

```
korcz-streamlit/
├── app.py                          ← Interfejs Streamlit (UI)
├── sheets_engine.py                ← Logika Google API (niezależna od UI)
├── requirements.txt
├── assets/
│   └── logo.png
└── .streamlit/
    └── secrets.toml.example        ← Wzorzec konfiguracji (skopiuj bez .example)
```

---

## Uruchomienie lokalne

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Konfiguracja (opcja A – plik secrets)
mkdir .streamlit
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# → uzupełnij TARGET_FOLDER_ID i GOOGLE_CREDENTIALS_JSON

# Konfiguracja (opcja B – zmienne środowiskowe)
export TARGET_FOLDER_ID="id_folderu"
export GOOGLE_CREDENTIALS="credentials.json"   # ścieżka do pliku

streamlit run app.py
# → http://localhost:8501
```

---

## Deploy na Streamlit Community Cloud (darmowy)

1. Wrzuć projekt na **GitHub** (bez `credentials.json` i `secrets.toml`!)
2. Wejdź na https://share.streamlit.io → **New app**
3. Wybierz repozytorium, branch `main`, plik `app.py`
4. Kliknij **Advanced settings → Secrets** i wklej:

```toml
TARGET_FOLDER_ID = "twoje_id_folderu"

GOOGLE_CREDENTIALS_JSON = '''
{ ... cała zawartość credentials.json ... }
'''
```

5. Kliknij **Deploy** — gotowe, link działa dla każdego z biura!

---

## Konfiguracja Google (wymagana raz)

1. https://console.cloud.google.com → włącz **Sheets API** i **Drive API**
2. **IAM → Service Accounts** → utwórz konto → pobierz klucz JSON
3. Na Google Drive: folder PPM → **Udostępnij** → wklej `client_email` z JSON → **Edytor**
4. Skopiuj ID folderu z URL i wklej jako `TARGET_FOLDER_ID`

---

## .gitignore

```
credentials.json
.streamlit/secrets.toml
__pycache__/
venv/
*.pyc
```
