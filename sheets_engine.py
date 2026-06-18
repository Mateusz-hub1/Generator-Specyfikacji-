"""
Cała logika Google Sheets / Drive API.
Niezależna od warstwy UI — działa tak samo w Flask i Streamlit.
"""

import os
import re
import json
import base64
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ── Autoryzacja ──────────────────────────────────────────────────────────────

def get_credentials(credentials_json: str | None = None,
                    credentials_file: str = "credentials.json"):
    """
    Obsługuje dwa tryby:
      1. credentials_json  – JSON jako string (Streamlit Secrets / env var)
      2. credentials_file  – ścieżka do pliku (tryb lokalny)
    """
    if credentials_json:
        info = json.loads(credentials_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return service_account.Credentials.from_service_account_file(credentials_file, scopes=SCOPES)


# ── Pomocnicze ───────────────────────────────────────────────────────────────

def parse_nr_prac(nr_prac: str):
    match = re.match(r"^(\d+)/(.+)$", nr_prac.strip())
    if not match:
        raise ValueError(f"Nieprawidłowy format: '{nr_prac}'. Oczekiwano np. '14/06/2026'")
    return match.group(1), match.group(2)


def cascade_nr(nr_prac: str, index: int) -> str:
    base, suffix = parse_nr_prac(nr_prac)
    return f"{base}-{index}/{suffix}"


def rgb_to_gsheets(hex_color: str):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


def batch_update(service, spreadsheet_id: str, requests_list: list):
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests_list},
    ).execute()


def cell_range(sheet_id: int, r1: int, c1: int, r2: int, c2: int):
    return {
        "sheetId": sheet_id,
        "startRowIndex": r1 - 1,
        "endRowIndex": r2,
        "startColumnIndex": c1 - 1,
        "endColumnIndex": c2,
    }


def border_style(style="SOLID", width=1, color="#000000"):
    return {"style": style, "width": width, "color": rgb_to_gsheets(color)}


# ── Formatowanie zakładki specyfikacja ──────────────────────────────────────

def format_specyfikacja(service, spreadsheet_id, sheet_id,
                        wagon_count, firma, data_usl, miejsce, nr_prac,
                        progress_cb=None):
    reqs = []
    GRAY_HEADER = "#1C3557"
    LIGHT_BLUE  = "#DCE8F5"
    ACCENT      = "#E8F0FA"

    col_widths = [30, 50, 220, 180, 60, 160, 160]
    for i, w in enumerate(col_widths):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize",
        }})

    row_heights = {1: 80, 2: 36, 3: 14, 4: 28, 5: 28, 6: 28, 7: 14, 8: 14, 9: 28}
    for row, h in row_heights.items():
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": row - 1, "endIndex": row},
            "properties": {"pixelSize": h}, "fields": "pixelSize",
        }})
    for row in range(10, 10 + wagon_count):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": row - 1, "endIndex": row},
            "properties": {"pixelSize": 26}, "fields": "pixelSize",
        }})

    merges = [(1,2,1,7),(2,1,2,7),(4,2,4,7),(5,4,5,5),(6,4,6,5)]
    for r1,c1,r2,c2 in merges:
        reqs.append({"mergeCells": {
            "range": cell_range(sheet_id, r1, c1, r2, c2),
            "mergeType": "MERGE_ALL",
        }})

    def bg_req(r1, c1, r2, c2, color):
        return {"repeatCell": {
            "range": cell_range(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": {"backgroundColor": rgb_to_gsheets(color)}},
            "fields": "userEnteredFormat.backgroundColor",
        }}

    reqs += [
        bg_req(1,2,1,7,"#FFFFFF"), bg_req(2,1,2,7,ACCENT),
        bg_req(4,2,4,7,GRAY_HEADER), bg_req(5,2,5,7,GRAY_HEADER),
        bg_req(9,2,9,7,GRAY_HEADER),
    ]
    for row in range(10, 10 + wagon_count, 2):
        reqs.append(bg_req(row, 2, row, 7, LIGHT_BLUE))

    border = border_style()
    for row in range(9, 10 + wagon_count):
        for col in range(2, 8):
            reqs.append({"updateBorders": {
                "range": cell_range(sheet_id, row, col, row, col),
                "top": border, "bottom": border,
                "left": border, "right": border,
            }})

    batch_update(service, spreadsheet_id, reqs)
    if progress_cb: progress_cb(0.4)

    sheet = service.spreadsheets()

    def write(range_, values_):
        sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=f"specyfikacja!{range_}",
            valueInputOption="USER_ENTERED",
            body={"values": values_},
        ).execute()

    write("B1", [["Serwis Pojazdów Kolejowych A. Korcz Sp. J."]])
    write("A2", [[f"SPECYFIKACJA WYKONANYCH PRAC nr {nr_prac}"]])
    write("B4", [["DANE PODSTAWOWE"]])
    write("B5", [["Lp.", "Seria i numer lokomotywy/wagonu:", "Data wykonania usługi:", None, "Miejsce naprawy"]])
    write("B6", [["1.", f"{wagon_count} szt.", data_usl, None, miejsce]])
    write("B9", [["Lp", "Numer wagonu", None, None, "kwota netto"]])
    write("B10", [[str(i+1), "", None, None, ""] for i in range(wagon_count)])

    def text_fmt(r1, c1, r2, c2, bold=False, size=10, fg="#000000", h_align="LEFT"):
        return {"repeatCell": {
            "range": cell_range(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": bold, "fontSize": size,
                               "foregroundColor": rgb_to_gsheets(fg)},
                "horizontalAlignment": h_align,
                "verticalAlignment": "MIDDLE",
                "wrapStrategy": "WRAP",
            }},
            "fields": ("userEnteredFormat.textFormat,"
                       "userEnteredFormat.horizontalAlignment,"
                       "userEnteredFormat.verticalAlignment,"
                       "userEnteredFormat.wrapStrategy"),
        }}

    batch_update(service, spreadsheet_id, [
        text_fmt(1,2,1,7, bold=True, size=13, fg="#1C3557"),
        text_fmt(2,1,2,7, bold=True, size=11, fg="#1C3557", h_align="CENTER"),
        text_fmt(4,2,4,7, bold=True, size=10, fg="#FFFFFF", h_align="CENTER"),
        text_fmt(5,2,5,7, bold=True, size=9,  fg="#FFFFFF", h_align="CENTER"),
        text_fmt(6,2,6,7, size=10, h_align="CENTER"),
        text_fmt(9,2,9,7, bold=True, size=9, fg="#FFFFFF", h_align="CENTER"),
        text_fmt(10,2,9+wagon_count,7, size=10, h_align="CENTER"),
    ])


# ── Formatowanie zakładek podrzędnych ───────────────────────────────────────

def format_sub_sheet(service, spreadsheet_id, sheet_id,
                     sheet_name, cascade_nr_val, firma, data_usl, miejsce):
    reqs = []
    for i, w in enumerate([40, 200, 150, 100, 100, 150]):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize",
        }})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                  "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 60}, "fields": "pixelSize",
    }})
    reqs.append({"repeatCell": {
        "range": cell_range(sheet_id, 1, 1, 1, 6),
        "cell": {"userEnteredFormat": {"backgroundColor": rgb_to_gsheets("#1C3557")}},
        "fields": "userEnteredFormat.backgroundColor",
    }})
    reqs.append({"mergeCells": {
        "range": cell_range(sheet_id, 1, 1, 1, 6),
        "mergeType": "MERGE_ALL",
    }})
    batch_update(service, spreadsheet_id, reqs)

    s = service.spreadsheets()
    for rng, val in [(f"{sheet_name}!A1", "Serwis Pojazdów Kolejowych A. Korcz Sp. J."),
                     (f"{sheet_name}!A2", cascade_nr_val)]:
        s.values().update(spreadsheetId=spreadsheet_id, range=rng,
                          valueInputOption="USER_ENTERED",
                          body={"values": [[val]]}).execute()

    batch_update(service, spreadsheet_id, [
        {"repeatCell": {
            "range": cell_range(sheet_id, 1, 1, 1, 6),
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "fontSize": 12,
                               "foregroundColor": rgb_to_gsheets("#FFFFFF")},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
            }},
            "fields": ("userEnteredFormat.textFormat,"
                       "userEnteredFormat.horizontalAlignment,"
                       "userEnteredFormat.verticalAlignment"),
        }},
        {"repeatCell": {
            "range": cell_range(sheet_id, 2, 1, 2, 3),
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "fontSize": 11,
                               "foregroundColor": rgb_to_gsheets("#1C3557")},
                "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
            }},
            "fields": ("userEnteredFormat.textFormat,"
                       "userEnteredFormat.horizontalAlignment,"
                       "userEnteredFormat.verticalAlignment"),
        }},
    ])



def insert_logo(service, drive_service, target_folder_id, spreadsheet_id, sheet_id):
    if not os.path.exists(LOGO_PATH):
        return

    # 1. Wgranie pliku na Dysk Google
    file_metadata = {
        'name': os.path.basename(LOGO_PATH),
        'parents': [target_folder_id],
    }
    media = MediaFileUpload(LOGO_PATH, mimetype='image/png')
    uploaded = drive_service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id',
        supportsAllDrives=True  # <--- I TU TEŻ
    ).execute()
    logo_id = uploaded.get('id')

    # 2. Generowanie formuły i wstawienie
    image_formula = f'=IMAGE("https://drive.google.com/uc?export=view&id={logo_id}")'
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="specyfikacja!A1",
        valueInputOption="USER_ENTERED",
        body={"values": [[image_formula]]},
    ).execute()

# ── Główna funkcja generująca arkusz ────────────────────────────────────────

def generate_spreadsheet(
    *,
    target_folder_id: str,
    credentials_json: str | None = None,
    credentials_file: str = "credentials.json",
    nazwa_pliku: str,
    firma: str,
    wagon_count: int,
    data_usl: str,
    miejsce: str,
    nr_prac: str,
    progress_cb=None,
) -> dict:
    """
    Tworzy arkusz Google Sheets i zwraca {'url': ..., 'filename': ..., 'sheets': ...}.
    Rzuca ValueError przy złych danych, HttpError przy błędach API.
    """
    parse_nr_prac(nr_prac)  # walidacja formatu

    creds  = get_credentials(credentials_json, credentials_file)
    drive  = build("drive",  "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    if progress_cb: progress_cb(0.1)

file_obj = drive.files().create(
        body={"name": nazwa_pliku,
              "mimeType": "application/vnd.google-apps.spreadsheet",
              "parents": [target_folder_id]},
        fields="id,webViewLink",
        supportsAllDrives=True  
    ).execute()
    sid = file_obj["id"]
    url = file_obj["webViewLink"]

    if progress_cb: progress_cb(0.2)

    meta = sheets.spreadsheets().get(spreadsheetId=sid).execute()
    main_sheet_id = meta["sheets"][0]["properties"]["sheetId"]

    sub_titles = [f"a{i+1}" for i in range(wagon_count)]
    rename_reqs = [{"updateSheetProperties": {
        "properties": {"sheetId": main_sheet_id, "title": "specyfikacja"},
        "fields": "title",
    }}]
    for t in sub_titles:
        rename_reqs.append({"addSheet": {"properties": {"title": t}}})
    batch_update(sheets, sid, rename_reqs)

    if progress_cb: progress_cb(0.3)

    meta2 = sheets.spreadsheets().get(spreadsheetId=sid).execute()
    id_map = {s["properties"]["title"]: s["properties"]["sheetId"]
              for s in meta2["sheets"]}

    format_specyfikacja(sheets, sid, id_map["specyfikacja"],
                        wagon_count, firma, data_usl, miejsce, nr_prac,
                        progress_cb=progress_cb)
    insert_logo(sheets, drive, target_folder_id, sid, id_map["specyfikacja"])

    if progress_cb: progress_cb(0.6)

    for i, title in enumerate(sub_titles):
        format_sub_sheet(sheets, sid, id_map[title], title,
                         cascade_nr(nr_prac, i+1), firma, data_usl, miejsce)
        if progress_cb:
            progress_cb(0.6 + 0.4 * (i + 1) / len(sub_titles))

    return {"url": url, "filename": nazwa_pliku, "sheets": len(sub_titles) + 1}
