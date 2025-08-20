# estrattore_cedolini_con_mappa_v1.py
import os
import re
import sqlite3
import pandas as pd
import pdfplumber
from tqdm import tqdm
from typing import Dict, Any, Optional, List

# --- CONFIGURAZIONE ---
DB_PATH = "gestionale_loves.db"
PDF_FOLDER = "cedolini1"
EXCEL_REPORT_PATH = "report_cedolini_finale.xlsx"

# ==============================================================================
# === MAPPA DEL LAYOUT (GENERATA DALL'ANALISI DELL'ALTRO LLM) ===
# ==============================================================================
LAYOUT_MAP = {
    'AZIENDA': { 'offset_x': -17.76, 'offset_y': 0.0, 'width': 18, 'height': 6 },
    'CODICE FISCALE': { 'offset_x': -39.46, 'offset_y': 9.60, 'width': 160, 'height': 9 },
    'PARTITA IVA': { 'offset_x': -23.13, 'offset_y': 9.60, 'width': 100, 'height': 9 },
    'MATRICOLA INPS': { 'offset_x': -44.28, 'offset_y': 11.52, 'width': 100, 'height': 9 },
    'POSIZIONE INAIL': { 'offset_x': -33.74, 'offset_y': 0.0, 'width': 80, 'height': 6 },
    'SEDE INAIL': { 'offset_x': -22.70, 'offset_y': 0.0, 'width': 80, 'height': 6 },
    'INDIRIZZO': { 'offset_x': 5, 'offset_y': 10, 'width': 200, 'height': 20 },
    'DIPENDENTE': { 'offset_x': -19.86, 'offset_y': 10, 'width': 150, 'height': 10 },
    'QUALIFICA': { 'offset_x': 5, 'offset_y': 10, 'width': 100, 'height': 10 },
    'MANSIONE': { 'offset_x': 5, 'offset_y': 10, 'width': 150, 'height': 10 },
    'LIVELLO': { 'offset_x': 5, 'offset_y': 10, 'width': 50, 'height': 10 },
    'CONTRATTO APPLICATO': { 'offset_x': 5, 'offset_y': 10, 'width': 300, 'height': 20 },
    'DATA ASSUNZIONE': { 'offset_x': 5, 'offset_y': 12, 'width': 80, 'height': 10 },
    'DATA CESSAZIONE': { 'offset_x': 5, 'offset_y': 12, 'width': 80, 'height': 10 },
    'DATA DI NASCITA': { 'offset_x': 5, 'offset_y': 10, 'width': 80, 'height': 10 },
    'MESE RETRIBUITO': { 'offset_x': -60.10, 'offset_y': -9.16, 'width': 80, 'height': 10 },
    'TOTALE COMPETENZE': { 'offset_x': 1, 'offset_y': 10, 'width': 80, 'height': 10 },
    'TOTALE RITENUTE': { 'offset_x': 1, 'offset_y': 10, 'width': 80, 'height': 10 },
    'RITENUTE INPS': { 'offset_x': 1, 'offset_y': 10, 'width': 80, 'height': 10 },
    'IMPONIBILE FISCALE': { 'offset_x': 1, 'offset_y': 10, 'width': 80, 'height': 10 },
    'NETTO IN BUSTA': { 'offset_x': 1, 'offset_y': 10, 'width': 80, 'height': 10 },
    'TFR DEL MESE': { 'offset_x': 1, 'offset_y': 10, 'width': 80, 'height': 10 }
}
# ==============================================================================

# --- FUNZIONI DI SUPPORTO E DB (Invariate) ---
def _parse_float(s: Optional[str]) -> Optional[float]:
    if not s: return None
    s = str(s).strip().replace(".", "").replace(",", ".")
    s = re.sub(r'[^\d.-]', '', s)
    try: return float(s)
    except (ValueError, TypeError): return None

def init_db(conn: sqlite3.Connection):
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS dim_dipendente (
        id INTEGER PRIMARY KEY AUTOINCREMENT, dipendente_cf TEXT NOT NULL UNIQUE, dipendente_nome TEXT,
        qualifica TEXT, mansione TEXT, livello TEXT, data_assunzione TEXT, data_nascita TEXT
    );""")
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS bi_labor_dettaglio (
        key_dipendente TEXT PRIMARY KEY, dipendente_id INTEGER, mese_retribuito TEXT, anno INTEGER,
        totale_competenze REAL, totale_trattenute REAL, netto_a_pagare REAL,
        imponibile_fiscale REAL, ritenute_inps REAL, tfr_mese REAL, source_file TEXT,
        FOREIGN KEY (dipendente_id) REFERENCES dim_dipendente (id)
    );""")
    conn.commit()

def load_data(conn: sqlite3.Connection, data: dict):
    cursor = conn.cursor()
    cf = data.get('codice_fiscale')
    if not cf: return
    cursor.execute("SELECT id FROM dim_dipendente WHERE dipendente_cf = ?", (cf,))
    res = cursor.fetchone()
    if res:
        dipendente_id = res[0]
    else:
        cursor.execute("INSERT INTO dim_dipendente (dipendente_cf, dipendente_nome, qualifica, mansione, livello, data_assunzione, data_nascita) VALUES (?,?,?,?,?,?,?)",
                       (cf, data.get('dipendente'), data.get('qualifica'), data.get('mansione'), data.get('livello'), data.get('data_assunzione'), data.get('data_di_nascita')))
        dipendente_id = cursor.lastrowid
    
    cursor.execute("""
    INSERT OR REPLACE INTO bi_labor_dettaglio (key_dipendente, dipendente_id, mese_retribuito, anno,
    totale_competenze, totale_trattenute, netto_a_pagare, imponibile_fiscale, ritenute_inps, tfr_mese, source_file)
    VALUES (?,?,?,?,?,?,?,?,?,?,?)
    """,(
        data.get('key_dipendente'), dipendente_id, data.get('mese_retribuito'), data.get('anno'),
        data.get('totale_competenze'), data.get('totale_trattenute'),
        data.get('netto_in_busta'), data.get('imponibile_fiscale'), data.get('ritenute_inps'),
        data.get('tfr_del_mese'), data.get('source_file')
    ))

def create_excel_report(conn: sqlite3.Connection):
    print(f"\nCreazione del report Excel: {EXCEL_REPORT_PATH}...")
    try:
        df_dettaglio = pd.read_sql_query("SELECT * FROM bi_labor_dettaglio", conn)
        df_dipendenti = pd.read_sql_query("SELECT * FROM dim_dipendente", conn)
        with pd.ExcelWriter(EXCEL_REPORT_PATH, engine='xlsxwriter') as writer:
            if not df_dettaglio.empty:
                df_dettaglio.to_excel(writer, sheet_name='Dettaglio Cedolini', index=False)
            if not df_dipendenti.empty:
                df_dipendenti.to_excel(writer, sheet_name='Anagrafica Dipendenti', index=False)
        print("Report Excel creato con successo.")
    except Exception as e:
        print(f"Errore durante la creazione del report Excel: {e}")

# --- LOGICA DI ESTRAZIONE CON MAPPA ---

def extract_data_with_layout(page, layout_map: Dict[str, Dict], source_file: str) -> Optional[Dict[str, Any]]:
    data = {'source_file': source_file}
    
    # Trova il codice fiscale come ancora principale per assicurarsi che sia un cedolino valido
    cf_match = re.search(r'\b([A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z])\b', page.extract_text(x_tolerance=2) or "")
    if not cf_match:
        return None
    
    for label, properties in layout_map.items():
        # Usiamo una ricerca flessibile per l'etichetta
        label_occurrences = page.search(label, case=False, whole_words=True)
        if not label_occurrences:
            continue
        
        label_bbox = label_occurrences[0]
        
        # Applica l'offset appreso per definire l'area del valore
        value_bbox = (
            label_bbox['x0'] + properties['offset_x'],
            label_bbox['top'] + properties['offset_y'],
            label_bbox['x0'] + properties['offset_x'] + properties['width'],
            label_bbox['top'] + properties['offset_y'] + properties['height'],
        )
        
        # Controlla che il bounding box sia valido e dentro la pagina
        if not (value_bbox[0] < value_bbox[2] and value_bbox[1] < value_bbox[3] and
                value_bbox[0] >= 0 and value_bbox[1] >=0 and
                value_bbox[2] <= page.width and value_bbox[3] <= page.height):
            continue

        cropped_page = page.crop(value_bbox)
        extracted_value = cropped_page.extract_text(x_tolerance=2, y_tolerance=2)
        
        if extracted_value:
            # Crea una chiave pulita per il dizionario (es. 'NETTO IN BUSTA' -> 'netto_in_busta')
            data_key = label.lower().replace(" ", "_").replace(".", "")
            data[data_key] = extracted_value.strip()

    # Post-elaborazione e pulizia
    if not data.get('codice_fiscale'):
        data['codice_fiscale'] = cf_match.group(1)
        
    anno_match = re.search(r'\b(20\d{2})\b', page.extract_text() or "")
    data['anno'] = int(anno_match.group(1)) if anno_match else None
    
    if not all([data.get('codice_fiscale'), data.get('anno'), data.get('mese_retribuito')]):
        return None # Salta se i dati chiave non sono stati estratti
        
    data['key_dipendente'] = f"{data['codice_fiscale']}_{data['anno']}{data['mese_retribuito']}"
    
    # Converte i valori numerici
    for key, value in data.items():
        if key in ['totale_competenze', 'totale_trattenute', 'netto_in_busta', 'imponibile_fiscale', 'ritenute_inps', 'tfr_del_mese']:
            data[key] = _parse_float(value)
            
    return data

# --- MAIN ---
def main():
    if not os.path.exists(PDF_FOLDER):
        print(f"ERRORE: La cartella '{PDF_FOLDER}' non esiste.")
        return

    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    
    pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith('.pdf')]
    
    for filename in tqdm(pdf_files, desc="Processing PDFs"):
        file_path = os.path.join(PDF_FOLDER, filename)
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    record = extract_data_with_layout(page, LAYOUT_MAP, filename)
                    if record:
                        load_data(conn, record)
        except Exception as e:
            print(f"\nErrore durante l'elaborazione del file {filename}: {e}")
    
    conn.commit()
    create_excel_report(conn)
    conn.close()
    print("\nProcesso completato.")

if __name__ == "__main__":
    main()