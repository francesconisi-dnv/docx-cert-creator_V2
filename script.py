import os
import shutil
import pandas as pd
from fpdf import FPDF

# --- CONFIGURAZIONI ---
SFONDO_PATH = "sfondo_pulito.svg"
OUTPUT_DIR = "Attestati_PDF"
LOGO_PATH = "logo.png"

FONT_REGULAR_PATH = "font_custom.ttf" 
FONT_BOLD_PATH = "font_custom_bold.ttf" 

COLORE_VIOLA = (123, 44, 191)
COLORE_TESTO = (26, 26, 26)

# --- SPOSTAMENTO GLOBALE ---
SPOSTAMENTO_VERTICALE = 0  

# --- COORDINATE VERTICALI RELATIVE (in millimetri) ---
# Spaziatura uniforme di 20mm tra ogni blocco per un allineamento centrale perfetto
Y_ATTESTATO = 55 + SPOSTAMENTO_VERTICALE
Y_CERTIFICA = 75 + SPOSTAMENTO_VERTICALE
Y_NOME = 95 + SPOSTAMENTO_VERTICALE
Y_ORE = 115 + SPOSTAMENTO_VERTICALE
Y_CORSO = 135 + SPOSTAMENTO_VERTICALE
Y_DATA = 165 + SPOSTAMENTO_VERTICALE  # Spazio maggiore qui per ospitare le 2 righe del corso

# --- COORDINATE LOGO SINISTRA ---
MARGINE_SUPERIORE = 15  
MARGINE_LATERALE = 15   

X_LOGO = MARGINE_LATERALE
Y_LOGO = MARGINE_SUPERIORE
LARGHEZZA_LOGO = 40     

FILE_DATI = {
    "PARTECIPANTI_LIVE_9ORE.xlsx": {
        "ore": "9",
        "date": "11, 20 maggio e 3 giugno 2026"
    },
    "PARTECIPANTI_PRESENZA_16ORE.xlsx": {
        "ore": "16",
        "date": "12, 13 e 14 ottobre 2026" 
    }
}

COLONNA_NOME = "NOME"
COLONNA_COGNOME = "COGNOME"

def genera_attestati_fpdf_allineato():
    print("--- INIZIO GENERAZIONE ATTESTATI ---")
    
    # Svuota e ricrea la cartella di output
    if os.path.exists(OUTPUT_DIR):
        try:
            print(f"Elimino la cartella precedente '{OUTPUT_DIR}'...")
            shutil.rmtree(OUTPUT_DIR)
        except Exception as e:
            print(f"ERRORE: Impossibile eliminare la cartella. Assicurati che i PDF siano chiusi. Dettagli: {e}")
            return
            
    os.makedirs(OUTPUT_DIR)
        
    conteggio_file = 0

    for file_excel, info_corso in FILE_DATI.items():
        if not os.path.exists(file_excel):
            continue
            
        ore = info_corso["ore"]
        date_corso = info_corso["date"]
            
        try:
            df = pd.read_excel(file_excel)
            df = df.dropna(subset=[COLONNA_NOME, COLONNA_COGNOME])
            
            for index, row in df.iterrows():
                nome_originale = str(row[COLONNA_NOME]).strip().title()
                cognome_originale = str(row[COLONNA_COGNOME]).strip().title()
                nome_completo = f"{cognome_originale} {nome_originale}"
                
                nome_sicuro = nome_originale.replace(' ', '_')
                cognome_sicuro = cognome_originale.replace(' ', '_')
                
                nome_file_base = f"Attestato_{nome_sicuro}_{cognome_sicuro}_{ore}"
                percorso_pdf = os.path.join(OUTPUT_DIR, f"{nome_file_base}.pdf")
                
                contatore = 1
                while os.path.exists(percorso_pdf):
                    percorso_pdf = os.path.join(OUTPUT_DIR, f"{nome_file_base}_{contatore}.pdf")
                    contatore += 1
                
                pdf = FPDF(orientation="L", unit="mm", format="A4")
                
                if os.path.exists(FONT_REGULAR_PATH):
                    pdf.add_font("CustomFont", style="", fname=FONT_REGULAR_PATH, uni=True)
                    font_family = "CustomFont"
                    font_style_bold = ""
                else:
                    font_family = "helvetica"
                    font_style_bold = "B"

                if os.path.exists(FONT_BOLD_PATH):
                    pdf.add_font("CustomFont", style="B", fname=FONT_BOLD_PATH, uni=True)
                    font_style_bold = "B"
                
                pdf.add_page()
                
                if os.path.exists(SFONDO_PATH):
                    pdf.image(SFONDO_PATH, x=0, y=0, w=297, h=210)
                
                pdf.set_font(font_family, style=font_style_bold, size=28)
                pdf.set_text_color(*COLORE_VIOLA)
                pdf.set_xy(0, Y_ATTESTATO)
                pdf.cell(w=297, h=10, text="ATTESTATO DI PARTECIPAZIONE", align="C")
                
                pdf.set_font(font_family, style="", size=18)
                pdf.set_text_color(*COLORE_TESTO)
                pdf.set_xy(0, Y_CERTIFICA)
                pdf.cell(w=297, h=10, text="Si certifica che", align="C")
                
                pdf.set_font(font_family, style=font_style_bold, size=34)
                pdf.set_text_color(0, 0, 0)
                pdf.set_xy(0, Y_NOME) 
                pdf.cell(w=297, h=10, text=nome_completo, align="C")
                
                pdf.set_font(font_family, style="", size=16)
                pdf.set_text_color(*COLORE_TESTO)
                pdf.set_xy(0, Y_ORE)
                testo_ore = f"ha partecipato al corso di {ore} ore"
                pdf.cell(w=297, h=10, text=testo_ore, align="C")
                
                # Calcolo esatto per centrare orizzontalmente il blocco multilinea:
                # Larghezza foglio (297) - Larghezza testo (220) / 2 = 38.5
                pdf.set_font(font_family, style=font_style_bold, size=22)
                pdf.set_text_color(*COLORE_VIOLA)
                pdf.set_xy(38.5, Y_CORSO) 
                titolo_corso = "Introduzione dell'Intelligenza Artificiale nell'Organizzazione e Gestione dei Servizi Comunali"
                pdf.multi_cell(w=220, h=9, text=titolo_corso, align="C")
                
                pdf.set_font(font_family, style="", size=16)
                pdf.set_text_color(*COLORE_TESTO)
                pdf.set_xy(0, Y_DATA)
                testo_data = f"tenutosi i giorni {date_corso}"
                pdf.cell(w=297, h=10, text=testo_data, align="C")
                
                if os.path.exists(LOGO_PATH):
                    pdf.image(LOGO_PATH, x=X_LOGO, y=Y_LOGO, w=LARGHEZZA_LOGO)
                
                pdf.output(percorso_pdf)
                conteggio_file += 1
                
        except Exception as e:
            print(f"ERRORE Excel o PDF: {e}")

    print(f"\n--- PROCEDURA COMPLETATA: Creati {conteggio_file} file PDF in '{OUTPUT_DIR}' ---")

if __name__ == "__main__":
    genera_attestati_fpdf_allineato()