import streamlit as st
import pandas as pd
import os, json
from datetime import datetime
from io import BytesIO
from pathlib import Path

# ——— Config ———
st.set_page_config(page_title="Sielte Gestione WIP", layout="wide")
BASE     = Path(__file__).parent
STYLE    = BASE / "style.css"
LOGO     = BASE / "sielte_logo.png"   # metti qui il logo Sielte
UTENTI   = BASE / "utenti.json"

# ——— Utility ———
def load_users():
    return json.loads(UTENTI.read_text(encoding="utf-8"))

def check_login(email, pwd):
    for u in load_users():
        if u["email"] == email and u["password"] == pwd:
            return u
    return None

def inject_css():
    # Carica e inietta lo stylesheet esterno
    if STYLE.exists():
        css = STYLE.read_text(encoding="utf-8")
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
    # Logo fisso
    if LOGO.exists():
        st.markdown(
            f'<img src="data:image/png;base64,{LOGO.read_bytes().hex()}" class="app-logo">',
            unsafe_allow_html=True
        )
    # Footer
    st.markdown(
        '<div class="app-powered">Powered by Ang3l-Dev</div>',
        unsafe_allow_html=True
    )

    st.markdown("""
    <style>
      /* selettore ultra-specifico per il button del form_submit_button */
      div[data-testid="stForm"] button {
        background-color: #00AEEF !important;
        color:            #FFFFFF !important;
        border:           2px solid #0077C8 !important;
        border-radius:    8px !important;
        padding:          0.75rem 1.5rem !important;
        font-size:        1rem !important;
        cursor:           pointer !important;
      }
      div[data-testid="stForm"] button:hover {
        background-color: #0077C8 !important;
      }
    </style>
    """, unsafe_allow_html=True)
    if LOGO.exists():
            import base64
            data = base64.b64encode(LOGO.read_bytes()).decode()
            st.markdown(f'<img src="data:image/png;base64,{data}" class="app-logo">', unsafe_allow_html=True)

    st.markdown('<div class="app-powered">Powered by Ang3l-Dev</div>', unsafe_allow_html=True)
# -------------------------------
# Funzioni principali
# -------------------------------
def unione_wip():
    st.subheader("Unione WIP")

    files = st.file_uploader("Carica esattamente 8 file TXT", type=["txt"], accept_multiple_files=True)
    if files and len(files) != 8:
        st.warning("⚠️ Devi caricare esattamente 8 file.")
        return

    if st.button("Unisci file"):
        colonne = [
            "Divisione", "Descr.Divisione", "Impresa", "Descr.Impresa", "Materiale",
            "Descr.Materiale", "Codice WBS", "Stato WBS", "Anno", "UM", "Quantità in lavorazione",
            "Valore in lavorazione", "Quantità RL Acq./Val.", "Valore RL Acq./Val.",
            "Quantità RL consunt.", "Valore RL consunt.", "Quantità Rientro da Lav.",
            "Valore Rientro da Lav.", "Valuta", "Numero proposta", "Tipo proposta", "Cup",
            "Cig", "Raggr.WBE"
        ]
        df_unificato = pd.DataFrame(columns=colonne)
        for f in files:
            lines = f.read().decode("cp1252").splitlines()
            dati = []
            for line in lines:
                line = line.strip()
                if line.startswith("---") or line.startswith("|Divisione") or line == "":
                    continue
                campi = [c.strip() for c in line.strip('|').split('|')]
                if len(campi) == 24:
                    dati.append(campi)
            if dati:
                df_temp = pd.DataFrame(dati, columns=colonne)
                df_unificato = pd.concat([df_unificato, df_temp], ignore_index=True)

        towrite = BytesIO()
        df_unificato.to_excel(towrite, index=False)
        st.success("✅ File unificato creato.")
        st.download_button("📥 Scarica file", data=towrite.getvalue(), file_name="WIP.xlsx")

def verifica_wbe():
    st.subheader("Verifica WBE e Matricole")

    file_unificato = st.file_uploader("Carica file unificato", type=["xlsx"])
    file_lut_wbe = st.file_uploader("Carica LUT WBE", type=["xlsx"])
    file_lut_nmu = st.file_uploader("Carica LUT NMU", type=["xlsx"])

    if st.button("Verifica"):
        if not file_unificato or not file_lut_wbe or not file_lut_nmu:
            st.warning("⚠️ Carica tutti i file richiesti.")
            return

        df_unificato = pd.read_excel(file_unificato)
        df_wbe = pd.read_excel(file_lut_wbe)
        df_nmu = pd.read_excel(file_lut_nmu)

        wbe_mancanti = set(df_unificato["Codice WBS"].astype(str)) - set(df_wbe["WBE"].astype(str))
        mat_mancanti = set(df_unificato["Materiale"].astype(str)) - set(df_nmu["Materiale"].astype(str))

        with BytesIO() as output:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if wbe_mancanti:
                    pd.DataFrame({"WBE mancanti": list(wbe_mancanti)}).to_excel(writer, index=False, sheet_name="WBE mancanti")
                else:
                    pd.DataFrame(["Nessuna WBE mancante"]).to_excel(writer, index=False, header=False, sheet_name="WBE mancanti")

                if mat_mancanti:
                    pd.DataFrame({"Matricole mancanti": list(mat_mancanti)}).to_excel(writer, index=False, sheet_name="Matricole mancanti")
                else:
                    pd.DataFrame(["Nessuna Matricola mancante"]).to_excel(writer, index=False, header=False, sheet_name="Matricole mancanti")

            st.download_button("📥 Scarica report", data=output.getvalue(), file_name="report_mancanti.xlsx")

        if wbe_mancanti or mat_mancanti:
            st.warning("⚠️ Alcune WBE o Matricole mancano nelle LUT.")
        else:
            st.success("✅ Tutto OK! Nessuna mancanza.")

def aggiorna_storico():
    st.subheader("Aggiorna Storico")

    file_unificato = st.file_uploader("File Unificato", type="xlsx", key="unificato")
    file_lut = st.file_uploader("File LUT_WBE", type="xlsx", key="lut")
    file_storico = st.file_uploader("File Storico", type="xlsx", key="storico")

    if st.button("Aggiorna"):
        if not file_unificato or not file_lut or not file_storico:
            st.warning("⚠️ Carica tutti i file richiesti.")
            return

        oggi = datetime.today().strftime('%Y-%m-%d')
        MAX_RIGHE_EXCEL = 1_000_000

        df = pd.read_excel(file_unificato)[['Codice WBS', 'Valore in lavorazione']]
        df['DataAggiornamento'] = oggi
        df_lut = pd.read_excel(file_lut)[['WBE', 'Area']]
        df = pd.merge(df, df_lut, left_on='Codice WBS', right_on='WBE', how='left')
        df_storico = pd.read_excel(file_storico)

        if len(df_storico) >= MAX_RIGHE_EXCEL:
            timestamp = datetime.now().strftime('%Y-%m-%d_%H%M%S')
            file_backup = f"storico_backup_{timestamp}.xlsx"
            df_storico.to_excel(file_backup, index=False)
            st.warning(f"⚠️ File storico troppo grande. Salvato backup: {file_backup}")
            df_finale = df
        else:
            df_finale = pd.concat([df_storico, df], ignore_index=True)

        output = BytesIO()
        df_finale.to_excel(output, index=False)
        st.success("✅ Storico aggiornato.")
        st.download_button("📥 Scarica storico aggiornato", data=output.getvalue(), file_name="storico_dati.xlsx")

# -------------------------------
# Login & Routing
# -------------------------------
def main():
    inject_css()

    # Login
    if "user" not in st.session_state:
        with st.form("login_form", clear_on_submit=False):
            st.title("Login")
            email    = st.text_input("Email")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("Accedi"):
                user = check_login(email, password)
                if user:
                    st.session_state.user = user
                    st.rerun()
                else:
                    st.error("Credenziali errate")
        return

    # Dopo il login
    st.title(f"Benvenuto, {st.session_state.user['email']}")

    # Primo menu
    if "step" not in st.session_state:
        st.session_state.step = "main"

    if st.session_state.step == "main":
        col1, col2 = st.columns(2)
        if col1.button("Gestione WIP"):
            st.session_state.step = "wip_menu"
            st.rerun()
        if col2.button("Logout"):
            del st.session_state.user
            del st.session_state.step
            st.rerun()
        return

    # Sub-menu WIP
    if st.session_state.step == "wip_menu":
        c1, c2, c3, c4 = st.columns(4)
        if c1.button("🔁 Unione WIP"):
            st.session_state.step = "unione"
            st.rerun()
        if c2.button("📊 Verifica WBE"):
            st.session_state.step = "verifica"
            st.rerun()
        if c3.button("🕒 Aggiorna Storico"):
            st.session_state.step = "storico"
            st.rerun()
        if c4.button("🔙 Torna al menu"):
            st.session_state.step = "main"
            st.rerun()
        return

    # Livello funzionale
    if st.session_state.step == "unione":
        unione_wip()
    elif st.session_state.step == "verifica":
        verifica_wbe()
    elif st.session_state.step == "storico":
        aggiorna_storico()

    if st.button("🔙 Torna al menu WIP"):
        st.session_state.step = "wip_menu"
        st.rerun()

if __name__ == "__main__":
    main()