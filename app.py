import streamlit as st
import pandas as pd
import json
from datetime import datetime
from io import BytesIO
from pathlib import Path
import base64

# ==============================
# CONFIGURAZIONE BASE
# ==============================
st.set_page_config(page_title="Sielte Gestione WIP", layout="wide")
BASE   = Path(__file__).parent
STYLE  = BASE / "style.css"
LOGO   = BASE / "sielte_logo.png"
UTENTI = BASE / "utenti.json"

COLONNE_WIP = [
    "Divisione", "Descr.Divisione", "Impresa", "Descr.Impresa", "Materiale", "Descr.Materiale",
    "Codice WBS", "Stato WBS", "Anno", "UM", "Quantità in lavorazione", "Valore in lavorazione",
    "Quantità RL Acq./Val.", "Valore RL Acq./Val.", "Quantità RL consunt.", "Valore RL consunt.",
    "Quantità Rientro da Lav.", "Valore Rientro da Lav.", "Valuta", "Numero proposta",
    "Tipo proposta", "Cup", "Cig", "Raggr.WBE"
]

# ==============================
# UTILITY
# ==============================
def load_users() -> list:
    try:
        if UTENTI.exists():
            return json.loads(UTENTI.read_text(encoding="utf-8"))
    except Exception:
        pass
    return []

def check_login(email: str, pwd: str):
    for u in load_users():
        if u.get("email") == email and u.get("password") == pwd:
            return u
    return None

def inject_css():
    # CSS custom
    if STYLE.exists():
        st.markdown(f"<style>{STYLE.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
    # Logo opzionale
    if LOGO.exists():
        data = base64.b64encode(LOGO.read_bytes()).decode()
        st.markdown(f'<img src="data:image/png;base64,{data}" class="app-logo">', unsafe_allow_html=True)
    st.markdown('<div class="app-powered">Powered by Ang3l-Dev</div>', unsafe_allow_html=True)

# ==============================
# PARSER TXT WIP
# ==============================
def parse_txt_file(uploaded_file) -> pd.DataFrame:
    """
    Converte un txt in DataFrame, ignorando righe di separatori/intestazioni.
    Non rimuove righe duplicate; valida che ci siano 24 campi separati da '|'.
    """
    try:
        raw = uploaded_file.read()
        text = raw.decode("cp1252", errors="ignore")
        rows = []
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("---") or line.startswith("|Divisione"):
                continue
            # rimuove l’eventuale '|' iniziale/finale e splitta
            campi = [c.strip() for c in line.strip("|").split("|")]
            if len(campi) == 24:
                rows.append(campi)
        if rows:
            df = pd.DataFrame(rows, columns=COLONNE_WIP)
            return df
    except Exception as e:
        st.error(f"Errore parsing TXT: {e}")
    return pd.DataFrame(columns=COLONNE_WIP)

# ==============================
# FUNZIONI PRINCIPALI
# ==============================
def unione_wip():
    st.subheader("Unione WIP")

    files = st.file_uploader("Carica esattamente 8 file TXT", type=["txt"], accept_multiple_files=True)
    if not files:
        st.info("⬆️ Carica 8 file .txt per abilitare l'unione.")
        return
    if len(files) != 8:
        st.warning("⚠️ Devi caricare esattamente 8 file.")
        return

    if st.button("Unisci e genera Excel"):
        # Parse di tutti i file
        frames = []
        for f in files:
            df = parse_txt_file(f)
            frames.append(df)
        df_unificato = pd.concat(frames, ignore_index=True)

        # Salvataggio locale + download
        stamp = datetime.today().strftime('%Y%m%d')
        filename = f"WIP_{stamp}.xlsx"

        # Scrivi su buffer per il download
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df_unificato.to_excel(writer, index=False, sheet_name="WIP")
        out.seek(0)

        st.download_button(
            "📥 Scarica WIP unificato",
            data=out.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # (opzionale) salva anche in locale nella cartella del progetto
        try:
            (BASE / filename).write_bytes(out.getvalue())
            st.success(f"✅ File unificato salvato: {filename}")
        except Exception:
            # se non può scrivere, va bene lo stesso: l’utente ha il download
            pass

def verifica_wbe():
    st.subheader("Verifica WBE e Matricole")

    file_unificato = st.file_uploader("Carica file unificato", type=["xlsx"])
    file_lut_wbe   = st.file_uploader("Carica LUT WBE", type=["xlsx"])
    file_lut_nmu   = st.file_uploader("Carica LUT NMU", type=["xlsx"])

    if st.button("Verifica"):
        if not file_unificato or not file_lut_wbe or not file_lut_nmu:
            st.warning("⚠️ Carica tutti i file richiesti.")
            return

        try:
            df_unificato = pd.read_excel(file_unificato)
            df_wbe = pd.read_excel(file_lut_wbe)
            df_nmu = pd.read_excel(file_lut_nmu)
        except Exception as e:
            st.error(f"Errore lettura Excel: {e}")
            return

        wbe_mancanti = set(df_unificato["Codice WBS"].astype(str)) - set(df_wbe["WBE"].astype(str))
        mat_mancanti = set(df_unificato["Materiale"].astype(str)) - set(df_nmu["Materiale"].astype(str))

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"WBE mancanti": sorted(wbe_mancanti) or ["Nessuna WBE mancante"]}).to_excel(
                writer, index=False, sheet_name="WBE mancanti"
            )
            pd.DataFrame({"Matricole mancanti": sorted(mat_mancanti) or ["Nessuna Matricola mancante"]}).to_excel(
                writer, index=False, sheet_name="Matricole mancanti"
            )
        output.seek(0)

        st.download_button(
            "📥 Scarica report",
            data=output.getvalue(),
            file_name="report_mancanti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if wbe_mancanti or mat_mancanti:
            st.warning("⚠️ Alcune WBE o Matricole mancano nelle LUT.")
        else:
            st.success("✅ Tutto OK! Nessuna mancanza.")

def aggiorna_storico():
    st.subheader("Aggiorna Storico")

    file_unificato = st.file_uploader("File Unificato", type="xlsx")
    file_lut       = st.file_uploader("File LUT_WBE", type="xlsx")
    file_storico   = st.file_uploader("File Storico", type="xlsx")

    if st.button("Aggiorna e genera Excel"):
        if not file_unificato or not file_lut or not file_storico:
            st.warning("⚠️ Carica tutti i file richiesti.")
            return

        oggi = datetime.today().strftime("%Y-%m-%d")

        try:
            df = pd.read_excel(file_unificato)[["Codice WBS", "Valore in lavorazione"]]
            df["DataAggiornamento"] = oggi

            df_lut = pd.read_excel(file_lut)[["WBE", "Area"]]
            df = pd.merge(df, df_lut, left_on="Codice WBS", right_on="WBE", how="left")

            df_storico = pd.read_excel(file_storico)
            df_storico = df_storico[df_storico["DataAggiornamento"] != oggi]

            df_finale = pd.concat([df_storico, df], ignore_index=True)
            df_finale.drop_duplicates(subset=["DataAggiornamento", "Codice WBS"], inplace=True)

            # Buffer per download
            out = BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df_finale.to_excel(writer, index=False, sheet_name="Storico")
            out.seek(0)

            st.download_button(
                "📥 Scarica storico aggiornato",
                data=out.getvalue(),
                file_name="storico_dati.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # (opzionale) salva anche in locale
            try:
                (BASE / "storico_dati.xlsx").write_bytes(out.getvalue())
                st.success("✅ Storico aggiornato e salvato in locale (storico_dati.xlsx)")
            except Exception:
                pass

        except Exception as e:
            st.error(f"Errore aggiornamento: {e}")

# ==============================
# LOGIN & ROUTING
# ==============================
def main():
    inject_css()

    if "user" not in st.session_state:
        with st.form("login_form", clear_on_submit=False):
            st.title("Login")
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("Accedi"):
                user = check_login(email, password)
                if user:
                    st.session_state.user = user
                    st.rerun()
                else:
                    st.error("Credenziali errate")
        return

    st.title(f"Benvenuto, {st.session_state.user['email']}")

    if "step" not in st.session_state:
        st.session_state.step = "main"

    if st.session_state.step == "main":
        col1, col2 = st.columns(2)
        if col1.button("Gestione WIP"):
            st.session_state.step = "wip_menu"; st.rerun()
        if col2.button("Logout"):
            for k in ("user", "step"):
                st.session_state.pop(k, None)
            st.rerun()
        return

    if st.session_state.step == "wip_menu":
        c1, c2, c3, c4 = st.columns(4)
        if c1.button("🔁 Unione WIP"):
            st.session_state.step = "unione"; st.rerun()
        if c2.button("📊 Verifica WBE"):
            st.session_state.step = "verifica"; st.rerun()
        if c3.button("🕒 Aggiorna Storico"):
            st.session_state.step = "storico"; st.rerun()
        if c4.button("🔙 Torna al menu"):
            st.session_state.step = "main"; st.rerun()
        return

    if st.session_state.step == "unione":
        unione_wip()
    elif st.session_state.step == "verifica":
        verifica_wbe()
    elif st.session_state.step == "storico":
        aggiorna_storico()

    if st.button("🔙 Torna al menu WIP"):
        st.session_state.step = "wip_menu"; st.rerun()

if __name__ == "__main__":
    main()

