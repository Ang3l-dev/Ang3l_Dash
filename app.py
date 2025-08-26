import streamlit as st
import pandas as pd
import json
from datetime import datetime
from io import BytesIO
from pathlib import Path

# ——— Config ———
st.set_page_config(page_title="Sielte Gestione WIP", layout="wide")
BASE   = Path(__file__).parent
STYLE  = BASE / "style.css"
LOGO   = BASE / "sielte_logo.png"   # metti qui il logo Sielte
UTENTI = BASE / "utenti.json"

# ——— Utility ———
def load_users() -> list:
    """Carica utenti da utenti.json, se manca ritorna lista vuota."""
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
    """Inietta CSS + logo (base64) + footer. Evita duplicati."""
    # CSS esterno
    if STYLE.exists():
        css = STYLE.read_text(encoding="utf-8")
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

    # Logo fisso (base64)
    if LOGO.exists():
        import base64
        data = base64.b64encode(LOGO.read_bytes()).decode()
        st.markdown(
            f'<img src="data:image/png;base64,{data}" class="app-logo">',
            unsafe_allow_html=True,
        )

    # Footer (una sola volta)
    st.markdown(
        '<div class="app-powered">Powered by Ang3l-Dev</div>',
        unsafe_allow_html=True,
    )

    # Stile per i bottoni nei form
    st.markdown(
        """
        <style>
          div[data-testid="stForm"] button {
            background-color: #00AEEF !important;
            color: #FFFFFF !important;
            border: 2px solid #0077C8 !important;
            border-radius: 8px !important;
            padding: 0.75rem 1.5rem !important;
            font-size: 1rem !important;
            cursor: pointer !important;
          }
          div[data-testid="stForm"] button:hover {
            background-color: #0077C8 !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


# -------------------------------
# Helper parsing TXT WIP
# -------------------------------
COLONNE_WIP = [
    "Divisione",
    "Descr.Divisione",
    "Impresa",
    "Descr.Impresa",
    "Materiale",
    "Descr.Materiale",
    "Codice WBS",
    "Stato WBS",
    "Anno",
    "UM",
    "Quantità in lavorazione",
    "Valore in lavorazione",
    "Quantità RL Acq./Val.",
    "Valore RL Acq./Val.",
    "Quantità RL consunt.",
    "Valore RL consunt.",
    "Quantità Rientro da Lav.",
    "Valore Rientro da Lav.",
    "Valuta",
    "Numero proposta",
    "Tipo proposta",
    "Cup",
    "Cig",
    "Raggr.WBE",
]


def parse_txt_file(uploaded_file) -> pd.DataFrame:
    """Parsa un file TXT in DataFrame con le 24 colonne attese."""
    try:
        raw = uploaded_file.read()
        # Rileggere ogni volta richiede reset del puntatore; evitiamo di riusare lo stesso file dopo
        text = raw.decode("cp1252", errors="ignore")
        rows = []
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("---") or line.startswith("|Divisione"):
                continue
            # strip dei bordi e split su pipe
            campi = [c.strip() for c in line.strip("|").split("|")]
            if len(campi) == 24:
                rows.append(campi)
        if rows:
            return pd.DataFrame(rows, columns=COLONNE_WIP)
    except Exception as e:
        st.error(f"Errore nel parsing di un file TXT: {e}")
    return pd.DataFrame(columns=COLONNE_WIP)


# -------------------------------
# Funzioni principali
# -------------------------------

def unione_wip():
    st.subheader("Unione WIP")

    files = st.file_uploader(
        "Carica esattamente 8 file TXT",
        type=["txt"],
        accept_multiple_files=True,
        key="txt_uploader",
    )

    # Evita azione se non ci sono file o numero errato
    if not files:
        st.info("⬆️ Carica 8 file .txt per abilitare l'unione.")
        return
    if len(files) != 8:
        st.warning("⚠️ Devi caricare esattamente 8 file.")
        return

    if st.button("Unisci file"):
        df_unificato = pd.DataFrame(columns=COLONNE_WIP)
        for f in files:
            df_temp = parse_txt_file(f)
            if not df_temp.empty:
                df_unificato = pd.concat([df_unificato, df_temp], ignore_index=True)

        # Scrittura XLSX su buffer
        towrite = BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            df_unificato.to_excel(writer, index=False, sheet_name="WIP")
        towrite.seek(0)

        st.success("✅ File unificato creato.")
        st.download_button(
            "📥 Scarica file",
            data=towrite.getvalue(),
            file_name="WIP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def verifica_wbe():
    st.subheader("Verifica WBE e Matricole")

    file_unificato = st.file_uploader("Carica file unificato", type=["xlsx"], key="unificato_xlsx")
    file_lut_wbe   = st.file_uploader("Carica LUT WBE", type=["xlsx"], key="lut_wbe")
    file_lut_nmu   = st.file_uploader("Carica LUT NMU", type=["xlsx"], key="lut_nmu")

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

        # Normalizzazione a stringa per i confronti
        wbe_mancanti = set(df_unificato["Codice WBS"].astype(str)) - set(df_wbe["WBE"].astype(str))
        mat_mancanti = set(df_unificato["Materiale"].astype(str)) - set(df_nmu["Materiale"].astype(str))

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            if wbe_mancanti:
                pd.DataFrame({"WBE mancanti": sorted(list(wbe_mancanti))}).to_excel(
                    writer, index=False, sheet_name="WBE mancanti"
                )
            else:
                pd.DataFrame(["Nessuna WBE mancante"]).to_excel(
                    writer, index=False, header=False, sheet_name="WBE mancanti"
                )

            if mat_mancanti:
                pd.DataFrame({"Matricole mancanti": sorted(list(mat_mancanti))}).to_excel(
                    writer, index=False, sheet_name="Matricole mancanti"
                )
            else:
                pd.DataFrame(["Nessuna Matricola mancante"]).to_excel(
                    writer, index=False, header=False, sheet_name="Matricole mancanti"
                )
        output.seek(0)

        st.download_button(
            "📥 Scarica report",
            data=output.getvalue(),
            file_name="report_mancanti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if wbe_mancanti or mat_mancanti:
            st.warning("⚠️ Alcune WBE o Matricole mancano nelle LUT.")
        else:
            st.success("✅ Tutto OK! Nessuna mancanza.")


def aggiorna_storico():
    st.subheader("Aggiorna Storico")

    file_unificato = st.file_uploader("File Unificato", type="xlsx", key="unificato_storico")
    file_lut       = st.file_uploader("File LUT_WBE", type="xlsx", key="lut_wbe_storico")
    file_storico   = st.file_uploader("File Storico", type="xlsx", key="storico_file")

    if st.button("Aggiorna"):
        if not file_unificato or not file_lut or not file_storico:
            st.warning("⚠️ Carica tutti i file richiesti.")
            return

        oggi = datetime.today().strftime("%Y-%m-%d")
        MAX_RIGHE_EXCEL = 1_000_000

        try:
            df = pd.read_excel(file_unificato)[["Codice WBS", "Valore in lavorazione"]]
            df["DataAggiornamento"] = oggi
            df_lut = pd.read_excel(file_lut)[["WBE", "Area"]]
            df = pd.merge(df, df_lut, left_on="Codice WBS", right_on="WBE", how="left")
            df_storico = pd.read_excel(file_storico)
        except Exception as e:
            st.error(f"Errore lettura/merge Excel: {e}")
            return

        if len(df_storico) >= MAX_RIGHE_EXCEL:
            # Se troppo grande, fai rollover: nuovo storico = df corrente,
            # e offri in download il backup del vecchio (non salvi su disco su Cloud)
            st.warning("⚠️ File storico troppo grande. Eseguo rollover (solo nuovo contenuto).")
            df_finale = df
        else:
            df_finale = pd.concat([df_storico, df], ignore_index=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_finale.to_excel(writer, index=False, sheet_name="Storico")
        output.seek(0)

        st.success("✅ Storico aggiornato.")
        st.download_button(
            "📥 Scarica storico aggiornato",
            data=output.getvalue(),
            file_name="storico_dati.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# -------------------------------
# Login & Routing
# -------------------------------

def main():
    inject_css()

    # Login
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
            for k in ("user", "step"):
                if k in st.session_state:
                    del st.session_state[k]
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
