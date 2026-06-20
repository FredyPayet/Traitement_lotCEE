import streamlit as st
import pandas as pd
import subprocess
import os
import tempfile
import io

st.set_page_config(page_title="Tableaux par client", layout="wide")

st.title("📊 Tableaux par client — Contrôles CEE")
st.markdown("Chargez votre fichier Excel (`.xls` ou `.xlsx`) pour générer automatiquement les tableaux par client.")

# ─── Labels courts pour les colonnes C–I et BS ──────────────────────────────
COL_LABELS = {
    "C": "Réf. EMMY",
    "D": "Réf. interne",
    "E": "Nom du site",
    "F": "Adresse",
    "G": "Code postal",
    "H": "Ville",
    "I": "Raison sociale bénéficiaire",
    "BS": "Conclusion de l'audit",
}

# ─── Helpers ────────────────────────────────────────────────────────────────

def convert_xls_to_xlsx(xls_path: str) -> str:
    """Convert legacy .xls to .xlsx via LibreOffice; return new path."""
    out_dir = tempfile.mkdtemp()
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "xlsx",
         "--outdir", out_dir, xls_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"Conversion LibreOffice échouée :\n{result.stderr}")
    base = os.path.splitext(os.path.basename(xls_path))[0]
    return os.path.join(out_dir, base + ".xlsx")


def load_dataframe(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Load the Excel file into a DataFrame, auto-converting .xls if needed."""
    ext = os.path.splitext(filename)[1].lower()

    if ext in (".xlsx", ".xlsm"):
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    else:
        # Legacy .xls → convert first
        with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        try:
            xlsx_path = convert_xls_to_xlsx(tmp_path)
            df = pd.read_excel(xlsx_path, header=None)
        finally:
            os.unlink(tmp_path)

    return df


def extract_tables(df: pd.DataFrame):
    """
    Returns:
      header_row  – index of the header row (row containing 'C', 'REFERENCE EMMY…')
      col_map     – dict mapping letter → column integer index (0-based)
      clients     – sorted list of unique client names
      data        – DataFrame with named columns (C-I + BS), client column = 'I'
    """
    # Find the header row: the row whose column B contains 'SIREN' (case-insensitive)
    header_idx = None
    for i, row in df.iterrows():
        for val in row:
            if isinstance(val, str) and "REFERENCE EMMY" in val.upper():
                header_idx = i
                break
        if header_idx is not None:
            break

    if header_idx is None:
        raise ValueError(
            "Impossible de trouver la ligne d'en-tête (cherche 'REFERENCE EMMY')."
        )

    # Build column map by letter
    # Column letters in Excel: A=0, B=1, … Z=25, AA=26 … BS = ?
    def col_letter_to_idx(letter: str) -> int:
        letter = letter.upper().strip()
        idx = 0
        for ch in letter:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1  # 0-based

    col_c  = col_letter_to_idx("C")
    col_d  = col_letter_to_idx("D")
    col_e  = col_letter_to_idx("E")
    col_f  = col_letter_to_idx("F")
    col_g  = col_letter_to_idx("G")
    col_h  = col_letter_to_idx("H")
    col_i  = col_letter_to_idx("I")
    col_bs = col_letter_to_idx("BS")

    needed_cols = [col_c, col_d, col_e, col_f, col_g, col_h, col_i, col_bs]
    max_col = max(needed_cols)

    if max_col >= df.shape[1]:
        raise ValueError(
            f"Le fichier n'a que {df.shape[1]} colonnes "
            f"(colonne BS = index {col_bs} attendue)."
        )

    # Data rows start after header
    data_raw = df.iloc[header_idx + 1 :, needed_cols].copy()
    data_raw.columns = ["C", "D", "E", "F", "G", "H", "I", "BS"]

    # Drop rows that are entirely empty
    data_raw = data_raw.dropna(subset=["I"], how="all")
    data_raw = data_raw[data_raw["I"].astype(str).str.strip() != ""]

    if data_raw.empty:
        raise ValueError(
            "Aucune ligne de données trouvée après l'en-tête. "
            "Vérifiez que le fichier contient bien des données (colonne I remplie)."
        )

    # Fill empty BS cells
    data_raw["BS"] = data_raw["BS"].fillna("non visité")
    data_raw["BS"] = data_raw["BS"].apply(
        lambda v: "non visité" if str(v).strip() == "" else v
    )

    clients = sorted(data_raw["I"].dropna().unique().tolist())
    return data_raw.reset_index(drop=True), clients


# ─── UI ─────────────────────────────────────────────────────────────────────

uploaded = st.file_uploader(
    "Choisir un fichier Excel (.xls / .xlsx)",
    type=["xls", "xlsx", "xlsm"],
)

if uploaded:
    with st.spinner("Lecture du fichier…"):
        try:
            file_bytes = uploaded.read()
            df_raw = load_dataframe(file_bytes, uploaded.name)
            data, clients = extract_tables(df_raw)
        except Exception as e:
            st.error(f"❌ Erreur lors de la lecture : {e}")
            st.stop()

    st.success(f"✅ Fichier chargé — **{len(data)}** ligne(s) · **{len(clients)}** client(s) détecté(s)")

    # Rename columns for display
    display_cols = {
        "C": COL_LABELS["C"],
        "D": COL_LABELS["D"],
        "E": COL_LABELS["E"],
        "F": COL_LABELS["F"],
        "G": COL_LABELS["G"],
        "H": COL_LABELS["H"],
        "I": COL_LABELS["I"],
        "BS": COL_LABELS["BS"],
    }

    # ── Sidebar: client filter ──────────────────────────────────────────────
    st.sidebar.header("Filtres")
    selected = st.sidebar.multiselect(
        "Clients à afficher",
        options=clients,
        default=clients,
    )

    # ── Export all clients to Excel ─────────────────────────────────────────
    def build_excel(df_in: pd.DataFrame, clients_list: list) -> bytes:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for client in clients_list:
                df_client = df_in[df_in["I"] == client].copy()
                df_client = df_client.rename(columns=display_cols)
                # Sheet name max 31 chars, no special chars
                sheet = str(client)[:31].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "").replace(":", "")
                df_client.to_excel(writer, sheet_name=sheet, index=False)
        return buf.getvalue()

    if st.sidebar.button("⬇️ Exporter tous les clients (Excel)"):
        xlsx_bytes = build_excel(data, clients)
        st.sidebar.download_button(
            label="Télécharger le fichier Excel",
            data=xlsx_bytes,
            file_name="tableaux_par_client.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Display per-client tables ───────────────────────────────────────────
    if not selected:
        st.info("Aucun client sélectionné dans le panneau de gauche.")
    else:
        for client in selected:
            df_client = data[data["I"] == client].copy()
            df_client = df_client.rename(columns=display_cols)
            df_client = df_client.reset_index(drop=True)
            df_client.index += 1  # start at 1

            with st.expander(f"🏢 {client}  ({len(df_client)} opération(s))", expanded=True):
                # Color-code the audit conclusion column
                def color_conclusion(val):
                    v = str(val).lower()
                    if "non satisfaisant" in v:
                        return "background-color: #ffd6d6; color: #8b0000;"
                    elif "satisfaisant" in v:
                        return "background-color: #d6f5d6; color: #145214;"
                    elif "non visité" in v:
                        return "background-color: #fff3cd; color: #7d5a00;"
                    elif "inaccessible" in v or "non vérifiable" in v:
                        return "background-color: #e0e0e0; color: #444;"
                    return ""

                styled = df_client.style.applymap(
                    color_conclusion, subset=[COL_LABELS["BS"]]
                )

                st.dataframe(styled, use_container_width=True, height=min(400, 55 + 35 * len(df_client)))

                # Per-client download button
                csv = df_client.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label=f"⬇️ Télécharger CSV — {client}",
                    data=csv,
                    file_name=f"{client[:40]}.csv",
                    mime="text/csv",
                    key=f"dl_{client}",
                )