import streamlit as st
import pandas as pd
import subprocess
import os
import re
import tempfile
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Tableaux par client", layout="wide")

st.title("📊 Tableaux par client — Contrôles CEE")
st.markdown("Chargez votre fichier Excel (`.xls` ou `.xlsx`) pour générer automatiquement un fichier Excel par client.")

# ─── Labels courts pour les colonnes D–I, BS, BY ────────────────────────────
COL_LABELS = {
    "D": "Réf. interne",
    "E": "Nom du site",
    "F": "Adresse",
    "G": "Code postal",
    "H": "Ville",
    "I": "Raison sociale bénéficiaire",
    "BS": "Conclusion de l'audit",
    "BY": "Conformité après correction",
}

# ─── Couleurs mise en forme Excel ───────────────────────────────────────────
COLORS = {
    "satisfaisant":    "C6EFCE",  # vert clair
    "non_satisfaisant":"FFCCCC",  # rouge clair
    "inaccessible":    "FFE0B2",  # orange clair
    "non_visite":      "E0E0E0",  # gris clair
    "header":          "2F5496",  # bleu foncé header
}

# ─── Helpers ────────────────────────────────────────────────────────────────

def col_letter_to_idx(letter: str) -> int:
    idx = 0
    for ch in letter.upper().strip():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def convert_xls_to_xlsx(xls_path: str) -> str:
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
    ext = os.path.splitext(filename)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(io.BytesIO(file_bytes), header=None)
    else:
        with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        try:
            xlsx_path = convert_xls_to_xlsx(tmp_path)
            return pd.read_excel(xlsx_path, header=None)
        finally:
            os.unlink(tmp_path)


def extract_tables(df: pd.DataFrame):
    # Trouver la ligne d'en-tête
    header_idx = None
    for i, row in df.iterrows():
        for val in row:
            if isinstance(val, str) and "REFERENCE EMMY" in val.upper():
                header_idx = i
                break
        if header_idx is not None:
            break

    if header_idx is None:
        raise ValueError("Impossible de trouver la ligne d'en-tête (cherche 'REFERENCE EMMY').")

    needed_cols = [col_letter_to_idx(l) for l in ["D", "E", "F", "G", "H", "I", "BS", "BY"]]
    max_col = max(needed_cols)

    if max_col >= df.shape[1]:
        raise ValueError(f"Le fichier n'a que {df.shape[1]} colonnes (colonne BY = index {col_letter_to_idx('BY')} attendue).")

    data_raw = df.iloc[header_idx + 1:, needed_cols].copy()
    data_raw.columns = ["D", "E", "F", "G", "H", "I", "BS", "BY"]

    # Supprimer lignes sans client
    data_raw = data_raw.dropna(subset=["I"], how="all")
    data_raw = data_raw[data_raw["I"].astype(str).str.strip() != ""]

    if data_raw.empty:
        raise ValueError("Aucune ligne de données trouvée. Vérifiez que la colonne I est remplie.")

    # Remplir cellules vides BS et BY
    for col in ["BS", "BY"]:
        data_raw[col] = data_raw[col].fillna("non visité")
        data_raw[col] = data_raw[col].apply(lambda v: "non visité" if str(v).strip() == "" else v)

    clients = sorted(data_raw["I"].dropna().unique().tolist())
    return data_raw.reset_index(drop=True), clients


def sanitize_filename(name: str) -> str:
    """Supprime les caractères interdits dans les noms de fichiers Windows."""
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


def get_conclusion_summary(df_client: pd.DataFrame) -> str:
    """Résumé des conclusions pour le nom de fichier."""
    bs_vals = df_client["BS"].astype(str).str.lower()
    if bs_vals.str.contains("non satisfaisant").any():
        return "non_satisfaisant"
    elif bs_vals.str.contains("inaccessible|non vérifiable").any():
        return "inaccessible"
    elif bs_vals.str.contains("satisfaisant").any():
        return "satisfaisant"
    return "non_visite"


def get_ref_prefix(df_client: pd.DataFrame) -> str:
    """6 premiers chiffres de la colonne D (première valeur non vide)."""
    for val in df_client["D"].dropna():
        digits = re.sub(r'\D', '', str(val))
        if digits:
            return digits[:6]
    return "000000"


def get_cell_color(value: str) -> str | None:
    v = str(value).lower()
    if "non satisfaisant" in v:
        return COLORS["non_satisfaisant"]
    elif "satisfaisant" in v:
        return COLORS["satisfaisant"]
    elif "inaccessible" in v or "non vérifiable" in v:
        return COLORS["inaccessible"]
    elif "non visité" in v:
        return COLORS["non_visite"]
    return None


def build_client_excel(df_client: pd.DataFrame, client_name: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = client_name[:31]

    col_keys = ["D", "E", "F", "G", "H", "I", "BS", "BY"]
    headers = [COL_LABELS[k] for k in col_keys]
    bs_col_idx = col_keys.index("BS") + 1  # 1-based
    by_col_idx = col_keys.index("BY") + 1

    # Styles
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    header_fill = PatternFill("solid", fgColor=COLORS["header"])
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(vertical="center", wrap_text=True)

    # En-tête
    ws.row_dimensions[1].height = 40
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    # Données
    for row_num, (_, row) in enumerate(df_client.iterrows(), 2):
        ws.row_dimensions[row_num].height = 20
        for col_num, key in enumerate(col_keys, 1):
            val = row[key]
            cell = ws.cell(row=row_num, column=col_num, value=val)
            cell.alignment = cell_align
            cell.border = border
            cell.font = Font(name="Arial", size=10)

            # Colorier BS et BY
            if col_num in (bs_col_idx, by_col_idx):
                color = get_cell_color(str(val))
                if color:
                    cell.fill = PatternFill("solid", fgColor=color)

    # Largeurs de colonnes
    col_widths = [20, 30, 35, 12, 20, 35, 30, 25]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_filename(client_name: str, df_client: pd.DataFrame) -> str:
    conclusion = get_conclusion_summary(df_client)
    ref = get_ref_prefix(df_client)
    name = sanitize_filename(client_name)[:50]
    return f"{name}_{conclusion}_{ref}.xlsx"


# ─── Messages automatiques ──────────────────────────────────────────────────

MESSAGES = {
    "Taux OK": (
        "Bonjour,\n\n"
        "Pour votre information, nous avons reçu le retour du lot de contrôle {lot}.\n\n"
        "Tous les taux réglementaires sont respectés. Toutes les opérations du lot relatif "
        "à la fiche travaux peuvent être finalisées.\n\n"
        "Vous trouverez ci-joint les résultats de contrôles pour vos opérations.\n\n"
        "Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
        "Cordialement,"
    ),
    "Taux NS KO": (
        "Bonjour,\n\n"
        "Pour votre information, nous avons reçu le retour du lot de contrôle {lot}.\n\n"
        "Le taux d'opérations contrôlées non satisfaisantes dépasse les 10 %. "
        "Nous ne pouvons finaliser que les opérations qui ont été contrôlées. "
        "Les opérations non visitées doivent être représentées dans un nouveau lot.\n\n"
        "Vous trouverez ci-joint les résultats de contrôles pour vos opérations.\n\n"
        "Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
        "Cordialement,"
    ),
    "Tous taux KO": (
        "Bonjour,\n\n"
        "Pour votre information, nous avons reçu le retour du lot de contrôle {lot}.\n\n"
        "Les taux réglementaires ne sont pas atteints. Nous ne pouvons finaliser aucune opération "
        "dans ce lot. Les opérations doivent être représentées dans un nouveau lot.\n\n"
        "Vous trouverez ci-joint les résultats de contrôles pour vos opérations.\n\n"
        "Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
        "Cordialement,"
    ),
}


def copy_button(text: str, key: str):
    """Affiche un bouton qui copie 'text' dans le presse-papiers via JavaScript."""
    # Échapper les caractères spéciaux pour l'intégration JS
    escaped = text.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")
    st.components.v1.html(
        f"""
        <button onclick="
            navigator.clipboard.writeText(`{escaped}`)
            .then(() => {{
                this.innerText = '✅ Copié !';
                setTimeout(() => this.innerText = '📋 Copier le message', 2000);
            }})
            .catch(() => alert('Copie impossible, utilisez Ctrl+C'));
        "
        style="
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 8px 18px;
            border-radius: 6px;
            font-size: 14px;
            cursor: pointer;
            font-family: Arial, sans-serif;
        ">
        📋 Copier le message
        </button>
        """,
        height=45,
    )


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

    # ── Informations lot ─────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📋 Informations du lot")

    col1, col2, col3 = st.columns(3)
    with col1:
        num_lot = st.text_input("Numéro de lot", placeholder="ex : LOT-2024-001")
    with col2:
        num_dossier = st.text_input("Numéro de dossier(s)", placeholder="ex : DOS-001, DOS-002")
    with col3:
        taux_choix = st.selectbox(
            "Résultat du lot",
            options=["Taux OK", "Taux NS KO", "Tous taux KO"],
        )

    st.markdown("---")

    # ── Sidebar ─────────────────────────────────────────────────────────────
    st.sidebar.header("Filtres & Export")
    selected = st.sidebar.multiselect("Clients à afficher", options=clients, default=clients)

    # Bouton ZIP — tous les clients
    if st.sidebar.button("⬇️ Télécharger tous les fichiers (ZIP)"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for client in clients:
                df_c = data[data["I"] == client].copy()
                xlsx_bytes = build_client_excel(df_c, client)
                filename = build_filename(client, df_c)
                zf.writestr(filename, xlsx_bytes)
        st.sidebar.download_button(
            label="📦 Télécharger le ZIP",
            data=zip_buf.getvalue(),
            file_name="tableaux_clients.zip",
            mime="application/zip",
        )

    # ── Tableaux par client ──────────────────────────────────────────────────
    if not selected:
        st.info("Aucun client sélectionné dans le panneau de gauche.")
    else:
        for client in selected:
            df_client = data[data["I"] == client].copy().reset_index(drop=True)
            df_client.index += 1

            filename = build_filename(client, df_client)

            with st.expander(f"🏢 {client}  ({len(df_client)} opération(s))", expanded=True):

                # ── Message automatique ──────────────────────────────────────
                lot_label = num_lot.strip() if num_lot.strip() else "[numéro de lot non renseigné]"
                message = MESSAGES[taux_choix].format(lot=lot_label)

                st.markdown("**✉️ Message à envoyer au client :**")

                # Couleur du bandeau selon le taux
                couleur_bandeau = {
                    "Taux OK":      "#e8f5e9",
                    "Taux NS KO":   "#fff3e0",
                    "Tous taux KO": "#ffebee",
                }[taux_choix]

                st.markdown(
                    f"<div style='background-color:{couleur_bandeau}; padding:16px; "
                    f"border-radius:8px; white-space:pre-wrap; font-family:Arial; font-size:14px;'>"
                    f"{message}</div>",
                    unsafe_allow_html=True,
                )

                copy_button(message, key=f"msg_{client}")

                st.markdown("---")

                # ── Tableau des opérations ───────────────────────────────────
                st.markdown("**📊 Opérations :**")
                df_display = df_client.rename(columns=COL_LABELS)

                def color_row(row):
                    styles = [""] * len(row)
                    for col_name in [COL_LABELS["BS"], COL_LABELS["BY"]]:
                        if col_name in row.index:
                            v = str(row[col_name]).lower()
                            if "non satisfaisant" in v:
                                bg = "#FFCCCC"
                            elif "satisfaisant" in v:
                                bg = "#C6EFCE"
                            elif "inaccessible" in v or "non vérifiable" in v:
                                bg = "#FFE0B2"
                            elif "non visité" in v:
                                bg = "#E0E0E0"
                            else:
                                bg = ""
                            idx = list(row.index).index(col_name)
                            styles[idx] = f"background-color: {bg};"
                    return styles

                styled = df_display.style.apply(color_row, axis=1)
                st.dataframe(styled, use_container_width=True, height=min(400, 55 + 35 * len(df_client)))

                # Bouton téléchargement Excel individuel
                xlsx_bytes = build_client_excel(df_client, client)
                st.download_button(
                    label=f"⬇️ Télécharger Excel — {filename}",
                    data=xlsx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{client}",
                )