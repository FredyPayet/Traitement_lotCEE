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

# ─── Labels colonnes ────────────────────────────────────────────────────────
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

# ─── Couleurs Excel ──────────────────────────────────────────────────────────
COLORS = {
    "satisfaisant":    "C6EFCE",
    "non_satisfaisant":"FFCCCC",
    "inaccessible":    "FFE0B2",
    "non_visite":      "E0E0E0",
    "header":          "2F5496",
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
        ["libreoffice", "--headless", "--convert-to", "xlsx", "--outdir", out_dir, xls_path],
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
    with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        xlsx_path = convert_xls_to_xlsx(tmp_path)
        return pd.read_excel(xlsx_path, header=None)
    finally:
        os.unlink(tmp_path)


def extract_tables(df: pd.DataFrame):
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
    if max(needed_cols) >= df.shape[1]:
        raise ValueError(f"Le fichier n'a que {df.shape[1]} colonnes (colonne BY attendue).")

    data_raw = df.iloc[header_idx + 1:, needed_cols].copy()
    data_raw.columns = ["D", "E", "F", "G", "H", "I", "BS", "BY"]
    data_raw = data_raw.dropna(subset=["I"], how="all")
    data_raw = data_raw[data_raw["I"].astype(str).str.strip() != ""]

    if data_raw.empty:
        raise ValueError("Aucune ligne de données trouvée. Vérifiez que la colonne I est remplie.")

    for col in ["BS", "BY"]:
        data_raw[col] = data_raw[col].fillna("non visité")
        data_raw[col] = data_raw[col].apply(lambda v: "non visité" if str(v).strip() == "" else v)

    clients = sorted(data_raw["I"].dropna().unique().tolist())
    return data_raw.reset_index(drop=True), clients


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


def get_cell_color(value: str):
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
    bs_col_idx = col_keys.index("BS") + 1
    by_col_idx = col_keys.index("BY") + 1

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    header_fill = PatternFill("solid", fgColor=COLORS["header"])
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 40
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    for row_num, (_, row) in enumerate(df_client.iterrows(), 2):
        ws.row_dimensions[row_num].height = 20
        for col_num, key in enumerate(col_keys, 1):
            val = row[key]
            cell = ws.cell(row=row_num, column=col_num, value=val)
            cell.alignment = cell_align
            cell.border = border
            cell.font = Font(name="Arial", size=10)
            if col_num in (bs_col_idx, by_col_idx):
                color = get_cell_color(str(val))
                if color:
                    cell.fill = PatternFill("solid", fgColor=color)

    col_widths = [20, 30, 35, 12, 20, 35, 30, 25]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_filename(client_name: str, num_dossier: str) -> str:
    name = sanitize_filename(client_name)[:50]
    dossier = sanitize_filename(num_dossier.strip()) if num_dossier.strip() else "sans_dossier"
    return f"{name}_{dossier}.xlsx"


def get_ns_adresses(df_client: pd.DataFrame) -> list:
    """Retourne la liste des adresses (col F + G + H) dont BS = non satisfaisant."""
    ns_rows = df_client[df_client["BS"].str.lower().str.contains("non satisfaisant", na=False)]
    adresses = []
    for _, row in ns_rows.iterrows():
        adresse = " ".join(filter(None, [
            str(row["F"]).strip() if pd.notna(row["F"]) else "",
            str(row["G"]).strip() if pd.notna(row["G"]) else "",
            str(row["H"]).strip() if pd.notna(row["H"]) else "",
        ]))
        if adresse.strip():
            adresses.append(adresse)
    return adresses


def build_message(taux_choix: str, lot_label: str, ns_adresses_causes: list) -> str:
    """
    Construit le message final.
    ns_adresses_causes : liste de tuples (adresse, [cause1, cause2, ...])
    """
    # Bloc adresses NS (inséré après 'Vous trouverez ci-joint...')
    bloc_ns = ""
    if ns_adresses_causes:
        lignes = []
        for adresse, causes in ns_adresses_causes:
            lignes.append(adresse)
            for cause in causes:
                c = cause.strip()
                if c:
                    lignes.append(f"  • {c}")
        bloc_ns = "\n\n" + "\n".join(lignes)

    corps = {
        "Taux OK": (
            f"Bonjour,\n\n\n"
            f"Pour votre information, nous avons reçu le retour du lot de contrôle {lot_label}.\n\n"
            f"Tous les taux réglementaires sont respectés. Toutes les opérations du lot relatif "
            f"à la fiche travaux peuvent être finalisées.\n\n"
            f"Vous trouverez ci-joint les résultats de contrôles pour vos opérations."
            f"{bloc_ns}\n\n"
            f"Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
            f"Cordialement,"
        ),
        "Taux NS KO": (
            f"Bonjour,\n\n\n"
            f"Pour votre information, nous avons reçu le retour du lot de contrôle {lot_label}.\n\n"
            f"Le taux d'opérations contrôlées non satisfaisantes dépasse les 10 %. "
            f"Nous ne pouvons finaliser que les opérations qui ont été contrôlées. "
            f"Les opérations non visitées doivent être représentées dans un nouveau lot.\n\n"
            f"Vous trouverez ci-joint les résultats de contrôles pour vos opérations."
            f"{bloc_ns}\n\n"
            f"Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
            f"Cordialement,"
        ),
        "Tous taux KO": (
            f"Bonjour,\n\n\n"
            f"Pour votre information, nous avons reçu le retour du lot de contrôle {lot_label}.\n\n"
            f"Les taux réglementaires ne sont pas atteints. Nous ne pouvons finaliser aucune opération "
            f"dans ce lot. Les opérations doivent être représentées dans un nouveau lot.\n\n"
            f"Vous trouverez ci-joint les résultats de contrôles pour vos opérations."
            f"{bloc_ns}\n\n"
            f"Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
            f"Cordialement,"
        ),
    }
    return corps[taux_choix]


def copy_button(text: str, key: str):
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
            background-color:#4CAF50; color:white; border:none;
            padding:8px 18px; border-radius:6px; font-size:14px;
            cursor:pointer; font-family:Arial,sans-serif;
        ">📋 Copier le message</button>
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
        num_dossier = st.text_input("Numéro de dossier(s)", placeholder="ex : DOS-001")
    with col3:
        taux_choix = st.selectbox(
            "Résultat du lot",
            options=["Taux OK", "Taux NS KO", "Tous taux KO"],
        )

    st.markdown("---")

    lot_label = num_lot.strip() if num_lot.strip() else "[numéro de lot non renseigné]"

    couleur_bandeau = {
        "Taux OK":      "#e8f5e9",
        "Taux NS KO":   "#fff3e0",
        "Tous taux KO": "#ffebee",
    }[taux_choix]

    # ── Sidebar ──────────────────────────────────────────────────────────────
    st.sidebar.header("Filtres & Export")
    selected = st.sidebar.multiselect("Clients à afficher", options=clients, default=clients)

    # ZIP — tous les clients
    if st.sidebar.button("⬇️ Télécharger tous les fichiers (ZIP)"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for client in clients:
                df_c = data[data["I"] == client].copy()
                xlsx_bytes = build_client_excel(df_c, client)
                fname = build_filename(client, num_dossier)
                zf.writestr(fname, xlsx_bytes)
        st.sidebar.download_button(
            label="📦 Télécharger le ZIP",
            data=zip_buf.getvalue(),
            file_name="tableaux_clients.zip",
            mime="application/zip",
        )

    # ── Section par client ───────────────────────────────────────────────────
    if not selected:
        st.info("Aucun client sélectionné dans le panneau de gauche.")
    else:
        for client in selected:
            df_client = data[data["I"] == client].copy().reset_index(drop=True)
            filename = build_filename(client, num_dossier)

            with st.expander(f"🏢 {client}  ({len(df_client)} opération(s))", expanded=True):

                # ── Saisie des causes NS ─────────────────────────────────────
                ns_adresses = get_ns_adresses(df_client)
                ns_adresses_causes = []

                if ns_adresses:
                    st.markdown("**⚠️ Opérations non satisfaisantes — saisir la/les cause(s) :**")
                    for i, adresse in enumerate(ns_adresses):
                        cause_raw = st.text_input(
                            f"Cause(s) pour : {adresse}",
                            placeholder="ex : Épaisseur insuffisante. Pare-vapeur manquant.",
                            key=f"cause_{client}_{i}",
                            help="Séparez plusieurs causes par un point ou un point-virgule.",
                        )
                        # Découper les causes sur . ou ;
                        causes = [c.strip() for c in re.split(r'[.;]+', cause_raw) if c.strip()]
                        ns_adresses_causes.append((adresse, causes))
                    st.markdown("---")

                # ── Message automatique ──────────────────────────────────────
                message = build_message(taux_choix, lot_label, ns_adresses_causes)

                st.markdown("**✉️ Message à envoyer au client :**")
                st.markdown(
                    f"<div style='background-color:{couleur_bandeau}; padding:16px; "
                    f"border-radius:8px; white-space:pre-wrap; font-family:Arial; font-size:14px;'>"
                    f"{message}</div>",
                    unsafe_allow_html=True,
                )
                copy_button(message, key=f"copy_{client}")

                st.markdown("---")

                # ── Export Excel ─────────────────────────────────────────────
                xlsx_bytes = build_client_excel(df_client, client)
                st.download_button(
                    label=f"⬇️ Télécharger Excel — {filename}",
                    data=xlsx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{client}",
                )