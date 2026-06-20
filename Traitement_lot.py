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

COLORS = {
    "satisfaisant":     "C6EFCE",
    "non_satisfaisant": "FFCCCC",
    "inaccessible":     "FFE0B2",
    "non_visite":       "E0E0E0",
    "header":           "2F5496",
}

# ─── Référentiel NS (intégré directement) ───────────────────────────────────

NS_REF = {
    "EN-101": {
        "Surface": (
            "Il a été constaté un écart de plus de 10 % entre la surface déclarée et celle relevée sur site.",
            "Merci de nous confirmer la prise en compte de la surface mesurée, ou nous transmettre un document justificatif type plans pour contredire les mesures."
        ),
        "Ecart au feu": (
            "Il a été constaté une absence d'écart au feu entre l'isolant et une source de chaleur.",
            "Merci de faire réaliser un écart au feu de 10 cm, et de nous transmettre une attestation de reprise et des photos."
        ),
        "Epaisseur": (
            "Il a été constaté une épaisseur insuffisante pour atteindre la résistance thermique imposée par la fiche standardisée.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Homogénéité": (
            "Il a été constaté que la pose de l'isolant n'est pas homogène. La performance thermique est donc discontinue.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Risque électrique": (
            "Il a été constaté que la VMC est en contact avec l'isolant.",
            "Merci de faire réaliser un écart de 10 cm entre l'isolant et la VMC, et de nous transmettre une attestation de reprise et des photos."
        ),
        "Rehausse": (
            "Il a été constaté l'absence de rehausse de trappe.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Non réalisés": (
            "Il a été constaté que les travaux n'ont pas été réalisés.",
            "Merci de nous confirmer l'annulation de la valorisation pour cette opération."
        ),
        "Pare vapeur": (
            "Il a été constaté l'absence de pare vapeur, alors que l'élément est nécessaire.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Humidité": (
            "Il a été constaté des traces d'humidité sur les murs.",
            "Merci de faire installer un pare vapeur, et de nous transmettre une attestation de reprise et des photos."
        ),
        "Non éligible": (
            "Il a été constaté que la configuration de l'installation n'est pas éligible au CEE.",
            "Merci de nous confirmer l'annulation de la valorisation pour cette opération."
        ),
        "Trappe bloquée": (
            "Il a été constaté que la trappe est bloquée par l'isolant.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Déclaratif": (
            "Il a été constaté que les éléments décrits sur la facture ne correspondent pas à ce qui a été installé sur site.",
            "Merci de faire reprendre la facture et de nous la transmettre."
        ),
    },
    "EN-102": {
        "Surface": (
            "Il a été constaté un écart de plus de 10 % entre la surface déclarée et celle relevée sur site.",
            "Merci de nous confirmer la prise en compte de la surface mesurée, ou nous transmettre un document justificatif type plans pour contredire les mesures."
        ),
        "Epaisseur": (
            "Il a été constaté une épaisseur insuffisante pour atteindre la résistance thermique imposée par la fiche standardisée.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Protection intempérie": (
            "Il a été constaté que l'isolant n'est pas protégé des intempéries sur toute la surface.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Finition (ITI)": (
            "Il a été constaté que l'isolant n'est pas protégé des intempéries sur toute la surface.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Non réalisés": (
            "Il a été constaté que les travaux n'ont pas été réalisés.",
            "Merci de nous confirmer l'annulation de la valorisation pour cette opération."
        ),
        "Distance sol": (
            "Il a été constaté qu'il n'y a pas d'espace entre le départ de l'isolant et le sol, ce qui peut entraîner des remontées capillaires.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Qualité non manifeste": (
            "Il a été observé des non-qualités manifestes des travaux.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Déclaratif": (
            "Il a été constaté que les éléments décrits sur la facture ne correspondent pas à ce qui a été installé sur site.",
            "Merci de faire reprendre la facture et de nous la transmettre."
        ),
    },
    "EN-103": {
        "Surface": (
            "Il a été constaté un écart de plus de 10 % entre la surface déclarée et celle relevée sur site.",
            "Merci de nous confirmer la prise en compte de la surface mesurée, ou nous transmettre un document justificatif type plans pour contredire les mesures."
        ),
        "Ecart au feu": (
            "Il a été constaté une absence d'écart au feu entre l'isolant et une source de chaleur.",
            "Merci de faire réaliser un écart au feu de 10 cm, et de nous transmettre une attestation de reprise et des photos."
        ),
        "Epaisseur": (
            "Il a été constaté une épaisseur insuffisante pour atteindre la résistance thermique imposée par la fiche standardisée.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Homogénéité": (
            "Il a été constaté que la pose de l'isolant n'est pas homogène. La performance thermique est donc discontinue.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Risque électrique": (
            "Il a été constaté que des luminaires sont en contact avec l'isolant.",
            "Merci de faire réaliser un écart de 10 cm entre l'isolant et les luminaires, et de nous transmettre une attestation de reprise et des photos."
        ),
        "Non réalisés": (
            "Il a été constaté que les travaux n'ont pas été réalisés.",
            "Merci de nous confirmer l'annulation de la valorisation pour cette opération."
        ),
        "Humidité": (
            "Il a été constaté des traces d'humidité sur les murs.",
            "Merci de faire installer un pare vapeur, et de nous transmettre une attestation de reprise et des photos."
        ),
        "Fixation": (
            "Il a été constaté un défaut de fixation sur les panneaux d'isolant.",
            "Merci de faire reprendre les travaux et de nous transmettre une attestation de reprise et des photos."
        ),
        "Morcellement (points particuliers non isolés)": (
            "Il a été constaté qu'une poutre n'est pas isolée et entraîne une discontinuité d'isolant.",
            "Merci de faire isoler la poutre pour enlever le pont thermique, et de nous transmettre une attestation de reprise et des photos. L'isolant peut être de fine couche."
        ),
        "Non éligible": (
            "Il a été constaté que la configuration de l'installation n'est pas éligible au CEE.",
            "Merci de nous confirmer l'annulation de la valorisation pour cette opération."
        ),
        "Déclaratif": (
            "Il a été constaté que les éléments décrits sur la facture ne correspondent pas à ce qui a été installé sur site.",
            "Merci de faire reprendre la facture et de nous la transmettre."
        ),
    },
    "EN-105": {
        "Surface": ("", ""),
        "Epaisseur": ("", ""),
        "Homogénéité": ("", ""),
        "Etanchéité": ("", ""),
        "Ecart au feu": ("", ""),
        "Non réalisés": ("", ""),
        "Document": ("", ""),
    },
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

    needed_cols = [col_letter_to_idx(l) for l in ["D", "E", "F", "G", "H", "I", "BS", "BY", "N", "O"]]
    if max(needed_cols) >= df.shape[1]:
        raise ValueError(f"Le fichier n'a que {df.shape[1]} colonnes (colonne BY attendue).")

    data_raw = df.iloc[header_idx + 1:, needed_cols].copy()
    data_raw.columns = ["D", "E", "F", "G", "H", "I", "BS", "BY", "N", "O"]
    data_raw = data_raw.dropna(subset=["I"], how="all")
    data_raw = data_raw[data_raw["I"].astype(str).str.strip() != ""]

    if data_raw.empty:
        raise ValueError("Aucune ligne de données trouvée. Vérifiez que la colonne I est remplie.")

    for col in ["BS", "BY"]:
        data_raw[col] = data_raw[col].fillna("non visité")
        data_raw[col] = data_raw[col].apply(lambda v: "non visité" if str(v).strip() == "" else v)

    for col in ["N", "O"]:
        data_raw[col] = pd.to_numeric(data_raw[col], errors="coerce").fillna(0)

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


def build_client_excel(df_client: pd.DataFrame, client_name: str, lot_destination: str = "") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = client_name[:31]

    col_keys = ["D", "E", "F", "G", "H", "I", "BS", "BY"]
    headers = [COL_LABELS[k] for k in col_keys]

    # Ajout colonne "Nouveau lot de contrôle" si lot_destination renseigné
    if lot_destination:
        col_keys_ext = col_keys + ["LOT_DEST"]
        headers_ext = headers + ["Nouveau lot de contrôle"]
    else:
        col_keys_ext = col_keys
        headers_ext = headers

    bs_col_idx = col_keys_ext.index("BS") + 1
    by_col_idx = col_keys_ext.index("BY") + 1

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    header_fill = PatternFill("solid", fgColor=COLORS["header"])
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 40
    for col_num, header in enumerate(headers_ext, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    for row_num, (_, row) in enumerate(df_client.iterrows(), 2):
        ws.row_dimensions[row_num].height = 20
        for col_num, key in enumerate(col_keys_ext, 1):
            val = lot_destination if key == "LOT_DEST" else row[key]
            cell = ws.cell(row=row_num, column=col_num, value=val)
            cell.alignment = cell_align
            cell.border = border
            cell.font = Font(name="Arial", size=10)
            if col_num in (bs_col_idx, by_col_idx):
                color = get_cell_color(str(val))
                if color:
                    cell.fill = PatternFill("solid", fgColor=color)

    col_widths = [20, 30, 35, 12, 20, 35, 30, 25] + ([25] if lot_destination else [])
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_filename(client_name: str, num_dossier: str, num_lot: str) -> str:
    name = sanitize_filename(client_name)[:50]
    dossier = sanitize_filename(num_dossier.strip()) if num_dossier.strip() else "sans_dossier"
    lot = sanitize_filename(num_lot.strip()) if num_lot.strip() else "lot_non_renseigné"
    return f"{name}_{dossier}_{lot}.xlsx"


def get_ns_adresses(df_client: pd.DataFrame) -> list:
    """Retourne liste de dicts {adresse, fiche} pour les lignes non satisfaisantes."""
    ns_rows = df_client[df_client["BS"].str.lower().str.contains("non satisfaisant", na=False)]
    result = []
    for _, row in ns_rows.iterrows():
        adresse = " ".join(filter(None, [
            str(row["F"]).strip() if pd.notna(row["F"]) else "",
            str(row["G"]).strip() if pd.notna(row["G"]) else "",
            str(row["H"]).strip() if pd.notna(row["H"]) else "",
        ]))
        # Extraire la fiche CEE depuis la colonne D (ex: "BAR-EN-101-xxx" → "EN-101")
        ref_d = str(row["D"]).strip() if pd.notna(row["D"]) else ""
        fiche = detect_fiche(ref_d)
        if adresse.strip():
            result.append({"adresse": adresse, "fiche": fiche})
    return result


def detect_fiche(ref: str) -> str | None:
    """Détecte la fiche CEE dans une référence (ex: BAR-EN-101-xxx → EN-101)."""
    m = re.search(r"(EN|BAR|BAT|RES|IND|AGR|TRA)-?\d{3}", ref, re.IGNORECASE)
    if m:
        return m.group(0).upper()
    # Cherche directement EN-101 style
    m2 = re.search(r"[A-Z]{2,3}-\d{3}", ref, re.IGNORECASE)
    return m2.group(0).upper() if m2 else None


def build_message(taux_choix: str, lot_label: str, ns_adresses_causes: list, lot_destination: str = "", delais_courts: bool = False) -> str:
    bloc_ns = ""
    if ns_adresses_causes:
        lignes = []
        for adresse, nc_items in ns_adresses_causes:
            lignes.append(f"• {adresse} :")
            for nc_label, msg_complet in nc_items:
                if msg_complet.strip():
                    lignes.append(f"\t• {msg_complet.strip()}")
        bloc_ns = "\n\n" + "\n".join(lignes)

    bloc_destination = ""
    if lot_destination.strip():
        bloc_destination = f"\n\nLes opérations concernées devront être représentées dans le lot de destination suivant : {lot_destination.strip()}."

    bloc_delais = ""
    if delais_courts:
        bloc_delais = "\n\nLa date de fin de travaux est inférieure à 3 mois. Nous ne pouvons plus intégrer le dossier dans un nouveau lot de contrôle pour validation. Merci de nous fournir un document du marché permettant de repousser cette date, ou de nous confirmer l'annulation du dossier."

    fin = f"{bloc_destination}{bloc_delais}\n\nCordialement,"

    corps = {
        "Taux OK": (
            f"Bonjour,\n\n\n"
            f"Pour votre information, nous avons reçu le retour du lot de contrôle {lot_label}.\n\n"
            f"Tous les taux réglementaires sont respectés. Toutes les opérations du lot relatif "
            f"à la fiche travaux peuvent être finalisées.\n\n"
            f"Vous trouverez ci-joint les résultats des contrôles pour vos opérations."
            f"{bloc_ns}\n\n"
            f"Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
            f"{fin}"
        ),
        "Taux NS KO": (
            f"Bonjour,\n\n\n"
            f"Pour votre information, nous avons reçu le retour du lot de contrôle {lot_label}.\n\n"
            f"Le taux d'opérations contrôlées non satisfaisantes dépasse les 10 %. "
            f"Nous ne pouvons finaliser que les opérations qui ont été contrôlées. "
            f"Les opérations non visitées doivent être représentées dans un nouveau lot.\n\n"
            f"Vous trouverez ci-joint les résultats des contrôles pour vos opérations."
            f"{bloc_ns}\n\n"
            f"Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
            f"{fin}"
        ),
        "Tous taux KO": (
            f"Bonjour,\n\n\n"
            f"Pour votre information, nous avons reçu le retour du lot de contrôle {lot_label}.\n\n"
            f"Les taux réglementaires ne sont pas atteints. Nous ne pouvons finaliser aucune opération "
            f"dans ce lot. Les opérations doivent être représentées dans un nouveau lot.\n\n"
            f"Vous trouverez ci-joint les résultats des contrôles pour vos opérations."
            f"{bloc_ns}\n\n"
            f"Les rapports de contrôle sont disponibles sur ODICEE, si vos opérations ont été contrôlées.\n\n"
            f"{fin}"
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

    col1, col2 = st.columns(2)
    with col1:
        num_lot = st.text_input("Numéro de lot", placeholder="ex : LOT-2024-001")
    with col2:
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

    # ZIP — tous les clients (utilise dossier vide par défaut pour le ZIP global)
    if st.sidebar.button("⬇️ Télécharger tous les fichiers (ZIP)"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for client in clients:
                df_c = data[data["I"] == client].copy()
                # Récupérer dossier et lot_destination depuis session_state si disponibles
                dossier_key = f"dossier_{client}"
                dest_active_key = f"dest_active_{client}"
                dest_key = f"lot_dest_{client}"
                num_dossier_c = st.session_state.get(dossier_key, "")
                lot_dest_c = st.session_state.get(dest_key, "") if st.session_state.get(dest_active_key, False) else ""
                xlsx_bytes = build_client_excel(df_c, client, lot_dest_c)
                fname = build_filename(client, num_dossier_c, num_lot)
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

            # Calcul du volume total (colonnes N + O) * 0.001
            volume = (pd.to_numeric(df_client["N"], errors="coerce").fillna(0) +
                      pd.to_numeric(df_client["O"], errors="coerce").fillna(0)).sum() * 0.001
            volume_str = f"{volume:,.3f} MWhc".replace(",", " ")
            volume_bold = volume >= 2000
            label_volume = f"**{volume_str}**" if volume_bold else volume_str

            with st.expander(f"🏢 {client}  ({len(df_client)} opération(s)) — {label_volume}", expanded=True):

                # ── Numéro de dossier par client ─────────────────────────────
                dossier_key = f"dossier_{client}"
                num_dossier = st.text_input(
                    "Numéro de dossier(s)",
                    placeholder="ex : DOS-001",
                    key=dossier_key,
                )

                # ── Lot de destination ───────────────────────────────────────
                dest_active_key = f"dest_active_{client}"
                dest_key = f"lot_dest_{client}"

                activer_destination = st.checkbox(
                    "Spécifier un lot de destination",
                    key=dest_active_key,
                )
                lot_destination = ""
                if activer_destination:
                    lot_destination = st.text_input(
                        "Lot de destination",
                        placeholder="ex : LOT-2024-002",
                        key=dest_key,
                    )

                delais_key = f"delais_{client}"
                delais_courts = st.checkbox(
                    "⏳ Délai restant pour contrôle < 3 mois",
                    key=delais_key,
                )

                filename = build_filename(client, num_dossier, num_lot)

                st.markdown("---")

                # ── Saisie des non-conformités NS via référentiel ────────────
                ns_adresses = get_ns_adresses(df_client)
                ns_adresses_causes = []

                if ns_adresses:
                    st.markdown("**⚠️ Opérations non satisfaisantes — saisir les non-conformités :**")

                    for i, item in enumerate(ns_adresses):
                        adresse = item["adresse"]
                        fiche   = item["fiche"]

                        st.markdown(f"📍 **{adresse}**")

                        # Sélection de la fiche CEE
                        fiches_dispo = sorted(NS_REF.keys()) if NS_REF else []
                        fiche_default_idx = fiches_dispo.index(fiche) if fiche in fiches_dispo else 0

                        fiche_sel = st.selectbox(
                            "Fiche CEE",
                            options=fiches_dispo if fiches_dispo else ["Référentiel non trouvé"],
                            index=fiche_default_idx,
                            key=f"fiche_sel_{client}_{i}",
                        )

                        # Liste des NC disponibles pour cette fiche
                        nc_list = []
                        if fiche_sel in NS_REF:
                            nc_list = list(NS_REF[fiche_sel].keys())

                        # Clé session_state pour le nombre de NC sélectionnées
                        count_key = f"nc_count_{client}_{i}"
                        if count_key not in st.session_state:
                            st.session_state[count_key] = 1

                        nc_items = []
                        for j in range(st.session_state[count_key]):
                            col_sel, col_btn = st.columns([10, 1])
                            with col_sel:
                                nc_sel = st.selectbox(
                                    f"Non-conformité {j+1}",
                                    options=["— Sélectionner —"] + nc_list,
                                    key=f"nc_sel_{client}_{i}_{j}",
                                    label_visibility="collapsed",
                                )
                            if j == st.session_state[count_key] - 1:
                                with col_btn:
                                    if st.button("➕", key=f"add_{client}_{i}_{j}", help="Ajouter une non-conformité"):
                                        st.session_state[count_key] += 1
                                        st.rerun()

                            if nc_sel and nc_sel != "— Sélectionner —" and nc_sel in NS_REF.get(fiche_sel, {}):
                                msg, act = NS_REF[fiche_sel][nc_sel]
                                msg_complet = f"{msg} {act}".strip()
                                nc_items.append((nc_sel, msg_complet))

                        ns_adresses_causes.append((adresse, nc_items))
                        st.markdown("")

                    st.markdown("---")

                # ── Message automatique ──────────────────────────────────────
                message = build_message(taux_choix, lot_label, ns_adresses_causes, lot_destination, delais_courts)

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
                xlsx_bytes = build_client_excel(df_client, client, lot_destination)
                st.download_button(
                    label=f"⬇️ Télécharger Excel — {filename}",
                    data=xlsx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{client}",
                )