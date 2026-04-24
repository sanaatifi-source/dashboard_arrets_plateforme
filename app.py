import io
from pathlib import Path
import math
from datetime import datetime
import base64

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle, Rectangle
from matplotlib.backends.backend_pdf import PdfPages


st.set_page_config(
    page_title="Dashboard Arrêts - Site Tizert",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

LOGO_PATH = Path("logoo_v2.png")

# =========================================================
# STYLE GLOBAL
# =========================================================
st.markdown("""
<style>
.stApp {
    background: #000000 !important;
    color: #e5eefc !important;
}

header[data-testid="stHeader"],
div[data-testid="stToolbar"],
div[data-testid="stDecoration"] {
    display: none !important;
}

section[data-testid="stSidebar"] {
    background: #050505 !important;
    border-right: 1px solid rgba(242, 178, 51, 0.20);
}

html, body, [class*="css"], [data-testid="stAppViewContainer"] {
    color: #e5eefc !important;
}

.block-container {
    padding-top: 0.7rem !important;
    padding-bottom: 1.5rem !important;
    max-width: 1550px;
}

h1, h2, h3, h4, h5, h6, p, label, span, div {
    color: #e5eefc !important;
}

.main-header {
    background: #000000;
    border: 1px solid rgba(242, 178, 51, 0.32);
    border-radius: 28px;
    padding: 34px 40px 38px 40px;
    margin-bottom: 28px;
    text-align: center;
    box-shadow: 0px 14px 34px rgba(0, 0, 0, 0.45);
}

.logo-center {
    display: flex;
    justify-content: center;
    align-items: center;
    margin-bottom: 18px;
}

.main-logo {
    width: 520px;
    max-width: 90%;
    height: auto;
    object-fit: contain;
    border-radius: 18px;
    background: transparent;
}

.header-title-center {
    font-size: 3.1rem;
    font-weight: 850;
    color: #f3f8ff !important;
    line-height: 1.08;
    margin-top: 4px;
}

.header-subtitle-center {
    font-size: 1.15rem;
    color: #a8a69c !important;
    margin-top: 12px;
    font-weight: 600;
}

.site-tizert {
    margin-top: 18px;
    font-size: 1.8rem;
    font-weight: 850;
    color: #f2b233 !important;
    letter-spacing: 0.08em;
}

.custom-card {
    background: #050505;
    border: 1px solid rgba(242, 178, 51, 0.18);
    border-radius: 18px;
    padding: 17px 19px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
    margin-bottom: 12px;
}

.card-title {
    font-size: 0.95rem;
    color: #c7b27a !important;
    font-weight: 700;
    margin-bottom: 6px;
}

.card-value {
    font-size: 1.75rem;
    font-weight: 850;
    color: #ffffff !important;
    line-height: 1.1;
}

.card-sub {
    font-size: 0.82rem;
    color: #9ca8bb !important;
    margin-top: 6px;
}

.executive-box {
    background: #050505;
    border: 1px solid rgba(242, 178, 51, 0.18);
    border-radius: 18px;
    padding: 18px 20px;
    margin-bottom: 16px;
    box-shadow: 0px 8px 22px rgba(0, 0, 0, 0.35);
}

.alert-box {
    background: #050505;
    border-left: 5px solid #f2b233;
    border-radius: 14px;
    padding: 14px 18px;
    margin-bottom: 10px;
}

.danger-box {
    background: #100505;
    border-left: 5px solid #f36c72;
    border-radius: 14px;
    padding: 14px 18px;
    margin-bottom: 10px;
}

.success-box {
    background: #061009;
    border-left: 5px solid #5ee08a;
    border-radius: 14px;
    padding: 14px 18px;
    margin-bottom: 10px;
}

.section-chip {
    display: inline-block;
    padding: 6px 12px;
    border-radius: 999px;
    background: rgba(242, 178, 51, 0.12);
    color: #f2b233 !important;
    border: 1px solid rgba(242, 178, 51, 0.18);
    font-size: 0.8rem;
    font-weight: 800;
    margin-bottom: 12px;
}

input, textarea {
    background-color: #111111 !important;
    color: #ffffff !important;
    border-radius: 10px !important;
}

div[data-baseweb="select"] > div {
    background-color: #111111 !important;
    color: #ffffff !important;
    border-radius: 10px !important;
    border: 1px solid rgba(242, 178, 51, 0.14) !important;
}

section[data-testid="stFileUploader"] {
    background: #111111 !important;
    border: 1px solid rgba(242, 178, 51, 0.14) !important;
    border-radius: 14px !important;
    padding: 10px !important;
}

div[data-testid="stDataFrame"] {
    background: #050505 !important;
    border-radius: 14px !important;
    border: 1px solid rgba(242, 178, 51, 0.12) !important;
    overflow: hidden !important;
}

button {
    border-radius: 10px !important;
}

div.stDownloadButton button,
div.stButton button {
    background: linear-gradient(90deg, #d99721 0%, #f2b233 100%) !important;
    color: #08101f !important;
    border: none !important;
    font-weight: 800 !important;
    box-shadow: 0 6px 18px rgba(242, 178, 51, 0.18);
}

div[role="radiogroup"] > label {
    background: #111111 !important;
    border-radius: 10px !important;
    padding: 8px 10px !important;
    margin-bottom: 6px !important;
    border: 1px solid rgba(242, 178, 51, 0.08);
}

hr {
    border: 1px solid rgba(242, 178, 51, 0.18) !important;
}

@media (max-width: 950px) {
    .main-logo {
        width: 330px;
    }
    .header-title-center {
        font-size: 2rem;
    }
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# OUTILS
# =========================================================
def get_logo_base64(path):
    if path.exists():
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def lower_text(value):
    return normalize_text(value).lower()


def duration_to_hours(value):
    if pd.isna(value):
        return 0.0

    if isinstance(value, pd.Timedelta):
        return value.total_seconds() / 3600

    if hasattr(value, "hour") and hasattr(value, "minute"):
        return value.hour + value.minute / 60 + getattr(value, "second", 0) / 3600

    if isinstance(value, (int, float)):
        val = float(value)
        return val * 24 if val <= 1.5 else val

    if isinstance(value, str):
        txt = value.strip().replace(",", ".")
        if txt == "":
            return 0.0
        try:
            if ":" in txt:
                parts = txt.split(":")
                h = float(parts[0])
                m = float(parts[1]) if len(parts) > 1 else 0
                s = float(parts[2]) if len(parts) > 2 else 0
                return h + m / 60 + s / 3600
            return float(txt)
        except Exception:
            return 0.0

    return 0.0


def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_hhmmss(hours):
    total_seconds = int(round(float(hours) * 3600)) if not pd.isna(hours) else 0
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def format_pct(value):
    return f"{value:.2f} %"


def format_h(value):
    return f"{value:.2f} h"


def render_info_card(title, value, subtitle=""):
    st.markdown(
        f"""
        <div class="custom-card">
            <div class="card-title">{title}</div>
            <div class="card-value">{value}</div>
            <div class="card-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def priority_from_duration(duree_h, total_h):
    if total_h <= 0:
        return "Faible"
    pct = duree_h / total_h * 100
    if pct >= 30:
        return "Critique"
    if pct >= 15:
        return "Élevée"
    if pct >= 7:
        return "Moyenne"
    return "Faible"


def action_recommendation(row):
    duree = float(row.get("Duree_h", 0))
    nb = int(row.get("Nb_arrets", 0)) if "Nb_arrets" in row else 0

    if nb >= 5 and duree >= 10:
        return "Analyser la récurrence, contrôler les organes critiques et planifier une action de fiabilisation."
    if duree >= 10:
        return "Analyser l'arrêt long, vérifier les causes racines et préparer une action corrective ciblée."
    if nb >= 5:
        return "Traiter la répétitivité des arrêts par inspection préventive et standardisation des interventions."
    return "Suivre l'équipement et vérifier si l'arrêt se répète sur les prochains mois."


# =========================================================
# CHARGEMENT DES DONNÉES
# =========================================================
def load_raw_data(uploaded_file, sheet_name=None):
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".csv":
        df = pd.read_csv(uploaded_file)
    else:
        excel = pd.ExcelFile(uploaded_file)
        chosen_sheet = sheet_name if sheet_name else excel.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=chosen_sheet, header=8)

    df = clean_columns(df)

    required_columns = [
        "Date",
        "Zone",
        "TAG",
        "Equipement",
        "Imputation arrêt",
        "Durée (H)",
        "Description Arrêt",
        "Type de panne",
        "Usine en arret/marche"
    ]

    optional_columns = ["Famille", "Causes majeures"]
    missing_required = [c for c in required_columns if c not in df.columns]

    for col in optional_columns:
        if col not in df.columns:
            df[col] = ""

    return df, missing_required


def prepare_data(df):
    df = clean_columns(df)
    df = df.dropna(how="all").copy()

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)

    text_cols = [
        "Zone", "Famille", "TAG", "Equipement",
        "Imputation arrêt", "Description Arrêt",
        "Type de panne", "Causes majeures", "Usine en arret/marche"
    ]

    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(normalize_text)

    df["Duree_h"] = df["Durée (H)"].apply(duration_to_hours) if "Durée (H)" in df.columns else 0.0

    return df


# =========================================================
# CLASSIFICATION
# =========================================================
def classify_planned(row):
    return "plan" in lower_text(row.get("Imputation arrêt", ""))


def classify_unplanned_stop(row):
    imp = lower_text(row.get("Imputation arrêt", ""))
    usine = lower_text(row.get("Usine en arret/marche", ""))
    return ("plan" not in imp) and usine in ["oui", "o", "yes", "y", "1"]


def classify_maintenance(row):
    imp = lower_text(row.get("Imputation arrêt", ""))
    usine = lower_text(row.get("Usine en arret/marche", ""))
    labels = ["mécanique", "mecanique", "electrique", "électrique", "instrumentation", "automatique"]
    return any(lbl in imp for lbl in labels) and usine in ["oui", "o", "yes", "y", "1"]


def maintenance_category(text):
    t = lower_text(text)
    if "mécanique" in t or "mecanique" in t:
        return "Mécanique"
    if "electrique" in t or "électrique" in t:
        return "Electrique"
    if "instrumentation" in t:
        return "Instrumentation"
    if "automatique" in t:
        return "Automatique"
    return "Autre"


# =========================================================
# CALCUL KPI
# =========================================================
def compute_kpis(df, annee, mois, temps_ouverture, cadence_theorique, tonnage_realise, taux_qualite):
    data = df.copy()

    data["is_planned"] = data.apply(classify_planned, axis=1)
    data["is_unplanned_stop"] = data.apply(classify_unplanned_stop, axis=1)
    data["is_maintenance"] = data.apply(classify_maintenance, axis=1)
    data["Maintenance_cat"] = data["Imputation arrêt"].apply(maintenance_category)

    arrets_planifies = round(data.loc[data["is_planned"], "Duree_h"].sum(), 2)
    temps_arrets_np = round(data.loc[data["is_unplanned_stop"], "Duree_h"].sum(), 2)

    temps_requis = round(max(temps_ouverture - arrets_planifies, 0), 2)
    temps_fonctionnement = round(max(temps_requis - temps_arrets_np, 0), 2)

    if cadence_theorique > 0:
        temps_theorique_production = tonnage_realise / cadence_theorique
        temps_ecart_cadence = round(max(temps_fonctionnement - temps_theorique_production, 0), 2)
    else:
        temps_theorique_production = 0.0
        temps_ecart_cadence = 0.0

    temps_net = round(max(temps_fonctionnement - temps_ecart_cadence, 0), 2)

    maintenance_df = data[data["is_maintenance"]].copy()
    nb_maintenance = int(len(maintenance_df))
    duree_maintenance = round(maintenance_df["Duree_h"].sum(), 2)

    disponibilite = round(((temps_requis - duree_maintenance) / temps_requis) * 100, 2) if temps_requis > 0 else 0.0
    mtbf = round((temps_requis - duree_maintenance) / nb_maintenance, 3) if nb_maintenance > 0 else 0.0
    mttr = round(duree_maintenance / nb_maintenance, 3) if nb_maintenance > 0 else 0.0
    trs = round((temps_net / temps_requis) * taux_qualite * 100, 2) if temps_requis > 0 else 0.0

    rep_maintenance = (
        maintenance_df.groupby("Maintenance_cat")["Duree_h"]
        .agg(["sum", "count"])
        .reset_index()
        .rename(columns={"sum": "Duree_h", "count": "Nb_arrets"})
    )

    categories = ["Mécanique", "Electrique", "Instrumentation", "Automatique"]
    rows = []
    total_maint = rep_maintenance["Duree_h"].sum() if not rep_maintenance.empty else 0.0

    for cat in categories:
        sub = rep_maintenance[rep_maintenance["Maintenance_cat"] == cat]
        duree = float(sub["Duree_h"].sum()) if not sub.empty else 0.0
        nb = int(sub["Nb_arrets"].sum()) if not sub.empty else 0
        pct = (duree / total_maint * 100) if total_maint > 0 else 0.0
        rows.append({
            "Categorie": cat,
            "Duree_h": round(duree, 2),
            "Duree_txt": to_hhmmss(duree),
            "Nb_arrets": nb,
            "Part maintenance (%)": round(pct, 2)
        })

    rep_maintenance_final = pd.DataFrame(rows)

    rep_zone = data.groupby("Zone", dropna=False)["Duree_h"].sum().sort_values(ascending=False).reset_index()
    top_equipements = data.groupby("Equipement", dropna=False)["Duree_h"].sum().sort_values(ascending=False).head(10).reset_index()

    equip_diag = (
        data.groupby(["TAG", "Equipement"], dropna=False)
        .agg(Duree_h=("Duree_h", "sum"), Nb_arrets=("Duree_h", "count"))
        .reset_index()
        .sort_values("Duree_h", ascending=False)
    )

    equip_diag["TAG_Equipement"] = equip_diag["TAG"].astype(str) + " | " + equip_diag["Equipement"].astype(str)

    total_arrets_all = equip_diag["Duree_h"].sum() if not equip_diag.empty else 0.0
    equip_diag["Part_globale_%"] = equip_diag["Duree_h"].apply(
        lambda x: round((x / total_arrets_all * 100), 2) if total_arrets_all > 0 else 0.0
    )
    equip_diag["Priorité"] = equip_diag["Duree_h"].apply(lambda x: priority_from_duration(x, total_arrets_all))
    equip_diag["Action recommandée"] = equip_diag.apply(action_recommendation, axis=1)

    pareto_tag = equip_diag[["TAG", "Equipement", "Duree_h", "TAG_Equipement"]].copy()
    if not pareto_tag.empty and pareto_tag["Duree_h"].sum() > 0:
        pareto_tag["Cum_%"] = pareto_tag["Duree_h"].cumsum() / pareto_tag["Duree_h"].sum() * 100
    else:
        pareto_tag["Cum_%"] = 0.0

    if "Date" in data.columns:
        temp = data.dropna(subset=["Date"]).copy()
        if not temp.empty:
            temp["Jour"] = temp["Date"].dt.date
            journalier = temp.groupby("Jour")["Duree_h"].sum().reset_index().sort_values("Jour")
        else:
            journalier = pd.DataFrame(columns=["Jour", "Duree_h"])
    else:
        journalier = pd.DataFrame(columns=["Jour", "Duree_h"])

    total_arrets = round(data["Duree_h"].sum(), 2)
    autres_arrets = round(max(total_arrets - arrets_planifies - temps_arrets_np, 0), 2)
    repartition_type = pd.DataFrame({
        "Type": ["Arrêts planifiés", "Arrêts NP usine arrêtée", "Autres arrêts"],
        "Durée_h": [arrets_planifies, temps_arrets_np, autres_arrets]
    })

    arrets_longs = data.sort_values("Duree_h", ascending=False).head(10).copy()

    bloc_kpi = pd.DataFrame({
        "Indicateur": [
            "Année", "Mois", "Temps d'ouverture", "Arrêts planifiés (hr)",
            "Temps Requis (hr)", "Temps arrêts NP", "Temps de fonctionnement",
            "Temps ecart cadence", "Temps Net", "Cadence théorique",
            "Temps théorique production", "Tonnage réalisé",
            "Somme des Arrêts maintenance", "Somme des durées maintenance",
            "Taux de qualité (%)"
        ],
        "Valeur": [
            annee, mois, round(temps_ouverture, 2), arrets_planifies,
            temps_requis, temps_arrets_np, temps_fonctionnement,
            temps_ecart_cadence, temps_net, round(cadence_theorique, 2),
            round(temps_theorique_production, 2), round(tonnage_realise, 2),
            nb_maintenance, duree_maintenance, round(taux_qualite * 100, 2)
        ]
    })

    return {
        "bloc_kpi": bloc_kpi,
        "trs": trs,
        "disponibilite": disponibilite,
        "mtbf": mtbf,
        "mttr": mttr,
        "temps_requis": temps_requis,
        "temps_fonctionnement": temps_fonctionnement,
        "temps_net": temps_net,
        "arrets_planifies": arrets_planifies,
        "temps_arrets_np": temps_arrets_np,
        "maintenance_total_h": duree_maintenance,
        "maintenance_total_txt": to_hhmmss(duree_maintenance),
        "maintenance_total_n": nb_maintenance,
        "rep_maintenance_final": rep_maintenance_final,
        "rep_zone": rep_zone,
        "top_equipements": top_equipements,
        "pareto_tag": pareto_tag,
        "journalier": journalier,
        "repartition_type": repartition_type,
        "equip_diag": equip_diag,
        "arrets_longs": arrets_longs,
        "data": data
    }


# =========================================================
# ANALYSE AUTOMATIQUE
# =========================================================
def generate_executive_summary(kpis):
    messages = []

    if kpis["trs"] >= 85:
        messages.append("Le TRS est à un très bon niveau, ce qui indique une performance globale satisfaisante.")
    elif kpis["trs"] >= 70:
        messages.append("Le TRS est acceptable, mais une marge d’amélioration reste possible sur les arrêts et la cadence.")
    else:
        messages.append("Le TRS est faible. Une analyse prioritaire des arrêts majeurs et des pertes de cadence est recommandée.")

    if kpis["disponibilite"] >= 90:
        messages.append("La disponibilité maintenance est très bonne.")
    elif kpis["disponibilite"] >= 75:
        messages.append("La disponibilité maintenance est moyenne. Les arrêts maintenance doivent être suivis de près.")
    else:
        messages.append("La disponibilité maintenance est critique. Les arrêts maintenance impactent fortement le temps requis.")

    if kpis["mttr"] > 4:
        messages.append("Le MTTR est élevé. Il faut analyser les causes de réparation longue et améliorer la réactivité maintenance.")
    elif kpis["mttr"] > 0:
        messages.append("Le MTTR reste maîtrisé, mais doit être suivi dans les prochains mois.")

    if kpis["mtbf"] > 0 and kpis["mtbf"] < 20:
        messages.append("Le MTBF est faible. La fréquence des pannes est élevée et nécessite une action de fiabilisation.")

    if kpis["maintenance_total_h"] > 0:
        messages.append("La priorité d’amélioration doit être orientée vers les équipements les plus pénalisants du Pareto.")

    return messages


def generate_alerts(kpis, obj_trs, obj_dispo, seuil_mttr, seuil_mtbf):
    alerts = []

    if obj_trs > 0:
        if kpis["trs"] < obj_trs:
            alerts.append(("danger", f"TRS inférieur à l’objectif saisi : {kpis['trs']} % < {obj_trs} %."))
        else:
            alerts.append(("success", f"TRS conforme à l’objectif saisi : {kpis['trs']} % ≥ {obj_trs} %."))

    if obj_dispo > 0:
        if kpis["disponibilite"] < obj_dispo:
            alerts.append(("danger", f"Disponibilité inférieure à l’objectif saisi : {kpis['disponibilite']} % < {obj_dispo} %."))
        else:
            alerts.append(("success", f"Disponibilité conforme à l’objectif saisi : {kpis['disponibilite']} % ≥ {obj_dispo} %."))

    if seuil_mttr > 0:
        if kpis["mttr"] > seuil_mttr:
            alerts.append(("danger", f"MTTR supérieur au seuil saisi : {kpis['mttr']} h > {seuil_mttr} h."))
        else:
            alerts.append(("success", f"MTTR conforme au seuil saisi : {kpis['mttr']} h ≤ {seuil_mttr} h."))

    if seuil_mtbf > 0:
        if kpis["mtbf"] > 0 and kpis["mtbf"] < seuil_mtbf:
            alerts.append(("danger", f"MTBF inférieur au seuil saisi : {kpis['mtbf']} h < {seuil_mtbf} h."))
        elif kpis["mtbf"] > 0:
            alerts.append(("success", f"MTBF conforme au seuil saisi : {kpis['mtbf']} h ≥ {seuil_mtbf} h."))

    if not alerts:
        alerts.append(("info", "Aucun seuil KPI n’a été saisi. Les alertes automatiques sont désactivées."))

    return alerts


# =========================================================
# VISUELS
# =========================================================
def draw_availability_gauge(value):
    fig, ax = plt.subplots(figsize=(6.2, 3.2), facecolor="#000000")
    ax.set_facecolor("#000000")
    ax.set_aspect("equal")
    ax.axis("off")

    segments = [
        (180, 144, "#f36c72"),
        (144, 108, "#f79a9a"),
        (108, 72, "#f0b47e"),
        (72, 36, "#d9b15f"),
        (36, 18, "#f2b233"),
        (18, 0, "#ffe08a"),
    ]

    for start, end, color in segments:
        ax.add_patch(Wedge((0, 0), 1.0, end, start, width=0.45, facecolor=color, edgecolor="none"))

    value = max(0, min(100, value))
    angle = 180 - (value / 100) * 180
    x = 0.75 * math.cos(math.radians(angle))
    y = 0.75 * math.sin(math.radians(angle))

    ax.plot([0, x], [0, y], linewidth=2.4, color="white")
    ax.add_patch(Circle((0, 0), 0.03, color="white"))
    ax.text(0, -0.12, f"{value:.2f}%", ha="center", va="center", fontsize=12, fontweight="bold", color="white")

    ax.set_xlim(-1.1, 1.1)
    ax.set_ylim(-0.2, 1.1)
    return fig


def draw_maintenance_columns(rep_df, total_h):
    fig, ax = plt.subplots(figsize=(8.2, 4.2), facecolor="#000000")
    ax.set_facecolor("#000000")
    ax.axis("off")

    categories = ["Mécanique", "Electrique", "Instrumentation", "Automatique"]
    colors = ["#f2b233", "#d99721", "#ffe08a", "#c7a64f"]
    x_positions = [0.12, 0.37, 0.62, 0.87]

    for x, cat, color in zip(x_positions, categories, colors):
        row = rep_df[rep_df["Categorie"] == cat]
        duree = float(row["Duree_h"].iloc[0]) if not row.empty else 0.0
        pct = float(row["Part maintenance (%)"].iloc[0]) if not row.empty else 0.0
        txt = row["Duree_txt"].iloc[0] if not row.empty else "00:00:00"

        height = pct / 100 if total_h > 0 else 0
        ax.add_patch(Rectangle((x - 0.06, 0.15), 0.12, 0.62, facecolor="#d8e2ea", edgecolor="#d8e2ea"))
        ax.add_patch(Circle((x, 0.77), 0.06, color="#bcc9d4"))
        ax.add_patch(Circle((x, 0.15), 0.06, color="#d8e2ea"))

        fill_h = 0.62 * height
        ax.add_patch(Rectangle((x - 0.06, 0.15), 0.12, fill_h, facecolor=color, edgecolor="none"))
        if fill_h > 0:
            ax.add_patch(Circle((x, 0.15 + fill_h), 0.06, color=color))

        ax.text(x, 0.61, txt, ha="center", va="center", fontsize=7.5, color="#08101f", fontweight="bold")
        ax.text(x, 0.37 if fill_h < 0.16 else 0.15 + fill_h / 2, f"{pct:.0f}%", ha="center", va="center", fontsize=10, fontweight="bold", color="#08101f")
        ax.text(x, 0.05, cat, ha="center", va="center", fontsize=8, color="white", fontweight="bold")

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    return fig


def make_dark_bar_plot(df, x_col, y_col, title, xlabel="", ylabel="", rotation=30, top_n=None):
    plot_df = df.copy()
    if top_n is not None:
        plot_df = plot_df.head(top_n)

    fig, ax = plt.subplots(figsize=(10, 4.5), facecolor="#000000")
    fig.patch.set_facecolor("#000000")
    ax.set_facecolor("#000000")

    if plot_df.empty:
        ax.text(0.5, 0.5, "Aucune donnée disponible", ha="center", va="center", color="white")
        ax.axis("off")
        return fig

    ax.bar(plot_df[x_col].astype(str), plot_df[y_col], color="#f2b233")
    ax.set_title(title, color="white", fontsize=13, pad=12, fontweight="bold")
    ax.set_xlabel(xlabel, color="#dbe7f5")
    ax.set_ylabel(ylabel, color="#dbe7f5")
    ax.tick_params(axis="x", colors="white", rotation=rotation)
    ax.tick_params(axis="y", colors="white")
    ax.grid(axis="y", linestyle="--", alpha=0.18, color="white")

    for spine in ax.spines.values():
        spine.set_color("#4c5f7a")

    plt.tight_layout()
    return fig


def make_dark_pie_plot(df, label_col, value_col, title):
    fig, ax = plt.subplots(figsize=(6.5, 4.5), facecolor="#000000")
    fig.patch.set_facecolor("#000000")
    ax.set_facecolor("#000000")

    plot_df = df[df[value_col] > 0].copy()

    if plot_df.empty:
        ax.text(0.5, 0.5, "Aucune donnée disponible", ha="center", va="center", color="white")
        ax.axis("off")
        return fig

    ax.pie(
        plot_df[value_col],
        labels=plot_df[label_col],
        autopct="%1.1f%%",
        startangle=90,
        textprops={"color": "white", "fontsize": 9}
    )
    ax.set_title(title, color="white", fontsize=13, fontweight="bold", pad=12)
    return fig


# =========================================================
# PDF
# =========================================================
def add_pdf_text_page(pdf, title, lines):
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    plt.axis("off")
    plt.text(0.05, 0.92, title, fontsize=22, fontweight="bold", color="#0b1326")

    y = 0.82
    for line in lines:
        plt.text(0.07, y, line, fontsize=12, color="#111827", wrap=True)
        y -= 0.055

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def add_pdf_table_page(pdf, title, df, max_rows=18):
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.axis("off")
    ax.text(0.03, 0.96, title, fontsize=18, fontweight="bold", transform=ax.transAxes)

    table_df = df.head(max_rows).copy()
    table = ax.table(cellText=table_df.values, colLabels=table_df.columns, loc="center", cellLoc="center")
    table.auto_set_font_size(False)
    table.set_fontsize(8.5)
    table.scale(1, 1.4)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(weight="bold", color="white")
            cell.set_facecolor("#132238")
        else:
            cell.set_facecolor("#f4f7fb")

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def make_pdf_bar_plot(df, x_col, y_col, title, top_n=10):
    fig, ax = plt.subplots(figsize=(11, 6))
    plot_df = df.copy().head(top_n)

    if plot_df.empty:
        ax.text(0.5, 0.5, "Aucune donnée disponible", ha="center", va="center")
        ax.axis("off")
        return fig

    ax.bar(plot_df[x_col].astype(str), plot_df[y_col])
    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_ylabel("Durée (h)")
    ax.tick_params(axis="x", rotation=35)
    ax.grid(axis="y", alpha=0.25)
    plt.tight_layout()
    return fig


def make_pdf_pareto_plot(pareto_df):
    fig, ax1 = plt.subplots(figsize=(11, 6))
    pareto = pareto_df.head(12).copy()

    if pareto.empty:
        ax1.text(0.5, 0.5, "Aucune donnée Pareto disponible", ha="center", va="center")
        ax1.axis("off")
        return fig

    ax1.bar(pareto["TAG_Equipement"], pareto["Duree_h"])
    ax1.set_ylabel("Durée (h)")
    ax1.tick_params(axis="x", rotation=40)
    ax1.set_title("Pareto par TAG | Equipement", fontsize=14, fontweight="bold")
    ax1.grid(axis="y", alpha=0.25)

    ax2 = ax1.twinx()
    ax2.plot(pareto["TAG_Equipement"], pareto["Cum_%"], marker="o")
    ax2.set_ylabel("Cumul (%)")
    ax2.set_ylim(0, 110)

    plt.tight_layout()
    return fig


def make_pdf_daily_plot(journalier):
    fig, ax = plt.subplots(figsize=(11, 6))

    if journalier.empty:
        ax.text(0.5, 0.5, "Aucune évolution journalière disponible", ha="center", va="center")
        ax.axis("off")
        return fig

    ax.plot(journalier["Jour"], journalier["Duree_h"], marker="o", linewidth=2)
    ax.set_title("Evolution journalière des arrêts", fontsize=14, fontweight="bold")
    ax.set_xlabel("Jour")
    ax.set_ylabel("Durée (h)")
    ax.tick_params(axis="x", rotation=35)
    ax.grid(alpha=0.25)
    plt.tight_layout()
    return fig


def generate_pdf_report(kpis, params, alerts):
    buffer = io.BytesIO()

    with PdfPages(buffer) as pdf:
        add_pdf_text_page(pdf, "Rapport KPI des arrêts - Managem", [
            "Rapport automatique généré par la plateforme Dashboard Arrêts - Managem.",
            "Site : Tizert",
            f"Période analysée : {params['mois']} {params['annee']}",
            f"Date de génération : {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            "",
            "KPI principaux :",
            f"TRS : {format_pct(kpis['trs'])}",
            f"Disponibilité maintenance : {format_pct(kpis['disponibilite'])}",
            f"MTBF : {format_h(kpis['mtbf'])}",
            f"MTTR : {format_h(kpis['mttr'])}",
        ])

        summary = generate_executive_summary(kpis)
        lines = [f"{i}. {msg}" for i, msg in enumerate(summary, 1)]
        lines.extend(["", "Alertes KPI :"])
        lines.extend([f"- {msg}" for _, msg in alerts])
        lines.extend([
            "",
            "Synthèse chiffrée :",
            f"Temps requis : {format_h(kpis['temps_requis'])}",
            f"Temps de fonctionnement : {format_h(kpis['temps_fonctionnement'])}",
            f"Temps net : {format_h(kpis['temps_net'])}",
            f"Arrêts planifiés : {format_h(kpis['arrets_planifies'])}",
            f"Temps arrêts non planifiés : {format_h(kpis['temps_arrets_np'])}",
            f"Maintenance totale : {format_h(kpis['maintenance_total_h'])}",
            f"Nombre d'arrêts maintenance : {kpis['maintenance_total_n']}",
        ])
        add_pdf_text_page(pdf, "Résumé exécutif et alertes", lines)

        add_pdf_table_page(pdf, "Bloc KPI détaillé", kpis["bloc_kpi"], max_rows=20)
        add_pdf_table_page(pdf, "Synthèse maintenance", kpis["rep_maintenance_final"], max_rows=10)
        add_pdf_table_page(
            pdf,
            "Diagnostic et plan d'action",
            kpis["equip_diag"][["TAG_Equipement", "Duree_h", "Nb_arrets", "Part_globale_%", "Priorité", "Action recommandée"]],
            max_rows=10
        )

        fig_zone = make_pdf_bar_plot(kpis["rep_zone"], "Zone", "Duree_h", "Répartition des arrêts par zone", top_n=12)
        pdf.savefig(fig_zone, bbox_inches="tight")
        plt.close(fig_zone)

        fig_eq = make_pdf_bar_plot(kpis["top_equipements"], "Equipement", "Duree_h", "Top équipements pénalisants", top_n=10)
        pdf.savefig(fig_eq, bbox_inches="tight")
        plt.close(fig_eq)

        fig_pareto = make_pdf_pareto_plot(kpis["pareto_tag"])
        pdf.savefig(fig_pareto, bbox_inches="tight")
        plt.close(fig_pareto)

        fig_daily = make_pdf_daily_plot(kpis["journalier"])
        pdf.savefig(fig_daily, bbox_inches="tight")
        plt.close(fig_daily)

        add_pdf_text_page(pdf, "Conclusion et recommandations", [
            "Ce rapport met en évidence les KPI principaux de performance et de maintenance.",
            "L'analyse Pareto permet d'identifier les équipements qui contribuent le plus aux pertes.",
            "",
            "Recommandations générales :",
            "1. Suivre mensuellement l'évolution du TRS, MTBF et MTTR.",
            "2. Prioriser les équipements du Pareto pour les actions de fiabilisation.",
            "3. Analyser les arrêts longs et récurrents.",
            "4. Mettre en place un plan d'action maintenance ciblé par zone et par équipement.",
            "5. Utiliser cette plateforme comme outil de suivi mensuel standardisé."
        ])

    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# HEADER
# =========================================================
logo_base64 = get_logo_base64(LOGO_PATH)

if logo_base64:
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" class="main-logo">'
else:
    logo_html = '<div style="color:#fca5a5;font-weight:800;">Logo introuvable : logoo_v2.png</div>'

st.markdown(f"""
<div class="main-header">
    <div class="logo-center">
        {logo_html}
    </div>
    <div class="header-title-center">
        Tableau de bord des arrêts
    </div>
    <div class="header-subtitle-center">
        Plateforme KPI maintenance - Groupe Managem
    </div>
    <div class="site-tizert">
        SITE TIZERT
    </div>
</div>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================
page = st.sidebar.radio(
    "Navigation",
    ["Dashboard principal", "Analyses complémentaires", "Diagnostic & Plan d’action", "Données", "Aide"]
)

uploaded_file = st.sidebar.file_uploader(
    "Importer le fichier des arrêts",
    type=["xlsx", "csv"]
)

sheet_name = None
if uploaded_file is not None and Path(uploaded_file.name).suffix.lower() != ".csv":
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("Choisir la feuille", xls.sheet_names)
        uploaded_file.seek(0)
    except Exception:
        uploaded_file.seek(0)

st.sidebar.markdown("---")
st.sidebar.subheader("Paramètres KPI")

annee = st.sidebar.number_input("Année", min_value=2000, max_value=2100, value=2026, step=1)

mois = st.sidebar.selectbox(
    "Mois",
    [
        "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
    ],
    index=1
)

temps_ouverture = st.sidebar.number_input("Temps d'ouverture (h)", min_value=0.0, value=672.0, step=1.0)
cadence_theorique = st.sidebar.number_input("Cadence théorique", min_value=0.0, value=450.0, step=1.0)
tonnage_realise = st.sidebar.number_input("Tonnage réalisé", min_value=0.0, value=214596.38, step=1.0)
taux_qualite_pct = st.sidebar.number_input("Taux de qualité (%)", min_value=0.0, max_value=200.0, value=102.0, step=0.1)
taux_qualite = taux_qualite_pct / 100

st.sidebar.markdown("---")
st.sidebar.subheader("Seuils KPI à définir par l'utilisateur")

obj_trs = st.sidebar.number_input("Saisir l’objectif TRS (%)", min_value=0.0, max_value=200.0, value=0.0, step=1.0)
obj_dispo = st.sidebar.number_input("Saisir l’objectif disponibilité (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
seuil_mttr = st.sidebar.number_input("Saisir le seuil MTTR max acceptable (h)", min_value=0.0, value=0.0, step=0.5)
seuil_mtbf = st.sidebar.number_input("Saisir le seuil MTBF min acceptable (h)", min_value=0.0, value=0.0, step=1.0)


if uploaded_file is None:
    st.info("Importez le fichier brut des arrêts depuis la barre latérale pour générer le dashboard.")
    st.markdown("""
    <div class="executive-box">
        <h3>Objectif de la plateforme</h3>
        <p>
        Cette application transforme automatiquement une base brute mensuelle des arrêts
        en tableau de bord KPI maintenance avec analyses, graphiques, diagnostic et rapport PDF professionnel.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


raw_df, missing_cols = load_raw_data(uploaded_file, sheet_name)

if missing_cols:
    st.error("Colonnes obligatoires manquantes dans le fichier importé.")
    st.write("Colonnes manquantes :")
    st.write(missing_cols)
    st.write("Colonnes détectées :")
    st.write(list(raw_df.columns))
    st.stop()

df = prepare_data(raw_df)

zones = sorted(df["Zone"].dropna().unique()) if "Zone" in df.columns else []
imputations = sorted(df["Imputation arrêt"].dropna().unique()) if "Imputation arrêt" in df.columns else []
tags = sorted(df["TAG"].dropna().unique()) if "TAG" in df.columns else []
types_panne = sorted(df["Type de panne"].dropna().unique()) if "Type de panne" in df.columns else []
equipements = sorted(df["Equipement"].dropna().unique()) if "Equipement" in df.columns else []

st.sidebar.markdown("---")
st.sidebar.subheader("Filtres avancés")

selected_zones = st.sidebar.multiselect("Filtrer Zone", zones, default=zones)
selected_imputations = st.sidebar.multiselect("Filtrer Imputation arrêt", imputations, default=imputations)
selected_tags = st.sidebar.multiselect("Filtrer TAG", tags, default=tags)
selected_types = st.sidebar.multiselect("Filtrer Type de panne", types_panne, default=types_panne)
selected_equipements = st.sidebar.multiselect("Filtrer Equipement", equipements, default=equipements)

df_filtered = df.copy()

if selected_zones:
    df_filtered = df_filtered[df_filtered["Zone"].isin(selected_zones)]
if selected_imputations:
    df_filtered = df_filtered[df_filtered["Imputation arrêt"].isin(selected_imputations)]
if selected_tags:
    df_filtered = df_filtered[df_filtered["TAG"].isin(selected_tags)]
if selected_types:
    df_filtered = df_filtered[df_filtered["Type de panne"].isin(selected_types)]
if selected_equipements:
    df_filtered = df_filtered[df_filtered["Equipement"].isin(selected_equipements)]

if "Date" in df_filtered.columns and df_filtered["Date"].notna().any():
    min_date = df_filtered["Date"].min().date()
    max_date = df_filtered["Date"].max().date()
    date_range = st.sidebar.date_input("Filtrer période", value=(min_date, max_date))
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        df_filtered = df_filtered[
            (df_filtered["Date"].dt.date >= start_date) &
            (df_filtered["Date"].dt.date <= end_date)
        ]

kpis = compute_kpis(
    df_filtered,
    annee=annee,
    mois=mois,
    temps_ouverture=temps_ouverture,
    cadence_theorique=cadence_theorique,
    tonnage_realise=tonnage_realise,
    taux_qualite=taux_qualite
)

alerts = generate_alerts(kpis, obj_trs, obj_dispo, seuil_mttr, seuil_mtbf)

params = {
    "annee": annee,
    "mois": mois,
    "temps_ouverture": temps_ouverture,
    "cadence_theorique": cadence_theorique,
    "tonnage_realise": tonnage_realise,
    "taux_qualite": taux_qualite
}


# =========================================================
# PAGES
# =========================================================
if page == "Dashboard principal":
    st.markdown('<div class="section-chip">Vue générale</div>', unsafe_allow_html=True)

    summary_lines = generate_executive_summary(kpis)

    st.markdown('<div class="executive-box">', unsafe_allow_html=True)
    st.subheader("Résumé exécutif automatique")
    for msg in summary_lines:
        st.write(f"- {msg}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.subheader("Alertes KPI")
    for level, msg in alerts:
        css_class = "danger-box" if level == "danger" else "success-box" if level == "success" else "alert-box"
        st.markdown(f'<div class="{css_class}">{msg}</div>', unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)

    with m1:
        render_info_card("TRS", f"{kpis['trs']} %", "Objectif saisi" if obj_trs > 0 else "Aucun objectif saisi")
    with m2:
        render_info_card("Disponibilité", f"{kpis['disponibilite']} %", "Objectif saisi" if obj_dispo > 0 else "Aucun objectif saisi")
    with m3:
        render_info_card("MTBF", f"{kpis['mtbf']} h", "Seuil saisi" if seuil_mtbf > 0 else "Aucun seuil saisi")
    with m4:
        render_info_card("MTTR", f"{kpis['mttr']} h", "Seuil saisi" if seuil_mttr > 0 else "Aucun seuil saisi")

    st.markdown("<br>", unsafe_allow_html=True)

    t1, t2 = st.columns([1.2, 1.1], gap="large")

    with t1:
        st.subheader("Bloc KPI")
        st.dataframe(kpis["bloc_kpi"], use_container_width=True, height=430)

    with t2:
        st.subheader("Synthèse maintenance")

        rep = kpis["rep_maintenance_final"]

        synthese = pd.DataFrame({
            "Indicateur": ["Maintenance totale", "Mécanique", "Electrique", "Instrumentation", "Automatique"],
            "Durée": [
                kpis["maintenance_total_txt"],
                rep.loc[rep["Categorie"] == "Mécanique", "Duree_txt"].iloc[0],
                rep.loc[rep["Categorie"] == "Electrique", "Duree_txt"].iloc[0],
                rep.loc[rep["Categorie"] == "Instrumentation", "Duree_txt"].iloc[0],
                rep.loc[rep["Categorie"] == "Automatique", "Duree_txt"].iloc[0],
            ],
            "Nombre arrêts": [
                kpis["maintenance_total_n"],
                int(rep.loc[rep["Categorie"] == "Mécanique", "Nb_arrets"].iloc[0]),
                int(rep.loc[rep["Categorie"] == "Electrique", "Nb_arrets"].iloc[0]),
                int(rep.loc[rep["Categorie"] == "Instrumentation", "Nb_arrets"].iloc[0]),
                int(rep.loc[rep["Categorie"] == "Automatique", "Nb_arrets"].iloc[0]),
            ],
            "Part maintenance (%)": [
                100.0 if kpis["maintenance_total_h"] > 0 else 0.0,
                float(rep.loc[rep["Categorie"] == "Mécanique", "Part maintenance (%)"].iloc[0]),
                float(rep.loc[rep["Categorie"] == "Electrique", "Part maintenance (%)"].iloc[0]),
                float(rep.loc[rep["Categorie"] == "Instrumentation", "Part maintenance (%)"].iloc[0]),
                float(rep.loc[rep["Categorie"] == "Automatique", "Part maintenance (%)"].iloc[0]),
            ]
        })

        st.dataframe(synthese, use_container_width=True, height=330)

        pdf_bytes = generate_pdf_report(kpis, params, alerts)

        st.download_button(
            "Télécharger le rapport PDF",
            data=pdf_bytes,
            file_name=f"rapport_kpi_arrets_{mois}_{annee}.pdf",
            mime="application/pdf"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    g1, g2 = st.columns([1.15, 1], gap="large")

    with g1:
        st.subheader("Répartition des arrêts maintenance")
        fig_col = draw_maintenance_columns(kpis["rep_maintenance_final"], kpis["maintenance_total_h"])
        st.pyplot(fig_col, use_container_width=True)

    with g2:
        st.subheader("Disponibilité maintenance")
        fig_g = draw_availability_gauge(kpis["disponibilite"])
        st.pyplot(fig_g, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Analyses rapides du mois")

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown("#### Top 5 équipements critiques")
        fig_top5 = make_dark_bar_plot(
            kpis["top_equipements"].head(5),
            x_col="Equipement",
            y_col="Duree_h",
            title="Top 5 équipements par durée d'arrêt",
            xlabel="Equipement",
            ylabel="Durée (h)",
            rotation=25
        )
        st.pyplot(fig_top5, use_container_width=True)

    with c2:
        st.markdown("#### Répartition des types d'arrêts")
        fig_type = make_dark_pie_plot(
            kpis["repartition_type"],
            label_col="Type",
            value_col="Durée_h",
            title="Répartition globale des arrêts"
        )
        st.pyplot(fig_type, use_container_width=True)


elif page == "Analyses complémentaires":
    st.markdown('<div class="section-chip">Analyses avancées</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.subheader("Répartition par zone")
        fig_zone = make_dark_bar_plot(
            kpis["rep_zone"],
            x_col="Zone",
            y_col="Duree_h",
            title="Durée d'arrêt par zone",
            xlabel="Zone",
            ylabel="Durée (h)",
            rotation=25,
            top_n=12
        )
        st.pyplot(fig_zone, use_container_width=True)

    with col2:
        st.subheader("Top équipements")
        fig_top = make_dark_bar_plot(
            kpis["top_equipements"],
            x_col="Equipement",
            y_col="Duree_h",
            title="Top 10 équipements",
            xlabel="Equipement",
            ylabel="Durée (h)",
            rotation=35,
            top_n=10
        )
        st.pyplot(fig_top, use_container_width=True)
        st.dataframe(kpis["top_equipements"], use_container_width=True, height=240)

    st.subheader("Pareto par TAG | Equipement")

    pareto = kpis["pareto_tag"].head(12).copy()
    fig2, ax1 = plt.subplots(figsize=(12, 5), facecolor="#000000")
    fig2.patch.set_facecolor("#000000")
    ax1.set_facecolor("#000000")

    if not pareto.empty:
        ax1.bar(pareto["TAG_Equipement"], pareto["Duree_h"], color="#f2b233")
        ax1.set_ylabel("Durée (h)", color="white")
        ax1.tick_params(axis="x", rotation=40, colors="white")
        ax1.tick_params(axis="y", colors="white")
        ax1.set_xlabel("TAG | Equipement", color="white")
        ax1.set_title("Pareto des arrêts par TAG | Equipement", color="white", fontsize=13, fontweight="bold", pad=12)
        ax1.grid(axis="y", linestyle="--", alpha=0.18, color="white")

        ax2 = ax1.twinx()
        ax2.plot(pareto["TAG_Equipement"], pareto["Cum_%"], marker="o", linewidth=2, color="#ffe08a")
        ax2.set_ylabel("Cumul (%)", color="white")
        ax2.tick_params(axis="y", colors="white")
        ax2.set_ylim(0, 110)

    st.pyplot(fig2, use_container_width=True)
    st.dataframe(pareto[["TAG", "Equipement", "Duree_h", "Cum_%"]], use_container_width=True, height=250)

    st.subheader("Évolution journalière")
    if not kpis["journalier"].empty:
        fig, ax = plt.subplots(figsize=(10, 4.2), facecolor="#000000")
        fig.patch.set_facecolor("#000000")
        ax.set_facecolor("#000000")
        ax.plot(kpis["journalier"]["Jour"], kpis["journalier"]["Duree_h"], marker="o", linewidth=2, color="#f2b233")
        ax.fill_between(kpis["journalier"]["Jour"], kpis["journalier"]["Duree_h"], alpha=0.16, color="#f2b233")
        ax.set_xlabel("Jour", color="white")
        ax.set_ylabel("Durée d'arrêt (h)", color="white")
        ax.set_title("Évolution journalière des arrêts", color="white", fontsize=13, fontweight="bold", pad=12)
        ax.tick_params(axis="x", rotation=45, colors="white")
        ax.tick_params(axis="y", colors="white")
        ax.grid(axis="y", linestyle="--", alpha=0.18, color="white")
        st.pyplot(fig, use_container_width=True)
    else:
        st.info("Pas de dates exploitables.")


elif page == "Diagnostic & Plan d’action":
    st.markdown('<div class="section-chip">Diagnostic maintenance</div>', unsafe_allow_html=True)

    st.subheader("Diagnostic automatique des équipements critiques")
    diag = kpis["equip_diag"][["TAG_Equipement", "Duree_h", "Nb_arrets", "Part_globale_%", "Priorité", "Action recommandée"]].head(15)
    st.dataframe(diag, use_container_width=True, height=420)

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.subheader("Arrêts les plus longs")
        cols = ["Date", "Zone", "TAG", "Equipement", "Imputation arrêt", "Duree_h", "Description Arrêt"]
        existing_cols = [c for c in cols if c in kpis["arrets_longs"].columns]
        st.dataframe(kpis["arrets_longs"][existing_cols], use_container_width=True, height=360)

    with c2:
        st.subheader("Priorités d’action")
        priority_table = diag.groupby("Priorité").agg(
            Nombre=("TAG_Equipement", "count"),
            Duree_totale_h=("Duree_h", "sum")
        ).reset_index()

        priority_table = priority_table.rename(columns={
            "Duree_totale_h": "Durée totale des arrêts (h)"
        })

        st.dataframe(priority_table, use_container_width=True, height=200)

        fig_priority = make_dark_bar_plot(
            priority_table,
            x_col="Priorité",
            y_col="Durée totale des arrêts (h)",
            title="Durée totale par niveau de priorité",
            xlabel="Priorité",
            ylabel="Durée (h)",
            rotation=0
        )
        st.pyplot(fig_priority, use_container_width=True)


elif page == "Données":
    st.markdown('<div class="section-chip">Exploration des données</div>', unsafe_allow_html=True)

    st.subheader("Données filtrées")
    st.dataframe(df_filtered, use_container_width=True, height=520)

    csv_bytes = df_filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Télécharger les données filtrées CSV",
        data=csv_bytes,
        file_name="donnees_filtrees.csv",
        mime="text/csv"
    )

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        render_info_card("Lignes filtrées", f"{len(df_filtered)}", "Nombre de lignes après filtres")

    with col_b:
        render_info_card("Zones", f"{df_filtered['Zone'].nunique() if 'Zone' in df_filtered.columns else 0}", "Nombre de zones visibles")

    with col_c:
        render_info_card("Equipements", f"{df_filtered['Equipement'].nunique() if 'Equipement' in df_filtered.columns else 0}", "Nombre d'équipements visibles")


elif page == "Aide":
    st.markdown('<div class="section-chip">Guide utilisateur</div>', unsafe_allow_html=True)

    st.subheader("Principe")
    st.write(
        "L'utilisateur importe seulement la base brute des arrêts. "
        "La plateforme calcule automatiquement les KPI, affiche les graphiques, génère un diagnostic et crée un rapport PDF."
    )

    st.subheader("Colonnes obligatoires")
    st.write([
        "Date", "Zone", "TAG", "Equipement", "Imputation arrêt",
        "Durée (H)", "Description Arrêt", "Type de panne",
        "Usine en arret/marche"
    ])

    st.subheader("Colonnes optionnelles")
    st.write(["Famille", "Causes majeures"])

    st.subheader("Seuils KPI")
    st.write(
        "Les seuils KPI ne sont pas fixés automatiquement. "
        "Ils doivent être saisis par l'utilisateur dans la barre latérale. "
        "Si aucun seuil n'est saisi, les alertes automatiques sont désactivées."
    )

    st.subheader("Formules KPI utilisées")
    st.write("Temps Requis = Temps d'ouverture - Arrêts planifiés")
    st.write("Temps de fonctionnement = Temps Requis - Temps arrêts NP")
    st.write("Temps ecart cadence = Temps de fonctionnement - (Tonnage réalisé / Cadence théorique)")
    st.write("Temps Net = Temps de fonctionnement - Temps ecart cadence")
    st.write("TRS = (Temps Net / Temps Requis) × Taux de qualité × 100")
    st.write("Disponibilité = ((Temps Requis - Somme des durées maintenance) / Temps Requis) × 100")
    st.write("MTBF = (Temps Requis - Somme des durées maintenance) / Nombre d'arrêts maintenance")
    st.write("MTTR = Somme des durées maintenance / Nombre d'arrêts maintenance")
