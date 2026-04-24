import io
from pathlib import Path
import math
from datetime import datetime

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle, Rectangle
from matplotlib.backends.backend_pdf import PdfPages


st.set_page_config(
    page_title="Tableau de bord des arrêts",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# STYLE PROFESSIONNEL SOMBRE
# =========================================================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(180deg, #08101f 0%, #0b1326 100%) !important;
    color: #e5eefc !important;
}

header[data-testid="stHeader"] {
    display: none !important;
}

div[data-testid="stToolbar"] {
    display: none !important;
}

div[data-testid="stDecoration"] {
    display: none !important;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1728 0%, #101b2f 100%) !important;
    border-right: 1px solid rgba(29, 212, 223, 0.12);
}

html, body, [class*="css"], [data-testid="stAppViewContainer"] {
    color: #e5eefc !important;
}

h1, h2, h3, h4, h5, h6 {
    color: #f3f8ff !important;
    letter-spacing: 0.2px;
}

h3 {
    min-height: 54px;
    line-height: 1.2 !important;
    margin-bottom: 0.6rem !important;
}

p, label, span, div {
    color: #dbe7f5 !important;
}

.block-container {
    padding-top: 0.6rem !important;
    padding-bottom: 1.5rem !important;
    max-width: 1500px;
}

.logo-box {
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 0px 0 6px 0;
    margin-bottom: 2px;
}

.custom-card {
    background: linear-gradient(180deg, #112036 0%, #0e1a2d 100%);
    border: 1px solid rgba(29, 212, 223, 0.14);
    border-radius: 18px;
    padding: 16px 18px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.22);
    margin-bottom: 12px;
}

.card-title {
    font-size: 0.95rem;
    color: #8fb7d8 !important;
    font-weight: 600;
    margin-bottom: 6px;
}

.card-value {
    font-size: 1.6rem;
    font-weight: 800;
    color: #ffffff !important;
    line-height: 1.1;
}

.card-sub {
    font-size: 0.8rem;
    color: #89a7c4 !important;
    margin-top: 6px;
}

.hero-box {
    background: linear-gradient(135deg, rgba(29,212,223,0.08) 0%, rgba(74,168,255,0.05) 100%);
    border: 1px solid rgba(29, 212, 223, 0.16);
    border-radius: 20px;
    padding: 18px 24px 12px 24px;
    margin-bottom: 18px;
}

div[data-testid="stMetric"] {
    background: linear-gradient(180deg, #112036 0%, #0e1a2d 100%) !important;
    border: 1px solid rgba(29, 212, 223, 0.14) !important;
    padding: 12px !important;
    border-radius: 16px !important;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.18) !important;
}

input, textarea {
    background-color: #132238 !important;
    color: #ffffff !important;
    border-radius: 10px !important;
}

div[data-baseweb="select"] > div {
    background-color: #132238 !important;
    color: #ffffff !important;
    border-radius: 10px !important;
    border: 1px solid rgba(29, 212, 223, 0.12) !important;
}

section[data-testid="stFileUploader"] {
    background: #132238 !important;
    border: 1px solid rgba(29, 212, 223, 0.14) !important;
    border-radius: 14px !important;
    padding: 10px !important;
}

div[data-testid="stDataFrame"] {
    background: #0f1c2f !important;
    border-radius: 14px !important;
    border: 1px solid rgba(29, 212, 223, 0.12) !important;
    overflow: hidden !important;
}

table {
    background-color: #132238 !important;
    color: #ffffff !important;
}

thead tr th {
    background-color: #1a2d47 !important;
    color: #ffffff !important;
}

tbody tr td {
    background-color: #132238 !important;
    color: #ffffff !important;
}

button {
    border-radius: 10px !important;
}

div.stDownloadButton button,
div.stButton button {
    background: linear-gradient(90deg, #1597c9 0%, #1dd4df 100%) !important;
    color: #08101f !important;
    border: none !important;
    font-weight: 700 !important;
    box-shadow: 0 6px 18px rgba(29, 212, 223, 0.18);
}

div[role="radiogroup"] > label {
    background: #132238 !important;
    border-radius: 10px !important;
    padding: 8px 10px !important;
    margin-bottom: 6px !important;
    border: 1px solid rgba(29, 212, 223, 0.08);
}

hr {
    border: 1px solid rgba(120, 150, 190, 0.18) !important;
}

.small-note {
    color: #8ea8c7 !important;
    font-size: 0.86rem;
}

.section-chip {
    display: inline-block;
    padding: 5px 10px;
    border-radius: 999px;
    background: rgba(29, 212, 223, 0.10);
    color: #8ee8ef !important;
    border: 1px solid rgba(29, 212, 223, 0.14);
    font-size: 0.78rem;
    font-weight: 700;
    margin-bottom: 10px;
}

.executive-box {
    background: linear-gradient(180deg, #10213a 0%, #0d192d 100%);
    border: 1px solid rgba(29, 212, 223, 0.16);
    border-radius: 16px;
    padding: 16px 18px;
    margin-bottom: 14px;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# LOGO
# =========================================================
LOGO_PATH = Path("logoo.v2.png")


# =========================================================
# OUTILS
# =========================================================
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
                if len(parts) >= 2:
                    h = float(parts[0])
                    m = float(parts[1])
                    s = float(parts[2]) if len(parts) > 2 else 0.0
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
    if pd.isna(hours):
        return "00:00:00"
    total_seconds = int(round(float(hours) * 3600))
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


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


def format_pct(value):
    return f"{value:.2f} %"


def format_h(value):
    return f"{value:.2f} h"


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

    optional_columns = [
        "Famille",
        "Causes majeures"
    ]

    missing_required = [c for c in required_columns if c not in df.columns]

    for col in optional_columns:
        if col not in df.columns:
            df[col] = ""

    return df, missing_required


def prepare_data(df):
    df = df.copy()
    df = clean_columns(df)
    df = df.dropna(how="all")

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

    if "Durée (H)" in df.columns:
        df["Duree_h"] = df["Durée (H)"].apply(duration_to_hours)
    else:
        df["Duree_h"] = 0.0

    return df


# =========================================================
# CLASSIFICATION
# =========================================================
def classify_planned(row):
    imp = lower_text(row.get("Imputation arrêt", ""))
    return "plan" in imp


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

    somme_arrets_maintenance = int(len(maintenance_df))
    somme_durees_maintenance = round(maintenance_df["Duree_h"].sum(), 2)

    disponibilite = round(
        ((temps_requis - somme_durees_maintenance) / temps_requis) * 100, 2
    ) if temps_requis > 0 else 0.0

    mtbf = round(
        (temps_requis - somme_durees_maintenance) / somme_arrets_maintenance, 3
    ) if somme_arrets_maintenance > 0 else 0.0

    mttr = round(
        somme_durees_maintenance / somme_arrets_maintenance, 3
    ) if somme_arrets_maintenance > 0 else 0.0

    trs = round(
        (temps_net / temps_requis) * taux_qualite * 100, 2
    ) if temps_requis > 0 else 0.0

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
            "Pct": round(pct, 2)
        })

    rep_maintenance_final = pd.DataFrame(rows)

    rep_zone = (
        data.groupby("Zone", dropna=False)["Duree_h"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    top_equipements = (
        data.groupby("Equipement", dropna=False)["Duree_h"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )

    pareto_tag = (
        data.groupby(["TAG", "Equipement"], dropna=False)["Duree_h"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    if not pareto_tag.empty and pareto_tag["Duree_h"].sum() > 0:
        pareto_tag["TAG_Equipement"] = pareto_tag["TAG"].astype(str) + " | " + pareto_tag["Equipement"].astype(str)
        pareto_tag["Cum_%"] = pareto_tag["Duree_h"].cumsum() / pareto_tag["Duree_h"].sum() * 100
    else:
        pareto_tag["TAG_Equipement"] = ""
        pareto_tag["Cum_%"] = 0.0

    if "Date" in data.columns:
        temp = data.copy()
        temp = temp.dropna(subset=["Date"])
        if not temp.empty:
            temp["Jour"] = temp["Date"].dt.date
            journalier = (
                temp.groupby("Jour")["Duree_h"]
                .sum()
                .reset_index()
                .sort_values("Jour")
            )
        else:
            journalier = pd.DataFrame(columns=["Jour", "Duree_h"])
    else:
        journalier = pd.DataFrame(columns=["Jour", "Duree_h"])

    bloc_kpi = pd.DataFrame({
        "Indicateur": [
            "Année",
            "Mois",
            "Temps d'ouverture",
            "Arrêts planifiés (hr)",
            "Temps Requis (hr)",
            "Temps arrêts NP",
            "Temps de fonctionnement",
            "Temps ecart cadence",
            "Temps Net",
            "Cadence théorique",
            "Temps théorique production",
            "Tonnage réalisé",
            "Somme des Arrêts maintenance",
            "Somme des durées maintenance",
            "Taux de qualité (%)"
        ],
        "Valeur": [
            annee,
            mois,
            round(temps_ouverture, 2),
            round(arrets_planifies, 2),
            round(temps_requis, 2),
            round(temps_arrets_np, 2),
            round(temps_fonctionnement, 2),
            round(temps_ecart_cadence, 2),
            round(temps_net, 2),
            round(cadence_theorique, 2),
            round(temps_theorique_production, 2),
            round(tonnage_realise, 2),
            somme_arrets_maintenance,
            round(somme_durees_maintenance, 2),
            round(taux_qualite * 100, 2)
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
        "maintenance_total_h": somme_durees_maintenance,
        "maintenance_total_txt": to_hhmmss(somme_durees_maintenance),
        "maintenance_total_n": somme_arrets_maintenance,
        "rep_maintenance_final": rep_maintenance_final,
        "rep_zone": rep_zone,
        "top_equipements": top_equipements,
        "pareto_tag": pareto_tag,
        "journalier": journalier,
        "data": data
    }


# =========================================================
# ANALYSE AUTOMATIQUE
# =========================================================
def generate_executive_summary(kpis):
    trs = kpis["trs"]
    dispo = kpis["disponibilite"]
    mttr = kpis["mttr"]
    mtbf = kpis["mtbf"]
    maint = kpis["maintenance_total_h"]

    messages = []

    if trs >= 85:
        messages.append("Le TRS est à un très bon niveau, ce qui indique une performance globale satisfaisante.")
    elif trs >= 70:
        messages.append("Le TRS est acceptable, mais une marge d’amélioration reste possible sur les arrêts et la cadence.")
    else:
        messages.append("Le TRS est faible. Une analyse prioritaire des arrêts majeurs et des pertes de cadence est recommandée.")

    if dispo >= 90:
        messages.append("La disponibilité maintenance est très bonne.")
    elif dispo >= 75:
        messages.append("La disponibilité maintenance est moyenne. Les arrêts maintenance doivent être suivis de près.")
    else:
        messages.append("La disponibilité maintenance est critique. Les arrêts maintenance impactent fortement le temps requis.")

    if mttr > 0 and mttr > 4:
        messages.append("Le MTTR est élevé. Il faut analyser les causes de réparation longue et améliorer la réactivité maintenance.")
    elif mttr > 0:
        messages.append("Le MTTR reste maîtrisé, mais doit être suivi dans les prochains mois.")

    if mtbf > 0 and mtbf < 20:
        messages.append("Le MTBF est faible. La fréquence des pannes est élevée et nécessite une action de fiabilisation.")

    if maint > 0:
        messages.append("La priorité d’amélioration doit être orientée vers les équipements les plus pénalisants du Pareto.")

    return messages


# =========================================================
# VISUELS STREAMLIT
# =========================================================
def draw_availability_gauge(value):
    fig, ax = plt.subplots(figsize=(6.2, 3.2), facecolor="#08101f")
    ax.set_facecolor("#08101f")
    ax.set_aspect("equal")
    ax.axis("off")

    segments = [
        (180, 144, "#f36c72"),
        (144, 108, "#f79a9a"),
        (108, 72, "#f0b47e"),
        (72, 36, "#9ad7df"),
        (36, 18, "#37c8d6"),
        (18, 0, "#1dd4df"),
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
    ax.text(-1.0, -0.02, "0", color="#9bb6d6", fontsize=9)
    ax.text(0.0, 1.02, "50", color="#9bb6d6", fontsize=9, ha="center")
    ax.text(1.0, -0.02, "100", color="#9bb6d6", fontsize=9, ha="center")
    ax.set_xlim(-1.1, 1.1)
    ax.set_ylim(-0.2, 1.1)
    return fig


def draw_maintenance_columns(rep_df, total_h):
    fig, ax = plt.subplots(figsize=(8.2, 4.2), facecolor="#08101f")
    ax.set_facecolor("#08101f")
    ax.axis("off")

    categories = ["Mécanique", "Electrique", "Instrumentation", "Automatique"]
    colors = ["#83d5e8", "#2b8bb8", "#1dd4df", "#59f0ff"]
    x_positions = [0.12, 0.37, 0.62, 0.87]

    for x, cat, color in zip(x_positions, categories, colors):
        row = rep_df[rep_df["Categorie"] == cat]
        duree = float(row["Duree_h"].iloc[0]) if not row.empty else 0.0
        pct = float(row["Pct"].iloc[0]) if not row.empty else 0.0
        txt = row["Duree_txt"].iloc[0] if not row.empty else "00:00:00"

        height = pct / 100 if total_h > 0 else 0

        ax.add_patch(Rectangle((x - 0.06, 0.15), 0.12, 0.62, facecolor="#d8e2ea", edgecolor="#d8e2ea"))
        ax.add_patch(Circle((x, 0.77), 0.06, color="#bcc9d4"))
        ax.add_patch(Circle((x, 0.15), 0.06, color="#d8e2ea"))

        fill_h = 0.62 * height
        ax.add_patch(Rectangle((x - 0.06, 0.15), 0.12, fill_h, facecolor=color, edgecolor="none", alpha=0.98))
        if fill_h > 0:
            ax.add_patch(Circle((x, 0.15 + fill_h), 0.06, color=color, alpha=0.98))

        ax.text(
            x, 0.61, txt,
            ha="center", va="center",
            fontsize=7.5,
            color="#08101f" if fill_h > 0.46 else "white",
            fontweight="bold"
        )

        ax.text(
            x,
            0.37 if fill_h < 0.16 else 0.15 + fill_h / 2,
            f"{pct:.0f}%",
            ha="center",
            va="center",
            fontsize=10,
            fontweight="bold",
            color="#08101f" if fill_h > 0.24 else "white"
        )

        ax.text(
            x, 0.05, cat,
            ha="center", va="center",
            fontsize=8,
            color="white",
            fontweight="bold"
        )

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    return fig


def make_dark_bar_plot(df, x_col, y_col, title, xlabel=None, ylabel=None, rotation=30, top_n=None):
    plot_df = df.copy()
    if top_n is not None:
        plot_df = plot_df.head(top_n)

    fig, ax = plt.subplots(figsize=(10, 4.5), facecolor="#08101f")
    fig.patch.set_facecolor("#08101f")
    ax.set_facecolor("#08101f")

    if plot_df.empty:
        ax.text(0.5, 0.5, "Aucune donnée disponible", ha="center", va="center", color="white")
        ax.axis("off")
        return fig

    ax.bar(plot_df[x_col].astype(str), plot_df[y_col], color="#1dd4df")
    ax.set_title(title, color="white", fontsize=13, pad=12, fontweight="bold")
    ax.set_xlabel(xlabel if xlabel else "", color="#dbe7f5")
    ax.set_ylabel(ylabel if ylabel else "", color="#dbe7f5")
    ax.tick_params(axis="x", colors="white", rotation=rotation)
    ax.tick_params(axis="y", colors="white")
    ax.grid(axis="y", linestyle="--", alpha=0.18, color="white")

    for spine in ax.spines.values():
        spine.set_color("#294566")

    plt.tight_layout()
    return fig


# =========================================================
# GRAPHIQUES POUR PDF
# =========================================================
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


# =========================================================
# PDF PROFESSIONNEL
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
    table = ax.table(
        cellText=table_df.values,
        colLabels=table_df.columns,
        loc="center",
        cellLoc="center"
    )
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


def generate_pdf_report(kpis, params):
    buffer = io.BytesIO()

    with PdfPages(buffer) as pdf:
        # Page 1 : couverture
        cover_lines = [
            "Rapport automatique généré par la plateforme Dashboard Arrêts - Managem",
            f"Période analysée : {params['mois']} {params['annee']}",
            f"Date de génération : {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            "",
            "Objectif du rapport :",
            "Présenter une synthèse claire des KPI maintenance, des arrêts, des équipements pénalisants",
            "et des axes prioritaires d'amélioration.",
            "",
            "KPI principaux :",
            f"TRS : {format_pct(kpis['trs'])}",
            f"Disponibilité maintenance : {format_pct(kpis['disponibilite'])}",
            f"MTBF : {format_h(kpis['mtbf'])}",
            f"MTTR : {format_h(kpis['mttr'])}",
        ]
        add_pdf_text_page(pdf, "Rapport KPI des arrêts - Managem", cover_lines)

        # Page 2 : résumé exécutif
        summary = generate_executive_summary(kpis)
        lines = []
        for i, msg in enumerate(summary, 1):
            lines.append(f"{i}. {msg}")

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
        add_pdf_text_page(pdf, "Résumé exécutif", lines)

        # Page 3 : bloc KPI
        add_pdf_table_page(pdf, "Bloc KPI détaillé", kpis["bloc_kpi"], max_rows=20)

        # Page 4 : maintenance
        add_pdf_table_page(pdf, "Synthèse maintenance", kpis["rep_maintenance_final"], max_rows=10)

        # Page 5 : zones
        fig_zone = make_pdf_bar_plot(kpis["rep_zone"], "Zone", "Duree_h", "Répartition des arrêts par zone", top_n=12)
        pdf.savefig(fig_zone, bbox_inches="tight")
        plt.close(fig_zone)

        # Page 6 : top équipements
        fig_eq = make_pdf_bar_plot(kpis["top_equipements"], "Equipement", "Duree_h", "Top équipements pénalisants", top_n=10)
        pdf.savefig(fig_eq, bbox_inches="tight")
        plt.close(fig_eq)

        # Page 7 : pareto
        fig_pareto = make_pdf_pareto_plot(kpis["pareto_tag"])
        pdf.savefig(fig_pareto, bbox_inches="tight")
        plt.close(fig_pareto)

        # Page 8 : évolution journalière
        fig_daily = make_pdf_daily_plot(kpis["journalier"])
        pdf.savefig(fig_daily, bbox_inches="tight")
        plt.close(fig_daily)

        # Page 9 : conclusion
        conclusion_lines = [
            "Conclusion automatique :",
            "",
            "Ce rapport met en évidence les KPI principaux de performance et de maintenance.",
            "L'analyse Pareto permet d'identifier les équipements qui contribuent le plus aux pertes.",
            "La priorité d'amélioration doit être donnée aux équipements ayant les durées d'arrêt les plus élevées.",
            "",
            "Recommandations générales :",
            "1. Suivre mensuellement l'évolution du TRS, MTBF et MTTR.",
            "2. Prioriser les équipements du Pareto pour les actions de fiabilisation.",
            "3. Analyser les arrêts longs et récurrents.",
            "4. Mettre en place un plan d'action maintenance ciblé par zone et par équipement.",
            "5. Utiliser cette plateforme comme outil de suivi mensuel standardisé."
        ]
        add_pdf_text_page(pdf, "Conclusion et recommandations", conclusion_lines)

    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# HEADER
# =========================================================
if LOGO_PATH.exists():
    st.markdown('<div class="logo-box">', unsafe_allow_html=True)
    col_logo1, col_logo2, col_logo3 = st.columns([0.1, 4.8, 0.1])
    with col_logo2:
        st.image(str(LOGO_PATH), width=700)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="hero-box">
    <h1 style='text-align: center; margin-bottom: 0.2rem;'>Tableau de bord des arrêts</h1>
    <p style='text-align: center; color: #9eb8d6; margin-top: 0; margin-bottom: 0.2rem;'>
        Plateforme KPI maintenance - Managem
    </p>
</div>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================
page = st.sidebar.radio(
    "Navigation",
    ["Dashboard principal", "Analyses complémentaires", "Données", "Aide"]
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


# =========================================================
# SI PAS DE FICHIER
# =========================================================
if uploaded_file is None:
    st.info("Importez le fichier brut des arrêts depuis la barre latérale pour générer le dashboard.")
    st.markdown("""
    <div class="executive-box">
        <h3>Objectif de la plateforme</h3>
        <p>
        Cette application permet de transformer automatiquement une base brute mensuelle des arrêts
        en tableau de bord KPI maintenance avec analyses, graphiques et rapport PDF.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# =========================================================
# LECTURE ET PRÉPARATION
# =========================================================
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

st.sidebar.markdown("---")
st.sidebar.subheader("Filtres")

selected_zones = st.sidebar.multiselect("Filtrer Zone", zones, default=zones)
selected_imputations = st.sidebar.multiselect("Filtrer Imputation arrêt", imputations, default=imputations)
selected_tags = st.sidebar.multiselect("Filtrer TAG", tags, default=tags)

df_filtered = df.copy()

if selected_zones:
    df_filtered = df_filtered[df_filtered["Zone"].isin(selected_zones)]

if selected_imputations:
    df_filtered = df_filtered[df_filtered["Imputation arrêt"].isin(selected_imputations)]

if selected_tags:
    df_filtered = df_filtered[df_filtered["TAG"].isin(selected_tags)]


kpis = compute_kpis(
    df_filtered,
    annee=annee,
    mois=mois,
    temps_ouverture=temps_ouverture,
    cadence_theorique=cadence_theorique,
    tonnage_realise=tonnage_realise,
    taux_qualite=taux_qualite
)

params = {
    "annee": annee,
    "mois": mois,
    "temps_ouverture": temps_ouverture,
    "cadence_theorique": cadence_theorique,
    "tonnage_realise": tonnage_realise,
    "taux_qualite": taux_qualite
}


# =========================================================
# PAGE 1 : DASHBOARD PRINCIPAL
# =========================================================
if page == "Dashboard principal":
    st.markdown('<div class="section-chip">Vue générale</div>', unsafe_allow_html=True)

    summary_lines = generate_executive_summary(kpis)

    st.markdown('<div class="executive-box">', unsafe_allow_html=True)
    st.subheader("Résumé exécutif automatique")
    for msg in summary_lines:
        st.write(f"- {msg}")
    st.markdown('</div>', unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)

    with m1:
        render_info_card("TRS", f"{kpis['trs']} %", "Taux de rendement synthétique")

    with m2:
        render_info_card("Disponibilité", f"{kpis['disponibilite']} %", "Disponibilité maintenance")

    with m3:
        render_info_card("MTBF", f"{kpis['mtbf']} h", "Temps moyen entre pannes")

    with m4:
        render_info_card("MTTR", f"{kpis['mttr']} h", "Temps moyen de réparation")

    st.markdown("<br>", unsafe_allow_html=True)

    t1, t2 = st.columns([1.2, 1.1], gap="large")

    with t1:
        st.subheader("Bloc KPI")
        st.dataframe(
            kpis["bloc_kpi"],
            use_container_width=True,
            height=430
        )

    with t2:
        st.subheader("Synthèse maintenance")

        rep = kpis["rep_maintenance_final"]

        synthese = pd.DataFrame({
            "Indicateur": [
                "Maintenance totale",
                "Mécanique",
                "Electrique",
                "Instrumentation",
                "Automatique"
            ],
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
            "Part (%)": [
                100.0 if kpis["maintenance_total_h"] > 0 else 0.0,
                float(rep.loc[rep["Categorie"] == "Mécanique", "Pct"].iloc[0]),
                float(rep.loc[rep["Categorie"] == "Electrique", "Pct"].iloc[0]),
                float(rep.loc[rep["Categorie"] == "Instrumentation", "Pct"].iloc[0]),
                float(rep.loc[rep["Categorie"] == "Automatique", "Pct"].iloc[0]),
            ]
        })

        st.dataframe(
            synthese,
            use_container_width=True,
            height=330
        )

        pdf_bytes = generate_pdf_report(kpis, params)

        st.download_button(
            "Télécharger le rapport PDF",
            data=pdf_bytes,
            file_name=f"rapport_kpi_arrets_{mois}_{annee}.pdf",
            mime="application/pdf"
        )

        st.markdown(
            "<div class='small-note'>Le rapport PDF contient les KPI, analyses, graphiques, Pareto et recommandations.</div>",
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    g1, g2 = st.columns([1.15, 1], gap="large")

    with g1:
        st.subheader("Répartition des arrêts maintenance")
        fig_col = draw_maintenance_columns(
            kpis["rep_maintenance_final"],
            kpis["maintenance_total_h"]
        )
        st.pyplot(fig_col, use_container_width=True)

    with g2:
        st.subheader("Disponibilité maintenance")
        fig_g = draw_availability_gauge(kpis["disponibilite"])
        st.pyplot(fig_g, use_container_width=True)


# =========================================================
# PAGE 2 : ANALYSES COMPLÉMENTAIRES
# =========================================================
elif page == "Analyses complémentaires":
    st.markdown('<div class="section-chip">Analyses avancées</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.subheader("Répartition par zone")
        if not kpis["rep_zone"].empty:
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
        else:
            st.info("Aucune donnée disponible.")

    with col2:
        st.subheader("Top équipements")
        if not kpis["top_equipements"].empty:
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
        else:
            st.info("Aucune donnée disponible.")

    st.subheader("Pareto par TAG | Equipement")
    if not kpis["pareto_tag"].empty:
        pareto = kpis["pareto_tag"].head(12).copy()
        fig2, ax1 = plt.subplots(figsize=(12, 5), facecolor="#08101f")
        fig2.patch.set_facecolor("#08101f")
        ax1.set_facecolor("#08101f")

        ax1.bar(pareto["TAG_Equipement"], pareto["Duree_h"], color="#1dd4df")
        ax1.set_ylabel("Durée (h)", color="white")
        ax1.tick_params(axis="x", rotation=40, colors="white")
        ax1.tick_params(axis="y", colors="white")
        ax1.set_xlabel("TAG | Equipement", color="white")
        ax1.set_title("Pareto des arrêts par TAG | Equipement", color="white", fontsize=13, fontweight="bold", pad=12)
        ax1.grid(axis="y", linestyle="--", alpha=0.18, color="white")

        for spine in ax1.spines.values():
            spine.set_color("#294566")

        if "Cum_%" in pareto.columns:
            ax2 = ax1.twinx()
            ax2.plot(pareto["TAG_Equipement"], pareto["Cum_%"], marker="o", linewidth=2, color="#4aa8ff")
            ax2.set_ylabel("Cumul (%)", color="white")
            ax2.tick_params(axis="y", colors="white")
            ax2.set_ylim(0, 110)

        st.pyplot(fig2, use_container_width=True)
        st.dataframe(pareto[["TAG", "Equipement", "Duree_h", "Cum_%"]], use_container_width=True, height=250)
    else:
        st.info("Pas de données Pareto exploitables.")

    st.subheader("Évolution journalière")
    if not kpis["journalier"].empty:
        fig, ax = plt.subplots(figsize=(10, 4.2), facecolor="#08101f")
        fig.patch.set_facecolor("#08101f")
        ax.set_facecolor("#08101f")
        ax.plot(kpis["journalier"]["Jour"], kpis["journalier"]["Duree_h"], marker="o", linewidth=2, color="#1dd4df")
        ax.fill_between(kpis["journalier"]["Jour"], kpis["journalier"]["Duree_h"], alpha=0.16, color="#1dd4df")
        ax.set_xlabel("Jour", color="white")
        ax.set_ylabel("Durée d'arrêt (h)", color="white")
        ax.set_title("Évolution journalière des arrêts", color="white", fontsize=13, fontweight="bold", pad=12)
        ax.tick_params(axis="x", rotation=45, colors="white")
        ax.tick_params(axis="y", colors="white")
        ax.grid(axis="y", linestyle="--", alpha=0.18, color="white")
        for spine in ax.spines.values():
            spine.set_color("#294566")
        st.pyplot(fig, use_container_width=True)
    else:
        st.info("Pas de dates exploitables.")


# =========================================================
# PAGE 3 : DONNÉES
# =========================================================
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


# =========================================================
# PAGE 4 : AIDE
# =========================================================
elif page == "Aide":
    st.markdown('<div class="section-chip">Guide utilisateur</div>', unsafe_allow_html=True)

    st.subheader("Principe")
    st.write(
        "L'utilisateur importe seulement la base brute des arrêts. "
        "La plateforme calcule automatiquement les KPI, affiche les graphiques et génère un rapport PDF."
    )

    st.subheader("Colonnes obligatoires")
    st.write([
        "Date",
        "Zone",
        "TAG",
        "Equipement",
        "Imputation arrêt",
        "Durée (H)",
        "Description Arrêt",
        "Type de panne",
        "Usine en arret/marche"
    ])

    st.subheader("Colonnes optionnelles")
    st.write([
        "Famille",
        "Causes majeures"
    ])

    st.subheader("Formules KPI utilisées")
    st.write("Temps Requis = Temps d'ouverture - Arrêts planifiés")
    st.write("Temps de fonctionnement = Temps Requis - Temps arrêts NP")
    st.write("Temps ecart cadence = Temps de fonctionnement - (Tonnage réalisé / Cadence théorique)")
    st.write("Temps Net = Temps de fonctionnement - Temps ecart cadence")
    st.write("TRS = (Temps Net / Temps Requis) × Taux de qualité × 100")
    st.write("Disponibilité = ((Temps Requis - Somme des durées maintenance) / Temps Requis) × 100")
    st.write("MTBF = (Temps Requis - Somme des durées maintenance) / Nombre d'arrêts maintenance")
    st.write("MTTR = Somme des durées maintenance / Nombre d'arrêts maintenance")

    st.subheader("Rapport PDF")
    st.write(
        "Le rapport PDF contient une couverture, un résumé exécutif, les KPI, la synthèse maintenance, "
        "les graphiques par zone, top équipements, Pareto TAG | Equipement, évolution journalière et recommandations."
    )
