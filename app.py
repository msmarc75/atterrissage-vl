# EN-TETE : modules nécessaires (PDF supprimé)
import streamlit as st
import pandas as pd
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io

# --- Fonction de formatage euro ---
def format_fr_euro(valeur):
    return f"{valeur:,.2f} €".replace(",", " ").replace(".", ",")

# --- Champ numérique simplifié ---
def champ_numerique(label, valeur):
    champ = st.sidebar.text_input(label, value=format_fr_euro(valeur))
    champ = champ.replace(" ", "").replace(",", ".").replace("€", "")
    return float(champ) if champ else 0.0

# === PARAMÈTRES INITIAUX ===
default_params = {
    "nom_fonds": "Nom du Fonds",
    "date_vl_connue": "31/12/2024",
    "date_fin_fonds": "31/12/2028",
    "anr_derniere_vl": 10_000_000.0,
    "nombre_parts": 10_000.0,
    "impacts": [
        ("Frais corporate", -50_000.0),
        ("Honoraires NIV", -30_000.0)
    ],
    "actifs": []
}

# === CHARGEMENT / PERSISTENCE DES PARAMÈTRES ===
st.sidebar.header("Importer / Exporter")
params_json = st.sidebar.file_uploader("Importer des paramètres JSON", type="json")

if 'params' not in st.session_state:
    st.session_state.params = default_params.copy()

if params_json is not None:
    st.session_state.params.update(json.load(params_json))

params = st.session_state.params

# === SAISIE UTILISATEUR ===
nom_fonds = st.sidebar.text_input("Nom du fonds", params['nom_fonds'])
date_vl_connue_str = st.sidebar.text_input("Date dernière VL connue (jj/mm/aaaa)", params['date_vl_connue'])
date_fin_fonds_str = st.sidebar.text_input("Date fin de fonds (jj/mm/aaaa)", params['date_fin_fonds'])
anr_derniere_vl = champ_numerique("ANR dernière VL connue (€)", params['anr_derniere_vl'])
nombre_parts = champ_numerique("Nombre de parts", params['nombre_parts'])

# Impacts personnalisés
st.sidebar.header("Impacts semestriels personnalisés")
impacts = []
nb_impacts = st.sidebar.number_input("Nombre d'impacts", min_value=0, value=len(params['impacts']), step=1)
for i in range(nb_impacts):
    if i < len(params['impacts']):
        libelle_defaut, montant_defaut = params['impacts'][i]
    else:
        libelle_defaut, montant_defaut = f"Impact {i+1}", 0.0
    libelle = st.sidebar.text_input(f"Libellé impact {i+1}", libelle_defaut)
    montant = champ_numerique(f"Montant semestriel impact {i+1} (€)", montant_defaut)
    impacts.append((libelle, montant))

# Actifs
st.sidebar.header("Ajouter des Actifs")
actifs = []
nb_actifs = st.sidebar.number_input("Nombre d'actifs à ajouter", min_value=1, value=max(1, len(params['actifs'])), step=1)
for i in range(nb_actifs):
    st.sidebar.subheader(f"Actif {i+1}")
    if i < len(params['actifs']):
        a = params['actifs'][i]
        nom_defaut = a['nom']
        pct_defaut = a['pct_detention'] * 100
        val_actuelle = a.get('valeur_actuelle', 1_000_000.0)
        val_proj = a.get('valeur_projetee', val_actuelle + 50_000)
    else:
        nom_defaut = f"Actif {i+1}"
        pct_defaut = 100.0
        val_actuelle = 1_000_000.0
        val_proj = 1_050_000.0

    nom_actif = st.sidebar.text_input(f"Nom Actif {i+1}", nom_defaut)
    pct_detention = st.sidebar.slider(f"% Détention Actif {i+1}", min_value=0.0, max_value=100.0, value=pct_defaut, step=1.0)
    valeur_actuelle = champ_numerique(f"Valeur actuelle Actif {i+1} (€)", val_actuelle)
    valeur_projetee = champ_numerique(f"Valeur projetée en S+1 Actif {i+1} (€)", val_proj)
    variation = (valeur_projetee - valeur_actuelle) * (pct_detention / 100)
    actifs.append({
        "nom": nom_actif,
        "pct_detention": pct_detention / 100,
        "valeur_actuelle": valeur_actuelle,
        "valeur_projetee": valeur_projetee,
        "variation": variation
    })

# === DATES & PROJECTION ===
date_vl_connue = datetime.strptime(date_vl_connue_str, "%d/%m/%Y")
date_fin_fonds = datetime.strptime(date_fin_fonds_str, "%d/%m/%Y")

# Générer les semestres
from dateutil.relativedelta import relativedelta
dates_semestres = [date_vl_connue]
while dates_semestres[-1] < date_fin_fonds:
    prochaine_date = dates_semestres[-1] + relativedelta(months=6)
    if prochaine_date <= date_fin_fonds:
        dates_semestres.append(prochaine_date)
    else:
        break

# === CALCUL PROJECTION DÉTAILLÉE ===
vl_semestres = []
anr_courant = anr_derniere_vl
projection_rows = []

for i, date in enumerate(dates_semestres):
    row = {"Date": date.strftime('%d/%m/%Y')}

    total_var_actifs = sum(a['variation'] if i == 1 else 0 for a in actifs)
    for a in actifs:
        var = a['variation'] if i == 1 else 0
        row[f"Actif - {a['nom']}"] = format_fr_euro(var)

    total_impacts = sum(montant for libelle, montant in impacts)
    for libelle, montant in impacts:
        row[f"Impact - {libelle}"] = format_fr_euro(montant)

    if i > 0:
        anr_courant += total_var_actifs + total_impacts

    vl = anr_courant / nombre_parts
    vl_semestres.append(vl)
    row["VL prévisionnelle (€)"] = format_fr_euro(vl)
    projection_rows.append(row)

# === AFFICHAGE TABLEAU ===
projection = pd.DataFrame(projection_rows)
st.title("Atterrissage VL")
st.subheader("VL prévisionnelle")
st.dataframe(projection)

# === GRAPHIQUE ===
couleur_bleue = "#0000DC"
fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(
    dates_semestres,
    vl_semestres,
    linewidth=2.5,
    marker='o',
    markersize=7,
    color=couleur_bleue,
    markerfacecolor=couleur_bleue,
    markeredgewidth=1,
    markeredgecolor=couleur_bleue
)
for i, txt in enumerate(vl_semestres):
    ax.annotate(
        format_fr_euro(txt),
        (dates_semestres[i], vl_semestres[i]),
        textcoords="offset points",
        xytext=(0, 10),
        ha='center',
        fontsize=9,
        color='white',
        bbox=dict(boxstyle="round,pad=0.3", fc=couleur_bleue, ec=couleur_bleue, alpha=0.9)
    )
ax.set_title(f"Atterrissage VL - {nom_fonds}", fontsize=16, fontweight='bold', color=couleur_bleue, pad=20)
ax.set_ylabel("VL (€)", fontsize=12, color=couleur_bleue)
ax.set_xticks(dates_semestres)
ax.set_xticklabels([d.strftime('%b-%y').capitalize() for d in dates_semestres], rotation=45, ha='right', fontsize=10, color=couleur_bleue)
ax.tick_params(axis='y', labelcolor=couleur_bleue)
ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:,.2f} €".replace(",", " ").replace(".", ",")))
fig.patch.set_facecolor('white')
ax.set_facecolor('white')
ax.spines['right'].set_visible(False)
ax.spines['top'].set_visible(False)
st.pyplot(fig)

# === PLACEHOLDER BOUTON PDF ===
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        label="📅 Exporter la projection avec graphique Excel",
        data=b"",  # Remplacer par un buffer si vous le générez
        file_name="projection_vl.xlsx",
        mime="application/vnd.ms-excel"
    )
with col2:
    st.write("\ud83d\udcca Export PDF désactivé.")
