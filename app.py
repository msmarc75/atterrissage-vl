import streamlit as st
import pandas as pd
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# === CHARGEMENT PARAMÈTRES DE BASE ===
default_params = {
    "nom_fonds": "Nom du Fonds",
    "date_vl_connue": "31/12/2024",
    "date_fin_fonds": "31/12/2028",
    "anr_derniere_vl": 10_000_000.0,
    "nombre_parts": 10_000.0,
    "impacts_recurrents": [],
    "impacts_specifiques": [],
    "actifs": []
}

params_json = st.sidebar.file_uploader("Importer des paramètres JSON", type="json")
if 'params' not in st.session_state:
    st.session_state.params = default_params.copy()
if params_json is not None:
    st.session_state.params.update(json.load(params_json))
params = st.session_state.params

# === FORMULAIRE PRINCIPAL ===
with st.sidebar.form("param_form"):
    st.markdown("### 📋 Paramètres du fonds")
    expand_all = st.checkbox("🧩 Tout déplier", value=False)

    nom_fonds = st.text_input("Nom du fonds", params['nom_fonds'])
    date_vl_connue_str = st.text_input("Date dernière VL connue (jj/mm/aaaa)", params['date_vl_connue'])
    date_fin_fonds_str = st.text_input("Date fin de fonds (jj/mm/aaaa)", params['date_fin_fonds'])
    anr_derniere_vl_str = st.text_input("ANR dernière VL connue (€)", f"{params['anr_derniere_vl']:,.2f}")
    nombre_parts_str = st.text_input("Nombre de parts", f"{params['nombre_parts']:,.2f}")

    submitted = st.form_submit_button("🧮 Calculer la projection")

if submitted:
    anr_derniere_vl = float(anr_derniere_vl_str.replace(" ", "").replace(",", ".").replace("€", "") or 0)
    nombre_parts = float(nombre_parts_str.replace(" ", "").replace(",", ".") or 0)

    date_vl_connue = datetime.strptime(date_vl_connue_str, "%d/%m/%Y")
    date_fin_fonds = datetime.strptime(date_fin_fonds_str, "%d/%m/%Y")
    dates_semestres = [date_vl_connue]
    y = date_vl_connue.year
    while datetime(y, 12, 31) <= date_fin_fonds:
        if datetime(y, 6, 30) > date_vl_connue:
            dates_semestres.append(datetime(y, 6, 30))
        if datetime(y, 12, 31) > date_vl_connue:
            dates_semestres.append(datetime(y, 12, 31))
        y += 1

    # === IMPACTS RÉCURRENTS ===
    with st.sidebar.expander("🔁 Impacts récurrents", expanded=expand_all):
        impacts_recurrents = []
        nb_impacts_rec = st.number_input("Nombre d'impacts récurrents", min_value=0, value=len(params['impacts_recurrents']), step=1)
        for i in range(nb_impacts_rec):
            libelle_defaut, montant_defaut = params['impacts_recurrents'][i] if i < len(params['impacts_recurrents']) else (f"Impact récurrent {i+1}", 0.0)
            libelle = st.text_input(f"Libellé impact récurrent {i+1}", libelle_defaut)
            montant_str = st.text_input(f"Montant impact récurrent {i+1} (€)", f"{montant_defaut:,.2f}")
            try:
                montant = float(montant_str.replace(" ", "").replace(",", ".").replace("€", "") or 0)
            except ValueError:
                montant = 0.0
            impacts_recurrents.append((libelle, montant))

    # === IMPACTS SPÉCIFIQUES ===
    with st.sidebar.expander("📅 Impacts spécifiques", expanded=expand_all):
        impacts_specifiques = []
        nb_impacts_spec = st.number_input("Nombre d'impacts spécifiques", min_value=0, value=len(params['impacts_specifiques']), step=1)
        for i in range(nb_impacts_spec):
            if i < len(params['impacts_specifiques']):
                imp = params['impacts_specifiques'][i]
                libelle_defaut = imp['libelle']
                montants_defaut = imp['montants']
            else:
                libelle_defaut = f"Impact spécifique {i+1}"
                montants_defaut = {}
            libelle = st.text_input(f"Libellé impact spécifique {i+1}", libelle_defaut)
            montants_par_semestre = {}
            for d in dates_semestres[1:]:
                key = d.strftime('%d/%m/%Y')
                val_def = montants_defaut.get(key, 0.0)
                montant_str = st.text_input(f"{libelle} ({key})", f"{val_def:,.2f}")
                try:
                    montant = float(montant_str.replace(" ", "").replace(",", ".").replace("€", "") or 0)
                except ValueError:
                    montant = 0.0
                montants_par_semestre[key] = montant
            impacts_specifiques.append({"libelle": libelle, "montants": montants_par_semestre})

    # === ACTIFS ===
    with st.sidebar.expander("🏢 Actifs du portefeuille", expanded=expand_all):
        actifs = []
        nb_actifs = st.number_input("Nombre d'actifs à ajouter", min_value=1, value=max(1, len(params['actifs'])), step=1)
        for i in range(nb_actifs):
            a = params['actifs'][i] if i < len(params['actifs']) else {}
            nom = st.text_input(f"Nom Actif {i+1}", a.get('nom', f"Actif {i+1}"))
            pct = st.slider(f"% Détention Actif {i+1}", 0.0, 100.0, float(a.get('pct_detention', 1.0) * 100), 1.0)
            val_actuelle_str = st.text_input(f"Valeur actuelle Actif {i+1} (€)", f"{a.get('valeur_actuelle', 1000000):,.2f}")
            val_proj_str = st.text_input(f"Valeur projetée S+1 Actif {i+1} (€)", f"{a.get('valeur_projetee', 1050000):,.2f}")
            try:
                val_actuelle = float(val_actuelle_str.replace(" ", "").replace(",", ".").replace("€", "") or 0)
                val_proj = float(val_proj_str.replace(" ", "").replace(",", ".").replace("€", "") or 0)
            except ValueError:
                val_actuelle = 0.0
                val_proj = 0.0
            actifs.append({
                "nom": nom,
                "pct_detention": pct / 100,
                "valeur_actuelle": val_actuelle,
                "valeur_projetee": val_proj,
                "variation": (val_proj - val_actuelle) * (pct / 100)
            })

    # === PROJECTION ===
    anr_courant = anr_derniere_vl
    vl_semestres = []
    projection_rows = []
    for i, date in enumerate(dates_semestres):
        row = {"Date": date.strftime('%d/%m/%Y')}
        total_var_actifs = sum(a['variation'] if i == 1 else 0 for a in actifs)
        for a in actifs:
            row[f"Actif - {a['nom']}"] = f"{a['variation'] if i == 1 else 0:,.2f} €".replace(",", " ").replace(".", ",")
        total_impacts = 0
        for libelle, montant in impacts_recurrents:
            row[f"Impact - {libelle} (R)"] = f"{montant:,.2f} €".replace(",", " ").replace(".", ",")
            total_impacts += montant
        for imp in impacts_specifiques:
            key = date.strftime('%d/%m/%Y')
            montant = imp['montants'].get(key, 0.0)
            row[f"Impact - {imp['libelle']} (S)"] = f"{montant:,.2f} €".replace(",", " ").replace(".", ",")
            total_impacts += montant
        if i > 0:
            anr_courant += total_var_actifs + total_impacts
        vl = anr_courant / nombre_parts
        vl_semestres.append(vl)
        row["VL prévisionnelle (€)"] = f"{vl:,.2f} €".replace(",", " ").replace(".", ",")
        projection_rows.append(row)

    projection = pd.DataFrame(projection_rows)
    st.subheader("📊 Résultat de la projection")
    st.dataframe(projection)

    # === GRAPHIQUE ===
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(dates_semestres, vl_semestres, marker='o', color='#0000DC', linewidth=2)
    for i, txt in enumerate(vl_semestres):
        ax.annotate(f"{txt:,.2f} €".replace(",", " ").replace(".", ","),
                    (dates_semestres[i], vl_semestres[i]),
                    textcoords="offset points", xytext=(0, 10), ha='center',
                    fontsize=9, color='white',
                    bbox=dict(boxstyle="round,pad=0.3", fc="#0000DC", ec="#0000DC"))
    ax.set_title(f"Atterrissage VL - {nom_fonds}", fontsize=14, color="#0000DC")
    ax.set_ylabel("VL (€)", fontsize=12)
    ax.set_xticks(dates_semestres)
    ax.set_xticklabels([d.strftime('%b-%y').capitalize() for d in dates_semestres], rotation=45, ha='right')
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:,.2f} €".replace(",", " ").replace(".", ",")))
    ax.grid(True, linestyle='--', alpha=0.5)
    st.pyplot(fig)

    # === EXPORTS ===
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        projection.to_excel(writer, index=False, sheet_name="Projection")
    buffer.seek(0)
    st.download_button("📥 Exporter Excel", buffer, file_name="projection_vl.xlsx", mime="application/vnd.ms-excel")
