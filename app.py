# ✅ VERSION COMPLÈTE avec rétrocompatibilité + coloration PDF/Excel

import streamlit as st
import pandas as pd
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm

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

# === CHARGEMENT / PERSISTENCE ===
params_json = st.sidebar.file_uploader("Importer des paramètres JSON", type="json")
if 'params' not in st.session_state:
    st.session_state.params = default_params.copy()
if params_json is not None:
    st.session_state.params.update(json.load(params_json))
params = st.session_state.params

# RÉTROCOMPATIBILITÉ
if 'impacts_recurrents' not in params:
    if 'impacts' in params and isinstance(params['impacts'], list):
        params['impacts_recurrents'] = params.pop('impacts')
    else:
        params['impacts_recurrents'] = []
if 'impacts_specifiques' not in params:
    params['impacts_specifiques'] = []

# === FORMULAIRE GLOBAL ===
with st.sidebar.form("param_form"):
    st.markdown("### 📋 Paramètres du fonds")
    expand_all = st.checkbox("🧩 Tout déplier", value=False)

    nom_fonds = st.text_input("Nom du fonds", params['nom_fonds'])
    date_vl_connue_str = st.text_input("Date dernière VL connue (jj/mm/aaaa)", params['date_vl_connue'])
    date_fin_fonds_str = st.text_input("Date fin de fonds (jj/mm/aaaa)", params['date_fin_fonds'])

    anr_derniere_vl = float(str(st.text_input("ANR dernière VL connue (€)", f"{params['anr_derniere_vl']:,.2f}"))
        .replace(" ", "").replace(",", ".").replace("€", "") or 0)
    nombre_parts = float(str(st.text_input("Nombre de parts", f"{params['nombre_parts']:,.2f}"))
        .replace(" ", "").replace(",", ".") or 0)

    submitted = st.form_submit_button("🧮 Calculer la projection")

# Exécution seulement si le formulaire est validé
if submitted:
    # === DATES SEMESTRES ===
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

# === SIDEBAR : IMPACTS RÉCURRENTS ===
with st.sidebar.expander("🔁 Impacts récurrents (même chaque semestre)", expanded=expand_all):
impacts_recurrents = []
nb_impacts_rec = st.sidebar.number_input("Nombre d'impacts récurrents", min_value=0, value=len(params['impacts_recurrents']), step=1)
for i in range(nb_impacts_rec):
    libelle_defaut, montant_defaut = params['impacts_recurrents'][i] if i < len(params['impacts_recurrents']) else (f"Impact récurrent {i+1}", 0.0)
    libelle = st.sidebar.text_input(f"Libellé impact récurrent {i+1}", libelle_defaut)
    montant = float(str(st.sidebar.text_input(f"Montant impact récurrent {i+1} (€)", f"{montant_defaut:,.2f}"))
                   .replace(" ", "").replace(",", ".").replace("€", "") or 0)
    impacts_recurrents.append((libelle, montant))

st.sidebar.markdown("---")

# === SIDEBAR : IMPACTS SPÉCIFIQUES ===
with st.sidebar.expander("📅 Impacts spécifiques (montants par semestre)", expanded=False):
    impacts_specifiques = []
    nb_impacts_spec = st.number_input("Nombre d'impacts spécifiques", min_value=0, value=len(params['impacts_specifiques']), step=1)
    for i in range(nb_impacts_spec):
        if i < len(params['impacts_specifiques']):
            imp = params['impacts_specifiques'][i]
            libelle_defaut = imp['libelle']
            montants_defaut = imp['montants']
        else:
            libelle_defaut, montants_defaut = f"Impact spécifique {i+1}", {}
        libelle = st.text_input(f"Libellé impact spécifique {i+1}", libelle_defaut)
        montants_par_semestre = {}
        for d in dates_semestres[1:]:
            key = d.strftime('%d/%m/%Y')
            val_def = montants_defaut.get(key, 0.0)
            montant = float(str(st.text_input(f"{libelle} ({key})", f"{val_def:,.2f}"))
                            .replace(" ", "").replace(",", ".").replace("€", "") or 0)
            montants_par_semestre[key] = montant
        impacts_specifiques.append({"libelle": libelle, "montants": montants_par_semestre})

# === ACTIFS ===
with st.sidebar.expander("🏢 Actifs du portefeuille", expanded=False):
    actifs = []
    nb_actifs = st.number_input("Nombre d'actifs à ajouter", min_value=1, value=max(1, len(params['actifs'])), step=1)
    for i in range(nb_actifs):
        a = params['actifs'][i] if i < len(params['actifs']) else {}
        nom = st.text_input(f"Nom Actif {i+1}", a.get('nom', f"Actif {i+1}"))
        pct = st.slider(f"% Détention Actif {i+1}", 0.0, 100.0, float(a.get('pct_detention', 1.0) * 100), 1.0)
        val_actuelle = float(str(st.text_input(f"Valeur actuelle Actif {i+1} (€)", f"{a.get('valeur_actuelle', 1000000):,.2f}"))
                             .replace(" ", "").replace(",", ".").replace("€", "") or 0)
        val_proj = float(str(st.text_input(f"Valeur projetée S+1 Actif {i+1} (€)", f"{a.get('valeur_projetee', val_actuelle + 50000):,.2f}"))
                         .replace(" ", "").replace(",", ".").replace("€", "") or 0)
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
st.dataframe(projection)

# === EXPORT PARAMÈTRES JSON ===
if submitted:
    if st.sidebar.button("📤 Exporter les paramètres JSON"):
        export_data = {
            "nom_fonds": nom_fonds,
            "date_vl_connue": date_vl_connue_str,
            "date_fin_fonds": date_fin_fonds_str,
            "anr_derniere_vl": anr_derniere_vl,
            "nombre_parts": nombre_parts,
            "impacts_recurrents": impacts_recurrents,
            "impacts_specifiques": impacts_specifiques,
            "actifs": actifs
        }
        json_export = json.dumps(export_data, indent=2).encode('utf-8')
        st.sidebar.download_button("Télécharger paramètres JSON", json_export, file_name="parametres_vl.json")

# === EXPORT PDF ET EXCEL ===
if submitted:
    col1, col2 = st.columns(2)

# Export Excel avec mise en forme conditionnelle
with col1:
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        projection.to_excel(writer, sheet_name='Projection', index=False, startrow=1, startcol=1)
        workbook = writer.book
        worksheet = writer.sheets['Projection']

        for idx, col in enumerate(projection.columns):
            worksheet.write(0, idx + 1, col)
            if "(R)" in col:
                worksheet.conditional_format(1, idx + 1, len(projection), idx + 1, {
                    'type': 'no_blanks',
                    'format': workbook.add_format({'bg_color': '#DDEEFF'})
                })
            elif "(S)" in col:
                worksheet.conditional_format(1, idx + 1, len(projection), idx + 1, {
                    'type': 'no_blanks',
                    'format': workbook.add_format({'bg_color': '#F0F0F0'})
                })
    st.download_button("📥 Exporter en Excel", data=excel_buffer.getvalue(), file_name="projection_vl.xlsx")

# Export PDF avec fond par colonne d'impact
with col2:
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph(f"Atterrissage VL - {nom_fonds}", styles['Title']))
    elements.append(Spacer(1, 12))

    headers = [Paragraph(h, styles['Heading4']) for h in projection.columns]
    data = [headers]
    for i in range(len(projection)):
        row = []
        for j, val in enumerate(projection.iloc[i]):
            style = styles['Normal']
            row.append(Paragraph(str(val), style))
        data.append(row)

    table = Table(data, repeatRows=1)
    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
    ])
    for col_idx, col_name in enumerate(projection.columns):
        if "(R)" in col_name:
            table_style.add('BACKGROUND', (col_idx, 1), (col_idx, -1), colors.HexColor('#E0F0FF'))
        elif "(S)" in col_name:
            table_style.add('BACKGROUND', (col_idx, 1), (col_idx, -1), colors.HexColor('#F5F5F5'))
    table.setStyle(table_style)
    elements.append(table)

    doc.build(elements)
    st.download_button("📄 Exporter en PDF", data=pdf_buffer.getvalue(), file_name="projection_vl.pdf")
