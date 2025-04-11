import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io

st.title("Atterrissage VL")

def format_fr_euro(valeur):
    # S'assurer que la valeur est arrondie à deux décimales
    valeur_arrondie = round(valeur, 2)
    return f"{valeur_arrondie:,.2f} €".replace(",", " ").replace(".", ",")

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
    "actifs": [],
    "impacts_specifiques": {}  # Nouveau dictionnaire pour impacts spécifiques
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
nb_impacts = st.sidebar.number_input("Nombre d'impacts récurrents", min_value=0, value=len(params['impacts']), step=1)
for i in range(nb_impacts):
    if i < len(params['impacts']):
        libelle_defaut, montant_defaut = params['impacts'][i]
    else:
        libelle_defaut, montant_defaut = f"Impact {i+1}", 0.0
    libelle = st.sidebar.text_input(f"Libellé impact {i+1}", libelle_defaut)
    montant = champ_numerique(f"Montant semestriel impact {i+1} (€)", montant_defaut)
    impacts.append((libelle, montant))

# Initialiser le dictionnaire des impacts spécifiques
if 'impacts_specifiques' not in params:
    params['impacts_specifiques'] = {}
impacts_specifiques = params.get('impacts_specifiques', {})

# Ajouter un expander pour les impacts spécifiques par semestre
with st.sidebar.expander("Impacts spécifiques par semestre"):
    st.write("Ajoutez ici des impacts ponctuels pour des semestres spécifiques")
    
    # Sélection du semestre
    semestres_formatted = [date.strftime('%d/%m/%Y') for date in dates_semestres[1:]]  # Exclure la première date
    if semestres_formatted:  # Si la liste n'est pas vide
        semestre_selectionne = st.selectbox("Sélectionner un semestre", semestres_formatted)
        
        # Champ pour le libellé et le montant
        libelle_specifique = st.text_input("Libellé de l'impact spécifique", "Impact ponctuel")
        
        # Récupérer la valeur précédente si elle existe
        valeur_precedente = 0.0
        for impact_key, impact_value in impacts_specifiques.get(semestre_selectionne, {}).items():
            if impact_key == libelle_specifique:
                valeur_precedente = impact_value
                break
        
        montant_specifique = st.number_input(
            "Montant de l'impact spécifique (€)", 
            value=float(valeur_precedente),
            step=1000.0, 
            format="%.2f"
        )
        
        # Bouton pour ajouter l'impact
        if st.button("Ajouter/Modifier cet impact spécifique"):
            if semestre_selectionne not in impacts_specifiques:
                impacts_specifiques[semestre_selectionne] = {}
            impacts_specifiques[semestre_selectionne][libelle_specifique] = montant_specifique
            st.success(f"Impact '{libelle_specifique}' de {montant_specifique} € ajouté pour le semestre {semestre_selectionne}")
    
    # Afficher les impacts spécifiques actuels
    if impacts_specifiques:
        st.write("Impacts spécifiques actuels :")
        for semestre, impacts_dict in impacts_specifiques.items():
            for libelle, montant in impacts_dict.items():
                st.write(f"- {semestre} : {libelle} ({montant} €)")
                
                # Ajouter un bouton de suppression
                if st.button(f"Supprimer '{libelle}' du {semestre}", key=f"del_{semestre}_{libelle}"):
                    if semestre in impacts_specifiques and libelle in impacts_specifiques[semestre]:
                        del impacts_specifiques[semestre][libelle]
                        # Si le dictionnaire est vide pour ce semestre, le supprimer également
                        if not impacts_specifiques[semestre]:
                            del impacts_specifiques[semestre]
                        st.success(f"Impact '{libelle}' supprimé pour le semestre {semestre}")
                        st.rerun()
    else:
        st.info("Aucun impact spécifique défini")

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

dates_semestres = [date_vl_connue]
y = date_vl_connue.year
while datetime(y, 12, 31) <= date_fin_fonds:
    if datetime(y, 6, 30) > date_vl_connue:
        dates_semestres.append(datetime(y, 6, 30))
    if datetime(y, 12, 31) > date_vl_connue:
        dates_semestres.append(datetime(y, 12, 31))
    y += 1

# === CALCUL PROJECTION DÉTAILLÉE ===
vl_semestres = []
anr_courant = anr_derniere_vl
projection_rows = []

for i, date in enumerate(dates_semestres):
    date_str = date.strftime('%d/%m/%Y')
    row = {"Date": date_str}

    # Variation par actif (S+1 uniquement)
    total_var_actifs = 0
    for a in actifs:
        var = a['variation'] if i == 1 else 0
        row[f"Actif - {a['nom']}"] = format_fr_euro(var)
        total_var_actifs += var

    # Impacts récurrents
    total_impacts = 0
    for libelle, montant in impacts:
        row[f"Impact - {libelle}"] = format_fr_euro(montant)
        total_impacts += montant
    
    # Impacts spécifiques pour ce semestre
    total_impacts_specifiques = 0
    if i > 0:  # On ne considère pas les impacts spécifiques pour le semestre initial
        date_str = date.strftime('%d/%m/%Y')
        if date_str in impacts_specifiques:
            for libelle, montant in impacts_specifiques[date_str].items():
                row[f"Impact spécifique - {libelle}"] = format_fr_euro(montant)
                total_impacts_specifiques += montant

    if i > 0:
        anr_courant += total_var_actifs + total_impacts + total_impacts_specifiques

    vl = anr_courant / nombre_parts
    # Arrondir à deux décimales
    vl = round(vl, 2)
    vl_semestres.append(vl)
    row["VL prévisionnelle (€)"] = format_fr_euro(vl)

    projection_rows.append(row)

# === AFFICHAGE TABLEAU ===
projection = pd.DataFrame(projection_rows)
st.subheader("VL prévisionnelle")
st.dataframe(projection)

# Définir Helvetica comme police globale
plt.rcParams['font.family'] = 'Open Sans'

# === GRAPHIQUE BLEU STYLÉ ===
couleur_bleue = "#0000DC"

fig, ax = plt.subplots(figsize=(10, 5))

# Courbe
ax.plot(
    dates_semestres,
    vl_semestres,
    linewidth=2.5,
    marker='o',
    markersize=7,
    color=couleur_bleue,
    markerfacecolor=couleur_bleue,  # Points remplis de couleur bleue
    markeredgewidth=1,
    markeredgecolor=couleur_bleue
)

# Annotations de chaque point
for i, txt in enumerate(vl_semestres):
    # S'assurer que chaque valeur est arrondie à 2 décimales
    txt_arrondi = round(txt, 2)
    ax.annotate(
        format_fr_euro(txt_arrondi),
        (dates_semestres[i], vl_semestres[i]),
        textcoords="offset points",
        xytext=(0, 10),
        ha='center',
        fontsize=9,
        color='white',
        bbox=dict(boxstyle="round,pad=0.3", fc=couleur_bleue, ec=couleur_bleue, alpha=0.9)
    )

# Titres et axes
ax.set_title(f"Atterrissage VL - {nom_fonds}", fontsize=16, fontweight='bold', color=couleur_bleue, pad=20)
ax.set_ylabel("VL (€)", fontsize=12, color=couleur_bleue)

# Ticks
ax.set_xticks(dates_semestres)
ax.set_xticklabels(
    [d.strftime('%b-%y').capitalize() for d in dates_semestres],
    rotation=45,
    ha='right',
    fontsize=10,
    color=couleur_bleue
)
ax.tick_params(axis='y', labelcolor=couleur_bleue)

ax.yaxis.set_major_formatter(
    ticker.FuncFormatter(lambda x, _: f"{round(x, 2):,.2f} €".replace(",", " ").replace(".", ","))
)

# Fond
fig.patch.set_facecolor('white')
ax.set_facecolor('white')

# Supprimer les contours inutiles
ax.spines['right'].set_visible(False)
ax.spines['top'].set_visible(False)

st.pyplot(fig)

# === EXPORT EXCEL AVEC GRAPHIQUE ===
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    # Exporter les données - nous n'utilisons pas to_excel directement
    # car nous voulons créer une marge d'une ligne et une colonne
    workbook = writer.book
    
    # Créer un onglet "Atterrissage VL" à la place de "Projection"
    worksheet = workbook.add_worksheet('Atterrissage VL')
    
    # Formats pour l'Excel
    header_format = workbook.add_format({
        'bold': True,
        'font_color': couleur_bleue,
        'bg_color': '#F0F0F0',
        'border': 0,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    # Format pour cellules texte alignées à droite (pour les valeurs monétaires)
    text_right_format = workbook.add_format({
        'align': 'right'
    })
    
    # Définir une marge pour laisser une ligne vide en haut et une colonne vide à gauche
    row_offset = 1  # Décalage d'une ligne
    col_offset = 1  # Décalage d'une colonne
    
    # Appliquer le format d'en-tête (avec décalage)
    for col_num, value in enumerate(projection.columns.values):
        worksheet.write(row_offset, col_num + col_offset, value, header_format)
    
    # Enlever le quadrillage
    worksheet.hide_gridlines(2)  # 2 = enlever complètement le quadrillage
    
    # Créer des formats pour les nombres et montants
    number_format = workbook.add_format({
        'align': 'right',
        'num_format': '#,##0.00',  # Format standard pour les nombres avec décimales
    })
    
    # Format pour les nombres négatifs en rouge
    negative_number_format = workbook.add_format({
        'align': 'right',
        'num_format': '#,##0.00',  # Format standard pour les nombres avec décimales
        'font_color': 'red',
    })
    
    # Créer un format pour les cellules monétaires
    money_format = workbook.add_format({
        'align': 'right',
        'num_format': '#,##0.00 €',  # Format monétaire sans zéros superflus
    })
    
    # Format pour les montants négatifs en rouge
    negative_money_format = workbook.add_format({
        'align': 'right',
        'num_format': '#,##0.00 €',  # Format monétaire sans zéros superflus
        'font_color': 'red',
    })
    
    # Mettre en forme les colonnes monétaires et ajuster les largeurs (avec décalage)
    for idx, col in enumerate(projection.columns):
        # Calculer la largeur optimale (plus précise)
        col_values = projection[col].astype(str)
        max_len = max(
            max([len(str(s)) for s in col_values]),
            len(col)
        ) + 3  # Marge supplémentaire
        
        # Limiter la largeur maximum
        max_len = min(max_len, 40)
        
        # Appliquer la largeur (avec décalage)
        worksheet.set_column(idx + col_offset, idx + col_offset, max_len)
        
                        # Appliquer le format monétaire si la colonne contient "€"
        if "€" in col or any(s in col for s in ["Impact", "Actif"]):
            # Écrire comme nombres réels avec format monétaire
            for row_num in range(len(projection)):
                cell_value = projection.iloc[row_num, idx]
                
                # Extraire la valeur numérique
                if isinstance(cell_value, str) and "€" in cell_value:
                    numeric_value = float(cell_value.replace(" ", "").replace("€", "").replace(",", "."))
                else:
                    try:
                        numeric_value = float(cell_value)
                    except (ValueError, TypeError):
                        numeric_value = 0
                
                # Appliquer le format approprié selon si le nombre est négatif ou positif (avec décalage)
                if numeric_value < 0:
                    worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, negative_money_format)
                else:
                    worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, money_format)
        # Format spécial pour les colonnes VL - nombre à deux décimales
        elif "VL" in col:
            for row_num in range(len(projection)):
                cell_value = projection.iloc[row_num, idx]
                
                # Extraire la valeur numérique
                if isinstance(cell_value, str) and "€" in cell_value:
                    numeric_value = float(cell_value.replace(" ", "").replace("€", "").replace(",", "."))
                else:
                    try:
                        numeric_value = float(cell_value)
                    except (ValueError, TypeError):
                        numeric_value = 0
                
                # Format nombre à deux décimales pour VL
                if numeric_value < 0:
                    worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, negative_number_format)
                else:
                    worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, number_format)
        else:
            # Pour les autres colonnes (non monétaires) - avec décalage
            for row_num in range(len(projection)):
                worksheet.write(row_num + row_offset + 1, idx + col_offset, projection.iloc[row_num, idx])
    
    # Aussi ajuster la largeur de la première colonne vide
    worksheet.set_column(0, 0, 3)  # Largeur de 3 pour la colonne vide
    
    # --- AJOUT DU GRAPHIQUE DIRECTEMENT DANS LA FEUILLE DE PROJECTION ---
    
    # Préparation des données pour le graphique (sans nouvelle feuille)
    # Nous allons ajouter ces données directement en-dessous du tableau principal
    
    # Déterminer où commencer à ajouter les données du graphique (après le tableau principal)
    # Décalé en fonction de l'ajout de la ligne vide
    start_row = len(projection) + 3 + row_offset  # +3 pour laisser un peu d'espace
    
    # Ajouter un titre pour la section graphique
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'font_color': couleur_bleue,
        'align': 'left'
    })
    worksheet.write(start_row, col_offset, f'Atterrissage VL - {nom_fonds}', title_format)
    
    # En-têtes pour les données du graphique (avec décalage)
    headers_format = workbook.add_format({
        'bold': True,
        'font_color': couleur_bleue,
        'bg_color': '#F0F0F0',
        'border': 0,
        'align': 'center'
    })
    worksheet.write(start_row + 1, col_offset, 'Date', headers_format)
    worksheet.write(start_row + 1, col_offset + 1, 'VL', headers_format)
    
    # Format pour les dates
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    
    # Écrire les données pour le graphique (avec décalage)
    for i, (date, vl) in enumerate(zip(dates_semestres, vl_semestres)):
        row = start_row + 2 + i
        # Écrire la date
        worksheet.write_datetime(row, col_offset, date, date_format)
        # Écrire la VL comme nombre
        worksheet.write_number(row, col_offset + 1, float(vl))
    
    # Créer le graphique simple avec style similaire à l'application
    chart = workbook.add_chart({'type': 'line'})
    
    # Configurer la série de données pour ressembler au graphique de l'application
    # Note: mise à jour des références aux cellules avec le décalage
    chart.add_series({
        'categories': [worksheet.name, start_row + 2, col_offset, start_row + 1 + len(dates_semestres), col_offset],
        'values': [worksheet.name, start_row + 2, col_offset + 1, start_row + 1 + len(dates_semestres), col_offset + 1],
        'marker': {
            'type': 'circle',
            'size': 8,
            'border': {'color': couleur_bleue},
            'fill': {'color': couleur_bleue}  # Points remplis de couleur bleue
        },
        'line': {
            'color': couleur_bleue,
            'width': 2.5
        },
        'data_labels': {
            'value': True,
            'position': 'above',
            'font': {
                'color': 'white',
                'bold': True,
                'size': 9
            },
            'num_format': '#,##0.00 "€"',  # Format sans zéros superflus
            'border': {'color': couleur_bleue},
            'fill': {'color': couleur_bleue}
        },
    })
    
    # Configurer le graphique pour qu'il ressemble à celui de l'application
    chart.set_title({
        'name': f'Atterrissage VL - {nom_fonds}',  # Même titre que dans l'interface
        'name_font': {
            'size': 16,  # Même taille que dans l'interface
            'color': couleur_bleue,
            'bold': True
        }
    })
    
    chart.set_x_axis({
        'name': '',  # Pas de titre d'axe
        'num_format': 'mmm-yy',  # Format "jun-25", "dec-25"
        'num_font': {
            'rotation': 45,
            'color': couleur_bleue
        },
        'line': {'color': couleur_bleue},
        'major_gridlines': {'visible': False}
    })
    
    chart.set_y_axis({
        'name': 'VL (€)',
        'name_font': {
            'color': couleur_bleue,
            'bold': True
        },
        'num_format': '#,##0.00 "€"',  # Format sans zéros superflus
        'num_font': {'color': couleur_bleue},
        'line': {'color': couleur_bleue},
        'major_gridlines': {'visible': True, 'line': {'color': '#E0E0E0', 'width': 0.5}}  # Lignes de grille légères
    })
    
    chart.set_legend({'none': True})
    chart.set_chartarea({'border': {'none': True}, 'fill': {'color': 'white'}})
    chart.set_plotarea({'border': {'none': True}, 'fill': {'color': 'white'}})
    
    # Insérer le graphique dans la feuille (avec décalage approprié pour la position du graphique)
    worksheet.insert_chart(start_row + 2, col_offset + 3, chart, {'x_scale': 1.5, 'y_scale': 1.2})

buffer.seek(0)

# Simplification des boutons d'exportation (retiré l'export PDF)
date_aujourd_hui = datetime.now().strftime("%Y%m%d")
nom_fichier_excel = f"{date_aujourd_hui} - Atterrissage VL - {nom_fonds}.xlsx"

st.download_button(
    label="📥 Exporter la projection avec graphique Excel",
    data=buffer,
    file_name=nom_fichier_excel,
    mime="application/vnd.ms-excel"
)

# === EXPORT PARAMÈTRES JSON ===
if st.sidebar.button("📤 Exporter les paramètres JSON"):
    export_data = {
        "nom_fonds": nom_fonds,
        "date_vl_connue": date_vl_connue_str,
        "date_fin_fonds": date_fin_fonds_str,
        "anr_derniere_vl": anr_derniere_vl,
        "nombre_parts": nombre_parts,
        "impacts": impacts,
        "actifs": actifs,
        "impacts_specifiques": impacts_specifiques
    }
    date_aujourd_hui = datetime.now().strftime("%Y%m%d")
    nom_fichier_json = f"{date_aujourd_hui} - Atterrissage VL - {nom_fonds}.json"
    json_export = json.dumps(export_data, indent=2).encode('utf-8')
    st.sidebar.download_button("Télécharger paramètres JSON", json_export, file_name=nom_fichier_json)

# === BOUTON RÉINITIALISATION ===
if st.sidebar.button("♻️ Réinitialiser les paramètres"):
    st.session_state.params = default_params.copy()
    st.rerun()