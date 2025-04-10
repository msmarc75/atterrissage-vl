import streamlit as st
import pandas as pd
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io

# Importations pour l'export PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm

st.title("Atterrissage VL")

def format_fr_euro(valeur):
    return f"{valeur:,.2f} €".replace(",", " ").replace(".", ",")

def champ_numerique(label, valeur):
    champ = st.sidebar.text_input(label, value=format_fr_euro(valeur))
    champ = champ.replace(" ", "").replace(",", ".").replace("€", "")
    return float(champ) if champ else 0.0

# === FONCTION D'EXPORT PDF ===
def export_pdf(projection, dates_semestres, vl_semestres, nom_fonds, couleur_bleue):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm, cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    import io
    import datetime
    
    # Créer un buffer pour le PDF
    buffer = io.BytesIO()
    
    # Configuration du document en format paysage
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=landscape(A4),
        leftMargin=1.5*cm,
        rightMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )
    
    # Préparer les styles
    styles = getSampleStyleSheet()
    
    # Style pour le titre principal
    styles.add(ParagraphStyle(
        name='TitrePrincipal',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        textColor=colors.HexColor(couleur_bleue),
        fontSize=18,
        spaceAfter=10,  # Réduit
        alignment=1,  # Centré
        leading=22,   # Réduit
    ))
    
    # Style pour le texte normal
    normal_style = ParagraphStyle(
        name='NormalStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,  # Légèrement réduit
        leading=11,
    )
    
    # Style pour les en-têtes de tableau
    header_style = ParagraphStyle(
        name='HeaderStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,  # Légèrement réduit
        textColor=colors.HexColor(couleur_bleue),
        alignment=1,  # Centré
        leading=11,
    )
    
    # Style pour les notes de bas de page
    footer_style = ParagraphStyle(
        name='FooterStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=7,  # Réduit
        textColor=colors.darkgrey,
        alignment=1,  # Centré
    )
    
    # Liste des éléments à inclure dans le PDF
    elements = []
    
    # === PREMIÈRE PAGE : TABLEAU ===
    # Groupe tous les éléments de la première page pour qu'ils restent ensemble
    page1_elements = []
    
    # Ajouter le titre principal
    title = Paragraph(f"Atterrissage VL - {nom_fonds}", styles['TitrePrincipal'])
    page1_elements.append(title)
    
    # Ajouter la date de génération
    date_generation = datetime.datetime.now().strftime("%d/%m/%Y")
    date_text = Paragraph(f"Document généré le {date_generation}", footer_style)
    page1_elements.append(date_text)
    page1_elements.append(Spacer(1, 8))  # Réduit
    
    # Préparer les données du tableau
    data = []
    
    # Convertir les en-têtes en paragraphes
    headers = []
    for header in projection.columns:
        headers.append(Paragraph(header, header_style))
    
    data.append(headers)
    
    # Format des données
    for i in range(len(projection)):
        row = []
        for j, val in enumerate(projection.iloc[i].tolist()):
            cell_style = normal_style
            # Appliquer une couleur rouge aux valeurs négatives
            if isinstance(val, str) and "€" in val and "-" in val:
                cell_style = ParagraphStyle(
                    name='NegativeValue',
                    parent=normal_style,
                    textColor=colors.red
                )
            
            # Formatage du texte pour le tableau
            if j == 0:  # Colonne de date
                row.append(Paragraph(str(val), normal_style))
            else:  # Colonnes numériques
                row.append(Paragraph(str(val), cell_style))
        
        data.append(row)
    
    # Largeurs de colonnes optimisées pour format paysage
    col_widths = [2.5*cm]  # Date
    
    # Ajuster les largeurs pour le format paysage
    for header in projection.columns[1:]:
        if 'Impact' in header or 'Actif' in header:
            col_widths.append(4.5*cm)  # Plus large pour les impacts/actifs
        else:
            col_widths.append(3.5*cm)  # Valeurs
    
    # Assurer que la largeur totale correspond à la page
    available_width = landscape(A4)[0] - 3*cm  # Largeur disponible avec marges
    if sum(col_widths) > available_width:
        scale_factor = available_width / sum(col_widths)
        col_widths = [w * scale_factor for w in col_widths]
    
    # Créer le tableau
    table = Table(data, repeatRows=1, colWidths=col_widths)
    
    # Style du tableau
    table_style = TableStyle([
        # En-têtes
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E6E6E6')),  # Gris clair
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor(couleur_bleue)),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),  # Réduit
        ('TOPPADDING', (0, 0), (-1, 0), 8),  # Réduit
        
        # Corps du tableau
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 1), (-1, -1), 4),  # Réduit
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),  # Réduit
        
        # Lignes alternées pour meilleure lisibilité
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#F9F9F9')),
        ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#F9F9F9')),
        ('BACKGROUND', (0, 6), (-1, 6), colors.HexColor('#F9F9F9')),
        
        # Bordures élégantes
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor(couleur_bleue)),
    ])
    
    # Appliquer le style
    table.setStyle(table_style)
    
    # Ajouter le tableau aux éléments de la première page
    page1_elements.append(table)
    
    # Ajouter un pied de page pour la première page
    page1_elements.append(Spacer(1, 8))  # Réduit
    footer_text = Paragraph(f"Document confidentiel - {nom_fonds} - Page 1/2", footer_style)
    page1_elements.append(footer_text)
    
    # Ajouter tous les éléments de la première page avec KeepTogether
    elements.append(KeepTogether(page1_elements))
    
    # Saut de page
    elements.append(PageBreak())
    
    # === DEUXIÈME PAGE : GRAPHIQUE CENTRÉ VERTICALEMENT ===
    # Utiliser également KeepTogether pour la deuxième page
    page2_elements = []
    
    # Ajouter le titre de la deuxième page
    title2 = Paragraph(f"Atterrissage VL - {nom_fonds}", styles['TitrePrincipal'])
    page2_elements.append(title2)
    
    # Espace avant le graphique pour le centrer verticalement
    page2_elements.append(Spacer(1, 4*cm))  # Ajusté
    
    # Créer un graphique adapté au format paysage
    plt.figure(figsize=(11, 5.2), dpi=120)  # Réduit légèrement
    
    # Style du graphique
    fig, ax = plt.subplots(figsize=(11, 5.2))
    
    # Couleur principale
    color_main = couleur_bleue
    
    # Tracer la courbe SANS les annotations d'abord
    ax.plot(
        dates_semestres,
        vl_semestres,
        linewidth=2.5,  # Réduit légèrement
        marker='o',
        markersize=8,  # Réduit légèrement
        color=color_main,
        markerfacecolor=color_main,
        markeredgewidth=1.5,
        markeredgecolor='white',
        zorder=5
    )
    
    # Ajouter une grille légère de fond
    ax.grid(True, linestyle='--', linewidth=0.5, color='#CCCCCC', alpha=0.7)
    
    # Style des axes
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_color('#BBBBBB')
    ax.spines['bottom'].set_color('#BBBBBB')
    
    # Formatage des dates sur l'axe X
    ax.set_xticks(dates_semestres)
    ax.set_xticklabels(
        [d.strftime('%b-%y').capitalize() for d in dates_semestres],
        rotation=45,
        ha='right',
        fontsize=10,
        color='#555555'
    )
    
    # Formatage des valeurs sur l'axe Y
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_formatter(
        ticker.FuncFormatter(lambda x, _: f"{x:,.2f} €".replace(",", " ").replace(".", ","))
    )
    ax.tick_params(axis='y', colors='#555555', labelsize=10)
    
    # Réglage des limites pour un meilleur cadrage
    y_min = min(vl_semestres) * 0.95
    y_max = max(vl_semestres) * 1.05
    ax.set_ylim(y_min, y_max)
    
    # Fond et esthétique
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#FAFAFA')
    
    # Style professionnel du graphique
    ax.set_title(f"Atterrissage VL - {nom_fonds}", fontsize=16, fontweight='bold', color=color_main, pad=20)
    ax.set_ylabel("VL (€)", fontsize=11, fontweight='bold', color=color_main)
    
    # IMPORTANT: Ajouter les annotations en DERNIER pour les mettre au premier plan
    for i, txt in enumerate(vl_semestres):
        formatted_value = f"{txt:,.2f} €".replace(",", " ").replace(".", ",")
        ax.annotate(
            formatted_value,
            (dates_semestres[i], vl_semestres[i]),
            textcoords="offset points",
            xytext=(0, 12),
            ha='center',
            fontsize=10,
            fontweight='bold',
            color='white',
            bbox=dict(
                boxstyle="round,pad=0.4",
                fc=color_main,
                ec='white',
                alpha=0.9,
                linewidth=1
            ),
            zorder=10  # Au premier plan
        )
    
    # Ajustement des marges
    plt.tight_layout(pad=1.5)  # Réduit
    
    # Sauvegarder le graphique en PNG haute définition
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
    img_buffer.seek(0)
    
    # Créer l'image avec taille optimale pour le format paysage
    img = Image(img_buffer, width=21*cm, height=10*cm)  # Réduit légèrement
    
    # Ajouter le graphique aux éléments de la deuxième page
    page2_elements.append(img)
    
    # Ajouter un pied de page pour la deuxième page
    page2_elements.append(Spacer(1, 8))  # Réduit
    footer_text2 = Paragraph(f"Document confidentiel - {nom_fonds} - Page 2/2", footer_style)
    page2_elements.append(footer_text2)
    
    # Ajouter tous les éléments de la deuxième page avec KeepTogether
    elements.append(KeepTogether(page2_elements))
    
    # Construire le PDF
    doc.build(elements)
    
    # Réinitialiser le buffer
    buffer.seek(0)
    plt.close(fig)
    
    return buffer

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
    row = {"Date": date.strftime('%d/%m/%Y')}

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

    if i > 0:
        anr_courant += total_var_actifs + total_impacts

    vl = anr_courant / nombre_parts
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
    ticker.FuncFormatter(lambda x, _: f"{x:,.2f} €".replace(",", " ").replace(".", ","))
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
    worksheet = workbook.add_worksheet('Projection')
    
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
        if "€" in col or any(s in col for s in ["Impact", "Actif", "VL"]):
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
col1, col2 = st.columns(2)

with col1:
    st.download_button(
        label="📥 Exporter la projection avec graphique Excel",
        data=buffer,
        file_name="projection_vl.xlsx",
        mime="application/vnd.ms-excel"
    )

with col2:
    # === EXPORT PDF ===
    if st.download_button(
        label="📊 Exporter la projection en PDF",
        data=export_pdf(projection, dates_semestres, vl_semestres, nom_fonds, couleur_bleue),
        file_name="projection_vl.pdf",
        mime="application/pdf"
    ):
        st.success("Le PDF a été généré avec succès !")

# === EXPORT PARAMÈTRES JSON ===
if st.sidebar.button("📤 Exporter les paramètres JSON"):
    export_data = {
        "nom_fonds": nom_fonds,
        "date_vl_connue": date_vl_connue_str,
        "date_fin_fonds": date_fin_fonds_str,
        "anr_derniere_vl": anr_derniere_vl,
        "nombre_parts": nombre_parts,
        "impacts": impacts,
        "actifs": actifs
    }
    json_export = json.dumps(export_data, indent=2).encode('utf-8')
    st.sidebar.download_button("Télécharger paramètres JSON", json_export, file_name="parametres_vl.json")

# === BOUTON RÉINITIALISATION ===
if st.sidebar.button("♻️ Réinitialiser les paramètres"):
    st.session_state.params = default_params.copy()
    st.rerun()