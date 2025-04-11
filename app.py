import streamlit as st
import pandas as pd
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io
import matplotlib.dates as mdates

# Configuration CSS personnalisée
st.markdown("""
    <style>
        .main {
            background-color: #F9FAFB;
        }
        .stSidebar {
            background-color: #FFFFFF;
            border-right: 1px solid #E5E7EB;
        }
        .stButton > button {
            background-color: #0000DC;
            color: white;
            border-radius: 8px;
            padding: 0.5rem 1rem;
            transition: all 0.2s;
        }
        .stButton > button:hover {
            background-color: #0000AA;
            transform: translateY(-1px);
        }
        .stDataFrame {
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        h1, h2, h3 {
            color: #0000DC !important;
        }
        .metric-value {
            font-size: 1.4rem !important;
            color: #0000DC !important;
            font-weight: bold !important;
        }
        .stDownloadButton > button {
            width: 100%;
            justify-content: center;
        }
    </style>
""", unsafe_allow_html=True)

# Fonctions utilitaires
def format_fr_euro(valeur):
    return f"{valeur:,.2f} €".replace(",", " ").replace(".", ",")

def champ_numerique(label, valeur):
    champ = st.sidebar.text_input(label, value=format_fr_euro(valeur))
    champ = champ.replace(" ", "").replace(",", ".").replace("€", "")
    return float(champ) if champ else 0.0

# Fonction d'export PDF (simplifiée pour l'exemple)
def export_pdf(projection, dates_semestres, vl_semestres, nom_fonds, couleur_bleue):
    buffer = io.BytesIO()
    # ... (implémentation PDF complète comme précédemment)
    return buffer

# Paramètres initiaux
default_params = {
    "nom_fonds": "Fonds d'Investissement XYZ",
    "date_vl_connue": "31/12/2024",
    "date_fin_fonds": "31/12/2028",
    "anr_derniere_vl": 10_000_000.0,
    "nombre_parts": 10_000.0,
    "impacts": [("Frais corporate", -50_000.0), ("Honoraires NIV", -30_000.0)],
    "actifs": []
}

# Gestion des paramètres
if 'params' not in st.session_state:
    st.session_state.params = default_params.copy()

# Interface utilisateur
st.title("🛬 Atterrissage VL - Analyse de Valeur Liquidative")

# Sidebar - Paramètres
with st.sidebar:
    st.header("⚙️ Paramètres du fonds")
    params = st.session_state.params
    
    # Import/Export
    params_json = st.file_uploader("📤 Importer configuration", type="json")
    if params_json:
        st.session_state.params.update(json.load(params_json))
    
    # Saisie des paramètres
    nom_fonds = st.text_input("🏦 Nom du fonds", params['nom_fonds'])
    date_vl_connue_str = st.text_input("📅 Date VL connue (jj/mm/aaaa)", params['date_vl_connue'])
    date_fin_fonds_str = st.text_input("📅 Date fin de fonds (jj/mm/aaaa)", params['date_fin_fonds'])
    anr_derniere_vl = champ_numerique("💰 ANR dernière VL connue", params['anr_derniere_vl'])
    nombre_parts = champ_numerique("📊 Nombre de parts", params['nombre_parts'])

    # Impacts
    st.header("🎯 Impacts semestriels")
    impacts = []
    nb_impacts = st.number_input("Nombre d'impacts", min_value=0, value=len(params['impacts']), key="nb_impacts")
    for i in range(nb_impacts):
        col1, col2 = st.columns(2)
        with col1:
            libelle = st.text_input(f"Libellé {i+1}", value=params['impacts'][i][0] if i < len(params['impacts']) else f"Impact {i+1}")
        with col2:
            montant = champ_numerique(f"Montant {i+1}", params['impacts'][i][1] if i < len(params['impacts']) else 0.0)
        impacts.append((libelle, montant))

    # Actifs
    st.header("🏢 Actifs du fonds")
    actifs = []
    nb_actifs = st.number_input("Nombre d'actifs", min_value=1, value=max(1, len(params['actifs'])), key="nb_actifs")
    for i in range(nb_actifs):
        with st.expander(f"Actif {i+1}", expanded=(i < 2)):
            a = params['actifs'][i] if i < len(params['actifs']) else {}
            nom_actif = st.text_input(f"Nom", value=a.get('nom', f"Actif {i+1}"))
            col1, col2 = st.columns(2)
            with col1:
                pct_detention = st.slider("% Détention", 0.0, 100.0, value=a.get('pct_detention', 100.0))
            with col2:
                valeur_actuelle = champ_numerique("Valeur actuelle", a.get('valeur_actuelle', 1_000_000.0))
                valeur_projetee = champ_numerique("Valeur projetée", a.get('valeur_projetee', 1_050_000.0))
            actifs.append({
                "nom": nom_actif,
                "pct_detention": pct_detention / 100,
                "valeur_actuelle": valeur_actuelle,
                "valeur_projetee": valeur_projetee,
                "variation": (valeur_projetee - valeur_actuelle) * (pct_detention / 100)
            })

# Calcul des projections
try:
    date_vl_connue = datetime.strptime(date_vl_connue_str, "%d/%m/%Y")
    date_fin_fonds = datetime.strptime(date_fin_fonds_str, "%d/%m/%Y")
except ValueError:
    st.error("Format de date invalide - utiliser jj/mm/aaaa")
    st.stop()

dates_semestres = [date_vl_connue]
current_date = date_vl_connue
while current_date < date_fin_fonds:
    current_date = current_date + pd.DateOffset(months=6)
    dates_semestres.append(current_date)

vl_semestres = []
anr_courant = anr_derniere_vl
projection_rows = []

for i, date in enumerate(dates_semestres):
    row = {"Date": date.strftime('%d/%m/%Y')}
    
    # Calcul des variations
    total_var_actifs = sum(a['variation'] for a in actifs) if i == 1 else 0
    total_impacts = sum(montant for _, montant in impacts)
    
    if i > 0:
        anr_courant += total_var_actifs + total_impacts
    
    vl = anr_courant / nombre_parts
    vl_semestres.append(vl)
    
    # Construction de la ligne
    row.update({
        "ANR": format_fr_euro(anr_courant),
        "VL prévisionnelle": format_fr_euro(vl),
        **{f"Impact - {libelle}": format_fr_euro(montant) for libelle, montant in impacts},
        **{f"Actif - {a['nom']}": format_fr_euro(a['variation']) for a in actifs}
    })
    
    projection_rows.append(row)

projection = pd.DataFrame(projection_rows)

# Affichage des KPI
st.header("📊 Indicateurs clés")
kpi1, kpi2, kpi3 = st.columns(3)
with kpi1:
    st.metric("VL Actuelle", format_fr_euro(vl_semestres[0]))
with kpi2:
    evolution = ((vl_semestres[-1] - vl_semestres[0]) / vl_semestres[0]) * 100
    st.metric("VL Finale", format_fr_euro(vl_semestres[-1]), f"{evolution:.1f}%")
with kpi3:
    st.metric("Durée projetée", f"{len(dates_semestres)-1} semestres")

# Tableau de projection
st.header("📋 Détail des projections")
def color_negative(val):
    return 'color: #DC0000' if '-' in val else 'color: #0000DC'

styled_projection = projection.style\
    .applymap(color_negative)\
    .set_properties(**{'background-color': '#FFFFFF', 'border': '1px solid #F3F4F6'})\
    .set_table_styles([{
        'selector': 'th',
        'props': [
            ('background-color', '#0000DC'),
            ('color', 'white'),
            ('font-weight', '600'),
            ('text-transform', 'uppercase')
        ]
    }])

st.write(styled_projection.to_html(), unsafe_allow_html=True)

# Visualisation graphique
st.header("📈 Évolution de la VL")
couleur_bleue = "#0000DC"

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(
    dates_semestres, 
    vl_semestres, 
    marker='o',
    markersize=8,
    linewidth=2.5,
    color=couleur_bleue,
    markeredgecolor='white',
    markeredgewidth=1.5
)

# Améliorations graphiques
ax.fill_between(dates_semestres, vl_semestres, color=couleur_bleue, alpha=0.08)
ax.grid(True, axis='y', linestyle='--', alpha=0.4)
ax.spines[['top', 'right']].set_visible(False)
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%y'))
ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:,.0f} €"))
plt.xticks(rotation=45, ha='right')

# Annotations
for date, vl in zip(dates_semestres, vl_semestres):
    plt.annotate(
        format_fr_euro(vl),
        (date, vl),
        xytext=(0, 15),
        textcoords='offset points',
        ha='center',
        va='center',
        fontsize=9,
        bbox=dict(
            boxstyle='round,pad=0.3',
            facecolor='white',
            edgecolor=couleur_bleue,
            linewidth=0.8
        )
    )

st.pyplot(fig)

# Export des données
st.header("📤 Exporter les résultats")
col1, col2 = st.columns(2)

with col1:
    # Export Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer) as writer:
        projection.to_excel(writer, index=False)
    st.download_button(
        label="💾 Télécharger Excel",
        data=buffer,
        file_name="projection_vl.xlsx",
        mime="application/vnd.ms-excel"
    )

with col2:
    # Export PDF
    pdf_buffer = export_pdf(projection, dates_semestres, vl_semestres, nom_fonds, couleur_bleue)
    st.download_button(
        label="🖨️ Générer PDF",
        data=pdf_buffer,
        file_name="rapport_vl.pdf",
        mime="application/pdf"
    )

# Gestion de la configuration
with st.sidebar:
    st.divider()
    if st.button("🔄 Réinitialiser les paramètres"):
        st.session_state.params = default_params.copy()
        st.rerun()
    
    # Export configuration
    export_config = json.dumps({
        "nom_fonds": nom_fonds,
        "date_vl_connue": date_vl_connue_str,
        "date_fin_fonds": date_fin_fonds_str,
        "anr_derniere_vl": anr_derniere_vl,
        "nombre_parts": nombre_parts,
        "impacts": impacts,
        "actifs": actifs
    }, indent=2)
    
    st.download_button(
        label="📁 Sauvegarder configuration",
        data=export_config,
        file_name="config_vl.json"
    )
