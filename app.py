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
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

st.set_page_config(
    page_title="Atterrissage VL",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
        .block-container { padding: 2rem; }
        .stDownloadButton button { background-color: #0000DC; color: white; font-weight: bold; }
        .stButton button { background-color: #f0f2f6; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Atterrissage VL Professionnel")

onglets = st.tabs(["⚙️ Paramètres", "📈 Projection & Graphique", "📤 Export"])

with onglets[0]:
    st.header("Paramètres du fonds")
    with st.expander("🗂 Informations générales", expanded=True):
        nom_fonds = st.text_input("Nom du fonds", "Nom du Fonds")
        date_vl_connue_str = st.text_input("Date dernière VL connue (jj/mm/aaaa)", "31/12/2024")
        date_fin_fonds_str = st.text_input("Date fin de fonds (jj/mm/aaaa)", "31/12/2028")
        anr_derniere_vl = st.number_input("ANR dernière VL connue (€)", value=10_000_000.0, step=10000.0)
        nombre_parts = st.number_input("Nombre de parts", value=10_000.0, step=100.0)

    with st.expander("📉 Impacts semestriels personnalisés"):
        nb_impacts = st.number_input("Nombre d'impacts", min_value=0, value=2, step=1)
        impacts = []
        for i in range(nb_impacts):
            st.markdown(f"**Impact {i+1}**")
            libelle = st.text_input(f"Libellé impact {i+1}", f"Impact {i+1}")
            montant = st.number_input(f"Montant semestriel impact {i+1} (€)", value=-50000.0)
            impacts.append((libelle, montant))

    with st.expander("🏢 Actifs du portefeuille"):
        nb_actifs = st.number_input("Nombre d'actifs", min_value=0, value=1, step=1)
        actifs = []
        for i in range(nb_actifs):
            st.markdown(f"**Actif {i+1}**")
            nom = st.text_input(f"Nom actif {i+1}", f"Actif {i+1}")
            pct = st.slider(f"% Détention actif {i+1}", 0.0, 100.0, 100.0, step=1.0)
            val_act = st.number_input(f"Valeur actuelle actif {i+1} (€)", value=1_000_000.0)
            val_proj = st.number_input(f"Valeur projetée S+1 actif {i+1} (€)", value=1_050_000.0)
            actifs.append({
                "nom": nom,
                "pct_detention": pct / 100,
                "valeur_actuelle": val_act,
                "valeur_projetee": val_proj,
                "variation": (val_proj - val_act) * (pct / 100)
            })

try:
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

    vl_semestres = []
    anr_courant = anr_derniere_vl
    projection_rows = []

    for i, date in enumerate(dates_semestres):
        row = {"Date": date.strftime('%d/%m/%Y')}
        total_var_actifs = sum(a['variation'] if i == 1 else 0 for a in actifs)
        for a in actifs:
            var = a['variation'] if i == 1 else 0
            row[f"Actif - {a['nom']}"] = var

        total_impacts = sum(m for _, m in impacts)
        for libelle, montant in impacts:
            row[f"Impact - {libelle}"] = montant

        if i > 0:
            anr_courant += total_var_actifs + total_impacts
        vl = anr_courant / nombre_parts
        vl_semestres.append(vl)
        row["VL prévisionnelle (€)"] = vl
        projection_rows.append(row)

    projection = pd.DataFrame(projection_rows)
    
# Convertir toutes les colonnes sauf 'Date' en float (si possible)
for col in projection.columns:
    if col != "Date":
        projection[col] = pd.to_numeric(projection[col], errors="coerce")

    with onglets[1]:
        st.header("Projection de la VL")
        st.dataframe(projection.style.format("{:.2f}"))

        fig, ax = plt.subplots(figsize=(10, 5))
        couleur_bleue = "#0000DC"
        ax.plot(dates_semestres, vl_semestres, marker='o', linewidth=2.5, color=couleur_bleue)
        for i, txt in enumerate(vl_semestres):
            ax.annotate(f"{txt:,.2f} €".replace(",", " ").replace(".", ","),
                        (dates_semestres[i], vl_semestres[i]),
                        textcoords="offset points", xytext=(0, 10), ha='center', fontsize=9,
                        bbox=dict(boxstyle="round,pad=0.3", fc=couleur_bleue, ec=couleur_bleue, alpha=0.9), color='white')
        ax.set_title(f"Atterrissage VL - {nom_fonds}", fontsize=16, fontweight='bold', color=couleur_bleue)
        ax.set_ylabel("VL (€)", fontsize=12, color=couleur_bleue)
        ax.set_xticks(dates_semestres)
        ax.set_xticklabels([d.strftime('%b-%y') for d in dates_semestres], rotation=45, ha='right')
        ax.tick_params(axis='y', labelcolor=couleur_bleue)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:,.2f} €".replace(",", " ").replace(".", ",")))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        st.pyplot(fig)

    with onglets[2]:
        st.header("Export des résultats")

        # Excel
        buffer = io.BytesIO()
        projection.to_excel(buffer, index=False)
        buffer.seek(0)
        st.download_button("📥 Télécharger Excel", buffer, file_name="projection_vl.xlsx")

        # PDF
        def generate_pdf():
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []

            title = Paragraph(f"Atterrissage VL - {nom_fonds}", styles['Title'])
            elements.append(title)
            elements.append(Spacer(1, 12))

            data = [projection.columns.tolist()] + projection.values.tolist()
            formatted_data = [[str(cell) for cell in row] for row in data]
            table = Table(formatted_data, hAlign='LEFT')
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#0000DC")),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ]))

            elements.append(table)
            elements.append(Spacer(1, 24))

            fig, ax = plt.subplots(figsize=(6, 3))
            ax.plot(dates_semestres, vl_semestres, marker='o', linewidth=2.5, color="#0000DC")
            for i, txt in enumerate(vl_semestres):
                ax.annotate(f"{txt:,.2f} €".replace(",", " ").replace(".", ","),
                            (dates_semestres[i], vl_semestres[i]),
                            textcoords="offset points", xytext=(0, 10), ha='center', fontsize=8,
                            bbox=dict(boxstyle="round,pad=0.3", fc="#0000DC", ec="#0000DC", alpha=0.9), color='white')
            ax.set_title("Projection de la VL", fontsize=12)
            ax.set_ylabel("VL (€)")
            ax.set_xticks(dates_semestres)
            ax.set_xticklabels([d.strftime('%b-%y') for d in dates_semestres], rotation=45, ha='right')
            ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:,.2f} €".replace(",", " ").replace(".", ",")))
            fig.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.spines['right'].set_visible(False)
            ax.spines['top'].set_visible(False)
            img_buffer = io.BytesIO()
            plt.tight_layout()
            plt.savefig(img_buffer, format='png')
            plt.close(fig)
            img_buffer.seek(0)

            elements.append(Image(img_buffer, width=400, height=200))
            doc.build(elements)
            buffer.seek(0)
            return buffer

        st.download_button("📄 Exporter en PDF", data=generate_pdf(), file_name="projection_vl.pdf", mime="application/pdf")

        # JSON paramètres
        export_data = {
            "nom_fonds": nom_fonds,
            "date_vl_connue": date_vl_connue_str,
            "date_fin_fonds": date_fin_fonds_str,
            "anr_derniere_vl": anr_derniere_vl,
            "nombre_parts": nombre_parts,
            "impacts": impacts,
            "actifs": actifs
        }
        st.download_button("📦 Exporter les paramètres JSON", data=json.dumps(export_data, indent=2).encode('utf-8'), file_name="parametres_vl.json")

except Exception as e:
    st.warning(f"Veuillez vérifier les champs de date ou d'entrée : {e}")
