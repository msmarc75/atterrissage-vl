# PDF réel avec graphique
        def generate_pdf():
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []

            title = Paragraph(f"Atterrissage VL - {nom_fonds}", styles['Title'])
            elements.append(title)
            elements.append(Spacer(1, 12))

            # Construire le tableau PDF
            data = [projection.columns.tolist()] + projection.values.tolist()
            formatted_data = []
            for row in data:
                formatted_data.append([str(cell) for cell in row])

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

            # Générer le graphique matplotlib
            fig, ax = plt.subplots(figsize=(6, 3))
            couleur_bleue = "#0000DC"
            ax.plot(dates_semestres, vl_semestres, marker='o', linewidth=2.5, color=couleur_bleue)
            for i, txt in enumerate(vl_semestres):
                ax.annotate(f"{txt:,.2f} €".replace(",", " ").replace(".", ","),
                            (dates_semestres[i], vl_semestres[i]),
                            textcoords="offset points", xytext=(0, 10), ha='center', fontsize=8,
                            bbox=dict(boxstyle="round,pad=0.3", fc=couleur_bleue, ec=couleur_bleue, alpha=0.9), color='white')
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

            img = Image(img_buffer, width=400, height=200)
            elements.append(img)

            doc.build(elements)
            buffer.seek(0)
            return buffer

        st.download_button(
            label="📄 Exporter en PDF",
            data=generate_pdf(),
            file_name="projection_vl.pdf",
            mime="application/pdf"
        )

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
        json_export = json.dumps(export_data, indent=2).encode('utf-8')
        st.download_button("📦 Exporter les paramètres JSON", data=json_export, file_name="parametres_vl.json")

except Exception as e:
    st.warning(f"Veuillez vérifier les champs de date ou d'entrée : {e}")
