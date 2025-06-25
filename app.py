import os
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO
import base64
import plotly.io as pio

st.set_page_config(layout="wide")
language = st.radio("Select Language / Sprache wählen", ["English", "Deutsch"])

translations = {
    "English": {
        "title": "📊 Innovation Benchmark Dashboard",
        "upload_prompt": "Upload your Benchmark Excel file",
        "select_company": "Select company to benchmark",
        "underperforming": "⚠️ Underperforming Metrics",
        "download_link": "📥 Download Full PDF Report",
        "industry_size": "Industry: {industry} | Size: {size}",
        "summary_title": "Summary of Underperforming Metrics",
        "category": "Category",
        "metric": "Metric",
        "top3": "Top 3 Avg",
        "top10": "Top 10 Avg",
        "overall": "Overall Avg"

    },
    "Deutsch": {
        "title": "📊 Innovations-Benchmark-Dashboard",
        "upload_prompt": "Laden Sie Ihre Benchmark-Excel-Datei hoch",
        "select_company": "Wählen Sie ein Unternehmen zum Benchmarking",
        "underperforming": "⚠️ Unterdurchschnittliche Kennzahlen",
        "download_link": "📥 Vollständigen PDF-Bericht herunterladen",
        "industry_size": "Branche: {industry} | Größe: {size}",
        "summary_title": "Zusammenfassung der unterdurchschnittlichen Kennzahlen",
        "category": "Kategorie",
        "metric": "Kennzahl",
        "top3": "Top 3 Durchschnitt",
        "top10": "Top 10 Durchschnitt",
        "overall": "Gesamtdurchschnitt"

    }
}

st.title(translations[language]["title"])

metric_translation = {
    "Top Management Innovation": "Top-Management-Innovation",
    "% Innovation Working Time": "Arbeitszeit für Innovation / % der Arbeitszeit",
    "Participation in Innovation Projects": "Mitwirkung bei Innovationsprojekten",
    "Strategy and Innovation": "Strategie und Innovation",
    "Innovation Strategy Content": "Inhalte der Innovationsstrategie",
    "Sustainability": "Nachhaltigkeit",
    "Digital Transformation": "Digitaler Wandel",

    "Climate Innovation": "Klima Innovation",
    "Innovation Culture": "Innovative Ausrichtung des Unternehmens",
    "Training": "Weiterbildung",
    "Free Space": "Freiräume",
    "Internal Venture Capital": 'Internes "Venture-Capital"',
    "Employee Participation": "Mitarbeiterbeteiligung",
    "Digitization of Internal Organization and Communication": "Digitalisierung der internen Organisation und Kommunikation",
    "Idea Management": "Ideenmanagement",
    "Creative Contributions from Employees": "Kreative Beiträge von Mitarbeitern",

    "Organization & Process Innovation": "Organisations- und Prozessinnovation",
    "Ongoing Market, Technology and Competition Monitoring": "Laufende Beobachtung von Markt, Technologie und Wettbewerb (Monitoring)",
    "Design of the Innovation Process": "Ausgestaltung des Innovationsprozesses",
    "Flexibility and Agility": "Flexibilität und Agilität",
    "Tools and Methods": "Instrumente und Methoden",
    "Digital Networking": "Digitale Vernetzung",
    "Use of Digital Technologies": "Einsatz digitaler Technologien",
    "Organization of Digital Transformation": "Organisation des digitalen Wandels",

    "External Orientation & Open Innovation": "Außenorientierung & Open Innovation",
    "Role of Marketing/Sales in the Innovation Process": "Stellung von Marketing/Vertrieb im Innovationsprozess",
    "Involvement in Individual Phases of the Innovation Process": "Mitwirkung in einzelnen Phasen des Innovationsprozesses",
    "Open Innovation Activites / Short Term": "Open Innovation Aktivitäten / Kurzfristig",
    "Open Innovation Activities / Long Term": "Open Innovation Aktivitäten / Langfristig",
    "Digital Sales / External Communication": "Digitaler Vertrieb/ externe Kommunikation"
}



def load_and_prepare(sheet, file, metric_cols, target_company):
    metric_cols = [col.strip() for col in metric_cols]
    df_raw = pd.read_excel(file, sheet_name=sheet, header=1)
    df_raw.columns = df_raw.columns.str.strip()  # clean up column names
    df = df_raw[["Company Name", "Company Industry", "Company Size"] + metric_cols].dropna(subset=["Company Name"])

    target_row = df[df["Company Name"] == target_company]
    if target_row.empty:
        return None, None

    target = target_row.iloc[0]
    industry = target["Company Industry"]
    size = target["Company Size"]

    peers_same_industry_size = df[(df["Company Industry"] == industry) & (df["Company Size"] == size)]
    peers_same_size = df[df["Company Size"] == size]

    benchmark = {
        "Metric": [],
        "Target": [],
        "Top 3 Avg": [],
        "Top 10 Avg": [],
        "Overall Avg": []
    }

    for col in metric_cols:
        target_val = target[col]
        ranked = peers_same_industry_size[[col]].dropna().sort_values(by=col, ascending=False)
        benchmark["Metric"].append(col)
        benchmark["Target"].append(target_val)
        benchmark["Top 3 Avg"].append(ranked.head(3)[col].mean())
        benchmark["Top 10 Avg"].append(ranked.head(10)[col].mean())
        benchmark["Overall Avg"].append(peers_same_size[col].mean())

    return pd.DataFrame(benchmark), {"industry": industry, "size": size}

def wrap_labels(labels, max_words=2):
    return ['<br>'.join(label.split(' ', max_words)) for label in labels]

def create_radar_chart(df, title, company, meta):
    if df is None or df.empty:
        return None, []
    
    if language == "Deutsch":
        df["Metric"] = df["Metric"].apply(lambda x: metric_translation.get(x, x))

    categories = wrap_labels(df["Metric"].tolist()) + [wrap_labels([df["Metric"].iloc[0]])[0]]
    target = df["Target"].tolist() + [df["Target"].iloc[0]]
    top3 = df["Top 3 Avg"].tolist() + [df["Top 3 Avg"].iloc[0]]
    top10 = df["Top 10 Avg"].tolist() + [df["Top 10 Avg"].iloc[0]]
    overall = df["Overall Avg"].tolist() + [df["Overall Avg"].iloc[0]]

    under_flags = df["Target"] < (df["Top 10 Avg"] * 0.75)
    under_flags_list = under_flags.tolist()
    under_vals = [val if flag else None for val, flag in zip(target, under_flags_list)]
    under_metrics = df["Metric"][under_flags].tolist()

    fig = go.Figure()

    fig.add_trace(go.Scatterpolar(r=overall, theta=categories, name="Overall Avg",
                                  line=dict(color="#C0C0C0", width=2, dash="dot"),
                                  marker=dict(size=6)))
    fig.add_trace(go.Scatterpolar(r=top10, theta=categories, name="Top 10 Avg",
                                  line=dict(color="#7F8C8D", width=2, dash="dot"),
                                  marker=dict(size=6)))
    fig.add_trace(go.Scatterpolar(r=top3, theta=categories, name="Top 3 Avg",
                                  line=dict(color="black", width=3),
                                  marker=dict(size=7)))
    fig.add_trace(go.Scatterpolar(r=target, theta=categories, name=company,
                                  line=dict(color="#D4AC0D", width=3),
                                  marker=dict(size=7)))
    fig.add_trace(go.Scatterpolar(r=under_vals + [under_vals[0]], theta=categories,
                                  mode='markers', name="Underperforming",
                                  marker=dict(size=11, color="red", symbol="circle-open-dot"),
                                  showlegend=True))

    fig.update_layout(
        template='simple_white',
        font=dict(family="Arial", size=13, color="black"),
        #title=f"<b>{title}</b>",#
        polar=dict(
            bgcolor="white",
            radialaxis=dict(visible=True, range=[0, 1.2], showticklabels=False, ticks=''),
            angularaxis=dict(tickfont=dict(size=12), rotation=90, direction="clockwise")
        ),
        showlegend=True,
        height=450,
        margin=dict(t=80, b=80, l=120, r=120)
    )

    return fig, under_metrics

from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
from PIL import Image
import streamlit as st

def get_pdf_download_link(figs, company_name, industry, size, language):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()

    # Add header
    story.append(Paragraph(f"<b>Company:</b> {company_name}", styles["Title"]))
    story.append(Paragraph(f"<b>Industry:</b> {industry} | <b>Size:</b> {size}", styles["Normal"]))
    story.append(Spacer(1, 12))

    for fig_data in figs:
        fig = fig_data["fig"]
        section_title = fig_data["title"]
        underperforming = fig_data["underperforming"]
        metrics_df = fig_data["metrics_df"]

        # Add section header
        story.append(Paragraph(f"<b>{section_title}</b>", styles["Heading2"]))
        story.append(Spacer(1, 12))

        # Convert Plotly figure to PNG bytes
        try:
            img_bytes = fig.to_image(format="png", width=800, height=600, engine="kaleido")
            img_buffer = BytesIO(img_bytes)
        except Exception as e:
            st.error("⚠️ Failed to render chart image. Please check Kaleido installation.")
            continue  # Skip this chart if it fails

        # Insert image into PDF
        rl_img = RLImage(img_buffer, width=400, height=300)
        story.append(rl_img)
        story.append(Spacer(1, 12))

        # Underperforming Metrics
        if underperforming:
            warning_text = f"⚠️ Underperforming Metrics: {', '.join(underperforming)}"
            story.append(Paragraph(warning_text, styles["Normal"]))
            story.append(Spacer(1, 12))

        # Metrics Table
        table_data = [[
            "Metric",
            company_name,
            "Overall Avg",
            "Top 10 Avg",
            "Top 3 Avg"
        ]]

        for _, row in metrics_df.iterrows():
            table_data.append([
                row["Metric"],
                f"{row['Target']:.2f}",
                f"{row['Overall Avg']:.2f}",
                f"{row['Top 10 Avg']:.2f}",
                f"{row['Top 3 Avg']:.2f}",
            ])

        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ]))
        story.append(table)
        story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)

    # Show download button
    st.download_button(
        label="📄 Download Full Dashboard as PDF",
        data=buffer,
        file_name=f"{company_name}_dashboard.pdf",
        mime="application/pdf"
    )


uploaded_file = st.file_uploader(translations[language]["upload_prompt"], type=["xlsx"])

if uploaded_file:
    company_df = pd.read_excel(uploaded_file, sheet_name="Overview", header=1)
    company_df.columns = company_df.columns.str.strip()
    company_names = company_df["Company Name"].dropna().unique()
    selected_company = st.selectbox(translations[language]["select_company"], sorted(company_names))

    chart_configs = [
        {
            "title": "Innovation Dimensions Overview",
            "sheet": "Overview",
            "metrics": [
                "Top Management Innovation", 
                "Climate Innovation", 
                "Organization & Process Innovation", 
                "External Orientation & Open Innovation"
            ]
        },
        {
            "title": "Top Management Innovation",
            "sheet": "Top Management Innovation",
            "metrics": [
                "% Innovation Working Time",
                "Participation in Innovation Projects ",
                "Strategy and Innovation",
                "Innovation Strategy Content",
                "Sustainability",
                "Digital Transformation",
                "Top Management Innovation"
            ]
        },
        {
            "title": "Climate Innovation",
            "sheet": "Climate Innovation",
            "metrics": [
                "Innovation Culture",
                "Training",
                "Free Space",
                "Internal Venture Capital",
                "Employee Participation",
                "Digitization of Internal Organization and Communication",
                "Idea Management",
                "Creative Contributions from Employees",
                "Climate Innovation"
            ]
        },
        {
            "title": "Organization & Process Innovation",
            "sheet": "Org. & Process Innovation",
            "metrics": [
                "Ongoing Market, Technology and Competition Monitoring",
                "Design of the Innovation Process",
                "Flexibility and Agility",
                "Tools and Methods",
                "Digital Networking",
                "Use of Digital Technologies",
                "Organization of Digital Transformation",
                "Organization & Process Innovation"
            ]
        },
        {
            "title": "External Orientation & Open Innovation",
            "sheet": "EO&OI",
            "metrics": [
                "Role of Marketing/Sales in the Innovation Process",
                "Involvement in Individual Phases of the Innovation Process",
                "Open Innovation Activites / Short Term",
                "Open Innovation Activities / Long Term",
                "Digital Networking",
                "Digital Sales / External Communication",
                "External Orientation & Open Innovation"
            ]
        }
    ]
    figs = []
    cols = st.columns(2)  # Create two columns
    col_index = 0  # Track which column to use
    section_titles = []
    underperforming_list = []

    for i, config in enumerate(chart_configs):
        section_title = config["title"]
        if language == "Deutsch":
            section_title = metric_translation.get(section_title, section_title)

        df, meta = load_and_prepare(config["sheet"], uploaded_file, config["metrics"], selected_company)
        fig, underperforming = create_radar_chart(df, config["title"], selected_company, meta)
        if fig:
            is_last = (i == len(chart_configs) - 1)
            if col_index == 1 and is_last:
                # Last chart, alone on final row → center it
                center = st.columns([1, 2, 1])
                with center[1]:
                    st.subheader(section_title)
                    st.markdown(f"**{selected_company}**")
                    st.markdown(f"_{translations[language]['industry_size'].format(industry=meta['industry'], size=meta['size'])}_")
                    st.plotly_chart(fig, use_container_width=False)
                    if underperforming:
                        st.warning(f"{translations[language]['underperforming']}: " + ", ".join(underperforming))
                    figs.append({
                        "fig": fig,
                        "title": section_title,
                        "underperforming": underperforming,
                        "metrics_df": df
                    })

                    section_titles.append(section_title)
                    underperforming_list.append(underperforming)
            else:
                with cols[col_index]:
                    st.subheader(section_title)
                    st.markdown(f"**{selected_company}**")
                    st.markdown(f"_{translations[language]['industry_size'].format(industry=meta['industry'], size=meta['size'])}_")
                    st.plotly_chart(fig, use_container_width=False)
                    if underperforming:
                        st.warning(f"{translations[language]['underperforming']}: " + ", ".join(underperforming))
                    figs.append({
                        "fig": fig,
                        "title": section_title,
                        "underperforming": underperforming,
                        "metrics_df": df
                    })
                    section_titles.append(section_title)
                    underperforming_list.append(underperforming)

                col_index += 1
                if col_index == 2:
                    cols = st.columns(2)
                    col_index = 0
        else:
            st.warning(f"No data available for {config['title']}")
    if figs:
        get_pdf_download_link(
            figs=figs,
            company_name=selected_company,
            industry=meta["industry"],
            size=meta["size"],
            language=language
        )

