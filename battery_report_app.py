
import io
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, Image as RLImage)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# -------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------

def load_data(uploaded_file: io.BytesIO) -> pd.DataFrame:
    """Read the first sheet of the uploaded Excel file (header row = 2)."""
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = xls.parse(sheet, header=1)
    return df


def compute_indicators(df: pd.DataFrame):
    """Return dataframe with numeric columns and dict of KPI."""
    df = df.copy()
    df['SoC_end_discharge'] = pd.to_numeric(
        df['SoC at End of Discharge [%] (Ah discharged - Ah charged in Discharge phase) / Cnom'],
        errors='coerce')
    df['SoC_end_charge'] = pd.to_numeric(df['SoC at End of Charge [%]'], errors='coerce')
    df['Tmax'] = pd.to_numeric(df['Max. Temperature At Cycle (℃)'], errors='coerce')
    df['Tmin'] = pd.to_numeric(df['Min. Temperature At Cycle (℃)'], errors='coerce')
    df['Ah_dis'] = pd.to_numeric(df['Ah Discharged'], errors='coerce')
    df['Ah_chg'] = pd.to_numeric(df['Ah Charged In Charge Phase'], errors='coerce')

    stats = {
        'total_cycles': len(df),
        'over_discharge': int((df['SoC_end_discharge'] < 20).sum()),
        'over_charge': int((df['SoC_end_charge'] > 105).sum()),
        'high_temp': int((df['Tmax'] > 45).sum()),
        'low_temp': int((df['Tmin'] < 0).sum()),
        'full_charges': int((df['SoC_end_charge'] >= 99).sum()),
        'efficiency': float(df['Ah_chg'].sum() / df['Ah_dis'].sum()) if df['Ah_dis'].sum() else float('nan'),
    }
    return df, stats


def create_charts(df: pd.DataFrame, out_dir: Path) -> dict:
    """Generate and save charts; return dict of file paths."""
    charts = {}
    fig, ax = plt.subplots(figsize=(8,3))
    ax.bar(df['Cycle Count']-0.2, df['Ah_dis'], width=0.4, label='Ah scaricati', color='red')
    ax.bar(df['Cycle Count']+0.2, df['Ah_chg'], width=0.4, label='Ah caricati', color='green')
    full = df['SoC_end_charge'] >= 99
    y_marker = max(df['Ah_dis'].max(), df['Ah_chg'].max()) * 1.05
    for x, complete in zip(df['Cycle Count'], full):
        ax.text(x, y_marker, '✓' if complete else '✗', color='green' if complete else 'red',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    ax.set_xlabel('Cycle')
    ax.set_ylabel('Ah')
    ax.set_title('Ah caricati/scaricati per ciclo')
    ax.legend()
    plt.tight_layout()
    file1 = out_dir / 'ah_cycle.png'
    fig.savefig(file1, dpi=300)
    plt.close(fig)
    charts['ah_cycle'] = file1

    fig, ax = plt.subplots(figsize=(8,3))
    ax.bar(df['Cycle Count'], df['Tmax'], color='orange')
    ax.axhline(45, linestyle='--', color='red')
    ax.set_xlabel('Cycle')
    ax.set_ylabel('Tmax (°C)')
    ax.set_title('Temperatura massima per ciclo')
    plt.tight_layout()
    file2 = out_dir / 'tmax_cycle.png'
    fig.savefig(file2, dpi=300)
    plt.close(fig)
    charts['tmax_cycle'] = file2
    return charts


def build_pdf(stats: dict, charts: dict, output_path: Path):
    """Generate a PDF report using ReportLab."""
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(output_path), pagesize=A4)
    flow = []

    flow.append(Paragraph('<b>Relazione tecnica – Batteria</b>', styles['Title']))
    flow.append(Spacer(1,12))

    data = [
        ['Indicatore','Valore','Soglia','Stato'],
        ['Scariche profonde (<20% SoC)', f"{stats['over_discharge']} / {stats['total_cycles']}", '0',
         'Critico' if stats['over_discharge']>0 else 'OK'],
        ['Sovra‑cariche (>105% SoC)', stats['over_charge'], '0', 'Critico' if stats['over_charge']>0 else 'OK'],
        ['Cariche complete (≥99% SoC)', stats['full_charges'], '>=95%', 'OK' if stats['full_charges']/stats['total_cycles']>=0.95 else 'Da migliorare'],
        ['Efficienza Ah', f"{stats['efficiency']:.2f}", '1.05‑1.10', 'OK' if 1.05<=stats['efficiency']<=1.10 else 'Check'],
    ]
    tbl = Table(data, hAlign='LEFT')
    tbl.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
        ('GRID',(0,0),(-1,-1),0.25,colors.black),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
    ]))
    flow.append(tbl)
    flow.append(Spacer(1,12))

    for desc, path in charts.items():
        flow.append(Paragraph(desc.replace('_',' ').title(), styles['Heading2']))
        flow.append(RLImage(str(path), width=6*72, height=3*72))
        flow.append(Spacer(1,12))

    doc.build(flow)

# -------------------------------------------------------------
# Streamlit app
# -------------------------------------------------------------

st.set_page_config(page_title='Battery Report Generator')
st.title('Battery Report Generator')

uploaded_file = st.file_uploader('Carica il file Excel di log batteria', type=['xls','xlsx'])

if uploaded_file is not None:
    df_raw = load_data(uploaded_file)
    df, stats = compute_indicators(df_raw)

    st.subheader('KPI principali')
    col1, col2, col3 = st.columns(3)
    col1.metric('Cicli totali', stats['total_cycles'])
    col2.metric('Scariche profonde', stats['over_discharge'])
    col3.metric('Cariche complete', stats['full_charges'])

    tmp_dir = Path('tmp')
    tmp_dir.mkdir(exist_ok=True)
    charts = create_charts(df, tmp_dir)

    for path in charts.values():
        st.image(str(path))

    pdf_path = tmp_dir / 'relazione_batteria.pdf'
    build_pdf(stats, charts, pdf_path)

    with open(pdf_path,'rb') as f:
        st.download_button('Scarica relazione PDF', f, file_name='relazione_batteria.pdf', mime='application/pdf')

    st.info('Il PDF viene generato localmente.')