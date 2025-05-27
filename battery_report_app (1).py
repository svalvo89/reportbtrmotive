import io
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, ListFlowable, ListItem, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------

def load_data(uploaded_file: io.BytesIO) -> pd.DataFrame:
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse(xls.sheet_names[0], header=1)
    return df


def compute_indicators(df: pd.DataFrame, c_nom: float, dod_thr: float = 0.8):
    df = df.copy()
    df['Ah_dis'] = pd.to_numeric(df['Ah Discharged'], errors='coerce')
    df['Ah_chg'] = pd.to_numeric(df['Ah Charged In Charge Phase'], errors='coerce')
    df['SoC_end_charge'] = pd.to_numeric(df['SoC at End of Charge [%]'], errors='coerce')
    df['Tmax'] = pd.to_numeric(df['Max. Temperature At Cycle (℃)'], errors='coerce')

    total = len(df)
    deep = (df['Ah_dis'] >= c_nom * dod_thr).sum()
    full = (df['SoC_end_charge'] >= 99).sum()
    partial = total - full
    eff = (df['Ah_chg'].sum() / df['Ah_dis'].sum()) if df['Ah_dis'].sum() else float('nan')
    cycles_above45 = (df['Tmax'] > 45).sum()
    tmax_avg = df['Tmax'].mean()

    stats = {
        'total': total,
        'deep': int(deep),
        'full': int(full),
        'partial': int(partial),
        'eff': float(eff),
        'cycles_above45': int(cycles_above45),
        'tmax_avg': float(tmax_avg)
    }
    return df, stats


def create_charts(df: pd.DataFrame, out_dir: Path, c_nom: float, dod_thr: float = 0.8):
    charts = {}

    # Ah chart
    fig, ax = plt.subplots(figsize=(8, 3))
    ax.bar(df['Cycle Count'] - 0.2, df['Ah_dis'], 0.4, color='#E74C3C', label='Ah scaricati')
    ax.bar(df['Cycle Count'] + 0.2, df['Ah_chg'], 0.4, color='#2ECC71', label='Ah caricati')
    y_max = df[['Ah_dis', 'Ah_chg']].values.max() * 1.05
    for x, complete in zip(df['Cycle Count'], df['SoC_end_charge'] >= 99):
        ax.text(x, y_max, '✓' if complete else '✗', color='green' if complete else 'red',
                ha='center', va='bottom', fontsize=8)
    ax.axhline(c_nom * dod_thr, linestyle='--', color='black', label=f'{int(dod_thr*100)}% DOD')
    ax.set_ylim(0, y_max * 1.15)
    ax.set_xlabel('Cycle')
    ax.set_ylabel('Ah')
    ax.set_title('Ah caricati / scaricati per ciclo')
    ax.legend()
    plt.tight_layout()
    ah_path = out_dir / 'ah_cycle.png'
    fig.savefig(ah_path, dpi=300)
    plt.close(fig)
    charts['ah_cycle'] = ah_path

    # Tmax chart
    fig, ax = plt.subplots(figsize=(8, 3))
    ax.bar(df['Cycle Count'], df['Tmax'], color='#F1C40F', label='Tmax ciclo')
    ax.axhline(45, linestyle='--', color='red', linewidth=1.5, label='Soglia 45 °C')
    ax.set_ylim(0, max(50, df['Tmax'].max() * 1.1))
    ax.set_xlabel('Cycle')
    ax.set_ylabel('Tmax (°C)')
    ax.set_title('Temperatura massima per ciclo')
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    tmax_path = out_dir / 'tmax_cycle.png'
    fig.savefig(tmax_path, dpi=300)
    plt.close(fig)
    charts['tmax_cycle'] = tmax_path

    return charts


def build_pdf(stats: dict, charts: dict, df: pd.DataFrame, output_path: Path,
              customer: str, battery_id: str, c_nom: float):
    styles = getSampleStyleSheet()
    body = ParagraphStyle('Body', parent=styles['Normal'], spaceAfter=6)
    sect = ParagraphStyle('Sect', parent=styles['Heading1'], textColor=colors.darkblue)

    flow = [
        Paragraph(f'Relazione tecnica – Batteria {battery_id}', styles['Title']),
        HRFlowable(width='100%', thickness=1, color=colors.darkblue),
        Spacer(1, 8),
        Paragraph(f'<b>Cliente:</b> {customer}', body),
        Paragraph(f'<b>Capacità nominale:</b> {int(c_nom)} Ah', body),
        Paragraph(f'<b>Cicli totali esaminati:</b> {stats['total']}', body),
        Spacer(1, 12),
        Paragraph('1. Indici chiave', sect)
    ]

    indici = [
        ['Indicatore', 'Valore'],
        [f'Scariche profonde (≥{int(c_nom*0.8)} Ah)', stats['deep']],
        ['Cariche parziali (<99% SoC)', stats['partial']],
        ['Efficienza Ah', f"{stats['eff']:.2f}"],
        ['Cicli Tmax >45 °C', stats['cycles_above45']],
        ['Tmax media (°C)', f"{stats['tmax_avg']:.1f}"]
    ]
    tbl = Table(indici, colWidths=[260, 100])
    tbl.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey)
    ]))
    flow += [tbl, Spacer(1, 12)]

    flow += [Paragraph('2. Grafici', sect),
             Paragraph('Ah caricati / scaricati per ciclo', styles['Heading3']),
             RLImage(str(charts['ah_cycle']), width=480, height=180),
             Spacer(1, 8),
             Paragraph('Temperatura massima per ciclo (soglia 45 °C)', styles['Heading3']),
             RLImage(str(charts['tmax_cycle']), width=480, height=180)]

    SimpleDocTemplate(str(output_path), pagesize=A4,
                      leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36).build(flow)

# -------------------------------------------------------------
# Streamlit app
# -------------------------------------------------------------

st.set_page_config(page_title='Battery Report Generator')

st.title('Battery Report Generator')

# Sidebar inputs
with st.sidebar:
    st.header('Dettagli relazione')
    customer_name = st.text_input('Nome cliente', '')
    battery_serial = st.text_input('Matricola batteria', '')
    capacity_nom = st.number_input('Capacità nominale (Ah)', min_value=10, max_value=2000, value=930, step=10)

uploaded_file = st.file_uploader('Carica il file Excel di log batteria', type=['xls', 'xlsx'])

if uploaded_file is not None and customer_name and battery_serial:
    df_raw = load_data(uploaded_file)
    df_proc, stats = compute_indicators(df_raw, capacity_nom)

    st.subheader('KPI principali')
    col1, col2, col3 = st.columns(3)
    col1.metric('Scariche profonde', stats['deep'])
    col2.metric('Cariche parziali', stats['partial'])
    col3.metric('Efficienza Ah', f"{stats['eff']:.2f}")

    tmp_dir = Path('tmp')
    tmp_dir.mkdir(exist_ok=True)
    charts = create_charts(df_proc, tmp_dir, capacity_nom)

    for img in charts.values():
        st.image(str(img))

    pdf_path = tmp_dir / 'relazione_batteria.pdf'
    build_pdf(stats, charts, df_proc, pdf_path, customer_name, battery_serial, capacity_nom)

    with open(pdf_path, 'rb') as f:
        st.download_button('Scarica relazione PDF', f, file_name='relazione_batteria.pdf', mime='application/pdf')
else:
    st.info('Compila i campi nella barra laterale e carica un file Excel per generare la relazione.')
