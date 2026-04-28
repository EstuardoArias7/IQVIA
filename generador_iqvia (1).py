import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import re

st.set_page_config(page_title="Generador IQVIA", page_icon="💊", layout="wide")
st.title("💊 Generador de Mercado IQVIA")
st.caption("Sube la sábana mensual · selecciona criterios · descarga el Excel con MAT dinámico, filtro de país y forma de administración.")

# ── ESTILOS EXCEL ─────────────────────────────────────────────────────────
DARK_BLUE  = PatternFill('solid', start_color='17375E')
MED_BLUE   = PatternFill('solid', start_color='1F4E79')
ALT_FILL   = PatternFill('solid', start_color='DEEAF1')
WHITE_FILL = PatternFill('solid', start_color='FFFFFF')
TOTAL_FILL = PatternFill('solid', start_color='D9E1F2')
GRAY_FILL  = PatternFill('solid', start_color='F2F2F2')

WHITE_BOLD  = Font(name='Arial', bold=True, color='FFFFFF', size=10)
DARK_BOLD   = Font(name='Arial', bold=True, color='000000', size=10)
DARK_NORM   = Font(name='Arial', size=9,  color='000000')
DARK_NORM10 = Font(name='Arial', size=10, color='000000')

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
RIGHT  = Alignment(horizontal='right',  vertical='center')
thin   = Side(style='thin', color='B8CCE4')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def hdr(ws, r, c, v, fill=DARK_BLUE, font=WHITE_BOLD, align=CENTER):
    cell = ws.cell(r, c, v)
    cell.fill = fill; cell.font = font; cell.alignment = align; cell.border = BORDER
    return cell

def dat(ws, r, c, v, font=DARK_NORM, align=LEFT, fill=None, fmt=None):
    cell = ws.cell(r, c, v)
    cell.font = font; cell.alignment = align; cell.border = BORDER
    if fill: cell.fill = fill
    if fmt:  cell.number_format = fmt
    return cell

def filter_label(ws, r, c, label, value):
    cl = ws.cell(r, c, label)
    cl.font = DARK_BOLD; cl.fill = GRAY_FILL; cl.alignment = LEFT
    cv = ws.cell(r, c+1, value)
    cv.font = DARK_NORM10; cv.fill = GRAY_FILL; cv.alignment = LEFT

# ── PARSEO DE MAT DINÁMICO ────────────────────────────────────────────────
MONTH_MAP    = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
                'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
NUM_TO_MON   = {v: k for k, v in MONTH_MAP.items()}

def parse_col_date(col_name):
    m = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', str(col_name))
    if m:
        return int(m.group(2)), MONTH_MAP[m.group(1)], m.group(1)
    return None

def compute_mat_periods(unit_cols, usd_cols):
    """
    MAT = mes cierre + 11 meses anteriores (12 meses corridos).
    MAT Feb 2026 = Mar 2025 → Feb 2026
    MAT Mar 2026 = Apr 2025 → Mar 2026
    Un MAT por año (el del último mes disponible de ese año).
    """
    dated = []
    for uc, dc in zip(unit_cols, usd_cols):
        parsed = parse_col_date(uc)
        if parsed:
            yr, mo, mo_name = parsed
            dated.append(((yr, mo), mo_name, uc, dc))
    dated.sort(key=lambda x: x[0])

    if len(dated) < 12:
        return []

    # Último índice por año
    last_by_year = {}
    for i, ((yr, mo), mo_name, uc, dc) in enumerate(dated):
        last_by_year[yr] = i

    periods = []
    for yr in sorted(last_by_year.keys()):
        end_i = last_by_year[yr]
        if end_i < 11:
            continue
        window = dated[end_i - 11: end_i + 1]
        if len(window) != 12:
            continue
        _, mo_name, _, _ = dated[end_i]
        label   = f"MAT {mo_name} {yr}"
        ucols12 = [w[2] for w in window]
        dcols12 = [w[3] for w in window]
        periods.append((label, ucols12, dcols12))

    return periods

# ── CARGA DE DATOS ────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Leyendo sábana IQVIA... puede tardar unos segundos.")
def cargar_sabana(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), header=None)
    headers_raw = list(df.iloc[0])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = headers_raw

    str_cols = ['Country Desc','Atc I','Atc IV','Molecule Desc','Prod Desc',
                'Pack Desc','Manu Desc','Pack Mark Desc','Pack Gene Desc','App1 Desc']
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    unit_cols = [c for c in df.columns if 'Sales Units Qty'          in str(c)]
    usd_cols  = [c for c in df.columns if 'Sales Value List Usd Amt'  in str(c)]

    for col in unit_cols + usd_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df, unit_cols, usd_cols

# ── HELPERS ───────────────────────────────────────────────────────────────
def fmt_launch(val):
    try:
        if isinstance(val, datetime): return val.strftime('%d/%m/%Y')
        if isinstance(val, (int, float)) and not np.isnan(float(val)):
            return (datetime(1899, 12, 30) + pd.Timedelta(days=int(val))).strftime('%d/%m/%Y')
        if isinstance(val, str) and val not in ('nan','NaT','','None'): return val
    except:
        pass
    return ''

# ── GENERADOR DE EXCEL ────────────────────────────────────────────────────
def generar_excel(df_filtrado, df_contexto, mat_periods, pais_label, forma_label):
    wb = Workbook()

    yr_short, yr_full = [], []
    for lbl, _, _ in mat_periods:
        parts = lbl.replace('MAT ', '').split()
        yr_full.append(f"{parts[0]} {parts[1]}")
        yr_short.append(parts[1])

    last_yr = yr_short[-1]
    prev_yr = yr_short[-2] if len(yr_short) >= 2 else None

    for df in [df_filtrado, df_contexto]:
        for lbl, ucols, dcols in mat_periods:
            yr = lbl.replace('MAT ', '').split()[-1]
            df[f'UNI {yr}'] = df[ucols].sum(axis=1)
            df[f'USD {yr}'] = df[dcols].sum(axis=1)

    uni_mat   = [f'UNI {y}' for y in yr_short]
    usd_mat   = [f'USD {y}' for y in yr_short]
    atc_label = df_filtrado['Atc IV'].mode()[0] if len(df_filtrado) else ''

    # ── HOJA 1: DATA MAT ─────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = 'DATA MAT'

    filter_label(ws1, 1, 1, 'PAIS',        pais_label)
    filter_label(ws1, 2, 1, 'FORMA ADMON',  forma_label)

    info_cols = ['Country Desc','Atc I','Atc IV','Molecule Desc','Prod Desc','Pack Desc',
                 'Manu Desc','Pack Mark Desc','Pack Gene Desc','App1 Desc','Pack Launch Dt']
    h1 = info_cols + \
         [f'Sales Units Qty\nMAT {yf}'          for yf in yr_full] + \
         [f'Sales Value List Usd Amt\nMAT {yf}'  for yf in yr_full]

    for ci, h in enumerate(h1, 1):
        hdr(ws1, 4, ci, h)
    ws1.row_dimensions[4].height = 42

    for ri, (_, row) in enumerate(df_filtrado.iterrows(), 5):
        fill = ALT_FILL if ri % 2 == 0 else WHITE_FILL
        for ci, col in enumerate(info_cols, 1):
            v = row.get(col, '')
            if col == 'Pack Launch Dt': v = fmt_launch(v)
            dat(ws1, ri, ci, v if str(v) not in ('nan','None') else '', fill=fill)
        for ci, col in enumerate(uni_mat, len(info_cols)+1):
            dat(ws1, ri, ci, round(float(row.get(col,0)),0), align=RIGHT, fill=fill, fmt='#,##0')
        for ci, col in enumerate(usd_mat, len(info_cols)+len(uni_mat)+1):
            dat(ws1, ri, ci, round(float(row.get(col,0)),3), align=RIGHT, fill=fill, fmt='#,##0.000')

    for i, w in enumerate([14,25,32,45,22,30,16,12,10,24,14]+[14]*len(mat_periods)*2, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = 'A5'

    # ── HOJA 2: 02_Mercado CT ─────────────────────────────────────────────
    ws2 = wb.create_sheet('02_Mercado CT')
    filter_label(ws2, 1, 1, 'PAIS',        pais_label)
    filter_label(ws2, 2, 1, 'FORMA ADMON',  forma_label)

    ct_hdrs = ['MERCADO','MOLECULA'] + \
              [f'UNI {y}' for y in yr_short] + \
              [f'US$ {y}' for y in yr_short] + ['crec%']
    for ci, h in enumerate(ct_hdrs, 1):
        hdr(ws2, 4, ci, h)
    ws2.row_dimensions[4].height = 30

    mol_agg = df_contexto.groupby('Molecule Desc')[uni_mat+usd_mat].sum()
    mol_agg = mol_agg.sort_values(f'USD {last_yr}', ascending=False).reset_index()

    ri = 5
    first = True
    tot_uni = {y: 0.0 for y in yr_short}
    tot_usd = {y: 0.0 for y in yr_short}

    for _, mrow in mol_agg.iterrows():
        fill = ALT_FILL if ri % 2 == 0 else WHITE_FILL
        dat(ws2, ri, 1, atc_label if first else '', font=DARK_BOLD if first else DARK_NORM, fill=fill)
        dat(ws2, ri, 2, mrow['Molecule Desc'], fill=fill)
        first = False
        for ci, y in enumerate(yr_short, 3):
            v = round(float(mrow[f'UNI {y}']), 0)
            dat(ws2, ri, ci, v, align=RIGHT, fill=fill, fmt='#,##0')
            tot_uni[y] += v
        for ci, y in enumerate(yr_short, 3+len(yr_short)):
            v = round(float(mrow[f'USD {y}']), 3)
            dat(ws2, ri, ci, v, align=RIGHT, fill=fill, fmt='#,##0.000')
            tot_usd[y] += v
        if prev_yr:
            u25 = float(mrow[f'USD {prev_yr}'])
            u26 = float(mrow[f'USD {last_yr}'])
            crec = (u26-u25)/u25 if u25 != 0 else None
            c = ws2.cell(ri, 3+2*len(yr_short), round(crec,6) if crec is not None else '')
            c.font = DARK_NORM; c.border = BORDER; c.alignment = RIGHT
            if crec is not None: c.number_format = '0.000000'
        ri += 1

    # Subtotal
    dat(ws2, ri, 1, f'Total {atc_label}', font=DARK_BOLD, fill=TOTAL_FILL)
    dat(ws2, ri, 2, '', fill=TOTAL_FILL)
    for ci, y in enumerate(yr_short, 3):
        dat(ws2, ri, ci, round(tot_uni[y],0), font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0')
    for ci, y in enumerate(yr_short, 3+len(yr_short)):
        dat(ws2, ri, ci, round(tot_usd[y],3), font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0.000')
    if prev_yr:
        u25t = tot_usd[prev_yr]; u26t = tot_usd[last_yr]
        crect = (u26t-u25t)/u25t if u25t else None
        c = ws2.cell(ri, 3+2*len(yr_short), round(crect,6) if crect else '')
        c.font=DARK_BOLD; c.fill=TOTAL_FILL; c.border=BORDER; c.alignment=RIGHT
        if crect: c.number_format='0.000000'
    ri += 1

    # Total general
    c=ws2.cell(ri,1,'Total general'); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=LEFT
    ws2.cell(ri,2,'').fill=MED_BLUE; ws2.cell(ri,2).border=BORDER
    for ci, y in enumerate(yr_short, 3):
        c=ws2.cell(ri,ci,round(tot_uni[y],0)); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT; c.number_format='#,##0'
    for ci, y in enumerate(yr_short, 3+len(yr_short)):
        c=ws2.cell(ri,ci,round(tot_usd[y],3)); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT; c.number_format='#,##0.000'
    if prev_yr:
        u25t=tot_usd[prev_yr]; u26t=tot_usd[last_yr]
        crect=(u26t-u25t)/u25t if u25t else None
        c=ws2.cell(ri,3+2*len(yr_short),round(crect,6) if crect else '')
        c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT
        if crect: c.number_format='0.000000'

    ws2.column_dimensions['A'].width = 36
    ws2.column_dimensions['B'].width = 55
    for ci in range(3, len(ct_hdrs)+1):
        ws2.column_dimensions[get_column_letter(ci)].width = 16

    # ── HOJA 3: 03_Mercado Molécula ──────────────────────────────────────
    ws3 = wb.create_sheet('03_Mercado Molécula')
    filter_label(ws3, 1, 1, 'PAIS',        pais_label)
    filter_label(ws3, 2, 1, 'FORMA ADMON',  forma_label)

    mol_hdrs = ['MERCADO','MOLECULA','MARCA','PRESENTACION'] + \
               [f'UNI {y}' for y in yr_short] + \
               [f'US$ {y}' for y in yr_short]
    for ci, h in enumerate(mol_hdrs, 1):
        hdr(ws3, 4, ci, h)
    ws3.row_dimensions[4].height = 30

    grp = df_filtrado.groupby(['Molecule Desc','Prod Desc','Pack Desc'])[uni_mat+usd_mat].sum().reset_index()
    grp = grp.sort_values(f'USD {last_yr}', ascending=False)

    ri = 5
    mol_totals = {}
    mol_first  = {}
    grand_uni  = {y: 0.0 for y in yr_short}
    grand_usd  = {y: 0.0 for y in yr_short}

    for _, grow in grp.iterrows():
        mol = grow['Molecule Desc']
        if mol not in mol_totals:
            mol_totals[mol] = {y: {'uni':0.0,'usd':0.0} for y in yr_short}
            mol_first[mol]  = True
        fill = ALT_FILL if ri % 2 == 0 else WHITE_FILL
        dat(ws3, ri, 1, atc_label if mol_first[mol] else '', font=DARK_BOLD if mol_first[mol] else DARK_NORM, fill=fill)
        dat(ws3, ri, 2, mol       if mol_first[mol] else '', fill=fill)
        mol_first[mol] = False
        dat(ws3, ri, 3, grow['Prod Desc'], fill=fill)
        dat(ws3, ri, 4, grow['Pack Desc'], fill=fill)
        for ci, y in enumerate(yr_short, 5):
            v = round(float(grow[f'UNI {y}']),0)
            dat(ws3, ri, ci, v, align=RIGHT, fill=fill, fmt='#,##0')
            mol_totals[mol][y]['uni'] += v
        for ci, y in enumerate(yr_short, 5+len(yr_short)):
            v = round(float(grow[f'USD {y}']),3)
            dat(ws3, ri, ci, v, align=RIGHT, fill=fill, fmt='#,##0.000')
            mol_totals[mol][y]['usd'] += v
        ri += 1

    for mol, tots in mol_totals.items():
        dat(ws3, ri, 1, '', fill=TOTAL_FILL)
        dat(ws3, ri, 2, f'Total {mol}', font=DARK_BOLD, fill=TOTAL_FILL)
        dat(ws3, ri, 3, '', fill=TOTAL_FILL)
        dat(ws3, ri, 4, '', fill=TOTAL_FILL)
        for ci, y in enumerate(yr_short, 5):
            v = round(tots[y]['uni'],0)
            dat(ws3, ri, ci, v, font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0')
            grand_uni[y] += v
        for ci, y in enumerate(yr_short, 5+len(yr_short)):
            v = round(tots[y]['usd'],3)
            dat(ws3, ri, ci, v, font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0.000')
            grand_usd[y] += v
        ri += 1

    dat(ws3, ri, 1, f'Total {atc_label}', font=DARK_BOLD, fill=TOTAL_FILL)
    for cc in [2,3,4]: dat(ws3, ri, cc, '', fill=TOTAL_FILL)
    for ci, y in enumerate(yr_short, 5):
        dat(ws3, ri, ci, round(grand_uni[y],0), font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0')
    for ci, y in enumerate(yr_short, 5+len(yr_short)):
        dat(ws3, ri, ci, round(grand_usd[y],3), font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0.000')
    ri += 1

    c=ws3.cell(ri,1,'Total general'); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=LEFT
    for cc in [2,3,4]: ws3.cell(ri,cc,'').fill=MED_BLUE; ws3.cell(ri,cc).border=BORDER
    for ci, y in enumerate(yr_short, 5):
        c=ws3.cell(ri,ci,round(grand_uni[y],0)); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT; c.number_format='#,##0'
    for ci, y in enumerate(yr_short, 5+len(yr_short)):
        c=ws3.cell(ri,ci,round(grand_usd[y],3)); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT; c.number_format='#,##0.000'

    ws3.column_dimensions['A'].width = 36
    ws3.column_dimensions['B'].width = 48
    ws3.column_dimensions['C'].width = 22
    ws3.column_dimensions['D'].width = 34
    for ci in range(5, len(mol_hdrs)+1):
        ws3.column_dimensions[get_column_letter(ci)].width = 15

    # ── HOJA 4: FECHA LANZAMIENTO ─────────────────────────────────────────
    ws4 = wb.create_sheet('FECHA LANZAMIENTO')
    filter_label(ws4, 1, 1, 'PAIS',        pais_label)
    filter_label(ws4, 2, 1, 'FORMA ADMON',  forma_label)

    fl_hdrs = ['MERCADO','MOLECULA','MARCA','FECHA LANZAMIENTO',
               f'UNI {last_yr}', f'US$ {last_yr}']
    for ci, h in enumerate(fl_hdrs, 1):
        hdr(ws4, 4, ci, h)
    ws4.row_dimensions[4].height = 30

    brands = df_filtrado.groupby(['Molecule Desc','Prod Desc','Pack Launch Dt']).agg(
        **{f'UNI {last_yr}': (f'UNI {last_yr}','sum'),
           f'USD {last_yr}': (f'USD {last_yr}','sum')}
    ).reset_index().sort_values(f'USD {last_yr}', ascending=False)

    ri = 5
    mol_first4 = {}
    for _, brow in brands.iterrows():
        mol = brow['Molecule Desc']
        if mol not in mol_first4: mol_first4[mol] = True
        fill = ALT_FILL if ri % 2 == 0 else WHITE_FILL
        dat(ws4, ri, 1, atc_label if mol_first4[mol] else '', font=DARK_BOLD if mol_first4[mol] else DARK_NORM, fill=fill)
        dat(ws4, ri, 2, mol       if mol_first4[mol] else '', fill=fill)
        mol_first4[mol] = False
        dat(ws4, ri, 3, brow['Prod Desc'], fill=fill)
        dat(ws4, ri, 4, fmt_launch(brow['Pack Launch Dt']), fill=fill)
        dat(ws4, ri, 5, round(float(brow[f'UNI {last_yr}']),0), align=RIGHT, fill=fill, fmt='#,##0')
        dat(ws4, ri, 6, round(float(brow[f'USD {last_yr}']),3), align=RIGHT, fill=fill, fmt='#,##0.000')
        ri += 1

    tot_u = round(float(df_filtrado[f'UNI {last_yr}'].sum()),0)
    tot_d = round(float(df_filtrado[f'USD {last_yr}'].sum()),3)
    dat(ws4, ri, 1, f'Total {atc_label}', font=DARK_BOLD, fill=TOTAL_FILL)
    for cc in [2,3,4]: dat(ws4, ri, cc, '', fill=TOTAL_FILL)
    dat(ws4, ri, 5, tot_u, font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0')
    dat(ws4, ri, 6, tot_d, font=DARK_BOLD, align=RIGHT, fill=TOTAL_FILL, fmt='#,##0.000')
    ri += 1
    c=ws4.cell(ri,1,'Total general'); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=LEFT
    for cc in [2,3,4]: ws4.cell(ri,cc,'').fill=MED_BLUE; ws4.cell(ri,cc).border=BORDER
    c=ws4.cell(ri,5,tot_u); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT; c.number_format='#,##0'
    c=ws4.cell(ri,6,tot_d); c.font=WHITE_BOLD; c.fill=MED_BLUE; c.border=BORDER; c.alignment=RIGHT; c.number_format='#,##0.000'

    ws4.column_dimensions['A'].width = 36
    ws4.column_dimensions['B'].width = 45
    ws4.column_dimensions['C'].width = 22
    ws4.column_dimensions['D'].width = 20
    ws4.column_dimensions['E'].width = 16
    ws4.column_dimensions['F'].width = 18

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── INTERFAZ STREAMLIT ────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Sábana IQVIA")
    archivo = st.file_uploader("Sube el archivo mensual (.xlsx)", type=['xlsx'])
    if archivo:
        st.success("Archivo cargado ✓")

if not archivo:
    st.info("👈 Sube la sábana mensual de IQVIA en el panel izquierdo para comenzar.")
    st.stop()

file_bytes = archivo.read()
df, unit_cols, usd_cols = cargar_sabana(file_bytes)
mat_periods = compute_mat_periods(unit_cols, usd_cols)

if not mat_periods:
    st.error("No se encontraron periodos MAT completos. Verifica que el archivo tenga al menos 12 meses de data.")
    st.stop()

st.success(f"✅ {len(df):,} registros · {len(mat_periods)} periodos MAT detectados: **{' · '.join([p[0] for p in mat_periods])}**")
st.divider()

# ── FILTROS DE BÚSQUEDA ───────────────────────────────────────────────────
st.subheader("🔍 Criterios de búsqueda")
st.caption("Selecciona uno o más valores. Puedes combinar filtros. Los filtros de país y forma de administración también aparecerán en el Excel generado.")

col1, col2, col3 = st.columns(3)

with col1:
    sel_mol = st.multiselect(
        "🧪 Molécula",
        options=sorted(df['Molecule Desc'].dropna().unique().tolist()),
        placeholder="Busca o selecciona moléculas..."
    )

with col2:
    sel_prod = st.multiselect(
        "💊 Producto / Marca",
        options=sorted(df['Prod Desc'].dropna().unique().tolist()),
        placeholder="Busca o selecciona productos..."
    )

with col3:
    sel_atc = st.multiselect(
        "📋 Clase terapéutica (ATC IV)",
        options=sorted(df['Atc IV'].dropna().unique().tolist()),
        placeholder="Busca o selecciona clase ATC..."
    )

col4, col5 = st.columns(2)
paises_disponibles = sorted(df['Country Desc'].dropna().unique().tolist())
formas_disponibles = sorted(df['App1 Desc'].dropna().unique().tolist())

with col4:
    sel_pais = st.multiselect(
        "🌎 País",
        options=paises_disponibles,
        default=paises_disponibles,
        placeholder="Selecciona países..."
    )

with col5:
    sel_forma = st.multiselect(
        "💉 Forma de administración",
        options=formas_disponibles,
        default=formas_disponibles,
        placeholder="Selecciona formas..."
    )

st.divider()
generar = st.button("🚀 Generar Excel", type="primary", use_container_width=True)

if generar:
    if not any([sel_mol, sel_prod, sel_atc]):
        st.warning("⚠️ Selecciona al menos una molécula, producto o clase terapéutica.")
        st.stop()
    if not sel_pais:
        st.warning("⚠️ Selecciona al menos un país.")
        st.stop()

    mask = pd.Series([True]*len(df), index=df.index)
    if sel_mol:  mask &= df['Molecule Desc'].isin(sel_mol)
    if sel_prod: mask &= df['Prod Desc'].isin(sel_prod)
    if sel_atc:  mask &= df['Atc IV'].isin(sel_atc)
    mask &= df['Country Desc'].isin(sel_pais)
    if len(sel_forma) < len(formas_disponibles):
        mask &= df['App1 Desc'].isin(sel_forma)

    df_filtrado = df[mask].copy()

    if len(df_filtrado) == 0:
        st.error("❌ No se encontraron registros con esa combinación de filtros.")
        st.stop()

    # Contexto del mercado: mismo ATC IV + mismos filtros de país y forma
    atc_iv = df_filtrado['Atc IV'].mode()[0]
    mask_ctx = df['Atc IV'] == atc_iv
    mask_ctx &= df['Country Desc'].isin(sel_pais)
    if len(sel_forma) < len(formas_disponibles):
        mask_ctx &= df['App1 Desc'].isin(sel_forma)
    df_contexto = df[mask_ctx].copy()

    pais_label  = ', '.join(sel_pais)  if len(sel_pais)  < len(paises_disponibles)  else '(Todas)'
    forma_label = ', '.join(sel_forma) if len(sel_forma) < len(formas_disponibles) else '(Todas)'

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros",  f"{len(df_filtrado):,}")
    c2.metric("Moléculas",  df_filtrado['Molecule Desc'].nunique())
    c3.metric("Productos",  df_filtrado['Prod Desc'].nunique())
    c4.metric("Países",     df_filtrado['Country Desc'].nunique())

    with st.spinner("Generando Excel..."):
        excel_buf = generar_excel(
            df_filtrado.copy(),
            df_contexto.copy(),
            mat_periods,
            pais_label,
            forma_label
        )

    nombre_base = (sel_mol[0] if sel_mol else sel_prod[0] if sel_prod else sel_atc[0])
    nombre_base = re.sub(r'[^A-Za-z0-9]', '_', nombre_base)[:30].upper()
    last_mat    = mat_periods[-1][0].replace(' ','_')
    filename    = f"MERCADO_{nombre_base}_{last_mat}.xlsx"

    st.download_button(
        label="⬇️ Descargar Excel",
        data=excel_buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.success(f"✅ **{filename}** listo para descargar.")
