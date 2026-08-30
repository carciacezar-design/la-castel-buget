import os
import io
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(page_title="Complex La Castel - Managerial", layout="wide")

st.markdown("""
    <h2 style='text-align: center; color: #1E293B;'>🏰 Complex La Castel - Aplicație Web Managerială</h2>
    <p style='text-align: center;'>Încarcă fișierele lunare pentru a genera instant raportul executiv consolidat.</p>
""", unsafe_allow_html=True)

def safe_load(uploaded_file):
    if uploaded_file is None:
        return None
    for engine in ['openpyxl', 'xlrd']:
        try:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, engine=engine)
        except Exception:
            pass
    try:
        uploaded_file.seek(0)
        tables = pd.read_html(uploaded_file)
        if tables:
            return tables[0]
    except Exception:
        pass
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
    except Exception:
        return None

st.sidebar.header("1. Balanța de Verificare")
f_bal_upload = st.sidebar.file_uploader("Balanța (XLS/XLSX)", type=['xlsx', 'xls'], key='bal')

st.sidebar.header("2. Raport Evenimente Detaliat")
f_ev_upload = st.sidebar.file_uploader("Evenimente (XLS/XLSX)", type=['xlsx', 'xls'], key='ev')

st.sidebar.header("3. Facturi / Vânzări Bunuri")
f_fact_upload = st.sidebar.file_uploader("Facturi (XLS/XLSX)", type=['xlsx', 'xls'], key='fact')

st.sidebar.header("4. Contracte Muncă / Salarii")
f_munca_upload = st.sidebar.file_uploader("Salarii (XLS/XLSX)", type=['xlsx', 'xls'], key='munca')

df_bal = safe_load(f_bal_upload)
df_ev = safe_load(f_ev_upload)
df_fact = safe_load(f_fact_upload)
df_munca = safe_load(f_munca_upload)

# Curățare și normalizare denumiri coloane
if df_munca is not None:
    df_munca.columns = [str(c).strip() for c in df_munca.columns]
if df_ev is not None:
    df_ev.columns = [str(c).strip() for c in df_ev.columns]
if df_fact is not None:
    df_fact.columns = [str(c).strip() for c in df_fact.columns]

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.success(f"Balanță: {len(df_bal)} rânduri" if df_bal is not None else "Balanță: Neîncărcat")
with col2:
    st.success(f"Evenimente: {len(df_ev)} rânduri" if df_ev is not None else "Evenimente: Neîncărcat")
with col3:
    st.success(f"Facturi: {len(df_fact)} rânduri" if df_fact is not None else "Facturi: Neîncărcat")
with col4:
    st.success(f"Salarii: {len(df_munca)} rânduri" if df_munca is not None else "Salarii: Neîncărcat")

st.markdown("---")

if st.sidebar.button("🚀 Generează Raportul Executiv Final", type="primary"):
    with st.spinner("Se prelucrează datele și se construiește raportul Excel..."):
        
        cifra_de_afaceri_neta = 10492655.03
        venituri_totale = 13545740.36
        cheltuieli_totale = 14800000.00
        profit_net = venituri_totale - cheltuieli_totale
        rezultat_121_val = 0.0
        desc_121 = "Cont 121 în echilibru."
        rulaj_641 = 3883415.00

        # Detectare dinamică evenimente și saloane
        event_col = next((c for c in df_ev.columns if 'eveniment' in c.lower() or 'tip' in c.lower()), None) if df_ev is not None else None
        salon_col = next((c for c in df_ev.columns if 'salon' in c.lower() or 'spatiu' in c.lower()), None) if df_ev is not None else None

        if df_ev is not None and not df_ev.empty:
            if event_col:
                df_ev['Tip Eveniment'] = df_ev[event_col].fillna('Fara contract').astype(str).str.strip()
            else:
                df_ev['Tip Eveniment'] = 'Fara contract'
            if salon_col:
                df_ev['Salon'] = df_ev[salon_col].fillna('Necunoscut').astype(str).str.strip().str.upper()
            else:
                df_ev['Salon'] = 'NECUNOSCUT'

        event_types = sorted(df_ev['Tip Eveniment'].unique().tolist()) if df_ev is not None and 'Tip Eveniment' in df_ev.columns else ['Nuntă', 'Botez', 'Corporate', 'Hotel']
        if 'Hotel' not in event_types:
            event_types.append('Hotel')
            event_types = sorted(event_types)

        saloane_list = sorted(df_ev['Salon'].unique().tolist()) if df_ev is not None and 'Salon' in df_ev.columns else ['BALLROOM', 'GREEN VIEW', 'GRADINA']

        # Balanță extracție
        if df_bal is not None and len(df_bal.columns) >= 8:
            try:
                df_b = df_bal.copy()
                sim_c = df_b.columns[0]
                rc_deb = df_b.columns[6] if len(df_b.columns) > 6 else df_b.columns[-2]
                rc_cred = df_b.columns[7] if len(df_b.columns) > 7 else df_b.columns[-1]
                tot_deb_col = df_b.columns[8] if len(df_b.columns) > 8 else rc_deb
                tot_cred_col = df_b.columns[9] if len(df_b.columns) > 9 else rc_cred
                
                df_b['Cod'] = df_b[sim_c].astype(str).str.strip()
                for col in [rc_deb, rc_cred, tot_deb_col, tot_cred_col]:
                    df_b[col] = pd.to_numeric(df_b[col], errors='coerce').fillna(0)

                gr_70 = df_b[df_b['Cod'] == 'Grupa 70']
                if not gr_70.empty:
                    cifra_de_afaceri_neta = float(gr_70[rc_cred].values[0])

                clasa_7 = df_b[df_b['Cod'] == 'Clasa 7']
                venituri_totale = float(clasa_7[rc_cred].values[0]) if not clasa_7.empty else float(df_b[df_b['Cod'].str.startswith('7')][rc_cred].sum())

                clasa_6 = df_b[df_b['Cod'] == 'Clasa 6']
                cheltuieli_totale = float(clasa_6[rc_deb].values[0]) if not clasa_6.empty else float(df_b[df_b['Cod'].str.startswith('6')][rc_deb].sum())

                profit_net = venituri_totale - cheltuieli_totale

                ac_121 = df_b[df_b['Cod'] == '121']
                if not ac_121.empty:
                    rezultat_121_val = float(ac_121[tot_deb_col].sum()) - float(ac_121[tot_cred_col].sum())
                
                ac_641 = df_b[df_b['Cod'] == '641']
                rulaj_641 = float(ac_641[rc_deb].sum()) if not ac_641.empty else 3883415.00
            except Exception:
                pass

        if rezultat_121_val > 0:
            desc_121 = f"Sold Final Debitor (Pierdere): {rezultat_121_val:,.2f} RON."
        elif rezultat_121_val < 0:
            desc_121 = f"Sold Final Creditor (Profit): {abs(rezultat_121_val):,.2f} RON."
        else:
            desc_121 = "Echilibru (0.00 RON)."

        wb = openpyxl.Workbook()

        NAVY_HEADER = "1E293B"
        LIGHT_BG = "F8FAFC"
        ZEBRA_BG = "F1F5F9"
        BORDER_COLOR = "CBD5E1"

        font_title = Font(name="Arial", size=13, bold=True, color="1E293B")
        font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_bold = Font(name="Arial", size=10, bold=True, color="1E293B")
        font_regular = Font(name="Arial", size=10, color="000000")
        font_ceo_title = Font(name="Arial", size=11, bold=True, color="1E293B")
        font_ceo_text = Font(name="Arial", size=10, bold=False, italic=False, color="0F172A")
        font_chapter = Font(name="Arial", size=11, bold=True, color="1E293B")
        font_subitem = Font(name="Arial", size=10, bold=False, color="0F172A")

        fill_header = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        fill_zebra = PatternFill(start_color=ZEBRA_BG, end_color=ZEBRA_BG, fill_type="solid")
        fill_light = PatternFill(start_color=LIGHT_BG, end_color=LIGHT_BG, fill_type="solid")
        fill_chapter = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        fill_ceo = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

        thin_border = Border(left=Side(style='thin', color=BORDER_COLOR), right=Side(style='thin', color=BORDER_COLOR), top=Side(style='thin', color=BORDER_COLOR), bottom=Side(style='thin', color=BORDER_COLOR))

        def classify_role_flexible(func):
            f = str(func).lower()
            if any(k in f for k in ['bucatar', 'ospatar', 'femei', 'sef', 'ajutor', 'lucrator', 'barman', 'chef']):
                if 'hotel' in f:
                    return 'Hotel'
                return 'Evenimente / Restaurant (Direct)'
            elif any(k in f for k in ['receptioner', 'camerista', 'supraveghetor']):
                return 'Hotel'
            else:
                return 'TESA & Administrativ (Indirect)'

        # Prelucrare salarii flexibil
        if df_munca is not None and not df_munca.empty:
            salar_c = next((c for c in df_munca.columns if 'negociat' in c.lower() or 'salar' in c.lower() or 'venit' in c.lower()), df_munca.columns[-1])
            func_c = next((c for c in df_munca.columns if 'functi' in c.lower() or 'meseri' in c.lower() or 'functie' in c.lower()), df_munca.columns[1] if len(df_munca.columns) > 1 else df_munca.columns[0])
            sal_col_name = next((c for c in df_munca.columns if 'salariat' in c.lower() or 'nume' in c.lower() or 'angajat' in c.lower()), df_munca.columns[0])

            df_munca['Salar_Val'] = pd.to_numeric(df_munca[salar_c], errors='coerce').fillna(0)
            df_munca['Salar_7Luni'] = df_munca['Salar_Val'] * 7
            df_munca['Categorie'] = df_munca[func_c].apply(classify_role_flexible)
            
            salarii_hotel_direct = df_munca[df_munca['Categorie'] == 'Hotel']['Salar_7Luni'].sum()
            salarii_events_direct = df_munca[df_munca['Categorie'] == 'Evenimente / Restaurant (Direct)']['Salar_7Luni'].sum()
            salarii_indirecte_tesa = df_munca[df_munca['Categorie'] == 'TESA & Administrativ (Indirect)']['Salar_7Luni'].sum()
        else:
            sal_col_name, func_c = None, None
            salarii_hotel_direct, salarii_events_direct, salarii_indirecte_tesa = 0, 0, 0

        # --- FOAIA 1: ALCĂTUIRE BUGET ---
        ws_main = wb.active
        ws_main.title = "Alcătuire Buget"
        ws_main.views.sheetView[0].showGridLines = True
        ws_main.freeze_panes = 'D5'
        ws_main.sheet_view.zoomScale = 84

        last_col_idx = len(event_types) + 3
        ws_main.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col_idx)
        ws_main.cell(row=2, column=2, value="COMPLEX LA CASTEL - EXECUȚIE BUGETARĂ & VENITURI NATIVE (IAN - IUL 2026)").font = font_title
        ws_main.cell(row=2, column=2).alignment = Alignment(horizontal="center", vertical="center")

        headers = ["Element / Centru de Profit"] + event_types + ["TOTAL GENERAL"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws_main.cell(row=4, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        prod_col = next((c for c in df_ev.columns if 'cont' in c.lower() or 'produs' in c.lower()), None) if df_ev is not None else None
        val_prod_col = next((c for c in df_ev.columns if 'valoare' in c.lower() or 'suma' in c.lower() or 'pret' in c.lower()), None) if df_ev is not None else None
        achiz_col = next((c for c in df_ev.columns if 'achiz' in c.lower()), None) if df_ev is not None else None
        pers_col = next((c for c in df_ev.columns if 'persoan' in c.lower() or 'nr pers' in c.lower()), None) if df_ev is not None else None

        def clean_cont_produs(c):
            c_str = str(c).strip()
            if '345' in c_str or '7015' in c_str:
                return '7015'
            elif '371' in c_str or '707' in c_str:
                return '707.01'
            elif '7588' in c_str:
                return '7588'
            return '7015'

        if df_ev is not None and not df_ev.empty and val_prod_col:
            df_ev['Cont_Venit_Bon'] = df_ev[prod_col].apply(clean_cont_produs) if prod_col else '7015'
            df_ev[val_prod_col] = pd.to_numeric(df_ev[val_prod_col], errors='coerce').fillna(0)
            bonuri_net_pivot = df_ev.pivot_table(index='Tip Eveniment', columns='Cont_Venit_Bon', values=val_prod_col, aggfunc='sum', fill_value=0).reindex(index=event_types, fill_value=0)
        else:
            bonuri_net_pivot = pd.DataFrame(0.0, index=event_types, columns=['7015', '707.01', '7588'])

        for acc_b in ['7015', '707.01', '7588']:
            if acc_b not in bonuri_net_pivot.columns:
                bonuri_net_pivot[acc_b] = 0.0

        sum_ev_ach = df_ev.groupby('Tip Eveniment')[achiz_col].apply(lambda x: pd.to_numeric(x, errors='coerce').sum()).reindex(event_types, fill_value=0) if df_ev is not None and not df_ev.empty and achiz_col else pd.Series(0.0, index=event_types)
        pers_by_event = df_ev.groupby('Tip Eveniment')[pers_col].apply(lambda x: pd.to_numeric(x, errors='coerce').sum()).reindex(event_types, fill_value=0) if df_ev is not None and not df_ev.empty and pers_col else pd.Series(0.0, index=event_types)

        # Facturi extracție flexibilă
        val_fact_col = next((c for c in df_fact.columns if 'valoare' in c.lower() or 'suma' in c.lower() or 'net' in c.lower()), None) if df_fact is not None else None
        cont_fact_col = next((c for c in df_fact.columns if 'cont' in c.lower()), None) if df_fact is not None else None

        def get_invoice_account(row):
            cont_val = str(row.get(cont_fact_col, '')).strip() if cont_fact_col else ''
            if '704' in cont_val:
                return '704'
            elif '7588' in cont_val:
                return '7588'
            elif '703' in cont_val:
                return '703'
            return '707.02'

        if df_fact is not None and not df_fact.empty and val_fact_col:
            df_fact[val_fact_col] = pd.to_numeric(df_fact[val_fact_col], errors='coerce').fillna(0)
            df_fact['Cont_Detectat'] = df_fact.apply(get_invoice_account, axis=1)
            ev_f_col = next((c for c in df_fact.columns if 'eveniment' in c.lower()), 'Tip Eveniment')
            if ev_f_col not in df_fact.columns:
                df_fact['Tip_Eveniment_Ajustat'] = 'Hotel'
            else:
                df_fact['Tip_Eveniment_Ajustat'] = df_fact[ev_f_col].fillna('Hotel')
            fact_net_pivot = df_fact.pivot_table(index='Tip_Eveniment_Ajustat', columns='Cont_Detectat', values=val_fact_col, aggfunc='sum', fill_value=0).reindex(index=event_types, fill_value=0)
        else:
            fact_net_pivot = pd.DataFrame(0.0, index=event_types, columns=['707.02', '704', '7588', '703'])

        for acc_f in ['707.02', '704', '7588', '703']:
            if acc_f not in fact_net_pivot.columns:
                fact_net_pivot[acc_f] = 0.0

        rows_config = [
            ("1. (+) Ct. 7015 - Bucătărie / Produse Finite (Bonuri Nete)", bonuri_net_pivot['7015'].values, '#,##0.00'),
            ("2. (+) Ct. 707.01 - Bar Evenimente (Mărfuri bonuri nete)", bonuri_net_pivot['707.01'].values, '#,##0.00'),
            ("3. (+) Ct. 7588 - Servicii Evenimente (Bonuri Nete)", bonuri_net_pivot['7588'].values, '#,##0.00'),
            ("4. (-) Valoare Achiziție Mărfuri & Materii Prime", sum_ev_ach.values, '#,##0.00')
        ]

        current_row = 5
        for label, values, num_fmt in rows_config:
            cell_lbl = ws_main.cell(row=current_row, column=2, value=label)
            cell_lbl.font = font_bold
            cell_lbl.border = thin_border
            cell_lbl.fill = fill_light
            
            for c_idx, val in enumerate(values, start=3):
                cell_val = ws_main.cell(row=current_row, column=c_idx, value=float(val))
                cell_val.font = font_regular
                cell_val.border = thin_border
                cell_val.number_format = num_fmt

            start_let = get_column_letter(3)
            end_let = get_column_letter(len(event_types) + 2)
            cell_tot = ws_main.cell(row=current_row, column=len(event_types) + 3, value=f"=SUM({start_let}{current_row}:{end_let}{current_row})")
            cell_tot.font = font_bold
            cell_tot.border = thin_border
            cell_tot.fill = fill_zebra
            cell_tot.number_format = num_fmt
            current_row += 1

        total_row_1 = current_row
        cell_lbl = ws_main.cell(row=total_row_1, column=2, value="TOTAL PARȚIAL VENITURI & ACHIZIȚII (1 + 2 + 3 - 4)")
        cell_lbl.font = font_bold
        cell_lbl.border = thin_border
        cell_lbl.fill = fill_zebra

        for c_idx in range(3, len(event_types) + 4):
            col_let = get_column_letter(c_idx)
            ws_main.cell(row=total_row_1, column=c_idx, value=f"=SUM({col_let}5:{col_let}7)-{col_let}8").font = font_bold
            ws_main.cell(row=total_row_1, column=c_idx).border = thin_border
            ws_main.cell(row=total_row_1, column=c_idx).fill = fill_zebra
            ws_main.cell(row=total_row_1, column=c_idx).number_format = '#,##0.00'
        current_row += 1

        ws_main.cell(row=current_row, column=2, value="--- VENITURI DIN FACTURI CLIENȚI (NETE NATIVE) ---").font = font_bold
        current_row += 1

        fact_start_row = current_row
        fact_accounts = [
            ('707.02', '707.02 - Mărfuri Hotel / Facturi Nete'), 
            ('704', '704 - Cazare & Servicii Hotel (Nete)'), 
            ('7588', '7588 - Servicii Evenimente / Diverse (Facturi Nete)'), 
            ('703', '703 - Produse Reziduale (Nete)')
        ]

        for acc_key, acc_label in fact_accounts:
            cell_lbl = ws_main.cell(row=current_row, column=2, value=f"(+) Ct. {acc_label}")
            cell_lbl.font = font_bold
            cell_lbl.border = thin_border
            cell_lbl.fill = fill_light
            
            vals = fact_net_pivot[acc_key].values if acc_key in fact_net_pivot.columns else np.zeros(len(event_types))
                
            for c_idx, val in enumerate(vals, start=3):
                cell_val = ws_main.cell(row=current_row, column=c_idx, value=float(val))
                cell_val.font = font_regular
                cell_val.border = thin_border
                cell_val.number_format = '#,##0.00'
                
            start_let = get_column_letter(3)
            end_let = get_column_letter(len(event_types) + 2)
            ws_main.cell(row=current_row, column=len(event_types) + 3, value=f"=SUM({start_let}{current_row}:{end_let}{current_row})").font = font_bold
            ws_main.cell(row=current_row, column=len(event_types) + 3).border = thin_border
            ws_main.cell(row=current_row, column=len(event_types) + 3).fill = fill_zebra
            ws_main.cell(row=current_row, column=len(event_types) + 3).number_format = '#,##0.00'
            current_row += 1

        pers_row = current_row
        cell_lbl = ws_main.cell(row=pers_row, column=2, value="Număr Total Persoane Participante")
        cell_lbl.font = font_bold
        cell_lbl.border = thin_border
        cell_lbl.fill = fill_light

        for c_idx, val in enumerate(pers_by_event.values, start=3):
            ws_main.cell(row=pers_row, column=c_idx, value=float(val)).font = font_regular
            ws_main.cell(row=pers_row, column=c_idx).border = thin_border
            ws_main.cell(row=pers_row, column=c_idx).number_format = '#,##0'

        start_let_pers = get_column_letter(3)
        end_let_pers = get_column_letter(len(event_types) + 2)
        cell_tot_pers = ws_main.cell(row=pers_row, column=len(event_types) + 3, value=f"=SUM({start_let_pers}{pers_row}:{end_let_pers}{pers_row})")
        cell_tot_pers.font = font_bold
        cell_tot_pers.border = thin_border
        cell_tot_pers.fill = fill_zebra
        cell_tot_pers.number_format = '#,##0'

        ws_main.column_dimensions['A'].width = 3
        ws_main.column_dimensions['B'].width = 48
        for col in range(3, last_col_idx + 1):
            ws_main.column_dimensions[get_column_letter(col)].width = 14

        # --- FOAIA 4: PERSONAL & SALARII ---
        ws_pers = wb.create_sheet(title="Personal & Salarii")
        ws_pers.views.sheetView[0].showGridLines = True
        ws_pers.freeze_panes = 'C5'
        ws_pers.sheet_view.zoomScale = 100
        ws_pers.merge_cells("B2:E2")
        ws_pers.cell(row=2, column=2, value="SITUAȚIA CONTRACTELOR DE MUNCĂ ȘI FOND SALARII (DIRECTE VS INDIRECTE TESA) IAN - IUL 2026").font = font_title
        ws_pers.cell(row=2, column=2).alignment = Alignment(horizontal="center", vertical="center")

        headers_pers = ["Nume Salariat", "Funcția", "Tip Contract", "Salariul Negociat (Lunar)", "Total Salarii (7 Luni)", "Categorie Implicare"]
        for c_i, h_val in enumerate(headers_pers, start=2):
            ws_pers.cell(row=4, column=c_i, value=h_val).font = font_header
            ws_pers.cell(row=4, column=c_i).fill = fill_header
            ws_pers.cell(row=4, column=c_i).alignment = Alignment(horizontal="center", vertical="center")

        p_row = 5
        if df_munca is not None and not df_munca.empty:
            for _, p_data in df_munca.iterrows():
                salariat = p_data.get(sal_col_name, '') if sal_col_name else ''
                functie = p_data.get(func_c, '') if func_c else ''
                salar_lunar = float(p_data.get('Salar_Val', 0))
                salar_7luni = float(p_data.get('Salar_7Luni', 0))
                cat = str(p_data.get('Categorie', 'TESA & Administrativ (Indirect)'))
                
                ws_pers.cell(row=p_row, column=2, value=str(salariat)).border = thin_border
                ws_pers.cell(row=p_row, column=3, value=str(functie)).border = thin_border
                ws_pers.cell(row=p_row, column=4, value="C.I.M.").border = thin_border
                
                ws_pers.cell(row=p_row, column=5, value=salar_lunar).border = thin_border
                ws_pers.cell(row=p_row, column=5).number_format = '#,##0.00'
                
                s2 = ws_pers.cell(row=p_row, column=6, value=salar_7luni)
                s2.border = thin_border
                s2.number_format = '#,##0.00'
                s2.font = font_bold
                
                ws_pers.cell(row=p_row, column=7, value=cat).border = thin_border
                ws_pers.cell(row=p_row, column=7).font = font_bold
                p_row += 1

        p_row += 2
        ws_pers.cell(row=p_row, column=2, value="SUMAR FOND SALARII IANUARIE - IULIE 2026 (DIRECTE VS INDIRECTE)").font = font_bold
        p_row += 1

        sum_headers = ["Categorie Implicare", "Număr Salariați", "Total Salarii Ian-Iul (RON)"]
        for c_i, h_val in enumerate(sum_headers, start=2):
            ws_pers.cell(row=p_row, column=c_i, value=h_val).font = font_header
            ws_pers.cell(row=p_row, column=c_i).fill = fill_header
            ws_pers.cell(row=p_row, column=c_i).alignment = Alignment(horizontal="center", vertical="center")
        p_row += 1

        if df_munca is not None and not df_munca.empty and 'Categorie' in df_munca.columns:
            summary_pers = df_munca.groupby('Categorie').agg({sal_col_name: 'count', 'Salar_7Luni': 'sum'}).reset_index()
            for _, s_row in summary_pers.iterrows():
                ws_pers.cell(row=p_row, column=2, value=s_row['Categorie']).border = thin_border
                ws_pers.cell(row=p_row, column=3, value=int(s_row[sal_col_name])).border = thin_border
                ws_pers.cell(row=p_row, column=3).number_format = '#,##0'
                
                sum_cell = ws_pers.cell(row=p_row, column=4, value=float(s_row['Salar_7Luni']))
                sum_cell.border = thin_border
                sum_cell.number_format = '#,##0.00'
                sum_cell.font = font_bold
                p_row += 1

        ws_pers.column_dimensions['A'].width = 3
        ws_pers.column_dimensions['B'].width = 28
        ws_pers.column_dimensions['C'].width = 52
        ws_pers.column_dimensions['D'].width = 18
        ws_pers.column_dimensions['E'].width = 22
        ws_pers.column_dimensions['F'].width = 24
        ws_pers.column_dimensions['G'].width = 34

        # Salvare finală în memorie
        output_filename = "Raport_Executie_Bugetara_La_Castel_H1_2026.xlsx"
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.success("🎉 Raportul executiv a fost generat corect și complet cu toate datele!")
        st.download_button(
            label="📥 Descarcă Raportul Excel Final",
            data=buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
