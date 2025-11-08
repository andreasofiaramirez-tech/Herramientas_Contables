# utils.py

import pandas as pd
import numpy as np
import re
import xlsxwriter
from io import BytesIO
import streamlit as st # Necesario para los decoradores de caché

def mapear_columnas(df, log_messages):
    """
    Mapea sinónimos de columnas a nombres estandarizados para el procesamiento.
    Añade columnas faltantes con valor cero si no se encuentran.
    """
    DEBITO_SYNONYMS = ['debito', 'debitos', 'débito', 'débitos', 'debe']
    CREDITO_SYNONYMS = ['credito', 'creditos', 'crédito', 'créditos', 'haber']
    BS_SYNONYMS = ['ves', 'bolivar', 'bolívar', 'local', 'bs']
    USD_SYNONYMS = ['dolar', 'dólar', 'dólares', 'usd', 'dolares', 'me']

    REQUIRED_COLUMNS = {
        'Débito Bolivar': (DEBITO_SYNONYMS, BS_SYNONYMS),
        'Crédito Bolivar': (CREDITO_SYNONYMS, BS_SYNONYMS),
        'Débito Dolar': (DEBITO_SYNONYMS, USD_SYNONYMS),
        'Crédito Dolar': (CREDITO_SYNONYMS, USD_SYNONYMS)
    }

    column_mapping, current_cols = {}, [col.strip() for col in df.columns]

    for req_col, (type_synonyms, curr_synonyms) in REQUIRED_COLUMNS.items():
        found = False
        for input_col in current_cols:
            normalized_input = re.sub(r'[^\w]', '', input_col.lower())
            is_type_match = any(syn in normalized_input for syn in type_synonyms)
            is_curr_match = any(syn in normalized_input for syn in curr_synonyms)

            if is_type_match and is_curr_match and input_col not in column_mapping.values():
                column_mapping[input_col] = req_col
                found = True
                break
        if not found and req_col not in df.columns:
            log_messages.append(f"⚠️ ADVERTENCIA: No se encontró columna para '{req_col}'. Se creará vacía.")
            df[req_col] = 0.0

    df.rename(columns=column_mapping, inplace=True)
    return df

def limpiar_numero_avanzado(texto):
    """
    Limpia y convierte una cadena de texto a un formato numérico estándar (punto decimal),
    manejando formatos europeos (coma decimal) y americanos.
    """
    if texto is None or str(texto).strip().lower() == 'nan': return '0.0'
    texto_limpio = re.sub(r'[^\d.,-]', '', str(texto).strip())
    if not texto_limpio: return '0.0'

    num_puntos, num_comas = texto_limpio.count('.'), texto_limpio.count(',')

    if num_comas == 1 and num_puntos > 0:  # Formato europeo: 1.234,56 -> 1234.56
        return texto_limpio.replace('.', '').replace(',', '.')
    elif num_puntos == 1 and num_comas > 0:  # Formato americano con comas: 1,234.56 -> 1234.56
        return texto_limpio.replace(',', '')
    else:  # Otros casos (ej: 1234.56 o 1234,56)
        return texto_limpio.replace(',', '.')

@st.cache_data
def cargar_y_limpiar_datos(uploaded_actual, uploaded_anterior, log_messages):
    """
    Carga los archivos Excel, los limpia, estandariza columnas, calcula montos netos
    y prepara el DataFrame unificado para la conciliación.
    """
    def procesar_excel(archivo_buffer):
        try:
            archivo_buffer.seek(0)
            df = pd.read_excel(archivo_buffer, engine='openpyxl', dtype={'Asiento': str})
        except Exception as e:
            log_messages.append(f"❌ Error al leer el archivo Excel: {e}")
            return None

        for col in ['Fuente', 'Nombre del Proveedor']:
            if col not in df.columns:
                df[col] = ''

        df.columns = [col.strip() for col in df.columns]
        df = mapear_columnas(df, log_messages).copy()

        df['Asiento'] = df['Asiento'].astype(str).str.strip()
        df['Referencia'] = df['Referencia'].astype(str).str.strip()
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')

        for col in ['Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']:
            if col in df.columns:
                temp_serie = df[col].apply(limpiar_numero_avanzado)
                df[col] = pd.to_numeric(temp_serie, errors='coerce').fillna(0.0).round(2)
        return df

    df_actual = procesar_excel(uploaded_actual)
    df_anterior = procesar_excel(uploaded_anterior)

    if df_actual is None or df_anterior is None:
        st.error("❌ ¡Error Fatal! No se pudo procesar uno o ambos archivos Excel.")
        return None

    df_full = pd.concat([df_anterior, df_actual], ignore_index=True)
    key_cols = ['Asiento', 'Referencia', 'Fecha', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
    df_full.drop_duplicates(subset=[col for col in key_cols if col in df_full.columns], keep='first', inplace=True)

    df_full['Monto_BS'] = (df_full['Débito Bolivar'] - df_full['Crédito Bolivar']).round(2)
    df_full['Monto_USD'] = (df_full['Débito Dolar'] - df_full['Crédito Dolar']).round(2)
    df_full[['Conciliado', 'Grupo_Conciliado', 'Referencia_Normalizada_Literal']] = [False, np.nan, np.nan]

    log_messages.append(f"✅ Datos de Excel cargados. Total movimientos: {len(df_full)}")
    return df_full

@st.cache_data
def generar_reporte_excel(df_full, df_saldos_abiertos, df_conciliados, _estrategia_actual, casa_seleccionada, cuenta_seleccionada):
    """
    Genera el archivo Excel de reporte con múltiples hojas, encabezados dinámicos y formato profesional.
    """
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- Definición de Formatos ---
        formato_encabezado_empresa = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        formato_encabezado_sub = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
        formato_header_tabla = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        formato_bs = workbook.add_format({'num_format': '#,##0.00'})
        formato_usd = workbook.add_format({'num_format': '$#,##0.00'})
        formato_tasa = workbook.add_format({'num_format': '#,##0.0000'})
        formato_texto = workbook.add_format({'align': 'left'})
        formato_total_label = workbook.add_format({'bold': True, 'align': 'right'})
        formato_total_usd = workbook.add_format({'bold': True, 'num_format': '$#,##0.00'})
        formato_total_bs = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
        formato_proveedor_header = workbook.add_format({'bold': True, 'fg_color': '#F2F2F2', 'border': 1})
        formato_subtotal_label = workbook.add_format({'bold': True, 'align': 'right', 'top': 1})
        formato_subtotal_usd = workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'top': 1})
        formato_subtotal_bs = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})

        # --- Preparación del Encabezado de Fecha ---
        fecha_maxima = df_full['Fecha'].dropna().max()
        if pd.notna(fecha_maxima):
            ultimo_dia_mes = fecha_maxima + pd.offsets.MonthEnd(0)
            meses_es = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
            texto_fecha_encabezado = f"PARA EL {ultimo_dia_mes.day} DE {meses_es[ultimo_dia_mes.month].upper()} DE {ultimo_dia_mes.year}"
        else:
            texto_fecha_encabezado = "FECHA NO DISPONIBLE"

        # --- HOJA 1: SALDOS PENDIENTES ---
        df_reporte_pendientes_prep = df_saldos_abiertos.copy()
        columnas_reporte = _estrategia_actual["columnas_reporte"]
        
        if _estrategia_actual['id'] == 'devoluciones_proveedores':
            df_reporte_pendientes_prep['Monto USD'] = pd.to_numeric(df_saldos_abiertos['Monto_USD'], errors='coerce').fillna(0)
            df_reporte_pendientes_prep['Monto Bs'] = pd.to_numeric(df_saldos_abiertos['Monto_BS'], errors='coerce').fillna(0)
        else:
            df_reporte_pendientes_prep['Monto Dólar'] = pd.to_numeric(df_saldos_abiertos['Monto_USD'], errors='coerce').fillna(0)
            df_reporte_pendientes_prep['Bs.'] = pd.to_numeric(df_saldos_abiertos['Monto_BS'], errors='coerce').fillna(0)
            monto_dolar_abs = np.abs(df_reporte_pendientes_prep['Monto Dólar'])
            monto_bolivar_abs = np.abs(df_reporte_pendientes_prep['Bs.'])
            df_reporte_pendientes_prep['Tasa'] = np.where(monto_dolar_abs != 0, monto_bolivar_abs / monto_dolar_abs, 0)

        df_reporte_pendientes_final = df_reporte_pendientes_prep.reindex(columns=columnas_reporte).sort_values(by='Fecha')
        if 'Fecha' in df_reporte_pendientes_final.columns:
            df_reporte_pendientes_final['Fecha'] = pd.to_datetime(df_reporte_pendientes_final['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
        
        nombre_hoja_pendientes = _estrategia_actual.get("nombre_hoja_excel", "Pendientes")
        worksheet_pendientes = workbook.add_worksheet(nombre_hoja_pendientes)

        num_cols = len(df_reporte_pendientes_final.columns)
        if num_cols > 0:
            worksheet_pendientes.merge_range(0, 0, 0, num_cols - 1, casa_seleccionada, formato_encabezado_empresa)
            worksheet_pendientes.merge_range(1, 0, 1, num_cols - 1, f"ESPECIFICACION DE LA CUENTA {cuenta_seleccionada.split(' - ')[0]}", formato_encabezado_sub)
            worksheet_pendientes.merge_range(2, 0, 2, num_cols - 1, texto_fecha_encabezado, formato_encabezado_sub)
        
        header_row = 4
        for col_num, value in enumerate(df_reporte_pendientes_final.columns.values):
            worksheet_pendientes.write(header_row, col_num, value, formato_header_tabla)

        start_row_data = 5
        
        def get_col_idx(df, potential_names):
            for name in potential_names:
                if name in df.columns:
                    return df.columns.get_loc(name)
            return -1

        dolar_col_idx = get_col_idx(df_reporte_pendientes_final, ['Monto Dólar', 'Monto USD'])
        bs_col_idx = get_col_idx(df_reporte_pendientes_final, ['Bs.', 'Monto Bs'])
        tasa_col_idx = get_col_idx(df_reporte_pendientes_final, ['Tasa'])

        for row_idx, row_data in enumerate(df_reporte_pendientes_final.itertuples(index=False)):
            current_excel_row = start_row_data + row_idx
            for col_idx, cell_value in enumerate(row_data):
                if col_idx == dolar_col_idx:
                    worksheet_pendientes.write_number(current_excel_row, col_idx, cell_value, formato_usd)
                elif col_idx == bs_col_idx:
                    worksheet_pendientes.write_number(current_excel_row, col_idx, cell_value, formato_bs)
                elif col_idx == tasa_col_idx:
                    worksheet_pendientes.write_number(current_excel_row, col_idx, cell_value, formato_tasa)
                else:
                    worksheet_pendientes.write(current_excel_row, col_idx, cell_value)
        
        if _estrategia_actual['id'] == 'devoluciones_proveedores':
            worksheet_pendientes.set_column('A:A', 12); worksheet_pendientes.set_column('B:B', 15); worksheet_pendientes.set_column('C:C', 30)
            worksheet_pendientes.set_column('D:D', 40); worksheet_pendientes.set_column('E:E', 18); worksheet_pendientes.set_column('F:F', 18)
        else:
            worksheet_pendientes.set_column('A:A', 15); worksheet_pendientes.set_column('B:B', 50); worksheet_pendientes.set_column('C:C', 12)
            worksheet_pendientes.set_column('D:D', 18); worksheet_pendientes.set_column('E:E', 15); worksheet_pendientes.set_column('F:F', 18)

        worksheet_pendientes.freeze_panes(5, 0)

        if _estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar', 'devoluciones_proveedores'] and not df_reporte_pendientes_final.empty:
            total_row_index = start_row_data + len(df_reporte_pendientes_final)
            first_data_row_num = start_row_data + 1
            last_data_row_num = total_row_index
            
            label_col_idx = dolar_col_idx - 1 if dolar_col_idx != -1 else (bs_col_idx - 1 if bs_col_idx != -1 else -1)
            if label_col_idx >= 0:
                 worksheet_pendientes.write(total_row_index, label_col_idx, 'TOTAL', formato_total_label)

            if dolar_col_idx != -1:
                dolar_col_letter = xlsxwriter.utility.xl_col_to_name(dolar_col_idx)
                formula_usd = f'=SUM({dolar_col_letter}{first_data_row_num}:{dolar_col_letter}{last_data_row_num})'
                worksheet_pendientes.write_formula(total_row_index, dolar_col_idx, formula_usd, formato_total_usd)
            
            if bs_col_idx != -1:
                bs_col_letter = xlsxwriter.utility.xl_col_to_name(bs_col_idx)
                formula_bs = f'=SUM({bs_col_letter}{first_data_row_num}:{bs_col_letter}{last_data_row_num})'
                worksheet_pendientes.write_formula(total_row_index, bs_col_idx, formula_bs, formato_total_bs)

        # --- HOJA 2: DETALLE DE CONCILIACIÓN ---
        if not df_conciliados.empty:
            if _estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar']:
                df_reporte_conciliados_final = df_conciliados.copy()
                df_reporte_conciliados_final['Débitos Dólares'] = df_reporte_conciliados_final['Monto_USD'].apply(lambda x: x if x > 0 else 0)
                df_reporte_conciliados_final['Créditos Dólares'] = df_reporte_conciliados_final['Monto_USD'].apply(lambda x: x if x < 0 else 0)
                df_reporte_conciliados_final['Débitos Bs'] = df_reporte_conciliados_final['Monto_BS'].apply(lambda x: x if x > 0 else 0)
                df_reporte_conciliados_final['Créditos Bs'] = df_reporte_conciliados_final['Monto_BS'].apply(lambda x: x if x < 0 else 0)
                df_reporte_conciliados_final['Grupo de Conciliación'] = df_reporte_conciliados_final['Grupo_Conciliado']
                
                columnas_conciliacion = ['Fecha', 'Asiento', 'Referencia', 'Débitos Dólares', 'Créditos Dólares', 'Débitos Bs', 'Créditos Bs', 'Grupo de Conciliación']
                df_reporte_conciliados_final = df_reporte_conciliados_final[columnas_conciliacion].sort_values(by=['Grupo de Conciliación', 'Fecha'])
                df_reporte_conciliados_final['Fecha'] = pd.to_datetime(df_reporte_conciliados_final['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
                
                worksheet_conciliados = workbook.add_worksheet("Conciliacion")
                worksheet_conciliados.merge_range(0, 0, 0, len(columnas_conciliacion) - 1, 'Detalle de Movimientos Conciliados', formato_encabezado_sub)
                for col_num, value in enumerate(df_reporte_conciliados_final.columns.values):
                    worksheet_conciliados.write(1, col_num, value, formato_header_tabla)

                deb_usd_idx = df_reporte_conciliados_final.columns.get_loc('Débitos Dólares')
                cre_usd_idx = df_reporte_conciliados_final.columns.get_loc('Créditos Dólares')
                deb_bs_idx = df_reporte_conciliados_final.columns.get_loc('Débitos Bs')
                cre_bs_idx = df_reporte_conciliados_final.columns.get_loc('Créditos Bs')

                for r_idx, row in enumerate(df_reporte_conciliados_final.itertuples(index=False), start=2):
                    for c_idx, value in enumerate(row):
                        if c_idx == deb_usd_idx or c_idx == cre_usd_idx:
                            worksheet_conciliados.write_number(r_idx, c_idx, value, formato_usd)
                        elif c_idx == deb_bs_idx or c_idx == cre_bs_idx:
                            worksheet_conciliados.write_number(r_idx, c_idx, value, formato_bs)
                        else:
                            worksheet_conciliados.write(r_idx, c_idx, value)
                
                worksheet_conciliados.set_column('A:A', 12); worksheet_conciliados.set_column('B:B', 15); worksheet_conciliados.set_column('C:C', 30)
                worksheet_conciliados.set_column('D:E', 18); worksheet_conciliados.set_column('F:G', 18)
                worksheet_conciliados.set_column('H:H', 40)
                
                worksheet_conciliados.freeze_panes(2, 0)

                if not df_reporte_conciliados_final.empty:
                    total_row = 2 + len(df_reporte_conciliados_final)
                    diff_row = total_row + 1
                    
                    worksheet_conciliados.write(total_row, deb_usd_idx - 1, 'TOTALES', formato_total_label)
                    for col_idx, fmt in [(deb_usd_idx, formato_total_usd), (cre_usd_idx, formato_total_usd), (deb_bs_idx, formato_total_bs), (cre_bs_idx, formato_total_bs)]:
                        col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
                        formula = f'=SUM({col_letter}3:{col_letter}{total_row})'
                        worksheet_conciliados.write_formula(total_row, col_idx, formula, fmt)

                    worksheet_conciliados.write(diff_row, deb_usd_idx - 1, 'DIFERENCIA (Débitos + Créditos)', formato_total_label)
                    deb_usd_cell = xlsxwriter.utility.xl_rowcol_to_cell(total_row, deb_usd_idx)
                    cre_usd_cell = xlsxwriter.utility.xl_rowcol_to_cell(total_row, cre_usd_idx)
                    worksheet_conciliados.write_formula(diff_row, cre_usd_idx, f'={deb_usd_cell}+{cre_usd_cell}', formato_total_usd)
                    deb_bs_cell = xlsxwriter.utility.xl_rowcol_to_cell(total_row, deb_bs_idx)
                    cre_bs_cell = xlsxwriter.utility.xl_rowcol_to_cell(total_row, cre_bs_idx)
                    worksheet_conciliados.write_formula(diff_row, cre_bs_idx, f'={deb_bs_cell}+{cre_bs_cell}', formato_total_bs)
            
            elif _estrategia_actual['id'] == 'devoluciones_proveedores':
                df_conciliados_prep = df_conciliados.rename(columns={'Monto_USD': 'Monto Dólar', 'Monto_BS': 'Monto Bs.', 'Grupo_Conciliado': 'Grupo de Conciliación'})
                columnas_prov = ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor', 'Monto Dólar', 'Monto Bs.', 'Grupo de Conciliación']
                df_reporte_conciliados_final = df_conciliados_prep[columnas_prov].sort_values(by=['Grupo de Conciliación', 'Fecha'])
                df_reporte_conciliados_final['Fecha'] = pd.to_datetime(df_reporte_conciliados_final['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
                
                worksheet_conciliados = workbook.add_worksheet("Conciliacion")
                worksheet_conciliados.merge_range(0, 0, 0, len(columnas_prov) - 1, 'Detalle de Movimientos Conciliados', formato_encabezado_sub)
                for col_num, value in enumerate(df_reporte_conciliados_final.columns.values):
                    worksheet_conciliados.write(1, col_num, value, formato_header_tabla)
                
                dolar_idx_conc = df_reporte_conciliados_final.columns.get_loc('Monto Dólar')
                bs_idx_conc = df_reporte_conciliados_final.columns.get_loc('Monto Bs.')
                
                for r_idx, row in enumerate(df_reporte_conciliados_final.itertuples(index=False), start=2):
                    for c_idx, value in enumerate(row):
                        if c_idx == dolar_idx_conc:
                            worksheet_conciliados.write_number(r_idx, c_idx, value, formato_usd)
                        elif c_idx == bs_idx_conc:
                            worksheet_conciliados.write_number(r_idx, c_idx, value, formato_bs)
                        else:
                            worksheet_conciliados.write(r_idx, c_idx, value)
                
                if not df_reporte_conciliados_final.empty:
                    total_row_idx = 2 + len(df_reporte_conciliados_final)
                    first_data_r = 3
                    last_data_r = total_row_idx

                    worksheet_conciliados.write(total_row_idx, dolar_idx_conc - 1, 'TOTAL', formato_total_label)

                    dolar_col_letter = xlsxwriter.utility.xl_col_to_name(dolar_idx_conc)
                    formula_usd = f'=SUM({dolar_col_letter}{first_data_r}:{dolar_col_letter}{last_data_r})'
                    worksheet_conciliados.write_formula(total_row_idx, dolar_idx_conc, formula_usd, formato_total_usd)
                    
                    bs_col_letter = xlsxwriter.utility.xl_col_to_name(bs_idx_conc)
                    formula_bs = f'=SUM({bs_col_letter}{first_data_r}:{bs_col_letter}{last_data_r})'
                    worksheet_conciliados.write_formula(total_row_idx, bs_idx_conc, formula_bs, formato_total_bs)
                
                worksheet_conciliados.freeze_panes(2, 0)

                worksheet_conciliados.set_column('A:A', 12); worksheet_conciliados.set_column('B:B', 15); worksheet_conciliados.set_column('C:C', 30)
                worksheet_conciliados.set_column('D:D', 40); worksheet_conciliados.set_column('E:E', 18); worksheet_conciliados.set_column('F:F', 18)
                worksheet_conciliados.set_column('G:G', 40)

        # --- HOJA 3: RESUMEN POR PROVEEDOR ---
        if _estrategia_actual['id'] == 'devoluciones_proveedores' and not df_saldos_abiertos.empty:
            worksheet_prov = workbook.add_worksheet("Resumen por Proveedor")
            
            worksheet_prov.merge_range('A1:E1', 'Detalle de Saldos Abiertos por Proveedor', formato_encabezado_sub)
            columnas_detalle_prov = ['Fecha', 'Fuente', 'Referencia', 'Monto USD', 'Monto Bs']
            for col_num, value in enumerate(columnas_detalle_prov):
                worksheet_prov.write(2, col_num, value, formato_header_tabla)
            
            df_saldos_abiertos_sorted = df_saldos_abiertos.sort_values(by='Nombre del Proveedor')
            
            current_row = 3 
            
            for proveedor, grupo in df_saldos_abiertos_sorted.groupby('Nombre del Proveedor'):
                if not grupo.empty:
                    worksheet_prov.merge_range(current_row, 0, current_row, 4, f"Proveedor: {proveedor}", formato_proveedor_header)
                    current_row += 1
                    
                    for _, movimiento in grupo.iterrows():
                        worksheet_prov.write(current_row, 0, pd.to_datetime(movimiento['Fecha']).strftime('%d/%m/%Y'))
                        worksheet_prov.write(current_row, 1, movimiento['Fuente'])
                        worksheet_prov.write(current_row, 2, movimiento['Referencia'])
                        worksheet_prov.write_number(current_row, 3, movimiento['Monto_USD'], formato_usd)
                        worksheet_prov.write_number(current_row, 4, movimiento['Monto_BS'], formato_bs)
                        current_row += 1
                    
                    subtotal_usd = grupo['Monto_USD'].sum()
                    subtotal_bs = grupo['Monto_BS'].sum()
                    worksheet_prov.write(current_row, 2, f"Subtotal {proveedor}", formato_subtotal_label)
                    worksheet_prov.write_number(current_row, 3, subtotal_usd, formato_subtotal_usd)
                    worksheet_prov.write_number(current_row, 4, subtotal_bs, formato_subtotal_bs)
                    current_row += 2
            
            current_row += 1
            worksheet_prov.write(current_row, 2, 'TOTAL GENERAL', formato_total_label)
            worksheet_prov.write_number(current_row, 3, df_saldos_abiertos['Monto_USD'].sum(), formato_total_usd)
            worksheet_prov.write_number(current_row, 4, df_saldos_abiertos['Monto_BS'].sum(), formato_total_bs)

            worksheet_prov.freeze_panes(3, 0)

            worksheet_prov.set_column('A:A', 12); worksheet_prov.set_column('B:B', 20); worksheet_prov.set_column('C:C', 40)
            worksheet_prov.set_column('D:E', 18)
            
    return output_excel.getvalue()
    
@st.cache_data
def generar_csv_saldos_abiertos(df_saldos_abiertos):
    """
    Genera el archivo Excel con los saldos pendientes para el próximo ciclo de conciliación,
    formateando los números con coma decimal.
    """
    columnas_originales_csv = ['Asiento', 'Referencia', 'Fecha', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar', 'Fuente', 'Nombre del Proveedor']
    df_saldos_a_exportar = df_saldos_abiertos[[col for col in columnas_originales_csv if col in df_saldos_abiertos.columns]].copy()
    
    if 'Fecha' in df_saldos_a_exportar.columns:
        df_saldos_a_exportar['Fecha'] = pd.to_datetime(df_saldos_a_exportar['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
        
    for col in ['Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']:
        if col in df_saldos_a_exportar.columns:
            df_saldos_a_exportar[col] = df_saldos_a_exportar[col].round(2).apply(lambda x: f"{x:.2f}".replace('.', ','))
            
    return df_saldos_a_exportar.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
