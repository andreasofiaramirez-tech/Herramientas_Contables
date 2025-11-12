# utils.py

import pandas as pd
import numpy as np
import re
import xlsxwriter
from io import BytesIO
import streamlit as st

@st.cache_data
def cargar_y_limpiar_datos(uploaded_actual, uploaded_anterior, log_messages):
    """
    Carga los archivos Excel, los limpia, estandariza columnas, calcula montos netos
    y prepara el DataFrame unificado para la conciliación.
    """
    def mapear_columnas(df, log_messages):
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
        if texto is None or str(texto).strip().lower() == 'nan': return '0.0'
        texto_limpio = re.sub(r'[^\d.,-]', '', str(texto).strip())
        if not texto_limpio: return '0.0'

        num_puntos, num_comas = texto_limpio.count('.'), texto_limpio.count(',')

        if num_comas == 1 and num_puntos > 0:
            return texto_limpio.replace('.', '').replace(',', '.')
        elif num_puntos == 1 and num_comas > 0:
            return texto_limpio.replace(',', '')
        else:
            return texto_limpio.replace(',', '.')

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
def generar_reporte_excel(_df_full, df_saldos_abiertos, df_conciliados, _estrategia_actual, casa_seleccionada, cuenta_seleccionada):
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        workbook = writer.book

        formato_encabezado_empresa = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        formato_encabezado_sub = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
        formato_header_tabla = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        formato_bs = workbook.add_format({'num_format': '#,##0.00'})
        formato_usd = workbook.add_format({'num_format': '$#,##0.00'})
        formato_tasa = workbook.add_format({'num_format': '#,##0.0000'})
        formato_total_label = workbook.add_format({'bold': True, 'align': 'right'})
        formato_total_usd = workbook.add_format({'bold': True, 'num_format': '$#,##0.00'})
        formato_total_bs = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
        formato_proveedor_header = workbook.add_format({'bold': True, 'fg_color': '#F2F2F2', 'border': 1})
        formato_subtotal_label = workbook.add_format({'bold': True, 'align': 'right', 'top': 1})
        formato_subtotal_usd = workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'top': 1})
        formato_subtotal_bs = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})

        fecha_maxima = _df_full['Fecha'].dropna().max()
        if pd.notna(fecha_maxima):
            ultimo_dia_mes = fecha_maxima + pd.offsets.MonthEnd(0)
            meses_es = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
            texto_fecha_encabezado = f"PARA EL {ultimo_dia_mes.day} DE {meses_es[ultimo_dia_mes.month].upper()} DE {ultimo_dia_mes.year}"
        else:
            texto_fecha_encabezado = "FECHA NO DISPONIBLE"

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
        
        for col_num, value in enumerate(df_reporte_pendientes_final.columns.values):
            worksheet_pendientes.write(4, col_num, value, formato_header_tabla)

        start_row_data = 5
        
        def get_col_idx(df, potential_names):
            return next((df.columns.get_loc(name) for name in potential_names if name in df.columns), -1)

        dolar_col_idx = get_col_idx(df_reporte_pendientes_final, ['Monto Dólar', 'Monto USD'])
        bs_col_idx = get_col_idx(df_reporte_pendientes_final, ['Bs.', 'Monto Bs'])
        tasa_col_idx = get_col_idx(df_reporte_pendientes_final, ['Tasa'])

        for row_idx, row_data in enumerate(df_reporte_pendientes_final.itertuples(index=False), start=start_row_data):
            for col_idx, cell_value in enumerate(row_data):
                fmt = formato_usd if col_idx == dolar_col_idx else formato_bs if col_idx == bs_col_idx else formato_tasa if col_idx == tasa_col_idx else None
                if fmt:
                    worksheet_pendientes.write_number(row_idx, col_idx, cell_value, fmt)
                else:
                    worksheet_pendientes.write(row_idx, col_idx, cell_value)
        
        worksheet_pendientes.set_column('A:A', 15); worksheet_pendientes.set_column('B:B', 50); worksheet_pendientes.set_column('C:C', 12)
        worksheet_pendientes.set_column('D:D', 18); worksheet_pendientes.set_column('E:E', 15); worksheet_pendientes.set_column('F:F', 18)
        if _estrategia_actual['id'] == 'devoluciones_proveedores':
             worksheet_pendientes.set_column('D:D', 40);

        worksheet_pendientes.freeze_panes(5, 0)

        if not df_reporte_pendientes_final.empty:
            total_row_index = start_row_data + len(df_reporte_pendientes_final)
            label_col_idx = dolar_col_idx - 1 if dolar_col_idx > 0 else bs_col_idx -1 if bs_col_idx > 0 else -1
            if label_col_idx >= 0:
                 worksheet_pendientes.write(total_row_index, label_col_idx, 'TOTAL', formato_total_label)

            if dolar_col_idx != -1:
                formula_usd = f'=SUM({xlsxwriter.utility.xl_col_to_name(dolar_col_idx)}{start_row_data + 1}:{xlsxwriter.utility.xl_col_to_name(dolar_col_idx)}{total_row_index})'
                worksheet_pendientes.write_formula(total_row_index, dolar_col_idx, formula_usd, formato_total_usd)
            
            if bs_col_idx != -1:
                formula_bs = f'=SUM({xlsxwriter.utility.xl_col_to_name(bs_col_idx)}{start_row_data + 1}:{xlsxwriter.utility.xl_col_to_name(bs_col_idx)}{total_row_index})'
                worksheet_pendientes.write_formula(total_row_index, bs_col_idx, formula_bs, formato_total_bs)

        if not df_conciliados.empty:
            worksheet_conciliados = workbook.add_worksheet("Conciliacion")
            
            if _estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar', 'cuentas_viajes']:
                df_reporte_conciliados_final = df_conciliados.copy()
                df_reporte_conciliados_final['Débitos Dólares'] = df_reporte_conciliados_final['Monto_USD'].apply(lambda x: x if x > 0 else 0)
                df_reporte_conciliados_final['Créditos Dólares'] = df_reporte_conciliados_final['Monto_USD'].apply(lambda x: x if x < 0 else 0)
                df_reporte_conciliados_final['Débitos Bs'] = df_reporte_conciliados_final['Monto_BS'].apply(lambda x: x if x > 0 else 0)
                df_reporte_conciliados_final['Créditos Bs'] = df_reporte_conciliados_final['Monto_BS'].apply(lambda x: x if x < 0 else 0)
                df_reporte_conciliados_final['Grupo de Conciliación'] = df_reporte_conciliados_final['Grupo_Conciliado']
                
                columnas_conciliacion = ['Fecha', 'Asiento', 'Referencia', 'Débitos Dólares', 'Créditos Dólares', 'Débitos Bs', 'Créditos Bs', 'Grupo de Conciliación']
                df_reporte_conciliados_final = df_reporte_conciliados_final[columnas_conciliacion].sort_values(by=['Grupo de Conciliación', 'Fecha'])
                df_reporte_conciliados_final['Fecha'] = pd.to_datetime(df_reporte_conciliados_final['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
                
                worksheet_conciliados.merge_range(0, 0, 0, len(columnas_conciliacion) - 1, 'Detalle de Movimientos Conciliados', formato_encabezado_sub)
                for col_num, value in enumerate(df_reporte_conciliados_final.columns.values):
                    worksheet_conciliados.write(1, col_num, value, formato_header_tabla)

                deb_usd_idx, cre_usd_idx = get_col_idx(df_reporte_conciliados_final, ['Débitos Dólares']), get_col_idx(df_reporte_conciliados_final, ['Créditos Dólares'])
                deb_bs_idx, cre_bs_idx = get_col_idx(df_reporte_conciliados_final, ['Débitos Bs']), get_col_idx(df_reporte_conciliados_final, ['Créditos Bs'])

                for r_idx, row in enumerate(df_reporte_conciliados_final.itertuples(index=False), start=2):
                    for c_idx, value in enumerate(row):
                        fmt = formato_usd if c_idx in [deb_usd_idx, cre_usd_idx] else formato_bs if c_idx in [deb_bs_idx, cre_bs_idx] else None
                        if fmt: worksheet_conciliados.write_number(r_idx, c_idx, value, fmt)
                        else: worksheet_conciliados.write(r_idx, c_idx, value)
                
                worksheet_conciliados.set_column('A:A', 12); worksheet_conciliados.set_column('B:B', 15); worksheet_conciliados.set_column('C:C', 30)
                worksheet_conciliados.set_column('D:G', 18); worksheet_conciliados.set_column('H:H', 40)
                worksheet_conciliados.freeze_panes(2, 0)

            elif _estrategia_actual['id'] == 'devoluciones_proveedores':
                df_conciliados_prep = df_conciliados.rename(columns={'Monto_USD': 'Monto Dólar', 'Monto_BS': 'Monto Bs.', 'Grupo_Conciliado': 'Grupo de Conciliación'})
                columnas_prov = ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor', 'Monto Dólar', 'Monto Bs.', 'Grupo de Conciliación']
                df_reporte_conciliados_final = df_conciliados_prep.reindex(columns=columnas_prov).sort_values(by=['Grupo de Conciliación', 'Fecha'])
                df_reporte_conciliados_final['Fecha'] = pd.to_datetime(df_reporte_conciliados_final['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
                
                worksheet_conciliados.merge_range(0, 0, 0, len(columnas_prov) - 1, 'Detalle de Movimientos Conciliados', formato_encabezado_sub)
                for col_num, value in enumerate(df_reporte_conciliados_final.columns.values):
                    worksheet_conciliados.write(1, col_num, value, formato_header_tabla)
                
                dolar_idx_conc, bs_idx_conc = get_col_idx(df_reporte_conciliados_final, ['Monto Dólar']), get_col_idx(df_reporte_conciliados_final, ['Monto Bs.'])
                
                for r_idx, row in enumerate(df_reporte_conciliados_final.itertuples(index=False), start=2):
                    for c_idx, value in enumerate(row):
                        fmt = formato_usd if c_idx == dolar_idx_conc else formato_bs if c_idx == bs_idx_conc else None
                        if fmt: worksheet_conciliados.write_number(r_idx, c_idx, value, fmt)
                        else: worksheet_conciliados.write(r_idx, c_idx, value)
                
                worksheet_conciliados.set_column('A:A', 12); worksheet_conciliados.set_column('B:B', 15); worksheet_conciliados.set_column('C:C', 30)
                worksheet_conciliados.set_column('D:D', 40); worksheet_conciliados.set_column('E:G', 18)
                worksheet_conciliados.freeze_panes(2, 0)

        if _estrategia_actual['id'] == 'devoluciones_proveedores' and not df_saldos_abiertos.empty:
            worksheet_prov = workbook.add_worksheet("Resumen por Proveedor")
            
            worksheet_prov.merge_range('A1:E1', 'Detalle de Saldos Abiertos por Proveedor', formato_encabezado_sub)
            columnas_detalle_prov = ['Fecha', 'Fuente', 'Referencia', 'Monto USD', 'Monto Bs']
            worksheet_prov.write_row(2, 0, columnas_detalle_prov, formato_header_tabla)
            
            df_saldos_abiertos_sorted = df_saldos_abiertos.sort_values(by=['Nombre del Proveedor', 'Fecha'])
            
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
                    
                    subtotal_usd, subtotal_bs = grupo['Monto_USD'].sum(), grupo['Monto_BS'].sum()
                    worksheet_prov.write(current_row, 2, f"Subtotal {proveedor}", formato_subtotal_label)
                    worksheet_prov.write_number(current_row, 3, subtotal_usd, formato_subtotal_usd)
                    worksheet_prov.write_number(current_row, 4, subtotal_bs, formato_subtotal_bs)
                    current_row += 2
            
            worksheet_prov.set_column('A:A', 12); worksheet_prov.set_column('B:B', 20); worksheet_prov.set_column('C:C', 40)
            worksheet_prov.set_column('D:E', 18)
            worksheet_prov.freeze_panes(3, 0)
            
    return output_excel.getvalue()
    
@st.cache_data
def generar_csv_saldos_abiertos(df_saldos_abiertos):
    """
    Genera el archivo CSV con los saldos pendientes para el próximo ciclo de conciliación.
    """
    columnas_csv = ['Asiento', 'Referencia', 'Fecha', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar', 'Fuente', 'Nombre del Proveedor', 'NIT']
    df_saldos_a_exportar = df_saldos_abiertos.reindex(columns=columnas_csv).copy()
    
    if 'Fecha' in df_saldos_a_exportar.columns:
        df_saldos_a_exportar['Fecha'] = pd.to_datetime(df_saldos_a_exportar['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
        
    for col in ['Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']:
        if col in df_saldos_a_exportar.columns:
            df_saldos_a_exportar[col] = df_saldos_a_exportar[col].round(2).apply(lambda x: f"{x:.2f}".replace('.', ','))
            
    return df_saldos_a_exportar.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

# ==============================================================================
# NUEVA FUNCIÓN DE REPORTE PARA LA HERRAMIENTA DE RETENCIONES
# ==============================================================================

def generar_reporte_retenciones(df_cp_results, df_galac_no_cp, df_cg, cuentas_map):
    """
    Genera el archivo Excel de reporte final, utilizando la nueva columna unificada
    'Validacion CG' y la lógica de clasificación de incidencias/éxitos actualizada.
    """
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        # --- Formatos (sin cambios) ---
        main_title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14, 'locked': False})
        group_title_format = workbook.add_format({'bold': True, 'italic': True, 'font_size': 12, 'locked': False})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center', 'locked': False})
        money_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center', 'locked': False})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center', 'locked': False})
        center_text_format = workbook.add_format({'align': 'center', 'valign': 'top', 'locked': False})
        long_text_format = workbook.add_format({'align': 'left', 'valign': 'top', 'locked': False, 'text_wrap': True})

        # --- HOJA 1: Relacion CP ---
        ws1 = workbook.add_worksheet('Relacion CP')
        ws1.hide_gridlines(2)
        
        # --- 1. SE ACTUALIZA LA PLANTILLA DE COLUMNAS ---
        final_order_cp = [
            'Asiento', 'Tipo', 'Fecha', 'Numero', 'Aplicacion', 'Subtipo', 'Monto', 
            'Cp Vs Galac', 'Validacion CG', 'RIF', 'Nombre Proveedor'
        ]
        
        df_reporte_cp = df_cp_results.copy()
        
        # --- 2. SE ACTUALIZA EL RENOMBRADO DE COLUMNAS ---
        df_reporte_cp.rename(columns={
            'Comprobante': 'Numero', 
            'CP_Vs_Galac': 'Cp Vs Galac', 
            'Validacion_CG': 'Validacion CG' # Se alinea con el nuevo nombre de logic.py
        }, inplace=True)
        
        if 'Fecha' in df_reporte_cp.columns: df_reporte_cp['Fecha'] = pd.to_datetime(df_reporte_cp['Fecha'], errors='coerce')
        
        # Este bucle ahora funcionará correctamente con la nueva lista de columnas
        for col in final_order_cp:
            if col not in df_reporte_cp.columns: df_reporte_cp[col] = ''
        
        # --- 3. SE ACTUALIZA LA LÓGICA DE FILTRADO ---
        condicion_exitosa = (
            (df_reporte_cp['Cp Vs Galac'] == 'Sí') &
            (df_reporte_cp['Validacion CG'] == 'Conciliado en CG')
        )
        condicion_anulado = (df_reporte_cp['Cp Vs Galac'] == 'Anulado')
        
        df_exitosos = df_reporte_cp[condicion_exitosa].copy()
        df_anulados = df_reporte_cp[condicion_anulado].copy()
        
        indices_exitosos_y_anulados = df_exitosos.index.union(df_anulados.index)
        df_incidencias = df_reporte_cp.drop(indices_exitosos_y_anulados)
        
        # Se ajusta el rango del título a 11 columnas (A hasta K)
        ws1.merge_range('A1:K1', 'Relacion de Retenciones CP', main_title_format)
        current_row = 2
        
        # --- Escritura de Incidencias (sin cambios en la lógica interna) ---
        ws1.write(current_row, 0, 'Incidencias Encontradas', group_title_format); current_row += 1
        ws1.write_row(current_row, 0, final_order_cp, header_format); current_row += 1
        if not df_incidencias.empty:
            for index, row in df_incidencias.iterrows():
                for col_idx, col_name in enumerate(final_order_cp):
                    value = row[col_name]
                    # Aplicamos formato especial a la nueva columna 'Validacion CG'
                    if col_name == 'Fecha' and pd.notna(value):
                        ws1.write_datetime(current_row, col_idx, value, date_format)
                    elif col_name == 'Monto':
                        ws1.write_number(current_row, col_idx, value, money_format)
                    elif col_name in ['Cp Vs Galac', 'Validacion CG'] and pd.notna(value):
                        ws1.write(current_row, col_idx, value, long_text_format)
                    elif pd.notna(value):
                        ws1.write(current_row, col_idx, value, center_text_format)
                current_row += 1
        
        current_row += 1
        
        # --- Escritura de Conciliaciones Exitosas (sin cambios en la lógica interna) ---
        ws1.write(current_row, 0, 'Conciliacion Exitosa', group_title_format); current_row += 1
        ws1.write_row(current_row, 0, final_order_cp, header_format); current_row += 1
        if not df_exitosos.empty:
            for index, row in df_exitosos.iterrows():
                for col_idx, col_name in enumerate(final_order_cp):
                    value = row[col_name]
                    if col_name == 'Fecha' and pd.notna(value):
                        ws1.write_datetime(current_row, col_idx, value, date_format)
                    elif col_name == 'Monto':
                        ws1.write_number(current_row, col_idx, value, money_format)
                    elif col_name in ['Cp Vs Galac', 'Validacion CG'] and pd.notna(value):
                         ws1.write(current_row, col_idx, value, long_text_format)
                    elif pd.notna(value):
                        ws1.write(current_row, col_idx, value, center_text_format)
                current_row += 1

        current_row += 1

        # --- Escritura de Anulados ---
        ws1.write(current_row, 0, 'Registros Anulados', group_title_format); current_row += 1
        ws1.write_row(current_row, 0, final_order_cp, header_format); current_row += 1
        if not df_anulados.empty:
            for index, row in df_anulados.iterrows():
                for col_idx, col_name in enumerate(final_order_cp):
                    value = row[col_name]
                    if col_name == 'Fecha' and pd.notna(value):
                        ws1.write_datetime(current_row, col_idx, value, date_format)
                    elif col_name == 'Monto':
                        ws1.write_number(current_row, col_idx, value, money_format)
                    elif col_name == 'Cp Vs Galac' and pd.notna(value):
                         ws1.write(current_row, col_idx, value, long_text_format)
                    elif pd.notna(value):
                        ws1.write(current_row, col_idx, value, center_text_format)
                current_row += 1

        # --- Bloque de autoajuste de ANCHO ---
        for i, col_name in enumerate(final_order_cp):
            column_data = df_reporte_cp[col_name].astype(str)
            max_data_len = column_data.map(len).max() if not column_data.empty else 0
            header_len = len(col_name)
            column_width = max(header_len, max_data_len) + 2
            column_width = min(column_width, 255)
            ws1.set_column(i, i, column_width)
            
        # --- HOJA 2: Análisis GALAC ---
        ws2 = workbook.add_worksheet('Análisis GALAC')
        ws2.hide_gridlines(2)
        ws2.merge_range('A1:H1', 'Análisis de Retenciones Oficiales (GALAC)', main_title_format)

        # 1. Preparar DFs de GALAC
        df_iva['Tipo'] = 'IVA'
        df_islr['Tipo'] = 'ISLR'
        df_municipal['Tipo'] = 'Municipal'

        # --- Lógica de Merge Separada ---
        df_iva_islr = pd.concat([df_iva, df_islr], ignore_index=True)

        # Merge para IVA e ISLR (usan las 3 claves)
        merge_keys_iva_islr = ['RIF_norm', 'Comprobante_norm', 'Factura_norm']
        df_merged_iva_islr = pd.merge(
            df_iva_islr,
            df_cp_results[['RIF_norm', 'Comprobante_norm', 'Factura_norm', 'Cp Vs Galac', 'Validacion CG']],
            on=merge_keys_iva_islr,
            how='left'
        )

        # Merge para Municipal (usan solo 2 claves)
        merge_keys_municipal = ['RIF_norm', 'Factura_norm']
        df_merged_municipal = pd.merge(
            df_municipal,
            df_cp_results[['RIF_norm', 'Factura_norm', 'Cp Vs Galac', 'Validacion CG']],
            on=merge_keys_municipal,
            how='left'
        )

        # Volvemos a unir los resultados
        df_merged_galac = pd.concat([df_merged_iva_islr, df_merged_municipal], ignore_index=True)

        # 3. Crear los tres grupos principales
        cond_exitosa_galac = (df_merged_galac['Cp Vs Galac'] == 'Sí') & (df_merged_galac['Validacion CG'] == 'Conciliado en CG')
        cond_no_encontrada_en_cp = df_merged_galac['Cp Vs Galac'].isna()

        df_exitosos_galac = df_merged_galac[cond_exitosa_galac].copy()
        df_no_encontrados_galac = df_merged_galac[cond_no_encontrada_en_cp].copy()
        
        indices_a_excluir = df_exitosos_galac.index.union(df_no_encontrados_galac.index)
        df_incidencias_galac = df_merged_galac.drop(indices_a_excluir)
        # Creamos la columna de error combinando los resultados
        df_incidencias_galac['Detalle del Error'] = df_incidencias_galac['Cp Vs Galac'].fillna('') + " | " + df_incidencias_galac['Validacion CG'].fillna('')

        # 4. Función Auxiliar para escribir los grupos en el Excel
        def escribir_grupo_galac(ws, title, df_grupo, headers, start_row):
            ws.write(start_row, 0, title, group_title_format)
            start_row += 1
            for tipo in ['IVA', 'ISLR', 'Municipal']:
                df_subgrupo = df_grupo[df_grupo['Tipo'] == tipo]
                if not df_subgrupo.empty:
                    ws.write(start_row, 1, f"Retenciones de {tipo}", group_title_format)
                    start_row += 1
                    ws.write_row(start_row, 0, headers, header_format)
                    start_row += 1
                    for index, row in df_subgrupo.iterrows():
                        # Aseguramos que las columnas existan antes de intentar acceder a ellas
                        fecha = row.get('Fecha', '')
                        comprobante = row.get('Comprobante', '')
                        factura = row.get('Factura', '')
                        rif = row.get('RIF', '')
                        proveedor = row.get('Nombre Proveedor', '')
                        monto = row.get('Monto', 0)
                        error = row.get('Detalle del Error', '')
                        
                        data_to_write = [fecha, comprobante, factura, rif, proveedor, monto]
                        if 'Detalle del Error' in headers: data_to_write.append(error)
                        
                        ws.write_row(start_row, 0, data_to_write)
                        start_row += 1
                    start_row += 1
            return start_row

        # 5. Escribir los grupos en la hoja
        current_row = 2
        headers_incidencias = ['Fecha', 'Comprobante', 'Nro Documento', 'RIF', 'Nombre Proveedor', 'Monto', 'Detalle del Error']
        headers_general = ['Fecha', 'Comprobante', 'Nro Documento', 'RIF', 'Nombre Proveedor', 'Monto']

        current_row = escribir_grupo_galac(ws2, "A. Incidencias de CP Reflejadas en GALAC", df_incidencias_galac, headers_incidencias, current_row)
        current_row += 2
        current_row = escribir_grupo_galac(ws2, "B. Retenciones en GALAC no encontradas en Relacion de CP", df_no_encontrados_galac, headers_general, current_row)
        current_row += 2
        current_row = escribir_grupo_galac(ws2, "C. Retenciones Conciliadas Correctamente", df_exitosos_galac, headers_general, current_row)
        
        # Formato de columnas para la Hoja 2
        ws2.set_column('A:A', 12, date_format)
        ws2.set_column('B:D', 20, center_text_format)
        ws2.set_column('E:E', 35)
        ws2.set_column('F:F', 18, money_format)
        ws2.set_column('G:G', 50, long_text_format) # Columna de errores


        # --- HOJA 2: Análisis GALAC ---
        ws3 = workbook.add_worksheet('Diario CG')
        ws3.hide_gridlines(2)
        ws3.merge_range('A1:I1', 'Asientos con Errores de Conciliación', main_title_format)
        
        # Las columnas originales de CG no cambian
        cg_original_cols = [c for c in ['ASIENTO', 'FUENTE', 'CUENTACONTABLE', 'DESCRIPCIONDELACUENTACONTABLE', 'REFERENCIA', 'DEBITOVES', 'CREDITOVES', 'RIF', 'NIT'] if c in df_cg.columns]
        cg_headers_final = cg_original_cols + ['Observacion']
        
        # La forma de encontrar los asientos con error no cambia
        asientos_con_error = df_incidencias['Asiento'].unique()
        df_cg_errores = df_cg[df_cg['ASIENTO'].isin(asientos_con_error)].copy()
        
        df_cg_errores.rename(columns={'ASIENTO': 'Asiento'}, inplace=True)

        df_error_cuenta = pd.DataFrame(columns=cg_headers_final)
        df_error_monto = pd.DataFrame(columns=cg_headers_final)
        
        if not df_incidencias.empty and not df_cg_errores.empty:
            # --- 1. SE MODIFICA EL MERGE PARA USAR LA NUEVA COLUMNA ---
            merged_errors = pd.merge(
                df_cg_errores, 
                df_incidencias[['Asiento', 'Validacion CG']], # Usamos 'Validacion CG'
                on='Asiento', 
                how='left'
            )
            
            merged_errors.rename(columns={'Asiento': 'ASIENTO'}, inplace=True)

            # --- 2. SE MODIFICA LA LÓGICA DE CLASIFICACIÓN DE ERRORES ---
            # Ahora buscamos los textos de error específicos dentro de la columna 'Validacion CG'
            conditions = [
                merged_errors['Validacion CG'].str.contains('Cuenta Contable no coincide', na=False),
                merged_errors['Validacion CG'].str.contains('Monto no coincide', na=False)
            ]
            choices = [
                'Cuenta Contable no corresponde al Subtipo', 
                'Monto en Diario no coincide con Relacion CP'
            ]
            merged_errors['Observacion'] = np.select(conditions, choices, default='Error de CG no clasificado')
            
            df_cg_final = merged_errors[cg_headers_final].drop_duplicates()
            df_error_cuenta = df_cg_final[df_cg_final['Observacion'] == 'Cuenta Contable no corresponde al Subtipo']
            df_error_monto = df_cg_final[df_cg_final['Observacion'] == 'Monto en Diario no coincide con Relacion CP']
        
        current_row = 2
        ws3.write(current_row, 0, 'INCIDENCIA: Cuenta Contable Incorrecta', group_title_format); current_row += 1
        ws3.write_row(current_row, 0, cg_headers_final, header_format); current_row += 1
        if not df_error_cuenta.empty:
             for r_idx, row in df_error_cuenta[cg_headers_final].iterrows():
                ws3.write_row(current_row, 0, row.fillna('').values); current_row += 1
        current_row += 1
        ws3.write(current_row, 0, 'INCIDENCIA: Monto del Diario vs. Relación CP', group_title_format); current_row += 1
        ws3.write_row(current_row, 0, cg_headers_final, header_format); current_row += 1
        if not df_error_monto.empty:
            for r_idx, row in df_error_monto[cg_headers_final].iterrows():
                ws3.write_row(current_row, 0, row.fillna('').values); current_row += 1
        ws3.set_column('A:E', 20); ws3.set_column('F:G', 15, money_format); ws3.set_column('H:H', 20); ws3.set_column('I:I', 40)

    return output_buffer.getvalue()
