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

         COLUMN_STANDARDIZATION_MAP = {
            'Asiento': ['ASIENTO', 'Asiento'],
            'Fuente': ['FUENTE', 'Fuente'],
            'Fecha': ['FECHA', 'Fecha'],
            'Referencia': ['REFERENCIA', 'Referencia'],
            'NIT': ['Nit', 'NIT', 'Rif', 'RIF'],
            'Descripcion NIT': ['Descripción Nit', 'Descripcion Nit', 'DESCRIPCION NIT', 'Descripción NIT'],
            'Nombre del Proveedor': ['Nombre del Proveedor', 'NOMBRE DEL PROVEEDOR', 'Nombre Proveedor']
        }

        rename_map = {}
        for standard_name, variations in COLUMN_STANDARDIZATION_MAP.items():
            for var in variations:
                if var in df.columns:
                    # Renombrar la variación encontrada al nombre estándar
                    rename_map[var] = standard_name
                    break # Pasar al siguiente nombre estándar

        if rename_map:
            df.rename(columns=rename_map, inplace=True)
            log_messages.append(f"✔️ Nombres de columna estandarizados. Mapeo aplicado: {rename_map}")
            
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
            df_reporte_pendientes_final = df_reporte_pendientes_final.fillna('')
        
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
            
            if _estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar', 'cuentas_viajes', 'deudores_empleados_me']:
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

            elif _estrategia_actual['id'] == 'cobros_viajeros':
                worksheet_conciliados.merge_range('A1:H1', 'Detalle de Movimientos Conciliados por Viajero (NIT)', formato_encabezado_sub)
                
                df_conciliados_prep = df_conciliados.copy()
                df_conciliados_prep['Débitos Dólares'] = df_conciliados_prep['Monto_USD'].apply(lambda x: x if x > 0 else 0)
                df_conciliados_prep['Créditos Dólares'] = df_conciliados_prep['Monto_USD'].apply(lambda x: abs(x) if x < 0 else 0)
                
                df_conciliados_final = df_conciliados_prep.sort_values(by=['NIT', 'Grupo_Conciliado', 'Fecha'])
                
                columnas_detalle_viajeros = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Débitos Dólares', 'Créditos Dólares']
                current_row = 2

                for nit, grupo_viajero in df_conciliados_final.groupby('NIT'):
                    nombre_viajero = grupo_viajero['Descripcion NIT'].iloc[0] if not grupo_viajero.empty else ''
                    
                    # Escribir encabezado del viajero/NIT
                    worksheet_conciliados.merge_range(current_row, 0, current_row, len(columnas_detalle_viajeros) - 1, f"NIT: {nit} - {nombre_viajero}", formato_proveedor_header)
                    current_row += 1
                    
                    # Escribir encabezados de la tabla
                    worksheet_conciliados.write_row(current_row, 0, columnas_detalle_viajeros, formato_header_tabla)
                    current_row += 1
                    
                    # Escribir filas de datos
                    for _, movimiento in grupo_viajero.iterrows():
                        worksheet_conciliados.write_datetime(current_row, 0, movimiento['Fecha'], date_format)
                        worksheet_conciliados.write(current_row, 1, movimiento['Asiento'])
                        worksheet_conciliados.write(current_row, 2, movimiento['Referencia'])
                        worksheet_conciliados.write(current_row, 3, movimiento['Fuente'])
                        worksheet_conciliados.write_number(current_row, 4, movimiento['Débitos Dólares'], formato_usd)
                        worksheet_conciliados.write_number(current_row, 5, movimiento['Créditos Dólares'], formato_usd)
                        current_row += 1
                    
                    # Escribir fila de subtotal
                    subtotal_deb = grupo_viajero['Débitos Dólares'].sum()
                    subtotal_cre = grupo_viajero['Créditos Dólares'].sum()
                    worksheet_conciliados.write(current_row, 3, "Subtotal NIT", formato_subtotal_label)
                    worksheet_conciliados.write_number(current_row, 4, subtotal_deb, formato_subtotal_usd)
                    worksheet_conciliados.write_number(current_row, 5, subtotal_cre, formato_subtotal_usd)
                    current_row += 2 # Espacio entre grupos

                worksheet_conciliados.set_column('A:A', 12); worksheet_conciliados.set_column('B:B', 15)
                worksheet_conciliados.set_column('C:D', 30); worksheet_conciliados.set_column('E:F', 18)

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
# REPORTE PARA LA HERRAMIENTA DE RETENCIONES
# ==============================================================================

def generar_reporte_retenciones(df_cp_results, df_galac_no_cp, df_cg, cuentas_map):
    """
    Genera el archivo Excel de reporte final, con formato y lógica actualizados.
    - Hoja 1: 'Relacion CP' con columna de validación de CG unificada.
    - Hoja 2: Eliminada.
    - Hoja 3: 'Diario CG' con título centrado y columnas autoajustadas.
    """
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        # --- Formatos ---
        main_title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14, 'locked': False})
        group_title_format = workbook.add_format({'bold': True, 'italic': True, 'font_size': 12, 'locked': False})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center', 'locked': False})
        money_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center', 'locked': False})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center', 'locked': False})
        center_text_format = workbook.add_format({'align': 'center', 'valign': 'top', 'locked': False})
        long_text_format = workbook.add_format({'align': 'left', 'valign': 'top', 'locked': False, 'text_wrap': True})

        # --- PREPARACIÓN DE DATOS ---
        df_reporte_cp = df_cp_results.copy()
        df_reporte_cp.rename(columns={'Comprobante': 'Numero', 'CP_Vs_Galac': 'Cp Vs Galac', 'Validacion_CG': 'Validacion CG'}, inplace=True)
        if 'Fecha' in df_reporte_cp.columns: df_reporte_cp['Fecha'] = pd.to_datetime(df_reporte_cp['Fecha'], errors='coerce')

        # --- HOJA 1: Relacion CP ---
        ws1 = workbook.add_worksheet('Relacion CP')
        ws1.hide_gridlines(2)
        
        final_order_cp = [
            'Asiento', 'Tipo', 'Fecha', 'Numero', 'Aplicacion', 'Subtipo', 'Monto', 
            'Cp Vs Galac', 'Validacion CG', 'RIF', 'Nombre Proveedor'
        ]
        
        for col in final_order_cp:
            if col not in df_reporte_cp.columns: df_reporte_cp[col] = ''
        
        condicion_exitosa = ((df_reporte_cp['Cp Vs Galac'] == 'Sí') & (df_reporte_cp['Validacion CG'] == 'Conciliado en CG'))
        condicion_anulado = (df_reporte_cp['Cp Vs Galac'] == 'Anulado')
        df_exitosos = df_reporte_cp[condicion_exitosa].copy()
        df_anulados = df_reporte_cp[condicion_anulado].copy()
        indices_exitosos_y_anulados = df_exitosos.index.union(df_anulados.index)
        df_incidencias = df_reporte_cp.drop(indices_exitosos_y_anulados)
        
        ws1.merge_range('A1:K1', 'Relacion de Retenciones CP', main_title_format)
        current_row = 2
        
        # Escritura de Incidencias
        ws1.write(current_row, 0, 'Incidencias Encontradas', group_title_format); current_row += 1
        ws1.write_row(current_row, 0, final_order_cp, header_format); current_row += 1
        if not df_incidencias.empty:
            for index, row in df_incidencias.iterrows():
                for col_idx, col_name in enumerate(final_order_cp):
                    value = row[col_name]
                    if col_name == 'Fecha' and pd.notna(value): ws1.write_datetime(current_row, col_idx, value, date_format)
                    elif col_name == 'Monto': ws1.write_number(current_row, col_idx, value, money_format)
                    elif col_name in ['Cp Vs Galac', 'Validacion CG'] and pd.notna(value): ws1.write(current_row, col_idx, value, long_text_format)
                    elif pd.notna(value): ws1.write(current_row, col_idx, value, center_text_format)
                current_row += 1
        current_row += 1
        
        # Escritura de Conciliaciones Exitosas
        ws1.write(current_row, 0, 'Conciliacion Exitosa', group_title_format); current_row += 1
        ws1.write_row(current_row, 0, final_order_cp, header_format); current_row += 1
        if not df_exitosos.empty:
            for index, row in df_exitosos.iterrows():
                for col_idx, col_name in enumerate(final_order_cp):
                    value = row[col_name]
                    if col_name == 'Fecha' and pd.notna(value): ws1.write_datetime(current_row, col_idx, value, date_format)
                    elif col_name == 'Monto': ws1.write_number(current_row, col_idx, value, money_format)
                    elif col_name in ['Cp Vs Galac', 'Validacion CG'] and pd.notna(value): ws1.write(current_row, col_idx, value, long_text_format)
                    elif pd.notna(value): ws1.write(current_row, col_idx, value, center_text_format)
                current_row += 1
        current_row += 1

        # Escritura de Anulados
        ws1.write(current_row, 0, 'Registros Anulados', group_title_format); current_row += 1
        ws1.write_row(current_row, 0, final_order_cp, header_format); current_row += 1
        if not df_anulados.empty:
            for index, row in df_anulados.iterrows():
                for col_idx, col_name in enumerate(final_order_cp):
                    value = row[col_name]
                    if col_name == 'Fecha' and pd.notna(value): ws1.write_datetime(current_row, col_idx, value, date_format)
                    elif col_name == 'Monto': ws1.write_number(current_row, col_idx, value, money_format)
                    elif col_name in ['Cp Vs Galac', 'Validacion CG'] and pd.notna(value): ws1.write(current_row, col_idx, value, long_text_format)
                    elif pd.notna(value): ws1.write(current_row, col_idx, value, center_text_format)
                current_row += 1

        # Bloque de autoajuste de ANCHO para Hoja 1
        for i, col_name in enumerate(final_order_cp):
            column_data = df_reporte_cp[col_name].astype(str)
            max_data_len = column_data.map(len).max() if not column_data.empty else 0
            header_len = len(col_name)
            column_width = max(header_len, max_data_len) + 2
            column_width = min(column_width, 50)
            ws1.set_column(i, i, column_width)

        # --- HOJA 3: Diario CG ---
        ws3 = workbook.add_worksheet('Diario CG')
        ws3.hide_gridlines(2)
        # 1. Título Centrado
        ws3.merge_range('A1:I1', 'Asientos con Errores de Conciliación', main_title_format)
        
        cg_original_cols = [c for c in ['ASIENTO', 'FUENTE', 'CUENTACONTABLE', 'DESCRIPCIONDELACUENTACONTABLE', 'REFERENCIA', 'DEBITOVES', 'CREDITOVES', 'RIF', 'NIT'] if c in df_cg.columns]
        cg_headers_final = cg_original_cols + ['Observacion']
        asientos_con_error = df_incidencias['Asiento'].unique()
        df_cg_errores = df_cg[df_cg['ASIENTO'].isin(asientos_con_error)].copy()
        
        df_cg_errores.rename(columns={'ASIENTO': 'Asiento'}, inplace=True)

        df_error_cuenta = pd.DataFrame(columns=cg_headers_final)
        df_error_monto = pd.DataFrame(columns=cg_headers_final)
        
        if not df_incidencias.empty and not df_cg_errores.empty:
            merged_errors = pd.merge(df_cg_errores, df_incidencias[['Asiento', 'Validacion CG']], on='Asiento', how='left')
            merged_errors.rename(columns={'Asiento': 'ASIENTO'}, inplace=True)
            conditions = [merged_errors['Validacion CG'].str.contains('Cuenta Contable no coincide', na=False), merged_errors['Validacion CG'].str.contains('Monto no coincide', na=False)]
            choices = ['Cuenta Contable no corresponde al Subtipo', 'Monto en Diario no coincide con Relacion CP']
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
        
        # 2. Bloque de autoajuste de ANCHO para Hoja 3
        df_cg_final_para_ancho = pd.concat([df_error_cuenta, df_error_monto])
        for i, col_name in enumerate(cg_headers_final):
            if col_name in df_cg_final_para_ancho.columns:
                column_data = df_cg_final_para_ancho[col_name].astype(str)
                max_data_len = column_data.map(len).max() if not column_data.empty else 0
                header_len = len(col_name)
                column_width = max(header_len, max_data_len) + 2
                column_width = min(column_width, 60)
                ws3.set_column(i, i, column_width)

    return output_buffer.getvalue()

# ==============================================================================
# FUNCIÓN DE REPORTE PARA ANÁLISIS DE PAQUETE CC
# ==============================================================================

def generar_reporte_paquete_cc(df_analizado):
    """
    Versión final del reporte de Análisis de Paquete CC.
    - Ordena las hojas numéricamente.
    - Añade un título descriptivo y coloreado en cada hoja.
    """
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- Formatos ---
        main_title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 16})
        # CORRECCIÓN: Nuevo formato para el título descriptivo como en la imagen
        descriptive_title_format = workbook.add_format({'bold': True, 'font_size': 14, 'fg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        subgroup_title_format = workbook.add_format({'bold': True, 'font_size': 11, 'fg_color': '#E0E0E0', 'border': 1})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        money_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        text_format = workbook.add_format({'border': 1})
        total_label_format = workbook.add_format({'bold': True, 'align': 'right', 'top': 2, 'font_color': '#003366'})
        total_money_format = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 2, 'bottom': 1})
        
        columnas_reporte = ['Asiento', 'Fecha', 'Fuente', 'Cuenta Contable', 'Descripción de Cuenta', 'Referencia', 'Débito Dolar', 'Crédito Dolar', 'Débito VES', 'Crédito VES']
        
        df_analizado['Grupo Principal'] = df_analizado['Grupo'].apply(lambda x: x.split(':')[0].strip())
        
        # --- CORRECCIÓN: Ordenamiento numérico ---
        def sort_key(group_name):
            if group_name.startswith('Grupo'):
                return (0, int(group_name.split()[1]))
            return (1, group_name) # Pone "No Clasificado" al final
            
        grupos_principales_ordenados = sorted(df_analizado['Grupo Principal'].unique(), key=sort_key)
        
        # --- PASO 1: Crear la hoja "Directorio" ---
        ws_dir = workbook.add_worksheet("Directorio")
        # ... (El código del directorio no cambia)
        ws_dir.merge_range('A1:B1', 'Directorio de Grupos', main_title_format)
        ws_dir.write('A2', 'Nombre de la Hoja', header_format)
        ws_dir.write('B2', 'Descripción del Contenido', header_format)
        dir_row = 2
        for grupo_principal in grupos_principales_ordenados:
            sheet_name = re.sub(r'[\\/*?:"\[\]]', '', grupo_principal)[:31]
            full_name_example = df_analizado[df_analizado['Grupo Principal'] == grupo_principal]['Grupo'].iloc[0]
            description = full_name_example.split(':', 1)[-1].strip() if ':' in full_name_example else full_name_example
            if grupo_principal in ["Grupo 3", "Grupo 9"]:
                description = f"{description.split('-')[0].strip()} (Varios Subgrupos)"
            ws_dir.write(dir_row, 0, sheet_name, text_format)
            ws_dir.write(dir_row, 1, description, text_format)
            dir_row += 1
        ws_dir.set_column('A:A', 25); ws_dir.set_column('B:B', 60)
        
        # --- PASO 2: Crear una hoja por cada grupo principal ---
        for grupo_principal_nombre in grupos_principales_ordenados:
            sheet_name = re.sub(r'[\\/*?:"\[\]]', '', grupo_principal_nombre)[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.hide_gridlines(2)
            
            ws.merge_range('A1:J1', 'Análisis de Asientos de Cuentas por Cobrar', main_title_format)
            
            df_grupo_completo = df_analizado[df_analizado['Grupo Principal'] == grupo_principal_nombre]
            subgrupos = sorted(df_grupo_completo['Grupo'].unique())
            
            # --- CORRECCIÓN: Usar el nombre completo del grupo para el título ---
            full_descriptive_title = subgrupos[0]
            if len(subgrupos) > 1: # Si es un grupo con sub-grupos
                 full_descriptive_title = f"{subgrupos[0].split(':')[0].strip()}: {subgrupos[0].split(':')[1].split('-')[0].strip()}"
            
            ws.merge_range('A3:J3', full_descriptive_title, descriptive_title_format)
            current_row = 4
            
            for subgrupo_nombre in subgrupos:
                df_subgrupo = df_grupo_completo[df_grupo_completo['Grupo'] == subgrupo_nombre]
                
                # Para N/C y Retenciones, mostrar el título del sub-grupo
                if len(subgrupos) > 1:
                    ws.merge_range(current_row, 0, current_row, len(columnas_reporte) - 1, subgrupo_nombre, subgroup_title_format)
                    current_row += 1
                
                ws.write_row(current_row, 0, columnas_reporte, header_format)
                current_row += 1
                
                start_data_row = current_row
                for _, row_data in df_subgrupo.iterrows():
                    # (El código para escribir las filas no cambia)
                    # ...
                    ws.write(current_row, 0, row_data.get('Asiento', ''), text_format)
                    ws.write_datetime(current_row, 1, row_data.get('Fecha', None), date_format)
                    ws.write(current_row, 2, row_data.get('Fuente', ''), text_format)
                    ws.write(current_row, 3, row_data.get('Cuenta Contable', ''), text_format)
                    ws.write(current_row, 4, row_data.get('Descripción de Cuenta', ''), text_format)
                    ws.write(current_row, 5, row_data.get('Referencia', ''), text_format)
                    ws.write_number(current_row, 6, row_data.get('Débito Dolar', 0), money_format)
                    ws.write_number(current_row, 7, row_data.get('Crédito Dolar', 0), money_format)
                    ws.write_number(current_row, 8, row_data.get('Débito VES', 0), money_format)
                    ws.write_number(current_row, 9, row_data.get('Crédito VES', 0), money_format)
                    current_row += 1

                if not df_subgrupo.empty:
                    # (El código para los totales no cambia)
                    # ...
                    ws.write(current_row, 5, f'TOTALES {subgrupo_nombre.split(":")[-1].strip()}', total_label_format)
                    ws.write_formula(current_row, 6, f'=SUM(G{start_data_row + 1}:G{current_row})', total_money_format)
                    ws.write_formula(current_row, 7, f'=SUM(H{start_data_row + 1}:H{current_row})', total_money_format)
                    ws.write_formula(current_row, 8, f'=SUM(I{start_data_row + 1}:I{current_row})', total_money_format)
                    ws.write_formula(current_row, 9, f'=SUM(J{start_data_row + 1}:J{current_row})', total_money_format)
                    current_row += 2

            # (El código para ajustar el ancho de las columnas no cambia)
            # ...
            ws.set_column('A:A', 12); ws.set_column('B:B', 12); ws.set_column('C:C', 15)
            ws.set_column('D:D', 18); ws.set_column('E:E', 40); ws.set_column('F:F', 50)
            ws.set_column('G:J', 15)

    return output_buffer.getvalue()
