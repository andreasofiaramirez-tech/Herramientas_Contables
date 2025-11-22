# utils.py

import pandas as pd
import numpy as np
import re
import xlsxwriter
from io import BytesIO
import streamlit as st

# ==============================================================================
# 1. FUNCIONES AUXILIARES Y DE LIMPIEZA
# ==============================================================================

def get_col_idx(df, possible_names):
    """
    Helper para encontrar el índice numérico de una columna en un DataFrame,
    necesario para aplicar formatos en xlsxwriter.
    """
    for idx, col in enumerate(df.columns):
        if col in possible_names:
            return idx
    return -1
    
@st.cache_data
def cargar_y_limpiar_datos(uploaded_actual, uploaded_anterior, log_messages):
    """
    Versión final que estandariza los nombres de las columnas principales
    para ser insensible a mayúsculas/minúsculas y variaciones comunes.
    """
    def mapear_columnas_financieras(df, log_messages):
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
            df = pd.read_excel(archivo_buffer, engine='openpyxl')
        except Exception as e:
            log_messages.append(f"❌ Error al leer el archivo Excel: {e}")
            return None

        COLUMN_STANDARDIZATION_MAP = {
            'Asiento': ['ASIENTO', 'Asiento'],
            'Fuente': ['FUENTE', 'Fuente'],
            'Fecha': ['FECHA', 'Fecha'],
            'Referencia': ['REFERENCIA', 'Referencia'],
            'NIT': ['Nit', 'NIT', 'Rif', 'RIF'],
            'Descripcion NIT': ['Descripción Nit', 'Descripcion Nit', 'DESCRIPCION NIT', 'Descripción NIT', 'Descripcion NIT'],
            'Nombre del Proveedor': ['Nombre del Proveedor', 'NOMBRE DEL PROVEEDOR', 'Nombre Proveedor']
        }
        rename_map = {}
        for standard_name, variations in COLUMN_STANDARDIZATION_MAP.items():
            for var in variations:
                if var in df.columns:
                    rename_map[var] = standard_name
                    break
        if rename_map:
            df.rename(columns=rename_map, inplace=True)
            log_messages.append(f"✔️ Nombres de columna estandarizados. Mapeo aplicado: {rename_map}")

        for col in ['Fuente', 'Nombre del Proveedor']:
            if col not in df.columns:
                df[col] = ''
        df = mapear_columnas_financieras(df, log_messages).copy()
        
        df['Asiento'] = df.get('Asiento', pd.Series(dtype='str')).astype(str).str.strip()
        df['Referencia'] = df.get('Referencia', pd.Series(dtype='str')).astype(str).str.strip()
        df['Fecha'] = pd.to_datetime(df.get('Fecha'), errors='coerce')

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

    df_full['Monto_BS'] = (df_full.get('Débito Bolivar', 0) - df_full.get('Crédito Bolivar', 0)).round(2)
    df_full['Monto_USD'] = (df_full.get('Débito Dolar', 0) - df_full.get('Crédito Dolar', 0)).round(2)
    df_full[['Conciliado', 'Grupo_Conciliado', 'Referencia_Normalizada_Literal']] = [False, np.nan, np.nan]

    log_messages.append(f"✅ Datos de Excel cargados. Total movimientos: {len(df_full)}")
    return df_full
    
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
# 2. LOGICA MODULAR PARA REPORTES EXCEL
# ==============================================================================

def _crear_formatos(workbook):
    """Centraliza la creación de estilos para el Excel."""
    return {
        'encabezado_empresa': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14}),
        'encabezado_sub': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 11}),
        'header_tabla': workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'}),
        'bs': workbook.add_format({'num_format': '#,##0.00'}),
        'usd': workbook.add_format({'num_format': '$#,##0.00'}),
        'tasa': workbook.add_format({'num_format': '#,##0.0000'}),
        'fecha': workbook.add_format({'num_format': 'dd/mm/yyyy'}),
        'total_label': workbook.add_format({'bold': True, 'align': 'right', 'top': 2}),
        'total_usd': workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'top': 2, 'bottom': 1}),
        'total_bs': workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 2, 'bottom': 1}),
        'proveedor_header': workbook.add_format({'bold': True, 'fg_color': '#F2F2F2', 'border': 1}),
        'subtotal_label': workbook.add_format({'bold': True, 'align': 'right', 'top': 1}),
        'subtotal_usd': workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'top': 1}),
        'subtotal_bs': workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})
    }

def _generar_hoja_pendientes(workbook, formatos, df_saldos, estrategia, casa, fecha_maxima):
    """Genera la hoja de 'Pendientes' o 'Saldos Abiertos'."""
    nombre_hoja = estrategia.get("nombre_hoja_excel", "Pendientes")
    ws = workbook.add_worksheet(nombre_hoja)
    cols = estrategia["columnas_reporte"]
    
    # Encabezados
    if pd.notna(fecha_maxima):
        ultimo_dia = fecha_maxima + pd.offsets.MonthEnd(0)
        meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        txt_fecha = f"PARA EL {ultimo_dia.day} DE {meses[ultimo_dia.month].upper()} DE {ultimo_dia.year}"
    else:
        txt_fecha = "FECHA NO DISPONIBLE"

    if cols:
        ws.merge_range(0, 0, 0, len(cols)-1, casa, formatos['encabezado_empresa'])
        ws.merge_range(1, 0, 1, len(cols)-1, f"ESPECIFICACION DE LA CUENTA {estrategia['nombre_hoja_excel']}", formatos['encabezado_sub'])
        ws.merge_range(2, 0, 2, len(cols)-1, txt_fecha, formatos['encabezado_sub'])
        ws.write_row(4, 0, cols, formatos['header_tabla'])

    if df_saldos.empty: return

    # Preparar datos
    df = df_saldos.copy()
    df['Monto Dólar'] = pd.to_numeric(df.get('Monto_USD'), errors='coerce').fillna(0)
    df['Bs.'] = pd.to_numeric(df.get('Monto_BS'), errors='coerce').fillna(0)
    df['Monto Bolivar'] = df['Bs.'] # Alias
    df['Tasa'] = np.where(df['Monto Dólar'].abs() != 0, df['Bs.'].abs() / df['Monto Dólar'].abs(), 0)
    df = df.sort_values(by=['NIT', 'Fecha'])

    current_row = 5
    # Indices para columnas
    usd_idx = get_col_idx(pd.DataFrame(columns=cols), ['Monto Dólar', 'Monto USD'])
    bs_idx = get_col_idx(pd.DataFrame(columns=cols), ['Bs.', 'Monto Bolivar', 'Monto Bs'])
    
    # Escritura agrupada por NIT
    for nit, grupo in df.groupby('NIT'):
        for _, row in grupo.iterrows():
            for c_idx, col_name in enumerate(cols):
                val = row.get(col_name)
                if col_name == 'Fecha' and pd.notna(val): ws.write_datetime(current_row, c_idx, val, formatos['fecha'])
                elif col_name in ['Monto Dólar', 'Monto USD']: ws.write_number(current_row, c_idx, val or 0, formatos['usd'])
                elif col_name in ['Bs.', 'Monto Bolivar', 'Monto Bs']: ws.write_number(current_row, c_idx, val or 0, formatos['bs'])
                elif col_name == 'Tasa': ws.write_number(current_row, c_idx, val or 0, formatos['tasa'])
                else: ws.write(current_row, c_idx, val if pd.notna(val) else '')
            current_row += 1
        
        # Subtotal por NIT (Etiqueta "Saldo")
        lbl_idx = max(0, (usd_idx if usd_idx != -1 else bs_idx) - 1)
        ws.write(current_row, lbl_idx, "Saldo", formatos['subtotal_label'])
        
        if usd_idx != -1: ws.write_number(current_row, usd_idx, grupo['Monto Dólar'].sum(), formatos['subtotal_usd'])
        if bs_idx != -1: ws.write_number(current_row, bs_idx, grupo['Bs.'].sum(), formatos['subtotal_bs'])
        current_row += 2

    # --- NUEVO BLOQUE: SALDO TOTAL AL FINAL DE LA HOJA ---
    current_row += 1 # Dejar una fila extra de espacio
    lbl_idx = max(0, (usd_idx if usd_idx != -1 else bs_idx) - 1)
    
    ws.write(current_row, lbl_idx, "SALDO TOTAL", formatos['total_label'])
    
    # Sumar toda la columna del DataFrame filtrado
    if usd_idx != -1: 
        ws.write_number(current_row, usd_idx, df['Monto Dólar'].sum(), formatos['total_usd'])
    if bs_idx != -1: 
        ws.write_number(current_row, bs_idx, df['Bs.'].sum(), formatos['total_bs'])
    # -----------------------------------------------------

    # Ajuste de anchos de columna
    ws.set_column(0, 0, 18)  # Asiento
    ws.set_column(1, 1, 55)  # Referencia
    ws.set_column(2, 2, 15)  # Fecha
    ws.set_column(3, 10, 20) # Montos

def _generar_hoja_conciliados_estandar(workbook, formatos, df_conciliados, estrategia):
    """Para cuentas: Tránsito, Depositar, Viajes, Devoluciones, Deudores."""
    ws = workbook.add_worksheet("Conciliacion")
    
    # Preparar DataFrame
    df = df_conciliados.copy()
    
    # Caso especial nombres de columnas para Devoluciones
    es_devolucion = estrategia['id'] == 'devoluciones_proveedores'
    
    if es_devolucion:
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor', 'Monto Dólar', 'Monto Bs.', 'Grupo de Conciliación']
        df['Monto Dólar'] = df['Monto_USD']
        df['Monto Bs.'] = df['Monto_BS']
        df['Grupo de Conciliación'] = df['Grupo_Conciliado']
    else:
        # Estándar
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Débitos Dólares', 'Créditos Dólares', 'Débitos Bs', 'Créditos Bs', 'Grupo de Conciliación']
        df['Débitos Dólares'] = df['Monto_USD'].apply(lambda x: x if x > 0 else 0)
        df['Créditos Dólares'] = df['Monto_USD'].apply(lambda x: x if x < 0 else 0)
        df['Débitos Bs'] = df['Monto_BS'].apply(lambda x: x if x > 0 else 0)
        df['Créditos Bs'] = df['Monto_BS'].apply(lambda x: x if x < 0 else 0)
        df['Grupo de Conciliación'] = df['Grupo_Conciliado']
    
    df = df.reindex(columns=columnas).sort_values(by=['Grupo de Conciliación', 'Fecha'])
    
    # Escribir
    ws.merge_range(0, 0, 0, len(columnas)-1, 'Detalle de Movimientos Conciliados', formatos['encabezado_sub'])
    ws.write_row(1, 0, columnas, formatos['header_tabla'])
    
    # Indices
    deb_usd_idx = get_col_idx(df, ['Débitos Dólares', 'Monto Dólar'])
    cre_usd_idx = get_col_idx(df, ['Créditos Dólares']) 
    deb_bs_idx = get_col_idx(df, ['Débitos Bs', 'Monto Bs.'])
    cre_bs_idx = get_col_idx(df, ['Créditos Bs'])

    current_row = 2
    for _, row in df.iterrows():
        for c_idx, val in enumerate(row):
            if c_idx in [deb_usd_idx, cre_usd_idx]: ws.write_number(current_row, c_idx, val, formatos['usd'])
            elif c_idx in [deb_bs_idx, cre_bs_idx]: ws.write_number(current_row, c_idx, val, formatos['bs'])
            elif pd.isna(val): ws.write(current_row, c_idx, '')
            elif isinstance(val, pd.Timestamp): ws.write_datetime(current_row, c_idx, val, formatos['fecha'])
            else: ws.write(current_row, c_idx, val)
        current_row += 1
    
    # --- TOTALES Y COMPROBACIÓN ---
    ws.write(current_row, 2, "TOTALES", formatos['total_label']) # Etiqueta en col Referencia
    
    sum_deb_usd = df.iloc[:, deb_usd_idx].sum() if deb_usd_idx != -1 else 0
    sum_cre_usd = df.iloc[:, cre_usd_idx].sum() if cre_usd_idx != -1 else 0
    sum_deb_bs = df.iloc[:, deb_bs_idx].sum() if deb_bs_idx != -1 else 0
    sum_cre_bs = df.iloc[:, cre_bs_idx].sum() if cre_bs_idx != -1 else 0

    if deb_usd_idx != -1: ws.write_number(current_row, deb_usd_idx, sum_deb_usd, formatos['total_usd'])
    if cre_usd_idx != -1: ws.write_number(current_row, cre_usd_idx, sum_cre_usd, formatos['total_usd'])
    if deb_bs_idx != -1: ws.write_number(current_row, deb_bs_idx, sum_deb_bs, formatos['total_bs'])
    if cre_bs_idx != -1: ws.write_number(current_row, cre_bs_idx, sum_cre_bs, formatos['total_bs'])
    
    # --- FILA DE COMPROBACIÓN ---
    # Como los créditos son negativos en la data, la suma algebraica (Débito + Crédito) debe dar 0.
    current_row += 1
    ws.write(current_row, 2, "Comprobacion", formatos['subtotal_label'])
    
    # Comprobación USD
    if deb_usd_idx != -1 and cre_usd_idx != -1:
        neto_usd = sum_deb_usd + sum_cre_usd 
        ws.write_number(current_row, deb_usd_idx, neto_usd, formatos['total_usd'])
        
    # Comprobación Bs (o en devoluciones si solo hay 1 columna de monto, no aplica comprobación D-C)
    if deb_bs_idx != -1 and cre_bs_idx != -1:
        neto_bs = sum_deb_bs + sum_cre_bs
        ws.write_number(current_row, deb_bs_idx, neto_bs, formatos['total_bs'])

    ws.set_column('A:H', 18)

def _generar_hoja_conciliados_agrupada(workbook, formatos, df_conciliados, estrategia):
    """Para cuentas agrupadas: Cobros Viajeros, Otras CxP."""
    ws = workbook.add_worksheet("Conciliacion")
    df = df_conciliados.copy()
    
    if estrategia['id'] == 'cobros_viajeros':
        df['Débitos Dólares'] = df['Monto_USD'].apply(lambda x: x if x > 0 else 0)
        df['Créditos Dólares'] = df['Monto_USD'].apply(lambda x: abs(x) if x < 0 else 0) # Aquí usamos abs para visualización si lo prefieres
        # NOTA: Si usamos abs en créditos para visualización, la comprobación debe ser RESTANDO.
        # Pero para mantener consistencia matemática con la base de datos, mantengamos el signo negativo en la suma
        # O si prefieres que se vean positivos en el reporte, cambiamos la lógica de comprobación.
        # Vamos a mantener el signo negativo para que la suma algebraica de 0.
        df['Créditos Dólares'] = df['Monto_USD'].apply(lambda x: x if x < 0 else 0) 
        
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Débitos Dólares', 'Créditos Dólares']
        df = df.sort_values(by=['NIT', 'Grupo_Conciliado', 'Fecha'])
        cols_sum = ['Débitos Dólares', 'Créditos Dólares']
        titulo = 'Detalle de Movimientos Conciliados por Viajero (NIT)'
    else: # Otras CxP
        df['Monto Bs.'] = df['Monto_BS']
        columnas = ['Fecha', 'Descripcion NIT', 'Numero_Envio', 'Monto Bs.']
        df = df.sort_values(by=['NIT', 'Numero_Envio', 'Fecha'])
        cols_sum = ['Monto Bs.']
        titulo = 'Detalle de Movimientos Conciliados por Proveedor y Envío'

    ws.merge_range(0, 0, 0, len(columnas)+1, titulo, formatos['encabezado_sub'])
    current_row = 2
    
    # Iterar por NIT
    grand_totals = {c: 0.0 for c in cols_sum}
    
    for nit, grupo in df.groupby('NIT'):
        nombre = grupo['Descripcion NIT'].iloc[0] if not grupo.empty and 'Descripcion NIT' in grupo else ''
        ws.merge_range(current_row, 0, current_row, len(columnas)-1, f"NIT: {nit} - {nombre}", formatos['proveedor_header'])
        current_row += 1
        ws.write_row(current_row, 0, columnas, formatos['header_tabla'])
        current_row += 1
        
        for _, row in grupo.iterrows():
            for c_idx, col_name in enumerate(columnas):
                val = row.get(col_name)
                if col_name == 'Fecha' and pd.notna(val): ws.write_datetime(current_row, c_idx, val, formatos['fecha'])
                elif col_name in ['Débitos Dólares', 'Créditos Dólares']: ws.write_number(current_row, c_idx, val, formatos['usd'])
                elif col_name in ['Monto Bs.']: ws.write_number(current_row, c_idx, val, formatos['bs'])
                else: ws.write(current_row, c_idx, val if pd.notna(val) else '')
            current_row += 1
        
        # Subtotal
        lbl_col = len(columnas) - len(cols_sum) - 1
        ws.write(current_row, lbl_col, "Subtotal", formatos['subtotal_label'])
        for i, c_sum in enumerate(cols_sum):
            suma = grupo[c_sum].sum()
            grand_totals[c_sum] += suma
            fmt = formatos['subtotal_usd'] if 'Dólares' in c_sum else formatos['subtotal_bs']
            ws.write_number(current_row, lbl_col + 1 + i, suma, fmt)
        current_row += 2

    # TOTALES
    lbl_col = len(columnas) - len(cols_sum) - 1
    ws.write(current_row, lbl_col, "TOTALES", formatos['total_label'])
    for i, c_sum in enumerate(cols_sum):
        fmt = formatos['total_usd'] if 'Dólares' in c_sum else formatos['total_bs']
        ws.write_number(current_row, lbl_col + 1 + i, grand_totals[c_sum], fmt)
        
    # COMPROBACION (Solo si hay débitos y créditos separados)
    if len(cols_sum) > 1:
        current_row += 1
        ws.write(current_row, lbl_col, "Comprobacion", formatos['subtotal_label'])
        # Asumiendo cols_sum = ['Debitos', 'Creditos'] donde Creditos son negativos
        neto = grand_totals[cols_sum[0]] + grand_totals[cols_sum[1]]
        ws.write_number(current_row, lbl_col + 1, neto, formatos['total_usd'])

    ws.set_column('A:F', 20)

def _generar_hoja_resumen_devoluciones(workbook, formatos, df_saldos):
    """Hoja extra específica para Devoluciones a Proveedores."""
    ws = workbook.add_worksheet("Resumen por Proveedor")
    cols = ['Fecha', 'Fuente', 'Referencia', 'Monto USD', 'Monto Bs']
    ws.merge_range('A1:E1', 'Detalle de Saldos Abiertos por Proveedor', formatos['encabezado_sub'])
    ws.write_row(2, 0, cols, formatos['header_tabla'])
    
    df = df_saldos.sort_values(by=['Nombre del Proveedor', 'Fecha'])
    current_row = 3
    
    for prov, grupo in df.groupby('Nombre del Proveedor'):
        ws.merge_range(current_row, 0, current_row, 4, f"Proveedor: {prov}", formatos['proveedor_header'])
        current_row += 1
        for _, row in grupo.iterrows():
            ws.write_datetime(current_row, 0, row['Fecha'], formatos['fecha'])
            ws.write(current_row, 1, row.get('Fuente', ''))
            ws.write(current_row, 2, row.get('Referencia', ''))
            ws.write_number(current_row, 3, row.get('Monto_USD', 0), formatos['usd'])
            ws.write_number(current_row, 4, row.get('Monto_BS', 0), formatos['bs'])
            current_row += 1
        
        ws.write(current_row, 2, f"Subtotal {prov}", formatos['subtotal_label'])
        ws.write_number(current_row, 3, grupo['Monto_USD'].sum(), formatos['subtotal_usd'])
        ws.write_number(current_row, 4, grupo['Monto_BS'].sum(), formatos['subtotal_bs'])
        current_row += 2
    ws.set_column('A:E', 18)

# ==============================================================================
# 3. FUNCIÓN PRINCIPAL (CONTROLADOR)
# ==============================================================================

@st.cache_data
def generar_reporte_excel(_df_full, df_saldos_abiertos, df_conciliados, _estrategia, casa_seleccionada, cuenta_seleccionada):
    """Controlador principal que orquesta la creación del Excel."""
    
    output_excel = BytesIO()
    
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        workbook = writer.book
        formatos = _crear_formatos(workbook)
        
        # 1. Generar Hoja de Pendientes (Común para todos)
        fecha_max = _df_full['Fecha'].dropna().max()
        
        # Pasamos _estrategia a las funciones auxiliares
        _generar_hoja_pendientes(workbook, formatos, df_saldos_abiertos, _estrategia, casa_seleccionada, fecha_max)
        
        # 2. Generar Hoja de Conciliados (Según estilo)
        if not df_conciliados.empty:
            # Lista de cuentas que usan el estilo agrupado por NIT
            cuentas_agrupadas = ['cobros_viajeros', 'otras_cuentas_por_pagar']
            
            if _estrategia['id'] in cuentas_agrupadas:
                _generar_hoja_conciliados_agrupada(workbook, formatos, df_conciliados, _estrategia)
            else:
                # Todas las demás usan el estilo estándar (lista plana)
                _generar_hoja_conciliados_estandar(workbook, formatos, df_conciliados, _estrategia)

        # 3. Generar Hoja Extra (Solo para Devoluciones)
        if _estrategia['id'] == 'devoluciones_proveedores' and not df_saldos_abiertos.empty:
            _generar_hoja_resumen_devoluciones(workbook, formatos, df_saldos_abiertos)

    return output_excel.getvalue()


# ==============================================================================
# 4. REPORTE PARA LA HERRAMIENTA DE RETENCIONES
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
# 5. REPORTE PARA ANÁLISIS DE PAQUETE CC
# ==============================================================================

def generar_reporte_paquete_cc(df_analizado):
    """
    Versión final del reporte compatible con versiones antiguas de XlsxWriter.
    """
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- Formatos ---
        main_title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 16})
        descriptive_title_format = workbook.add_format({'bold': True, 'font_size': 14, 'fg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        subgroup_title_format = workbook.add_format({'bold': True, 'font_size': 11, 'fg_color': '#E0E0E0', 'border': 1})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        
        # Formatos estándar
        money_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        text_format = workbook.add_format({'border': 1})
        
        # Formatos para filas con incidencia (rojo)
        incidencia_text_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
        incidencia_money_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '#,##0.00', 'border': 1})
        incidencia_date_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': 'dd/mm/yyyy', 'border': 1})

        # Formatos para totales
        total_label_format = workbook.add_format({'bold': True, 'align': 'right', 'top': 2, 'font_color': '#003366'})
        total_money_format = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 2, 'bottom': 1})

        columnas_reporte = [
            'Asiento', 'Fecha', 'Fuente', 'Cuenta Contable', 'Descripción de Cuenta', 
            'Referencia', 'Débito Dolar', 'Crédito Dolar', 'Débito VES', 'Crédito VES'
        ]
        
        df_analizado['Grupo Principal'] = df_analizado['Grupo'].apply(lambda x: x.split(':')[0].strip())
        def sort_key(group_name):
            if group_name.startswith('Grupo'): return (0, int(group_name.split()[1]))
            return (1, group_name)
        grupos_principales_ordenados = sorted(df_analizado['Grupo Principal'].unique(), key=sort_key)
        
        ws_dir = workbook.add_worksheet("Directorio")
        ws_dir.merge_range('A1:C1', 'Directorio de Grupos y Resumen de Auditoría', main_title_format)
        ws_dir.write('A2', 'Nombre de la Hoja', header_format)
        ws_dir.write('B2', 'Descripción del Contenido', header_format)
        ws_dir.write('C2', 'Observaciones', header_format)
        
        dir_row = 2
        for grupo_principal in grupos_principales_ordenados:
            sheet_name = re.sub(r'[\\/*?:"\[\]]', '', grupo_principal)[:31]
            df_grupo_dir = df_analizado[df_analizado['Grupo Principal'] == grupo_principal]
            full_name_example = df_grupo_dir['Grupo'].iloc[0]
            description = full_name_example.split(':', 1)[-1].strip() if ':' in full_name_example else full_name_example
            if grupo_principal in ["Grupo 3", "Grupo 9", "Grupo 8", "Grupo 6", "Grupo 7"]:
                description = f"{description.split('-')[0].strip()} (Varios Subgrupos)"
            
            observacion = "Incidencia Encontrada" if (df_grupo_dir['Estado'] != 'Conciliado').any() else "Conciliado"
            
            ws_dir.write(dir_row, 0, sheet_name, text_format)
            ws_dir.write(dir_row, 1, description, text_format)
            ws_dir.write(dir_row, 2, observacion, text_format)
            dir_row += 1
            
        ws_dir.set_column('A:A', 25); ws_dir.set_column('B:B', 60); ws_dir.set_column('C:C', 25)
        
        for grupo_principal_nombre in grupos_principales_ordenados:
            sheet_name = re.sub(r'[\\/*?:"\[\]]', '', grupo_principal_nombre)[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.hide_gridlines(2)
            
            ws.merge_range('A1:J1', 'Análisis de Asientos de Cuentas por Cobrar', main_title_format)
            
            df_grupo_completo = df_analizado[df_analizado['Grupo Principal'] == grupo_principal_nombre]
            subgrupos = sorted(df_grupo_completo['Grupo'].unique())
            
            full_descriptive_title = subgrupos[0]
            if len(subgrupos) > 1:
                full_descriptive_title = f"{subgrupos[0].split(':')[0].strip()}: {subgrupos[0].split(':')[1].split('-')[0].strip()}"
            
            ws.merge_range('A3:J3', full_descriptive_title, descriptive_title_format)
            current_row = 4
            
            for subgrupo_nombre in subgrupos:
                df_subgrupo = df_grupo_completo[df_grupo_completo['Grupo'] == subgrupo_nombre]
                
                if len(subgrupos) > 1:
                    ws.merge_range(current_row, 0, current_row, len(columnas_reporte) - 1, subgrupo_nombre, subgroup_title_format)
                    current_row += 1
                
                ws.write_row(current_row, 0, columnas_reporte, header_format)
                current_row += 1
                
                start_data_row = current_row
                for _, row_data in df_subgrupo.iterrows():
                    formato_fila_texto = text_format
                    formato_fila_numero = money_format
                    formato_fila_fecha = date_format
                    if row_data.get('Estado', 'Conciliado') != 'Conciliado':
                        formato_fila_texto = incidencia_text_format
                        formato_fila_numero = incidencia_money_format
                        formato_fila_fecha = incidencia_date_format
                    else:
                        formato_fila_texto = text_format
                        formato_fila_numero = money_format
                        formato_fila_fecha = date_format
                    
                    # Escribir filas usando los formatos predefinidos
                    ws.write(current_row, 0, row_data.get('Asiento', ''), formato_fila_texto)
                    ws.write_datetime(current_row, 1, row_data.get('Fecha', None), formato_fila_fecha)
                    ws.write(current_row, 2, row_data.get('Fuente', ''), formato_fila_texto)
                    ws.write(current_row, 3, row_data.get('Cuenta Contable', ''), formato_fila_texto)
                    ws.write(current_row, 4, row_data.get('Descripción de Cuenta', ''), formato_fila_texto)
                    ws.write(current_row, 5, row_data.get('Referencia', ''), formato_fila_texto)
                    ws.write_number(current_row, 6, row_data.get('Débito Dolar', 0), formato_fila_numero)
                    ws.write_number(current_row, 7, row_data.get('Crédito Dolar', 0), formato_fila_numero)
                    ws.write_number(current_row, 8, row_data.get('Débito VES', 0), formato_fila_numero)
                    ws.write_number(current_row, 9, row_data.get('Crédito VES', 0), formato_fila_numero)
                    current_row += 1
                
                if not df_subgrupo.empty:
                    ws.write(current_row, 5, f'TOTALES {subgrupo_nombre.split(":")[-1].strip()}', total_label_format)
                    ws.write_formula(current_row, 6, f'=SUM(G{start_data_row + 1}:G{current_row})', total_money_format)
                    ws.write_formula(current_row, 7, f'=SUM(H{start_data_row + 1}:H{current_row})', total_money_format)
                    ws.write_formula(current_row, 8, f'=SUM(I{start_data_row + 1}:I{current_row})', total_money_format)
                    ws.write_formula(current_row, 9, f'=SUM(J{start_data_row + 1}:J{current_row})', total_money_format)
                    current_row += 2

            ws.set_column('A:A', 12); ws.set_column('B:B', 12); ws.set_column('C:C', 15)
            ws.set_column('D:D', 18); ws.set_column('E:E', 40); ws.set_column('F:F', 50)
            ws.set_column('G:J', 15)

    return output_buffer.getvalue()
