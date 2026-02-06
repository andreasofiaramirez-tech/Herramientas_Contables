# utils.py

import pandas as pd
import numpy as np
import re
import xlsxwriter
from io import BytesIO
import streamlit as st    
import unicodedata

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
    """Carga, limpia y unifica los archivos de Excel."""
    
    # --- FUNCIONES AUXILIARES INTERNAS ---
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
        
        # Limpieza inicial
        t = re.sub(r'[^\d.,-]', '', str(texto).strip())
        if not t: return '0.0'
        
        # Detección inteligente
        idx_punto = t.rfind('.')
        idx_coma = t.rfind(',')

        if idx_punto > idx_coma:
            # Formato "81,268.96" -> Eliminar comas
            return t.replace(',', '')
            
        elif idx_coma > idx_punto:
            # Formato "81.268,96" -> Eliminar puntos, coma a punto
            return t.replace('.', '').replace(',', '.')
            
        return t.replace(',', '.') # Fallback estándar

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

    # --- EJECUCIÓN PRINCIPAL ---
    df_actual = procesar_excel(uploaded_actual)
    df_anterior = procesar_excel(uploaded_anterior)

    if df_actual is None or df_anterior is None:
        st.error("❌ ¡Error Fatal! No se pudo procesar uno o ambos archivos Excel.")
        return None

    # Concatenar
    df_full = pd.concat([df_anterior, df_actual], ignore_index=True)
    
    # --- NUEVO BLOQUE DE LIMPIEZA DE "BASURA" (TOTALES/VACÍOS) ---
    # 1. Eliminamos filas donde TODAS las columnas sean nulas
    df_full.dropna(how='all', inplace=True)
    
    # 2. Eliminamos filas que parecen Totales o Subtotales
    # Buscamos en las columnas de texto palabras clave como "TOTAL", "SALDO", "SUMA"
    #cols_texto = ['Asiento', 'Referencia', 'NIT', 'Descripcion NIT', 'Nombre del Proveedor']
    #for col in cols_texto:
    #    if col in df_full.columns:
            # Convertimos a string mayúscula y buscamos la palabra "TOTAL" o "SALDO"
            # Pero CUIDADO: No borrar "SALDO INICIAL" si es un asiento legítimo.
            # Borramos solo si la celda es EXACTAMENTE "TOTAL", "GRAN TOTAL", "SALDO TOTAL"
    #        mask_basura = df_full[col].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAN TOTAL', 'SALDO TOTAL', 'SUBTOTAL', 'TOTALES'])
    #        if mask_basura.any():
    #            filas_borradas = mask_basura.sum()
    #            df_full = df_full[~mask_basura]
                # log_messages.append(f"🧹 Se eliminaron {filas_borradas} filas de totales en columna {col}.")
    
    # --- ¡IMPORTANTE! ELIMINACIÓN DE DUPLICADOS DESACTIVADA ---
    # Se comenta esta línea para evitar pérdida de datos legítimos idénticos
    # key_cols = ['Asiento', 'Referencia', 'Fecha', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
    # df_full.drop_duplicates(subset=[col for col in key_cols if col in df_full.columns], keep='first', inplace=True)
    # ----------------------------------------------------------

    df_full['Monto_BS'] = (df_full.get('Débito Bolivar', 0) - df_full.get('Crédito Bolivar', 0)).round(2)
    df_full['Monto_USD'] = (df_full.get('Débito Dolar', 0) - df_full.get('Crédito Dolar', 0)).round(2)
    df_full[['Conciliado', 'Grupo_Conciliado', 'Referencia_Normalizada_Literal']] = [False, np.nan, np.nan]

    # --- LOG DE VERIFICACIÓN (AHORA EN EL LUGAR CORRECTO) ---
    log_messages.append(f"✅ Datos cargados. Filas archivo anterior: {len(df_anterior)}, Actual: {len(df_actual)}. Total consolidado: {len(df_full)}")
    
    return df_full

@st.cache_data
def cargar_datos_cofersa(uploaded_actual, uploaded_anterior, log_messages):
    """
    Función de carga EXCLUSIVA para COFERSA.
    MEJORA: Ignora acentos en encabezados y procesa comas decimales correctamente.
    """
    import unicodedata

    def normalizar_texto(texto):
        """Quita acentos y convierte a mayúsculas."""
        if not isinstance(texto, str): return str(texto)
        return ''.join(c for c in unicodedata.normalize('NFD', texto)
                      if unicodedata.category(c) != 'Mn').upper().strip()

    def limpiar_monto_robusto(val):
        """Convierte montos (con comas o puntos) a float puro."""
        if pd.isna(val) or str(val).strip() in ['', '-', 'nan']: return 0.0
        if isinstance(val, (int, float)): return float(val)
        
        t = str(val).strip().replace('Bs', '').replace('$', '').replace(' ', '')
        # Caso: 1.234.567,89 -> 1234567.89
        if ',' in t and '.' in t:
            if t.rfind(',') > t.rfind('.'): t = t.replace('.', '').replace(',', '.')
            else: t = t.replace(',', '')
        # Caso: 1234,56 -> 1234.56
        elif ',' in t:
            t = t.replace(',', '.')
            
        t = re.sub(r'[^\d.-]', '', t)
        try: return float(t)
        except: return 0.0

    def procesar_excel_cofersa(archivo_buffer):
        try:
            archivo_buffer.seek(0)
            df = pd.read_excel(archivo_buffer, engine='openpyxl')
        except Exception as e:
            log_messages.append(f"❌ Error al leer Excel COFERSA: {e}")
            return None

        # 1. Mapeo Robusto de Columnas (Quitando acentos)
        rename_map = {}
        for col in df.columns:
            norm_col = normalizar_texto(col)
            
            if 'DEBITO' in norm_col and 'LOCAL' in norm_col: rename_map[col] = 'Débito Bolivar'
            elif 'CREDITO' in norm_col and 'LOCAL' in norm_col: rename_map[col] = 'Crédito Bolivar'
            elif 'DEBITO' in norm_col and 'DOLAR' in norm_col: rename_map[col] = 'Débito Dolar'
            elif 'CREDITO' in norm_col and 'DOLAR' in norm_col: rename_map[col] = 'Crédito Dolar'
            
            elif 'ASIENTO' in norm_col: rename_map[col] = 'Asiento'
            elif 'FECHA' in norm_col: rename_map[col] = 'Fecha'
            elif 'REFERENCIA' in norm_col: rename_map[col] = 'Referencia'
            elif 'TIPO' in norm_col: rename_map[col] = 'Tipo'
            elif 'FUENTE' in norm_col: rename_map[col] = 'Fuente'
            elif 'NIT' in norm_col or 'RIF' in norm_col: rename_map[col] = 'NIT'
            elif 'DESCRIPCI' in norm_col: rename_map[col] = 'Descripcion NIT'

        df.rename(columns=rename_map, inplace=True)
        # Eliminar columnas duplicadas si las hay después del rename
        df = df.loc[:, ~df.columns.duplicated()]
        
        # 2. Limpieza de Datos
        if 'Fecha' in df.columns: df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        
        # 3. Limpieza de montos usando el nuevo motor robusto
        cols_fin = ['Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
        for c in cols_fin:
            if c in df.columns:
                df[c] = df[c].apply(limpiar_monto_robusto)
            else:
                df[c] = 0.0
        
        return df

    # --- EJECUCIÓN ---
    df_act = procesar_excel_cofersa(uploaded_actual)
    df_ant = procesar_excel_cofersa(uploaded_anterior)

    if df_act is None or df_ant is None: return None

    df_full = pd.concat([df_ant, df_act], ignore_index=True)
    
    # Calcular NETOS (Ya son números seguros)
    df_full['Neto Local'] = (df_full['Débito Bolivar'] - df_full['Crédito Bolivar']).round(2)
    df_full['Neto Dólar'] = (df_full['Débito Dolar'] - df_full['Crédito Dolar']).round(2)
    
    df_full['Monto_BS'] = df_full['Neto Local']
    df_full['Monto_USD'] = df_full['Neto Dólar']
    df_full['Conciliado'] = False
    
    log_messages.append(f"✅ Datos COFERSA cargados exitosamente (Columnas con acento detectadas).")
    return df_full

@st.cache_data
def generar_excel_saldos_abiertos(df_saldos_abiertos):
    """
    Genera el archivo Excel (.xlsx) con los saldos pendientes para el próximo ciclo.
    Mantiene el formato numérico correcto para que la herramienta lo lea bien el próximo mes.
    """
    output = BytesIO()
    
    # Definir las columnas estándar que espera la herramienta al cargar
    columnas_exportar = [
        'Asiento', 'Referencia', 'Fecha', 
        'Débito Bolivar', 'Crédito Bolivar', 
        'Débito Dolar', 'Crédito Dolar', 
        'Fuente', 'Nombre del Proveedor', 'NIT', 'Descripcion NIT'
    ]
    
    # Filtrar solo las columnas que existen en el DF
    cols_existentes = [c for c in columnas_exportar if c in df_saldos_abiertos.columns]
    df_export = df_saldos_abiertos[cols_existentes].copy()
    
    # Asegurar que la fecha tenga formato fecha (sin hora)
    if 'Fecha' in df_export.columns:
        df_export['Fecha'] = pd.to_datetime(df_export['Fecha']).dt.date
        
    # Crear el Excel
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='SaldosAnteriores')
        
        # Ajustar anchos de columna para mejor visualización
        workbook = writer.book
        worksheet = writer.sheets['SaldosAnteriores']
        
        # Formatos
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        num_format = workbook.add_format({'num_format': '#,##0.00'})
        
        # Aplicar formatos
        for idx, col in enumerate(cols_existentes):
            # Ancho base
            worksheet.set_column(idx, idx, 15)
            
            # Formato específico
            if col == 'Fecha':
                worksheet.set_column(idx, idx, 12, date_format)
            elif 'Débito' in col or 'Crédito' in col:
                worksheet.set_column(idx, idx, 15, num_format)
            elif col in ['Referencia', 'Nombre del Proveedor', 'Descripcion NIT']:
                worksheet.set_column(idx, idx, 40)

    return output.getvalue()

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
        'colones': workbook.add_format({'num_format': '₡#,##0.00'}),    
        'usd': workbook.add_format({'num_format': '$#,##0.00'}),
        'tasa': workbook.add_format({'num_format': '#,##0.0000'}),
        'fecha': workbook.add_format({'num_format': 'dd/mm/yyyy'}),
        'text': workbook.add_format({'align': 'left'}), 
        'total_label': workbook.add_format({'bold': True, 'align': 'right', 'top': 2}),
        'total_usd': workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'top': 2, 'bottom': 1}),
        'total_bs': workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 2, 'bottom': 1}),
        'total_colones': workbook.add_format({'bold': True, 'num_format': '₡#,##0.00', 'top': 2}),
        'proveedor_header': workbook.add_format({'bold': True, 'fg_color': '#F2F2F2', 'border': 1}),
        'subtotal_label': workbook.add_format({'bold': True, 'align': 'right', 'top': 1}),
        'subtotal_usd': workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'top': 1}),
        'subtotal_bs': workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})
    }

def _generar_hoja_pendientes(workbook, formatos, df_saldos, estrategia, casa, fecha_maxima):
    """
    Genera la hoja de pendientes AGRUPADA POR NIT.
    CORREGIDO: Orden de limpieza para no borrar filas SIN_NIT.
    """
    nombre_hoja = estrategia.get("nombre_hoja_excel", "Pendientes")
    ws = workbook.add_worksheet(nombre_hoja)
    ws.hide_gridlines(2)
    cols = estrategia["columnas_reporte"]
    fmt_moneda_local = formatos['colones'] if "COFERSA" in casa else formatos['bs']
    
    # Encabezados
    if pd.notna(fecha_maxima):
        ultimo_dia = fecha_maxima + pd.offsets.MonthEnd(0)
        meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        txt_fecha = f"PARA EL {ultimo_dia.day} DE {meses[ultimo_dia.month].upper()} DE {ultimo_dia.year}"
    else:
        txt_fecha = "FECHA NO DISPONIBLE"

    ws.merge_range(0, 0, 0, len(cols)-1, casa, formatos['encabezado_empresa'])
    ws.merge_range(1, 0, 1, len(cols)-1, f"ESPECIFICACION DE LA CUENTA {estrategia['nombre_hoja_excel']}", formatos['encabezado_sub'])
    ws.merge_range(2, 0, 2, len(cols)-1, txt_fecha, formatos['encabezado_sub'])
    ws.write_row(4, 0, cols, formatos['header_tabla'])

    if df_saldos.empty: return

    df = df_saldos.copy()

    # --- PASO 1: SALVAR LAS FILAS SIN NIT (RELLENAR ANTES DE FILTRAR) ---
    if 'NIT' in df.columns: 
        # Rellenar nulos
        df['NIT'] = df['NIT'].fillna('SIN_NIT')
        # Rellenar espacios vacíos
        df['NIT'] = df['NIT'].replace(r'^\s*$', 'SIN_NIT', regex=True)
        # Convertir a string
        df['NIT'] = df['NIT'].astype(str)
        # Reemplazar 'nan' literal (insensible a mayúsculas usando regex (?i))
        df['NIT'] = df['NIT'].replace(r'(?i)^nan$', 'SIN_NIT', regex=True)
    
    # --- PASO 2: FILTRO DE BASURA (AHORA ES SEGURO) ---
    if 'NIT' in df.columns:
        # Solo borramos si dice explícitamente TOTAL o CONTABILIDAD
        # Quitamos 'NAN' de la lista negra porque ya limpiamos arriba
        mask_basura = df['NIT'].str.upper().str.contains('CONTABILIDAD|TOTAL', na=False)
        df = df[~mask_basura]
    # --------------------------------------------------

    # Conversión numérica
    df['Monto Dólar'] = pd.to_numeric(df.get('Monto_USD'), errors='coerce').fillna(0)
    df['Bs.'] = pd.to_numeric(df.get('Monto_BS'), errors='coerce').fillna(0)
    df['Monto Bolivar'] = df['Bs.']
    df['Tasa'] = np.where(df['Monto Dólar'].abs() != 0, df['Bs.'].abs() / df['Monto Dólar'].abs(), 0)
    
    # Ordenamiento
    if estrategia['id'] == 'haberes_clientes':
        df = df.sort_values(by=['Fecha', 'NIT'], ascending=[True, True])
    else:
        df = df.sort_values(by=['NIT', 'Fecha'], ascending=[True, True])
    
    current_row = 5
    
    # Indices
    col_df_ref = pd.DataFrame(columns=cols)
    usd_idx = get_col_idx(col_df_ref, ['Monto Dólar', 'Monto USD'])
    bs_idx = get_col_idx(col_df_ref, ['Bs.', 'Monto Bolivar', 'Monto Bs'])
    ref_idx = get_col_idx(col_df_ref, ['Referencia'])

    # BUCLE AGRUPADO
    for nit, grupo in df.groupby('NIT', sort=False):
        for _, row in grupo.iterrows():
            for c_idx, col_name in enumerate(cols):
                
                # --- MAPEO DE ALIAS DE COLUMNAS ---
                val = None
                
                # Caso Haberes
                if col_name == 'Fecha Origen Acreencia': val = row.get('Fecha')
                elif col_name == 'Numero de Documento': val = row.get('Fuente')
                
                # Caso Proveedores Costos
                elif col_name == 'PROVEEDOR Y DESCRIPCION': val = row.get('Referencia') # O Nombre según convenga
                elif col_name == 'FECHA COMPROB.': val = row.get('Fecha')
                elif col_name == 'EMB': 
                    val = row.get('Numero_Embarque', '')
                    if val == 'NO_EMB': val = ''
                elif col_name == 'MONEDA EXTRANJERA': val = row.get('Monto_USD')
                elif col_name == 'CAMBIO': val = row.get('Tasa')
                elif col_name == 'Bs.': val = row.get('Monto_BS')
                elif col_name == 'OBSERVACION': val = row.get('Fuente')
                
                # Caso Default
                else: val = row.get(col_name)

                # Escritura
                if col_name in ['Fecha', 'Fecha Origen Acreencia'] and pd.notna(val): 
                    ws.write_datetime(current_row, c_idx, val, formatos['fecha'])
                elif col_name in ['Monto Dólar', 'Monto USD']: 
                    ws.write_number(current_row, c_idx, val or 0, formatos['usd'])
                elif col_name in ['Bs.', 'Monto Bolivar', 'Monto Bs']: 
                    ws.write_number(current_row, c_idx, val or 0, formatos['bs'])
                elif col_name == 'Tasa': 
                    ws.write_number(current_row, c_idx, val or 0, formatos['tasa'])
                else: 
                    ws.write(current_row, c_idx, val if pd.notna(val) else '')
            current_row += 1
        
        # Subtotal por NIT
        if ref_idx != -1: lbl_idx = ref_idx
        else:
            indices_monedas = [i for i in [usd_idx, bs_idx] if i != -1]
            lbl_idx = max(0, min(indices_monedas) - 1) if indices_monedas else 0

        ws.write(current_row, lbl_idx, "Saldo", formatos['subtotal_label'])
        if usd_idx != -1: ws.write_number(current_row, usd_idx, grupo['Monto Dólar'].sum(), formatos['subtotal_usd'])
        if bs_idx != -1: ws.write_number(current_row, bs_idx, grupo['Bs.'].sum(), formatos['subtotal_bs'])
        current_row += 2
        
    # SALDO TOTAL AL FINAL
    current_row += 1
    if ref_idx != -1: lbl_idx = ref_idx
    else:
        indices_monedas = [i for i in [usd_idx, bs_idx] if i != -1]
        lbl_idx = max(0, min(indices_monedas) - 1) if indices_monedas else 0

    ws.write(current_row, lbl_idx, "SALDO TOTAL", formatos['total_label'])
    if usd_idx != -1: ws.write_number(current_row, usd_idx, df['Monto Dólar'].sum(), formatos['total_usd'])
    if bs_idx != -1: ws.write_number(current_row, bs_idx, df['Bs.'].sum(), formatos['total_bs'])


    ws.set_column(0, 0, 18) # NIT
    ws.set_column(1, 1, 55) # Descripción
    ws.set_column(2, 2, 18) # Fecha (Fecha Origen)
    ws.set_column(3, 3, 20) # Fuente (Num Documento)
    ws.set_column(4, 10, 20) # Resto

def _generar_hoja_conciliados_estandar(workbook, formatos, df_conciliados, estrategia):
    """Para cuentas: Tránsito, Depositar, Viajes, Devoluciones, Deudores."""
    ws = workbook.add_worksheet("Conciliacion")
    ws.hide_gridlines(2)
    
    # Preparar DataFrame
    df = df_conciliados.copy()
    
    # Identificadores de caso
    es_devolucion = estrategia['id'] == 'devoluciones_proveedores'
    is_cofersa = estrategia['id'] == "fondos_transito_cofersa"
    
    # 1. DETERMINAR ESTRUCTURA DE COLUMNAS (Jerarquía protegida)
    if is_cofersa:
        # Estructura de COLUMNA ÚNICA para FONDOS COFERSA
        # El crédito ya es negativo por la lógica Neto = D - C definida en el cargador
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Monto Dólar', 'Monto Colones', 'Grupo de Conciliación']
        df['Monto Colones'] = df['Monto_BS']
        df['Monto Dólar'] = df['Monto_USD']
        fmt_local = formatos['colones']
        fmt_total_local = formatos['total_colones']
        
    elif es_devolucion:
        # Estructura para Devoluciones
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor', 'Monto Dólar', 'Monto Bs.', 'Grupo de Conciliación']
        df['Monto Dólar'] = df['Monto_USD']
        df['Monto Bs.'] = df['Monto_BS']
        fmt_local = formatos['bs']
        fmt_total_local = formatos['total_bs']
        
    else:
        # Estructura Estándar de DOBLE COLUMNA (Débitos / Créditos)
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Débitos Dólares', 'Créditos Dólares', 'Débitos Bs', 'Créditos Bs', 'Grupo de Conciliación']
        df['Débitos Dólares'] = df['Monto_USD'].apply(lambda x: x if x > 0 else 0)
        df['Créditos Dólares'] = df['Monto_USD'].apply(lambda x: abs(x) if x < 0 else 0)
        df['Débitos Bs'] = df['Monto_BS'].apply(lambda x: x if x > 0 else 0)
        df['Créditos Bs'] = df['Monto_BS'].apply(lambda x: abs(x) if x < 0 else 0)
        fmt_local = formatos['bs']
        fmt_total_local = formatos['total_bs']

    # Estandarizar nombre del grupo y reordenar
    df['Grupo de Conciliación'] = df['Grupo_Conciliado']
    df = df.reindex(columns=columnas).sort_values(by=['Grupo de Conciliación', 'Fecha'])
    
    # Escribir Encabezados
    ws.merge_range(0, 0, 0, len(columnas)-1, 'Detalle de Movimientos Conciliados', formatos['encabezado_sub'])
    ws.write_row(1, 0, columnas, formatos['header_tabla'])
    
    # 2. ESCRITURA DE DATOS (Detección por nombre de columna para aplicar formatos)
    current_row = 2
    for _, row in df.iterrows():
        for c_idx, col_name in enumerate(columnas):
            val = row[col_name]
            
            # Formatos de Moneda
            if 'Dólar' in col_name or 'Dólares' in col_name:
                ws.write_number(current_row, c_idx, val, formatos['usd'])
            elif 'Colones' in col_name:
                ws.write_number(current_row, c_idx, val, formatos['colones'])
            elif 'Bs' in col_name:
                ws.write_number(current_row, c_idx, val, formatos['bs'])
            
            # Formatos de Texto y Fecha
            elif pd.isna(val):
                ws.write(current_row, c_idx, '')
            elif isinstance(val, pd.Timestamp):
                ws.write_datetime(current_row, c_idx, val, formatos['fecha'])
            else:
                ws.write(current_row, c_idx, val)
        current_row += 1
    
    # 3. TOTALES GENERALES
    ws.write(current_row, 2, "TOTALES", formatos['total_label'])
    for c_idx, col_name in enumerate(columnas):
        if any(k in col_name for k in ['Dólar', 'Dólares', 'Colones', 'Bs']):
            suma = df[col_name].sum()
            # Asignar formato de total correcto
            if 'Dólar' in col_name or 'Dólares' in col_name: f = formatos['total_usd']
            elif 'Colones' in col_name: f = formatos['total_colones']
            else: f = formatos['total_bs']
            ws.write_number(current_row, c_idx, suma, f)
    
    # 4. FILA DE COMPROBACIÓN (Solo para doble columna estándar)
    if not is_cofersa and not es_devolucion:
        current_row += 1
        ws.write(current_row, 2, "Comprobacion", formatos['subtotal_label'])
        for c_idx, col_name in enumerate(columnas):
            if 'Débitos Dólares' in col_name:
                # En la data original los créditos son negativos, por lo que Sum(D)+Sum(C) debe ser 0
                ws.write_number(current_row, c_idx, df['Débitos Dólares'].sum() - df['Créditos Dólares'].sum(), formatos['total_usd'])
            if 'Débitos Bs' in col_name:
                ws.write_number(current_row, c_idx, df['Débitos Bs'].sum() - df['Créditos Bs'].sum(), formatos['total_bs'])

    ws.set_column(0, len(columnas), 18)

def _generar_hoja_conciliados_agrupada(workbook, formatos, df_conciliados, estrategia):
    """Para cuentas agrupadas: Cobros Viajeros, Otras CxP, Deudores, Haberes y Factoring."""
    ws = workbook.add_worksheet("Conciliacion")
    ws.hide_gridlines(2)
    
    df = df_conciliados.copy()
    
    # Variables de control por defecto
    mostrar_saldo_linea = False
    col_saldo_idx = -1
    fmt_moneda = formatos['bs']
    fmt_total = formatos['total_bs']

    # --- 1. COBROS VIAJEROS ---
    if estrategia['id'] == 'cobros_viajeros':
        df['Débitos'] = df['Monto_USD'].apply(lambda x: x if x > 0 else 0)
        df['Créditos'] = df['Monto_USD'].apply(lambda x: abs(x) if x < 0 else 0)
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Débitos', 'Créditos']
        cols_sum = ['Débitos', 'Créditos']
        titulo = 'Detalle de Movimientos Conciliados por Viajero (NIT)'
        fmt_moneda = formatos['usd']
        fmt_total = formatos['total_usd']
        
    # --- 2. OTRAS CUENTAS POR PAGAR ---
    elif estrategia['id'] == 'otras_cuentas_por_pagar':
        df['Monto Bs.'] = df['Monto_BS']
        columnas = ['Fecha', 'Descripcion NIT', 'Numero_Envio', 'Monto Bs.']
        cols_sum = ['Monto Bs.']
        titulo = 'Detalle de Movimientos Conciliados por Proveedor y Envío'
        
    # --- 3. DEUDORES EMPLEADOS (ME y BS) ---
    elif estrategia['id'] in ['deudores_empleados_me', 'deudores_empleados_bs']:
        is_usd = estrategia['id'] == 'deudores_empleados_me'
        col_origen = 'Monto_USD' if is_usd else 'Monto_BS'
        fmt_moneda = formatos['usd'] if is_usd else formatos['bs']
        fmt_total = formatos['total_usd'] if is_usd else formatos['total_bs']
        
        df['Débitos'] = df[col_origen].apply(lambda x: x if x > 0 else 0)
        df['Créditos'] = df[col_origen].apply(lambda x: abs(x) if x < 0 else 0)
        
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Débitos', 'Créditos', 'Saldo']
        cols_sum = ['Débitos', 'Créditos']
        titulo = 'Detalle de Movimientos Conciliados por Empleado'
        mostrar_saldo_linea = True
        col_saldo_idx = 5

    # --- 4. HABERES DE CLIENTES ---
    elif estrategia['id'] == 'haberes_clientes':
        df['Monto Bs.'] = df['Monto_BS']
        # Usamos los nombres personalizados que pediste
        columnas = ['Fecha', 'Fuente', 'Referencia', 'Monto Bs.'] 
        cols_sum = ['Monto Bs.']
        titulo = 'Detalle de Movimientos Conciliados por Cliente (NIT)'
        # Mapeo de nombres para el writer abajo
        # (Fecha y Fuente ya se llaman así en el DF, no necesitamos mapeo especial aquí,
        #  pero visualmente en el Excel el header será 'Fecha' y 'Fuente')

    # --- 5. CDC FACTORING ---
    elif estrategia['id'] == 'cdc_factoring':
        df['Débitos'] = df['Monto_USD'].apply(lambda x: x if x > 0 else 0)
        df['Créditos'] = df['Monto_USD'].apply(lambda x: abs(x) if x < 0 else 0)
        columnas = ['Fecha', 'Contrato', 'Fuente', 'Referencia', 'Débitos', 'Créditos']
        cols_sum = ['Débitos', 'Créditos']
        titulo = 'Detalle de Movimientos Conciliados por NIT (Factoring)'
        fmt_moneda = formatos['usd']
        fmt_total = formatos['total_usd']

    # --- 6. PROVEEDORES COSTOS ---
    elif estrategia['id'] == 'proveedores_costos':
        df['Débitos'] = df['Monto_USD'].apply(lambda x: x if x > 0 else 0)
        df['Créditos'] = df['Monto_USD'].apply(lambda x: abs(x) if x < 0 else 0)
        df['Monto Bs.'] = df['Monto_BS'] # Columna informativa
        columnas = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Débitos', 'Créditos', 'Monto Bs.']
        cols_sum = ['Débitos', 'Créditos', 'Monto Bs.']
        titulo = 'Detalle de Movimientos Conciliados por Proveedor (USD/BS)'
        fmt_moneda = formatos['usd']
        fmt_total = formatos['total_usd']    

    # --------------------------------------------------
    # Definimos el orden por defecto para no afectar a las otras cuentas
    criterios_orden = ['NIT', 'Fecha']
    
    # Solo si es la cuenta de Costos Y existe la columna de embarque, cambiamos el orden
    if estrategia['id'] == 'proveedores_costos' and 'Numero_Embarque' in df.columns:
        criterios_orden = ['NIT', 'Numero_Embarque', 'Fecha']
        
    df = df.sort_values(by=criterios_orden, ascending=True)

    # --- PROCESO DE ESCRITURA ---
    # Encabezado Principal
    ws.merge_range(0, 0, 0, len(columnas)-1, titulo, formatos['encabezado_sub']) # Ajustado len -1 para merge correcto
    current_row = 2
    
    grand_totals = {c: 0.0 for c in cols_sum}
    
    for nit, grupo in df.groupby('NIT'):
        col_nombre = 'Descripcion NIT' if 'Descripcion NIT' in grupo.columns else 'Nombre del Proveedor'
        nombre = grupo[col_nombre].iloc[0] if not grupo.empty and col_nombre in grupo else 'NO DEFINIDO'
        
        ws.merge_range(current_row, 0, current_row, len(columnas)-1, f"NIT: {nit} - {nombre}", formatos['proveedor_header'])
        current_row += 1
        ws.write_row(current_row, 0, columnas, formatos['header_tabla'])
        current_row += 1
        
        sum_deb = 0
        sum_cre = 0

        for _, row in grupo.iterrows():
            for c_idx, col_name in enumerate(columnas):
                val = row.get(col_name)
                
                # Verificación de sangría estricta aquí:
                if col_name == 'Fecha' and pd.notna(val): 
                    ws.write_datetime(current_row, c_idx, val, formatos['fecha'])
                elif col_name in ['Débitos', 'Créditos']: 
                    ws.write_number(current_row, c_idx, val, fmt_moneda)
                elif col_name == 'Monto Bs.': 
                    ws.write_number(current_row, c_idx, val, formatos['bs'])
                elif col_name == 'Saldo': 
                    pass
                else: 
                    ws.write(current_row, c_idx, val if pd.notna(val) else '')
            
            if mostrar_saldo_linea:
                sum_deb += row.get('Débitos', 0)
                sum_cre += row.get('Créditos', 0)

            current_row += 1
        
        # Subtotal
        lbl_col = len(columnas) - len(cols_sum) - (1 if mostrar_saldo_linea else 1)
        if mostrar_saldo_linea: lbl_col -= 1

        ws.write(current_row, lbl_col, "Subtotal", formatos['subtotal_label'])
        
        for i, c_sum in enumerate(cols_sum):
            suma = grupo[c_sum].sum()
            grand_totals[c_sum] += suma
            col_idx_sum = lbl_col + 1 + i
            ws.write_number(current_row, col_idx_sum, suma, fmt_moneda)
        
        if mostrar_saldo_linea:
            saldo_neto = sum_deb - sum_cre
            ws.write_number(current_row, col_saldo_idx, saldo_neto, fmt_total)

        current_row += 2

    # Totales Generales
    lbl_col_tot = len(columnas) - len(cols_sum) - (1 if mostrar_saldo_linea else 1)
    if mostrar_saldo_linea: lbl_col_tot -= 1

    ws.write(current_row, lbl_col_tot, "TOTALES", formatos['total_label'])
    for i, c_sum in enumerate(cols_sum):
        col_idx_sum = lbl_col_tot + 1 + i
        ws.write_number(current_row, col_idx_sum, grand_totals[c_sum], fmt_total)
    
    if mostrar_saldo_linea:
        neto_global = grand_totals.get('Débitos', 0) - grand_totals.get('Créditos', 0)
        ws.write_number(current_row, col_saldo_idx, neto_global, fmt_total)
        
    ws.set_column('A:F', 18)

def _generar_hoja_resumen_devoluciones(workbook, formatos, df_saldos):
    """Hoja extra específica para Devoluciones a Proveedores."""
    ws = workbook.add_worksheet("Resumen por Proveedor")
    ws.hide_gridlines(2)
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

def _generar_hoja_pendientes_resumida(workbook, formatos, df_saldos, estrategia, casa, fecha_maxima):
    """
    Genera una hoja de saldos RESUMIDA (una línea por NIT).
    CAMBIOS: Sin columna Fecha, Sin líneas de división.
    """
    nombre_hoja = estrategia.get("nombre_hoja_excel", "Saldos Por Empleado")
    ws = workbook.add_worksheet(nombre_hoja)
    ws.hide_gridlines(2) # <--- Ocultar celdas de fondo
    
    # Encabezados
    if pd.notna(fecha_maxima):
        ultimo_dia = fecha_maxima + pd.offsets.MonthEnd(0)
        meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        txt_fecha = f"PARA EL {ultimo_dia.day} DE {meses[ultimo_dia.month].upper()} DE {ultimo_dia.year}"
    else:
        txt_fecha = "FECHA NO DISPONIBLE"

    ws.merge_range('A1:F1', casa, formatos['encabezado_empresa'])
    ws.merge_range('A2:F2', f"ESPECIFICACION DE LA CUENTA {estrategia['nombre_hoja_excel']}", formatos['encabezado_sub'])
    ws.merge_range('A3:F3', txt_fecha, formatos['encabezado_sub'])

    # --- CAMBIO: Eliminada columna FECHA ---
    # Antes: ['SUB-CTA', 'NIT', 'NOMBRE', '$', 'FECHA', 'Bs.', 'Tasa']
    headers = ['SUB-CTA', 'NIT', 'NOMBRE', '$', 'Bs.', 'Tasa']
    ws.write_row('A5', headers, formatos['header_tabla'])

    if df_saldos.empty: return

    # Lógica de Agrupación
    col_nombre = 'Descripcion NIT' if 'Descripcion NIT' in df_saldos.columns else 'Nombre del Proveedor'
    if col_nombre not in df_saldos.columns:
        df_saldos['Nombre_Final'] = 'NO DEFINIDO'
    else:
        df_saldos['Nombre_Final'] = df_saldos[col_nombre].fillna('NO DEFINIDO')

    resumen = df_saldos.groupby('NIT').agg({
        'Nombre_Final': 'first',
        'Monto_USD': 'sum',
        'Monto_BS': 'sum'
    }).reset_index()

    resumen['Tasa_Impl'] = np.where(
        resumen['Monto_USD'].abs() > 0.01, 
        (resumen['Monto_BS'] / resumen['Monto_USD']).abs(), 
        0
    )

    # Escritura
    current_row = 5
    sub_cta = estrategia['nombre_hoja_excel'].split('.')[-1][:4]

    for _, row in resumen.iterrows():
        ws.write(current_row, 0, sub_cta, formatos['encabezado_sub'])
        ws.write(current_row, 1, row['NIT'])
        ws.write(current_row, 2, row['Nombre_Final'])
        ws.write_number(current_row, 3, row['Monto_USD'], formatos['usd'])
        # Eliminada columna fecha, rodamos índices
        ws.write_number(current_row, 4, row['Monto_BS'], formatos['bs'])
        ws.write_number(current_row, 5, row['Tasa_Impl'], formatos['tasa'])
        current_row += 1

    # Totales
    ws.write(current_row, 2, "TOTALES", formatos['total_label'])
    ws.write_number(current_row, 3, resumen['Monto_USD'].sum(), formatos['total_usd'])
    ws.write_number(current_row, 4, resumen['Monto_BS'].sum(), formatos['total_bs'])

    # Ajuste de anchos
    ws.set_column('A:A', 10)
    ws.set_column('B:B', 15)
    ws.set_column('C:C', 45)
    ws.set_column('D:F', 15)
    
def _generar_hoja_pendientes_cdc(workbook, formatos, df_saldos, estrategia, casa, fecha_maxima):
    """
    Genera hoja de pendientes para Factoring agrupada por NIT -> Contrato.
    Muestra subtotales por contrato como se solicitó.
    """
    ws = workbook.add_worksheet(estrategia.get("nombre_hoja_excel", "Pendientes"))
    ws.hide_gridlines(2)
    
    # 1. Encabezados del Reporte
    if pd.notna(fecha_maxima):
        ultimo_dia = fecha_maxima + pd.offsets.MonthEnd(0)
        meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        txt_fecha = f"PARA EL {ultimo_dia.day} DE {meses[ultimo_dia.month].upper()} DE {ultimo_dia.year}"
    else:
        txt_fecha = "FECHA NO DISPONIBLE"

    ws.merge_range('A1:H1', casa, formatos['encabezado_empresa'])
    ws.merge_range('A2:H2', f"ESPECIFICACIÓN DE LA CUENTA {estrategia['nombre_hoja_excel']}", formatos['encabezado_sub'])
    ws.merge_range('A3:H3', txt_fecha, formatos['encabezado_sub'])

    # Encabezados de la Tabla (Sin títulos en A y B para limpieza visual en filas de datos)
    headers = ['NIT', 'Descripción NIT', 'FECHA', 'CONTRATO', 'DOCUMENTO', 'MONEDA ($)', 'TASA', 'MONTO (Bs)']
    ws.write_row('A5', headers, formatos['header_tabla'])

    if df_saldos.empty: return

    # 3. Preparación de Datos
    df = df_saldos.copy()
    df['Monto_BS'] = pd.to_numeric(df['Monto_BS'], errors='coerce').fillna(0)
    df['Monto_USD'] = pd.to_numeric(df['Monto_USD'], errors='coerce').fillna(0)
    df['Tasa_Impl'] = np.where(df['Monto_USD'].abs() > 0.01, (df['Monto_BS'] / df['Monto_USD']).abs(), 0)
    
    # Detección de columnas
    col_nombre = None
    for col in ['Descripcion NIT', 'Descripción Nit', 'Nombre del Proveedor', 'Nombre']:
        if col in df.columns: col_nombre = col; break
    if not col_nombre: df['N'] = 'NO DEFINIDO'; col_nombre = 'N'
        
    col_nit = None
    for col in ['NIT', 'Nit', 'NIT_Normalizado']:
        if col in df.columns: col_nit = col; break
    if not col_nit: df['I'] = 'SIN_NIT'; col_nit = 'I'

    # Limpieza y Orden
    df[col_nombre] = df[col_nombre].astype(str).replace(r'^\s*$', 'NO DEFINIDO', regex=True).fillna('NO DEFINIDO')
    df[col_nit] = df[col_nit].astype(str).replace(r'^\s*$', 'SIN_NIT', regex=True).fillna('SIN_NIT')
    
    # Ordenamos por Proveedor, luego por Contrato y luego por Fecha
    df = df.sort_values(by=[col_nombre, 'Contrato', 'Fecha'])
    
    current_row = 5 
    grand_total_usd = 0
    grand_total_bs = 0

    # 4. BUCLE PRINCIPAL: Por Proveedor
    for (nombre_prov, nit_prov), grupo_prov in df.groupby([col_nombre, col_nit]):
        
        # Encabezado visual del PROVEEDOR
        ws.merge_range(current_row, 0, current_row, 7, f"{nit_prov} - {nombre_prov}", formatos['proveedor_header'])
        current_row += 1
        
        # 5. BUCLE ANIDADO: Por Contrato
        for contrato, grupo_contrato in grupo_prov.groupby('Contrato'):
            
            subtotal_contrato_usd = 0
            subtotal_contrato_bs = 0
            
            for _, row in grupo_contrato.iterrows():
                fecha = row['Fecha']
                fuente = row.get('Fuente', '') # Documento
                monto_usd = row['Monto_USD']
                monto_bs = row['Monto_BS']
                
                subtotal_contrato_usd += monto_usd
                subtotal_contrato_bs += monto_bs

                # Escribir fila
                ws.write(current_row, 0, nit_prov)
                ws.write(current_row, 1, nombre_prov)
                
                if pd.notna(fecha): ws.write_datetime(current_row, 2, fecha, formatos['fecha'])
                else: ws.write(current_row, 2, '-')
                
                ws.write(current_row, 3, contrato)
                ws.write(current_row, 4, fuente)
                ws.write_number(current_row, 5, monto_usd, formatos['usd'])
                ws.write_number(current_row, 6, row['Tasa_Impl'], formatos['tasa'])
                ws.write_number(current_row, 7, monto_bs, formatos['bs'])
                
                current_row += 1
            
            # TOTAL CONTRATO (Justo debajo del bloque)
            ws.write(current_row, 4, "Total Contrato", formatos['subtotal_label'])
            ws.write_number(current_row, 5, subtotal_contrato_usd, formatos['subtotal_usd'])
            # Opcional: Mostrar total Bs si se desea
            # ws.write_number(current_row, 7, subtotal_contrato_bs, formatos['subtotal_bs'])
            
            current_row += 2 # Espacio entre contratos
            
            # Acumulamos para el gran total
            grand_total_usd += subtotal_contrato_usd
            grand_total_bs += subtotal_contrato_bs

    # 6. Gran Total Final
    ws.write(current_row, 4, "TOTAL GENERAL", formatos['total_label'])
    ws.write_number(current_row, 5, grand_total_usd, formatos['total_usd'])
    ws.write_number(current_row, 7, grand_total_bs, formatos['total_bs'])
    
    # 7. Ajuste de Anchos
    ws.set_column('A:A', 15)
    ws.set_column('B:B', 40)
    ws.set_column('C:C', 12)
    ws.set_column('D:D', 18) # Contrato
    ws.set_column('E:E', 15) # Documento
    ws.set_column('F:H', 18)

def _generar_hoja_pendientes_corrida(workbook, formatos, df_saldos, estrategia, casa, fecha_maxima):
    """
    Genera hoja de pendientes como LISTADO CORRIDO (Cronológico).
    CORREGIDO: Formato de fecha, limpieza de filas basura y alias de columnas.
    """
    nombre_hoja = estrategia.get("nombre_hoja_excel", "Pendientes")
    ws = workbook.add_worksheet(nombre_hoja)
    ws.hide_gridlines(2)
    cols = estrategia["columnas_reporte"]
    
    # Encabezados
    if pd.notna(fecha_maxima):
        ultimo_dia = fecha_maxima + pd.offsets.MonthEnd(0)
        meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        txt_fecha = f"PARA EL {ultimo_dia.day} DE {meses[ultimo_dia.month].upper()} DE {ultimo_dia.year}"
    else:
        txt_fecha = "FECHA NO DISPONIBLE"

    ws.merge_range(0, 0, 0, len(cols)-1, casa, formatos['encabezado_empresa'])
    ws.merge_range(1, 0, 1, len(cols)-1, f"ESPECIFICACION DE LA CUENTA {estrategia['nombre_hoja_excel']}", formatos['encabezado_sub'])
    ws.merge_range(2, 0, 2, len(cols)-1, txt_fecha, formatos['encabezado_sub'])
    ws.write_row(4, 0, cols, formatos['header_tabla'])

    if df_saldos.empty: return

    df = df_saldos.copy()
    
    # --- 1. LIMPIEZA DE FILAS BASURA (Total Contabilidad, nan, etc) ---
    # Convertimos a string y mayúsculas para filtrar
    if 'NIT' in df.columns:
        mask_basura = df['NIT'].astype(str).str.upper().str.contains('CONTABILIDAD|TOTAL|NAN|SIN_NIT', na=False)
        # También verificamos si la descripción tiene "TOTAL"
        if 'Descripcion NIT' in df.columns:
            mask_basura |= df['Descripcion NIT'].astype(str).str.upper().str.contains('TOTAL', na=False)
            
        df = df[~mask_basura]
    # ------------------------------------------------------------------

    df['Monto Dólar'] = pd.to_numeric(df.get('Monto_USD'), errors='coerce').fillna(0)
    df['Bs.'] = pd.to_numeric(df.get('Monto_BS'), errors='coerce').fillna(0)
    df['Monto Bolivar'] = df['Bs.']
    df['Tasa'] = np.where(df['Monto Dólar'].abs() != 0, df['Bs.'].abs() / df['Monto Dólar'].abs(), 0)
    
    # ORDENAMIENTO
    if estrategia['id'] == 'haberes_clientes':
        df = df.sort_values(by=['Fecha', 'NIT'], ascending=[True, True])
    else:
        df = df.sort_values(by=['Fecha', 'Asiento'], ascending=[True, True])

    current_row = 5
    
    # Buscamos índices para el total
    # Creamos un DF dummy con las columnas para buscar índices
    dummy_df = pd.DataFrame(columns=cols)
    usd_idx = get_col_idx(dummy_df, ['Monto Dólar', 'Monto USD'])
    bs_idx = get_col_idx(dummy_df, ['Bs.', 'Monto Bolivar', 'Monto Bs'])
    
    for _, row in df.iterrows():
        for c_idx, col_name in enumerate(cols):
            
            # --- MAPEO DE ALIAS ---
            val = None
            if col_name == 'Fecha Origen Acreencia': val = row.get('Fecha')
            elif col_name == 'Numero de Documento': val = row.get('Fuente')
            else: val = row.get(col_name)
            # ----------------------
            
            # Escritura con formato
            if col_name in ['Fecha', 'Fecha Origen Acreencia']:
                # Asegurar que es fecha válida para Excel
                val_dt = pd.to_datetime(val, errors='coerce')
                if pd.notna(val_dt):
                    ws.write_datetime(current_row, c_idx, val_dt, formatos['fecha'])
                else:
                    ws.write(current_row, c_idx, str(val) if val else "", formatos['text'])
                    
            elif col_name in ['Monto Dólar', 'Monto USD']: 
                ws.write_number(current_row, c_idx, val or 0, formatos['usd'])
            elif col_name in ['Bs.', 'Monto Bolivar', 'Monto Bs']: 
                ws.write_number(current_row, c_idx, val or 0, formatos['bs'])
            elif col_name == 'Tasa': 
                ws.write_number(current_row, c_idx, val or 0, formatos['tasa'])
            else: 
                ws.write(current_row, c_idx, val if pd.notna(val) else '')
        current_row += 1
        
    # SALDO TOTAL AL FINAL
    current_row += 1
    # Ubicar etiqueta antes de los montos
    indices_montos = [i for i in [usd_idx, bs_idx] if i != -1]
    lbl_idx = max(0, min(indices_montos) - 1) if indices_montos else 0
    
    ws.write(current_row, lbl_idx, "SALDO TOTAL", formatos['total_label'])
    
    if usd_idx != -1: ws.write_number(current_row, usd_idx, df['Monto Dólar'].sum(), formatos['total_usd'])
    if bs_idx != -1: ws.write_number(current_row, bs_idx, df['Bs.'].sum(), formatos['total_bs'])

    # Ajuste anchos
    ws.set_column(0, 0, 18)
    ws.set_column(1, 1, 55)
    ws.set_column(2, 2, 15)
    ws.set_column(3, 10, 20)

def _generar_hoja_pendientes_proveedores(workbook, formatos, df_saldos, estrategia, casa, fecha_maxima):
    """
    Hoja 1: Resumen de saldos abiertos por Embarque.
    La descripción mostrada es la referencia del movimiento ACREEDOR (monto negativo).
    """
    ws = workbook.add_worksheet(estrategia.get("nombre_hoja_excel", "Pendientes"))
    ws.hide_gridlines(2)
    
    # 1. ENCABEZADOS DE REPORTE
    if pd.notna(fecha_maxima):
        meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        txt_fecha = f"PARA EL {fecha_maxima.day} DE {meses[fecha_maxima.month].upper()} DE {fecha_maxima.year}"
    else: txt_fecha = "FECHA NO DISPONIBLE"

    ws.merge_range('A1:E1', casa, formatos['encabezado_empresa'])
    ws.merge_range('A2:E2', f"ESPECIFICACION DE LA CUENTA {estrategia['nombre_hoja_excel']}", formatos['encabezado_sub'])
    ws.merge_range('A3:E3', txt_fecha, formatos['encabezado_sub'])
    ws.write_row('A5', ['NIT', 'Referencia / Descripción', 'Embarque', 'Saldo USD', 'Saldo Bs.'], formatos['header_tabla'])

    df = df_saldos[df_saldos['Conciliado'] == False].copy()
    if df.empty: return

    # 2. PREPARACIÓN (CORREGIDA PARA DETECTAR EL NOMBRE CORRECTAMENTE)
    col_nit = 'NIT_Reporte' if 'NIT_Reporte' in df.columns else 'NIT_Norm'
    
    # Lista de posibles nombres que la herramienta pudo haber asignado
    posibles_nombres = ['Descripcion NIT', 'Descripción Nit', 'Descripcion Nit', 'DESCRIPCION NIT', 'Nombre del Proveedor']
    col_nombre = None
    
    for c in posibles_nombres:
        if c in df.columns:
            col_nombre = c
            break
    
    # Si no encontró ninguna de las anteriores, usamos el NIT como emergencia
    if not col_nombre:
        col_nombre = col_nit

    df[col_nit] = df[col_nit].astype(str).replace(['nan', 'NaN', 'None', 'ND', '0'], 'SIN NIT')
    df[col_nombre] = df[col_nombre].astype(str).replace(['nan', 'NaN', 'None', '0', ''], 'PROVEEDOR NO IDENTIFICADO')

    fmt_header_prov = workbook.add_format({'bold': True, 'bg_color': '#FFFFFF', 'bottom': 1})
    current_row = 5
    gran_total_usd = 0
    gran_total_bs = 0

    # 3. BUCLE POR PROVEEDOR
    for nit_val, grupo_prov in df.groupby(col_nit, sort=False):
        
        if abs(round(grupo_prov['Monto_USD'].sum(), 2)) <= 0.01:
            continue

        # Encabezado del Proveedor
        nombre_disp = grupo_prov[col_nombre].iloc[0]
        ws.write(current_row, 0, nit_val, fmt_header_prov)
        ws.write(current_row, 1, nombre_disp, fmt_header_prov)
        ws.write_row(current_row, 2, ["", "", ""], fmt_header_prov)
        current_row += 1

        # 4. BUCLE POR EMBARQUE (Agrupación para la fila única)
        for emb_id, grupo_emb in grupo_prov.groupby('Numero_Embarque'):
            s_usd = grupo_emb['Monto_USD'].sum()
            s_bs = grupo_emb['Monto_BS'].sum()

            if abs(round(s_usd, 2)) > 0.01:
                # --- LÓGICA DE REFERENCIA ACREEDORA ---
                # Buscamos la fila donde el monto sea negativo (Acreedor)
                fila_acreedora = grupo_emb[grupo_emb['Monto_USD'] < 0]
                
                if not fila_acreedora.empty:
                    # Usamos la referencia de la primera fila negativa encontrada
                    ref_acreedora = fila_acreedora.iloc[0]['Referencia']
                else:
                    # Fallback: Si por alguna razón no hay negativos, usamos la primera disponible
                    ref_acreedora = grupo_emb.iloc[0]['Referencia']
                
                ws.write(current_row, 0, "", formatos['text']) 
                ws.write(current_row, 1, str(ref_acreedora), formatos['text'])
                ws.write(current_row, 2, emb_id if emb_id != 'NO_EMB' else "", formatos['text'])
                ws.write_number(current_row, 3, s_usd, formatos['usd'])
                ws.write_number(current_row, 4, s_bs, formatos['bs'])
                
                gran_total_usd += s_usd
                gran_total_bs += s_bs
                current_row += 1

        current_row += 1 # Espacio entre proveedores

    # 5. TOTAL GENERAL
    current_row += 1
    ws.write(current_row, 2, "TOTAL GENERAL", formatos['total_label'])
    ws.write_number(current_row, 3, gran_total_usd, formatos['total_usd'])
    ws.write_number(current_row, 4, gran_total_bs, formatos['total_bs'])
    
    # Ajuste de anchos
    ws.set_column('A:A', 15) # NIT
    ws.set_column('B:B', 55) # Referencia
    ws.set_column('C:C', 15) # Embarque
    ws.set_column('D:E', 18) # Montos

def _generar_hoja_detalle_especificacion_proveedores(workbook, formatos, df_saldos):
    """
    Hoja 2: Detalle analítico con fila de totales al final de cada embarque.
    Incluye validación de fechas NaT.
    """
    ws = workbook.add_worksheet("Detalle Especificacion")
    ws.hide_gridlines(2)
    
    df = df_saldos[df_saldos['Conciliado'] == False].copy()
    if df.empty: return
    
    df = df.sort_values(by=['NIT_Reporte', 'Numero_Embarque', 'Fecha'])
    
    columnas = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Monto USD', 'Monto Bs.']
    ws.merge_range(0, 0, 0, 5, "DETALLE ANALÍTICO DE PARTIDAS PENDIENTES", formatos['encabezado_sub'])
    
    current_row = 2
    fmt_total_emb = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'top': 1, 'num_format': '#,##0.00'})

    for (nit, emb), grupo in df.groupby(['NIT_Reporte', 'Numero_Embarque'], sort=False):
        ws.merge_range(current_row, 0, current_row, 5, f"NIT: {nit} | EMBARQUE: {emb}", formatos['proveedor_header'])
        current_row += 1
        ws.write_row(current_row, 0, columnas, formatos['header_tabla'])
        current_row += 1
        
        for _, row in grupo.iterrows():
            # Validación de Fecha NaT
            fec = row.get('Fecha')
            if pd.notna(fec): ws.write_datetime(current_row, 0, fec, formatos['fecha'])
            else: ws.write(current_row, 0, '-', formatos['text'])
            
            ws.write(current_row, 1, str(row.get('Asiento', '')), formatos['text'])
            ws.write(current_row, 2, str(row.get('Referencia', '')), formatos['text'])
            ws.write(current_row, 3, str(row.get('Fuente', '')), formatos['text'])
            ws.write_number(current_row, 4, row['Monto_USD'], formatos['usd'])
            ws.write_number(current_row, 5, row['Monto_BS'], formatos['bs'])
            current_row += 1
        
        # --- FILA DE TOTALIZACIÓN DEL GRUPO (NUEVO) ---
        ws.write(current_row, 3, "Total Embarque:", formatos['subtotal_label'])
        ws.write_number(current_row, 4, grupo['Monto_USD'].sum(), fmt_total_emb)
        ws.write_number(current_row, 5, grupo['Monto_BS'].sum(), fmt_total_emb)
        current_row += 2

    ws.set_column('A:F', 18); ws.set_column('C:C', 40)


#@st.cache_data
def generar_reporte_excel(_df_full, df_saldos_abiertos, df_conciliados, _estrategia, casa_seleccionada, cuenta_seleccionada):
    """
    Controlador principal que orquesta la creación del Excel.
    Nombres corregidos para evitar NameError.
    """
    output_excel = BytesIO()
    
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        workbook = writer.book
        formatos = _crear_formatos(workbook)
        fecha_max = _df_full['Fecha'].dropna().max()
        
        # ============================================================
        # 1. SELECCIÓN DE HOJA DE PENDIENTES
        # ============================================================
        cuentas_resumen = ['deudores_empleados_me', 'deudores_empleados_bs']
        cuentas_corridas = ['fondos_transito', 'fondos_depositar', 'haberes_clientes']

        if _estrategia['id'] == 'proveedores_costos':
            # HOJA 1: RESUMEN (Nombre corregido aquí)
            _generar_hoja_pendientes_proveedores(workbook, formatos, df_saldos_abiertos, _estrategia, casa_seleccionada, fecha_max)
            
            # NUEVA HOJA: DETALLE ANALÍTICO
            _generar_hoja_detalle_especificacion_proveedores(workbook, formatos, df_saldos_abiertos)
            
        elif _estrategia['id'] in cuentas_resumen:
            _generar_hoja_pendientes_resumida(workbook, formatos, df_saldos_abiertos, _estrategia, casa_seleccionada, fecha_max)
            
        elif _estrategia['id'] in cuentas_corridas:
            _generar_hoja_pendientes_corrida(workbook, formatos, df_saldos_abiertos, _estrategia, casa_seleccionada, fecha_max)
            
        elif _estrategia['id'] == 'cdc_factoring':
            _generar_hoja_pendientes_cdc(workbook, formatos, df_saldos_abiertos, _estrategia, casa_seleccionada, fecha_max)
            
        else:
            _generar_hoja_pendientes(workbook, formatos, df_saldos_abiertos, _estrategia, casa_seleccionada, fecha_max)
        
        # ============================================================
        # 2. SELECCIÓN DE HOJA DE CONCILIADOS
        # ============================================================
        if _estrategia['id'] in cuentas_resumen:
            datos_conciliacion = _df_full.copy()
        else:
            datos_conciliacion = df_conciliados.copy()

        if not datos_conciliacion.empty:
            # Filtro para que la Hoja 2 no tenga los ajustes menores de 1$
            if _estrategia['id'] == 'proveedores_costos':
                datos_h2 = datos_conciliacion[~datos_conciliacion['Grupo_Conciliado'].astype(str).str.contains('REQUIERE_AJUSTE', na=False)]
            else:
                datos_h2 = datos_conciliacion

            if not datos_h2.empty:
                cuentas_agrupadas_conc = [
                    'cobros_viajeros', 
                    'otras_cuentas_por_pagar', 
                    'deudores_empleados_me',
                    'deudores_empleados_bs',
                    'haberes_clientes',
                    'cdc_factoring',
                    'proveedores_costos'
                ]
                
                if _estrategia['id'] in cuentas_agrupadas_conc:
                    _generar_hoja_conciliados_agrupada(workbook, formatos, datos_h2, _estrategia)
                else:
                    _generar_hoja_conciliados_estandar(workbook, formatos, datos_h2, _estrategia)

        # ============================================================
        # 3. HOJA DE AJUSTES (Para Costos)
        # ============================================================
        if _estrategia['id'] == 'proveedores_costos' and not df_conciliados.empty:
            df_ajustes = df_conciliados[df_conciliados['Grupo_Conciliado'].astype(str).str.contains('REQUIERE_AJUSTE', na=False)]
            if not df_ajustes.empty:
                _generar_hoja_ajustes_menores(workbook, formatos, df_ajustes)

        # ============================================================
        # 4. HOJAS ADICIONALES (Devoluciones)
        # ============================================================
        if _estrategia['id'] == 'devoluciones_proveedores' and not df_saldos_abiertos.empty:
            _generar_hoja_resumen_devoluciones(workbook, formatos, df_saldos_abiertos)

    return output_excel.getvalue()
    
def _generar_hoja_ajustes_menores(workbook, formatos, df_ajustes):
    """
    Genera la 3ra hoja incluyendo la columna de Bolívares a modo informativo.
    """
    ws = workbook.add_worksheet("Para Asiento de Ajuste")
    ws.hide_gridlines(2)
    
    df = df_ajustes.sort_values(by=['NIT_Reporte', 'Numero_Embarque', 'Fecha'])
    # Añadimos Monto Bs. a los encabezados
    columnas = ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Monto USD', 'Monto Bs.']
    ws.merge_range(0, 0, 0, 5, "EMBARQUES PENDIENTES POR AJUSTE MENOR A 1$", formatos['encabezado_sub'])
    
    current_row = 2
    fmt_diff_usd = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'num_format': '$#,##0.00', 'border': 1})
    fmt_diff_bs = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'num_format': '#,##0.00', 'border': 1})
    
    total_ajustes_usd = 0
    total_ajustes_bs = 0

    for (nit, emb), grupo in df.groupby(['NIT_Reporte', 'Numero_Embarque'], sort=False):
        diferencia_usd = round(grupo['Monto_USD'].sum(), 2)
        diferencia_bs = round(grupo['Monto_BS'].sum(), 2)

        ws.merge_range(current_row, 0, current_row, 5, f"NIT: {nit} | EMBARQUE: {emb}", formatos['proveedor_header'])
        current_row += 1
        ws.write_row(current_row, 0, columnas, formatos['header_tabla'])
        current_row += 1
        
        for _, row in grupo.iterrows():
            fec = row.get('Fecha')
            if pd.notna(fec): ws.write_datetime(current_row, 0, fec, formatos['fecha'])
            else: ws.write(current_row, 0, '-', formatos['text'])

            ws.write(current_row, 1, str(row.get('Asiento', '')), formatos['text'])
            ws.write(current_row, 2, str(row.get('Referencia', '')), formatos['text'])
            ws.write(current_row, 3, str(row.get('Fuente', '')), formatos['text'])
            ws.write_number(current_row, 4, row['Monto_USD'], formatos['usd'])
            ws.write_number(current_row, 5, row['Monto_BS'], formatos['bs']) # Nueva columna
            current_row += 1
        
        ws.write(current_row, 3, "DIFERENCIA A AJUSTAR:", formatos['subtotal_label'])
        ws.write_number(current_row, 4, diferencia_usd, fmt_diff_usd)
        ws.write_number(current_row, 5, diferencia_bs, fmt_diff_bs) # Diferencia en Bs.
        
        total_ajustes_usd += diferencia_usd
        total_ajustes_bs += diferencia_bs
        current_row += 2

    current_row += 1
    ws.write(current_row, 3, "TOTAL GENERAL AJUSTES:", formatos['total_label'])
    ws.write_number(current_row, 4, total_ajustes_usd, formatos['total_usd'])
    ws.write_number(current_row, 5, total_ajustes_bs, formatos['total_bs'])

    ws.set_column('A:B', 15)
    ws.set_column('C:C', 40)
    ws.set_column('D:F', 18)

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

def generar_reporte_paquete_cc(df_analizado, nombre_casa):
    """
    Genera reporte de análisis de Paquete CC.
    Versión actualizada: Elimina columna 'Nombre', mantiene 'NIT'.
    """
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- CÁLCULO DEL TÍTULO DINÁMICO ---
        if 'Fecha' in df_analizado.columns and not df_analizado['Fecha'].empty:
            fecha_max = pd.to_datetime(df_analizado['Fecha']).max()
            meses_es = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
            texto_fecha = f"{meses_es[fecha_max.month]} {fecha_max.year}"
        else:
            texto_fecha = "PERIODO NO DEFINIDO"
        titulo_reporte = f"Análisis de Asientos de Cuentas por Cobrar {nombre_casa} {texto_fecha}"

        # --- ESTILOS ---
        main_title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 16})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        text_format = workbook.add_format({'border': 1})
        money_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        incidencia_text_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
        incidencia_money_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '#,##0.00', 'border': 1})
        incidencia_date_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': 'dd/mm/yyyy', 'border': 1})
        header_corr_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#BDD7EE', 'border': 1, 'align': 'center'})
        descriptive_title_format = workbook.add_format({'bold': True, 'font_size': 14, 'fg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        subgroup_title_format = workbook.add_format({'bold': True, 'font_size': 11, 'fg_color': '#E0E0E0', 'border': 1})
        total_label_format = workbook.add_format({'bold': True, 'align': 'right', 'top': 2, 'font_color': '#003366'})
        total_money_format = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 2, 'bottom': 1})

        columnas_reporte = [
            'Asiento',                  # 0
            'Fecha',                    # 1
            'NIT',                      # 2 
            'Fuente',                   # 3
            'Cuenta Contable',          # 4
            'Descripción de Cuenta',    # 5
            'Referencia',               # 6
            'Débito Dolar',             # 7
            'Crédito Dolar',            # 8
            'Débito VES',               # 9
            'Crédito VES'               # 10
        ]
        
        df_analizado['Grupo Principal'] = df_analizado['Grupo'].apply(lambda x: x.split(':')[0].strip())
        grupos_principales_ordenados = sorted(df_analizado['Grupo Principal'].unique(), key=lambda x: (0, int(x.split()[1])) if x.startswith('Grupo') else (1, x))
        
        # --- HOJA 1: DIRECTORIO ---
        ws_dir = workbook.add_worksheet("Directorio")
        ws_dir.merge_range('A1:C1', titulo_reporte, main_title_format) 
        ws_dir.write('A2', 'Nombre de la Hoja', header_format)
        ws_dir.write('B2', 'Descripción del Contenido', header_format)
        ws_dir.write('C2', 'Observaciones', header_format)
        dir_row = 2
        for grupo_principal in grupos_principales_ordenados:
            sheet_name = re.sub(r'[\\/*?:"\[\]]', '', grupo_principal)[:31]
            df_grupo_dir = df_analizado[df_analizado['Grupo Principal'] == grupo_principal]
            full_name_example = df_grupo_dir['Grupo'].iloc[0]
            description = full_name_example.split(':', 1)[-1].strip() if ':' in full_name_example else full_name_example
            if grupo_principal in ["Grupo 3", "Grupo 9", "Grupo 8", "Grupo 6", "Grupo 7"]: description = f"{description.split('-')[0].strip()} (Varios Subgrupos)"
            
            tiene_error = False
            for est in df_grupo_dir['Estado']:
                if not str(est).startswith('Conciliado'):
                    tiene_error = True; break
            observacion = "Incidencia Encontrada" if tiene_error else "Conciliado"
            ws_dir.write(dir_row, 0, sheet_name, text_format)
            ws_dir.write(dir_row, 1, description, text_format)
            ws_dir.write(dir_row, 2, observacion, text_format)
            dir_row += 1
        ws_dir.set_column('A:A', 25); ws_dir.set_column('B:B', 60); ws_dir.set_column('C:C', 25)

        # --- HOJA 2: LISTADO CORRELATIVO ---
        ws_corr = workbook.add_worksheet("Listado Correlativo")
        ws_corr.merge_range('A1:G1', titulo_reporte, main_title_format) 
        cols_corr = ['Asiento', 'Estado Global', 'Grupo', 'Fecha', 'Fuente', 'Total Asiento ($)', 'Total Asiento (Bs)']
        ws_corr.write_row('A2', cols_corr, header_corr_format)
        ws_corr.freeze_panes(2, 0)
        resumen_data = []
        for asiento, grupo in df_analizado.groupby('Asiento'):
            tiene_incidencia = False
            for est in grupo['Estado']:
                if not str(est).startswith('Conciliado'):
                    tiene_incidencia = True; break
            estado_final = "Incidencia (Revisar)" if tiene_incidencia else grupo['Estado'].iloc[0]
            primera_fila = grupo.iloc[0]
            resumen_data.append({
                'Asiento': asiento, 'Estado Global': estado_final, 'Grupo': primera_fila['Grupo'].split(':')[0],
                'Fecha': primera_fila['Fecha'], 'Fuente': primera_fila['Fuente'],
                'Total Asiento ($)': grupo['Débito Dolar'].sum(), 'Total Asiento (Bs)': grupo['Débito VES'].sum()
            })
        df_resumen = pd.DataFrame(resumen_data).sort_values('Asiento')
        curr_row = 2
        for _, row in df_resumen.iterrows():
            is_incidencia = not str(row['Estado Global']).startswith('Conciliado')
            fmt_txt = incidencia_text_format if is_incidencia else text_format
            fmt_num = incidencia_money_format if is_incidencia else money_format
            fmt_date = incidencia_date_format if is_incidencia else date_format
            ws_corr.write(curr_row, 0, row['Asiento'], fmt_txt)
            ws_corr.write(curr_row, 1, row['Estado Global'], fmt_txt)
            ws_corr.write(curr_row, 2, row['Grupo'], fmt_txt)
            ws_corr.write_datetime(curr_row, 3, row['Fecha'], fmt_date)
            ws_corr.write(curr_row, 4, row['Fuente'], fmt_txt)
            ws_corr.write_number(curr_row, 5, row['Total Asiento ($)'], fmt_num)
            ws_corr.write_number(curr_row, 6, row['Total Asiento (Bs)'], fmt_num)
            curr_row += 1
        ws_corr.set_column('A:A', 15); ws_corr.set_column('B:B', 20); ws_corr.set_column('C:C', 15)
        ws_corr.set_column('D:E', 12); ws_corr.set_column('F:G', 18)

        # --- HOJAS 3...N: DETALLE POR GRUPOS ---
        for grupo_principal_nombre in grupos_principales_ordenados:
            sheet_name = re.sub(r'[\\/*?:"\[\]]', '', grupo_principal_nombre)[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.hide_gridlines(2)
            # Combinar hasta K (se redujo una columna)
            ws.merge_range('A1:K1', titulo_reporte, main_title_format) 
            
            df_grupo_completo = df_analizado[df_analizado['Grupo Principal'] == grupo_principal_nombre]
            subgrupos = sorted(df_grupo_completo['Grupo'].unique())
            full_descriptive_title = subgrupos[0]
            if len(subgrupos) > 1: full_descriptive_title = f"{subgrupos[0].split(':')[0].strip()}: {subgrupos[0].split(':')[1].split('-')[0].strip()}"
            ws.merge_range('A3:K3', full_descriptive_title, descriptive_title_format)
            current_row = 4
            for subgrupo_nombre in subgrupos:
                df_subgrupo = df_grupo_completo[df_grupo_completo['Grupo'] == subgrupo_nombre]
                if len(subgrupos) > 1: ws.merge_range(current_row, 0, current_row, len(columnas_reporte) - 1, subgrupo_nombre, subgroup_title_format); current_row += 1
                ws.write_row(current_row, 0, columnas_reporte, header_format); current_row += 1
                start_data_row = current_row
                for _, row_data in df_subgrupo.iterrows():
                    estado_fila = str(row_data.get('Estado', 'Conciliado'))
                    is_incidencia = not estado_fila.startswith('Conciliado')
                    
                    fmt_txt = incidencia_text_format if is_incidencia else text_format
                    fmt_num = incidencia_money_format if is_incidencia else money_format
                    fmt_date = incidencia_date_format if is_incidencia else date_format
                    
                    # --- CAMBIO: Escritura de columnas reajustada (Sin Nombre) ---
                    ws.write(current_row, 0, row_data.get('Asiento', ''), fmt_txt)
                    ws.write_datetime(current_row, 1, row_data.get('Fecha', None), fmt_date)
                    ws.write(current_row, 2, row_data.get('NIT', ''), fmt_txt)
                    # Col 3 ya no es Nombre, ahora es Fuente
                    ws.write(current_row, 3, row_data.get('Fuente', ''), fmt_txt)
                    ws.write(current_row, 4, row_data.get('Cuenta Contable', ''), fmt_txt)
                    ws.write(current_row, 5, row_data.get('Descripción de Cuenta', ''), fmt_txt)
                    ws.write(current_row, 6, row_data.get('Referencia', ''), fmt_txt)
                    ws.write_number(current_row, 7, row_data.get('Débito Dolar', 0), fmt_num)
                    ws.write_number(current_row, 8, row_data.get('Crédito Dolar', 0), fmt_num)
                    ws.write_number(current_row, 9, row_data.get('Débito VES', 0), fmt_num)
                    ws.write_number(current_row, 10, row_data.get('Crédito VES', 0), fmt_num)
                    current_row += 1
                if not df_subgrupo.empty:
                    # Ajuste de totales (Columna 6 es referencia, montos empiezan en 7/H)
                    ws.write(current_row, 6, f'TOTALES {subgrupo_nombre.split(":")[-1].strip()}', total_label_format)
                    ws.write_formula(current_row, 7, f'=SUM(H{start_data_row + 1}:H{current_row})', total_money_format) # Deb $
                    ws.write_formula(current_row, 8, f'=SUM(I{start_data_row + 1}:I{current_row})', total_money_format) # Cre $
                    ws.write_formula(current_row, 9, f'=SUM(J{start_data_row + 1}:J{current_row})', total_money_format) # Deb Bs
                    ws.write_formula(current_row, 10, f'=SUM(K{start_data_row + 1}:K{current_row})', total_money_format) # Cre Bs
                    current_row += 2
            
            # Ajuste de anchos final
            ws.set_column('A:A', 12) # Asiento
            ws.set_column('B:B', 12) # Fecha
            ws.set_column('C:C', 15) # NIT
            ws.set_column('D:D', 15) # Fuente
            ws.set_column('E:E', 18) # Cuenta
            ws.set_column('F:F', 40) # Desc Cuenta
            ws.set_column('G:G', 40) # Referencia
            ws.set_column('H:K', 15) # Montos
            
    return output_buffer.getvalue()

# ==============================================================================
# 6. REPORTE PARA AUDITORIA CB-CG
# ==============================================================================

def generar_reporte_cuadre(df_resultado, df_huerfanos, nombre_empresa):
    """
    Genera el Excel del Cuadre CB-CG.
    Hoja 1: Resumen General (Con Totales).
    Hoja 2: Análisis de Descuadres (Con Código CB y Cuenta CG).
    Hoja 3: Cuentas No Configuradas.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- ESTILOS ---
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        header_fmt = workbook.add_format({'bold': True, 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        text_fmt = workbook.add_format({'border': 1})
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '#,##0.00', 'border': 1})
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '#,##0.00', 'border': 1})
        
        group_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1})
        
        # Estilos Totales Hoja 1
        total_label_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'right'})
        total_val_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'num_format': '#,##0.00', 'border': 1})
        total_red_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '#,##0.00', 'border': 1})
        total_green_fmt = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '#,##0.00', 'border': 1})
        
        # ==========================================
        # HOJA 1: RESUMEN GENERAL
        # ==========================================
        ws1 = workbook.add_worksheet('Resumen General')
        ws1.hide_gridlines(2)
        
        cols_resumen = ['Banco (Tesorería)', 'Cuenta Contable', 'Descripción', 'Saldo Final CB', 'Saldo Final CG', 'Diferencia', 'Estado']
        
        current_row = 0
        ws1.merge_range('A1:G1', f"CUADRE DE DISPONIBILIDAD BANCARIA (CB vs CG) - {nombre_empresa}", title_fmt)
        current_row += 2
        
        for moneda, grupo in df_resultado.groupby('Moneda'):
            ws1.merge_range(current_row, 0, current_row, 6, f"MONEDA: {moneda}", group_fmt)
            current_row += 1
            
            ws1.write_row(current_row, 0, cols_resumen, header_fmt)
            current_row += 1
            
            # Variables para Totales
            sum_cb = 0.0
            sum_cg = 0.0
            sum_dif = 0.0
            
            for _, row in grupo.iterrows():
                ws1.write(current_row, 0, row['Banco (Tesorería)'], text_fmt)
                ws1.write(current_row, 1, row['Cuenta Contable'], text_fmt)
                ws1.write(current_row, 2, row['Descripción'], text_fmt)
                ws1.write_number(current_row, 3, row['Saldo Final CB'], money_fmt)
                ws1.write_number(current_row, 4, row['Saldo Final CG'], money_fmt)
                
                dif = row['Diferencia']
                fmt_dif = red_fmt if dif != 0 else green_fmt
                ws1.write_number(current_row, 5, dif, fmt_dif)
                
                ws1.write(current_row, 6, row['Estado'], text_fmt)
                
                # Acumular
                sum_cb += row['Saldo Final CB']
                sum_cg += row['Saldo Final CG']
                sum_dif += dif
                current_row += 1
            
            # Fila de Totales
            ws1.write(current_row, 2, f"TOTAL {moneda}", total_label_fmt)
            ws1.write_number(current_row, 3, sum_cb, total_val_fmt)
            ws1.write_number(current_row, 4, sum_cg, total_val_fmt)
            fmt_tot_dif = total_red_fmt if abs(sum_dif) > 0.01 else total_green_fmt
            ws1.write_number(current_row, 5, sum_dif, fmt_tot_dif)
            ws1.write(current_row, 6, "", total_val_fmt) # Borde vacío
            
            current_row += 2

        ws1.set_column('A:B', 15); ws1.set_column('C:C', 40); ws1.set_column('D:F', 18); ws1.set_column('G:G', 12)

        # ==========================================
        # HOJA 2: ANÁLISIS DE DESCUADRES
        # ==========================================
        ws2 = workbook.add_worksheet('Análisis de Descuadres')
        ws2.hide_gridlines(2)
        
        df_descuadres = df_resultado[df_resultado['Estado'] == 'DESCUADRE'].copy()
        ws2.merge_range('A1:K1', f"DETALLE DE DESCUADRES - {nombre_empresa}", title_fmt)
        
        if not df_descuadres.empty:
            # --- CAMBIO AQUÍ: AGREGADAS COLUMNAS CÓDIGO Y CUENTA ---
            headers_det = [
                'Moneda', 
                'Código CB',       # <--- Nueva
                'Cuenta Contable', # <--- Nueva
                'Descripción', 
                'CB Inicial', 'CB Débitos', 'CB Créditos',
                'CG Inicial', 'CG Débitos', 'CG Créditos',
                'DIFERENCIA FINAL'
            ]
            ws2.write_row(2, 0, headers_det, header_fmt)
            
            curr_row = 3
            for _, row in df_descuadres.iterrows():
                ws2.write(curr_row, 0, row['Moneda'], text_fmt)
                ws2.write(curr_row, 1, row['Banco (Tesorería)'], text_fmt) # Código CB
                ws2.write(curr_row, 2, row['Cuenta Contable'], text_fmt)   # Cuenta CG
                ws2.write(curr_row, 3, row['Descripción'], text_fmt)
                
                # CB Saldos (Indices +1 por las nuevas columnas)
                ws2.write_number(curr_row, 4, row.get('CB Inicial', 0), money_fmt)
                ws2.write_number(curr_row, 5, row.get('CB Débitos', 0), money_fmt)
                ws2.write_number(curr_row, 6, row.get('CB Créditos', 0), money_fmt)
                
                # CG Saldos
                ws2.write_number(curr_row, 7, row.get('CG Inicial', 0), money_fmt)
                ws2.write_number(curr_row, 8, row.get('CG Débitos', 0), money_fmt)
                ws2.write_number(curr_row, 9, row.get('CG Créditos', 0), money_fmt)
                
                # Diferencia
                ws2.write_number(curr_row, 10, row['Diferencia'], red_fmt)
                curr_row += 1
            
            # Ajuste de Anchos
            ws2.set_column('A:A', 10) # Moneda
            ws2.set_column('B:C', 18) # Codigos
            ws2.set_column('D:D', 35) # Descripcion
            ws2.set_column('E:K', 15) # Montos
        else:
            ws2.write('A3', "¡Felicidades! No hay descuadres en saldos finales.")

        # ==========================================
        # HOJA 3: CUENTAS NO CONFIGURADAS
        # ==========================================
        if not df_huerfanos.empty:
            ws3 = workbook.add_worksheet('⚠️ Cuentas Sin Configurar')
            ws3.hide_gridlines(2)
            warning_fmt = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'center', 'font_size': 12})
            ws3.merge_range('A1:E1', "¡ALERTA! Se encontraron movimientos en cuentas que NO están en el diccionario", warning_fmt)
            
            headers_huerfanos = ['Origen', 'Código/Cuenta', 'Descripción/Nombre', 'Saldo Final', 'Mensaje']
            ws3.write_row(2, 0, headers_huerfanos, header_fmt)
            
            curr_row = 3
            for _, row in df_huerfanos.iterrows():
                ws3.write(curr_row, 0, row['Origen'], text_fmt)
                ws3.write(curr_row, 1, row['Código/Cuenta'], text_fmt)
                ws3.write(curr_row, 2, row['Descripción/Nombre'], text_fmt)
                ws3.write(curr_row, 3, row['Saldo Final'], text_fmt)
                ws3.write(curr_row, 4, row['Mensaje'], text_fmt)
                curr_row += 1
            ws3.set_column('A:B', 20); ws3.set_column('C:C', 40); ws3.set_column('D:E', 30)
            ws3.set_tab_color('red')

    return output.getvalue()

# ==============================================================================
# UTILS PARA IMPRENTA
# ==============================================================================

def generar_archivo_txt(lineas):
    """Crea el archivo TXT para descargar."""
    output = BytesIO()
    # Usamos \r\n para mayor compatibilidad con Windows (Galac)
    contenido = "\r\n".join(lineas)
    output.write(contenido.encode('latin-1', errors='ignore')) 
    return output.getvalue()

def generar_reporte_imprenta(df_resultado):
    """Excel para Validación (Pestaña 1)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Resultados')
        workbook = writer.book
        ws = writer.sheets['Resultados']
        
        red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        
        ws.conditional_format('E2:E5000', {'type': 'text', 'criteria': 'containing', 'value': 'ERROR', 'format': red})
        ws.conditional_format('E2:E5000', {'type': 'text', 'criteria': 'containing', 'value': 'OK', 'format': green})
        ws.set_column('A:E', 20)
    return output.getvalue()

def generar_reporte_auditoria_txt(df_audit):
    """Excel para Generación de Imprenta con IVA GALAC y % de Retención."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_audit.to_excel(writer, index=False, sheet_name='Auditoria')
        workbook = writer.book
        ws = writer.sheets['Auditoria']
        
        # Estilos mejorados
        header_fmt = workbook.add_format({'bold': True, 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})
        pct_fmt = workbook.add_format({'num_format': '0.0%'}) # Formato de porcentaje con 1 decimal
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE'})
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE'})
        
        for idx, col in enumerate(df_audit.columns):
            ws.write(0, idx, col, header_fmt)
            
            # Anchos
            if 'Nombre' in col: ws.set_column(idx, idx, 35)
            elif 'Referencia' in col: ws.set_column(idx, idx, 30)
            elif 'Estatus' in col: ws.set_column(idx, idx, 22)
            else: ws.set_column(idx, idx, 15)

            # Formatos específicos de datos
            if col in ['IVA Origen Softland', 'IVA GALAC (Base)', 'Monto Retenido GALAC']:
                ws.set_column(idx, idx, 16, money_fmt)
            elif col == '% Retención':
                ws.set_column(idx, idx, 12, pct_fmt)
        
        # Aplicar colores de estatus
        for r_idx, row in df_audit.iterrows():
            status = str(row['Estatus'])
            fmt = green_fmt if 'OK' in status else red_fmt
            ws.write(r_idx + 1, 0, status, fmt)
            
    return output.getvalue()

def generar_reporte_pensiones(df_agrupado, df_base, df_asiento, resumen_validacion, nombre_empresa, tasa_cambio, fecha_cierre):
    """
    Genera Excel Profesional de Pensiones.
    Hoja 1: Dos tablas comparativas (Por Cuenta y Por Centro de Costo) + Validación.
    Hoja 3: Asiento Contable (Fondo Blanco).
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- ESTILOS GENERALES ---
        header_green = workbook.add_format({'bold': True, 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'})
        money_bold = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bold': True, 'valign': 'vcenter'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        text_center = workbook.add_format({'align': 'center', 'border': 1, 'valign': 'vcenter'})
        text_left = workbook.add_format({'align': 'left', 'border': 1, 'valign': 'vcenter'})
        
        # Estilos Validación
        fmt_red = workbook.add_format({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '#,##0.00', 'border': 1})
        fmt_green = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '#,##0.00', 'border': 1})

        # Estilos Títulos Hoja 1
        fmt_main_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        fmt_sub_title = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'left', 'valign': 'vcenter'})
        fmt_periodo = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'right', 'valign': 'vcenter'})
        fmt_table_title = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2', 'border': 1})

        # --- ESTILOS ASIENTO (HOJA 3) ---
        title_company = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
        fmt_title_label = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter'})
        fmt_company = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bottom': 1})
        fmt_code_company = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bottom': 1})
        fmt_input = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'bold': True})
        fmt_date_calc = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'bold': True, 'num_format': 'dd/mm/yyyy'})
        fmt_usd_4 = workbook.add_format({'num_format': '#,##0.0000', 'border': 1, 'valign': 'vcenter'})
        fmt_calc = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'bold': True,'num_format': '#,##0.00'})
        fmt_calc_usd = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'bold': True, 'num_format': '#,##0.0000'})
        fmt_calc_ves = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1,'align': 'center', 'bold': True, 'num_format': '#,##0.00'})
        box_header = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#FFFFFF'})
        box_data_center = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        box_data_left = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        box_money = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'valign': 'vcenter'})
        box_money_bold = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'valign': 'vcenter', 'bold': True})
        small_text = workbook.add_format({'font_size': 9, 'italic': True, 'align': 'left'})
        
        # ==========================================
        # HOJA 1: CÁLCULO Y BASE
        # ==========================================
        ws1 = workbook.add_worksheet('1. Calculo y Base')
        ws1.hide_gridlines(2)
        
        # 1. ENCABEZADO CORPORATIVO
        if fecha_cierre:
            meses_es = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
            periodo_txt = f"{meses_es.get(fecha_cierre.month, '')} {fecha_cierre.year}"
        else:
            periodo_txt = "PERIODO NO DEFINIDO"

        ws1.merge_range('A1:I1', "CÁLCULO LEY DE PROTECCIÓN DE PENSIONES (9%)", fmt_main_title)
        ws1.merge_range('A2:D2', f"EMPRESA: {nombre_empresa}", fmt_sub_title)
        ws1.merge_range('G2:I2', f"PERIODO: {periodo_txt}", fmt_periodo)
        
        # 2. TÍTULOS DE LAS TABLAS (FILA 4)
        # Tabla Izquierda (Detallada)
        ws1.merge_range('A4:D4', "SUMATORIA POR CUENTA CONTABLE", fmt_table_title)
        # Tabla Derecha (Resumida) - Dejamos Columna E como separador
        ws1.merge_range('G4:I4', "SUMATORIA POR CENTRO DE COSTO", fmt_table_title)

        # 3. ENCABEZADOS DE COLUMNAS (FILA 5)
        headers_left = ['Centro de Costo', 'Cuenta Contable', 'Base', 'Impuesto (9%)']
        ws1.write_row('A5', headers_left, header_green)

        headers_right = ['Centro de Costo', 'Total Nomina', 'Total Aporte']
        ws1.write_row('G5', headers_right, header_green)
        
        # ---------------------------------------------------------
        # TABLA IZQUIERDA (DETALLE POR TIPO NOMINA)
        # ---------------------------------------------------------
        row_left = 5
        
        # Bloque Nómina
        nomina = df_agrupado[df_agrupado['Cuenta Contable'].astype(str).str.contains('7.1.1.01', na=False)]
        if not nomina.empty:
            for _, row in nomina.iterrows():
                ws1.write(row_left, 0, row['Centro de Costo (Padre)'], text_left)
                ws1.write(row_left, 1, row['Cuenta Contable'], text_center)
                ws1.write_number(row_left, 2, row['Base_Neta'], money_fmt)
                ws1.write_number(row_left, 3, row['Impuesto (9%)'], money_fmt)
                row_left += 1
            ws1.write(row_left, 1, "Total Nomina", total_fmt)
            ws1.write_number(row_left, 2, nomina['Base_Neta'].sum(), money_bold)
            ws1.write_number(row_left, 3, nomina['Impuesto (9%)'].sum(), money_bold)
            row_left += 2

        # Bloque Cestaticket
        ticket = df_agrupado[df_agrupado['Cuenta Contable'].astype(str).str.contains('7.1.1.09', na=False)]
        if not ticket.empty:
            for _, row in ticket.iterrows():
                ws1.write(row_left, 0, row['Centro de Costo (Padre)'], text_left)
                ws1.write(row_left, 1, row['Cuenta Contable'], text_center)
                ws1.write_number(row_left, 2, row['Base_Neta'], money_fmt)
                ws1.write_number(row_left, 3, row['Impuesto (9%)'], money_fmt)
                row_left += 1
            ws1.write(row_left, 1, "Total Cestaticket", total_fmt)
            ws1.write_number(row_left, 2, ticket['Base_Neta'].sum(), money_bold)
            ws1.write_number(row_left, 3, ticket['Impuesto (9%)'].sum(), money_bold)
            row_left += 2

        # Total General Izquierda
        ws1.write(row_left, 1, "Total General", total_fmt)
        ws1.write_number(row_left, 2, df_agrupado['Base_Neta'].sum(), money_bold)
        ws1.write_number(row_left, 3, df_agrupado['Impuesto (9%)'].sum(), money_bold)
        
        # ---------------------------------------------------------
        # TABLA DERECHA (RESUMEN POR CENTRO DE COSTO)
        # ---------------------------------------------------------
        # Preparamos datos agrupados
        df_cc = df_agrupado.groupby('Centro de Costo (Padre)')[['Base_Neta', 'Impuesto (9%)']].sum().reset_index()
        
        row_right = 5
        for _, row in df_cc.iterrows():
            ws1.write(row_right, 6, row['Centro de Costo (Padre)'], text_left) # Col G
            ws1.write_number(row_right, 7, row['Base_Neta'], money_fmt)       # Col H
            ws1.write_number(row_right, 8, row['Impuesto (9%)'], money_fmt)   # Col I
            row_right += 1
        
        # Total General Derecha
        ws1.write(row_right, 6, "Total General", total_fmt)
        ws1.write_number(row_right, 7, df_cc['Base_Neta'].sum(), money_bold)
        ws1.write_number(row_right, 8, df_cc['Impuesto (9%)'].sum(), money_bold)

        # ---------------------------------------------------------
        # TABLA DE VALIDACIÓN (Al final de la más larga)
        # ---------------------------------------------------------
        current_row = max(row_left, row_right) + 3

        ws1.merge_range(current_row, 1, current_row, 4, "VALIDACIÓN CRUZADA DETALLADA (CONTABILIDAD vs NÓMINA)", header_green)
        current_row += 1
        
        headers_val = ['CONCEPTO', 'SEGÚN CONTABILIDAD', 'SEGÚN NÓMINA (ARCHIVO)', 'DIFERENCIA']
        ws1.write_row(current_row, 1, headers_val, header_green)
        current_row += 1
        
        # 1. Salarios
        dif_sal = resumen_validacion['dif_salario']
        fmt_sal = fmt_green if abs(dif_sal) < 1 else fmt_red
        ws1.write(current_row, 1, "Salario (7.1.1.01)", text_center)
        ws1.write_number(current_row, 2, resumen_validacion['salario_cont'], money_fmt)
        ws1.write_number(current_row, 3, resumen_validacion['salario_nom'], money_fmt)
        ws1.write_number(current_row, 4, dif_sal, fmt_sal)
        current_row += 1
        
        # 2. Tickets
        dif_tkt = resumen_validacion['dif_ticket']
        fmt_tkt = fmt_green if abs(dif_tkt) < 1 else fmt_red
        ws1.write(current_row, 1, "Ticket (7.1.1.09)", text_center)
        ws1.write_number(current_row, 2, resumen_validacion['ticket_cont'], money_fmt)
        ws1.write_number(current_row, 3, resumen_validacion['ticket_nom'], money_fmt)
        ws1.write_number(current_row, 4, dif_tkt, fmt_tkt)
        current_row += 1
        
        # 3. Total Base
        dif_tot = resumen_validacion['dif_base_total']
        fmt_tot = fmt_green if abs(dif_tot) < 1 else fmt_red
        ws1.write(current_row, 1, "Total General Base", total_fmt)
        ws1.write_number(current_row, 2, resumen_validacion['total_base_cont'], money_bold)
        ws1.write_number(current_row, 3, resumen_validacion['total_base_nom'], money_bold)
        ws1.write_number(current_row, 4, dif_tot, fmt_tot)
        current_row += 1
        
        # 4. Impuesto
        dif_imp = resumen_validacion['dif_imp']
        fmt_imp = fmt_green if abs(dif_imp) < 1 else fmt_red
        ws1.write(current_row, 1, "Impuesto (Apartado)", total_fmt)
        ws1.write_number(current_row, 2, resumen_validacion['imp_calc'], money_bold)
        ws1.write_number(current_row, 3, resumen_validacion['imp_nom'], money_bold)
        ws1.write_number(current_row, 4, dif_imp, fmt_imp)
        
        # Ajuste de Anchos
        ws1.set_column('A:B', 20)
        ws1.set_column('C:E', 18) # Aumentamos E de 2 a 18 para mostrar la Diferencia
        ws1.set_column('F:F', 2)  # Mantenemos solo la F como separador estrecho
        ws1.set_column('G:I', 18)

        # ==========================================
        # HOJA 2: DETALLE MAYOR
        # ==========================================
        if df_base is not None:
            cols_drop = ['CC_Agrupado', 'Monto_Deb', 'Monto_Cre', 'Base_Neta']
            df_clean = df_base.drop(columns=cols_drop, errors='ignore')
            
            # --- CORRECCIÓN FECHA: Convertir a Texto Corto ---
            # Buscamos la columna de fecha (puede llamarse FECHA, Fecha, etc.)
            col_fecha = next((c for c in df_clean.columns if 'FECHA' in c.upper()), None)
            
            if col_fecha:
                # Convertimos a string dd/mm/yyyy para evitar que excel ponga la hora o #####
                df_clean[col_fecha] = pd.to_datetime(df_clean[col_fecha], errors='coerce').dt.strftime('%d/%m/%Y')
                # Rellenamos NaT con vacío
                df_clean[col_fecha] = df_clean[col_fecha].fillna('')
            # -------------------------------------------------

            df_clean.to_excel(writer, sheet_name='2. Detalle Mayor', index=False)
            
            # Aumentamos ancho de columna para que quepa todo
            writer.sheets['2. Detalle Mayor'].set_column('A:Z', 20)
            
        # ==========================================
        # HOJA 3: ASIENTO CONTABLE
        # ==========================================
        if df_asiento is not None:
            ws3 = workbook.add_worksheet('3. Asiento Contable')
            ws3.hide_gridlines(2)
            
            # --- MAPEO DE CÓDIGOS DINÁMICO ---
            mapa_codigos = {
                "FEBECA": "004",
                "BEVAL": "207",
                "PRISMA": "298",
                "QUINCALLA": "071",
                "SILLACA": "071"
            }
            codigo_empresa = "000" # Default
            # Buscamos coincidencias (ej: "FEBECA, C.A" contiene "FEBECA")
            for k, v in mapa_codigos.items():
                if k in str(nombre_empresa).upper():
                    codigo_empresa = v
                    break
            # --------------------------------

            ws3.write('A1', "COMPAÑÍA:", fmt_title_label)
            ws3.merge_range('C1:F1', nombre_empresa, fmt_company)
            ws3.write('G1', "Nº.", workbook.add_format({'bold': True, 'align': 'right'}))
            
            # USO DEL CÓDIGO DINÁMICO
            ws3.write('H1', codigo_empresa, fmt_code_company)

            ws3.write('B3', "PARA ASENTAR EN DIARIO Y CUENTAS:", fmt_title_label)
            ws3.write('B4', "1) Escríbase con máquina de escribir.", small_text)
            ws3.write('B5', "2) Entréguese a Contabilidad.", small_text)
            ws3.write('B6', "3) Anéxese documentación original, si la hay.", small_text)
            ws3.write('B7', "4) En caso de no anexarla. Indíquese dónde se archiva.", small_text)

            ws3.merge_range('G3:H3', "A S E N T A D O", box_header)
            ws3.write('G4', "Operación No.: _______", workbook.add_format({'align': 'right', 'valign': 'vcenter'}))
            ws3.write('H4', fecha_cierre if fecha_cierre else "DD/MM/AAAA", fmt_date_calc)
            ws3.write('G5', "Comprob. N°.: _______", workbook.add_format({'align': 'right', 'valign': 'vcenter'}))
            ws3.write('H5', "", fmt_input)

            start_row = 8
            ws3.merge_range(start_row, 0, start_row, 2, "NUMERO DE CUENTA", box_header)
            ws3.merge_range(start_row, 3, start_row, 5, "TITULO DE CUENTA", box_header)
            ws3.merge_range(start_row, 6, start_row, 7, "MONTO BOLÍVARES", box_header)
            ws3.merge_range(start_row, 8, start_row, 9, "MONTO DOLARES", box_header)
            
            ws3.write(start_row+1, 0, "OFIC.", box_header)
            ws3.write(start_row+1, 1, "CENTRO DE COSTO", box_header)
            ws3.write(start_row+1, 2, "CTA.", box_header)
            ws3.merge_range(start_row+1, 3, start_row+1, 5, "", box_header)
            ws3.write(start_row+1, 6, "DEBE (D)", box_header)
            ws3.write(start_row+1, 7, "HABER (H)", box_header)
            ws3.write(start_row+1, 8, "DEBE (D)", box_header)
            ws3.write(start_row+1, 9, "HABER (H)", box_header)
            
            row_idx = start_row + 2
            
            for _, row in df_asiento.iterrows():
                ws3.write(row_idx, 0, "01", box_data_center)
                ws3.write(row_idx, 1, row['Centro Costo'], box_data_center)
                ws3.write(row_idx, 2, row['Cuenta Contable'], box_data_center)
                ws3.merge_range(row_idx, 3, row_idx, 5, row['Descripción'], box_data_left)
                
                d_v = row['Débito VES']; h_v = row['Crédito VES']
                d_u = row['Débito USD']; h_u = row['Crédito USD']    
        
                ws3.write(row_idx, 6, d_v if d_v > 0 else "", box_money)
                ws3.write(row_idx, 7, h_v if h_v > 0 else "", box_money)

                # Escribir Debe USD (Columna 8): solo si es mayor a 0, de lo contrario vacío ""
                if d_u > 0:
                    ws3.write_number(row_idx, 8, d_u, fmt_usd_4)
                else:
                    ws3.write(row_idx, 8, "", fmt_usd_4)

                # Escribir Haber USD (Columna 9): solo si es mayor a 0, de lo contrario vacío ""
                if h_u > 0:
                    ws3.write_number(row_idx, 9, h_u, fmt_usd_4)
                else:
                    ws3.write(row_idx, 9, "", fmt_usd_4)
                row_idx += 1
            
            ws3.write(row_idx, 6, df_asiento['Débito VES'].sum(), box_money_bold)
            ws3.write(row_idx, 7, df_asiento['Crédito VES'].sum(), box_money_bold)
            ws3.write(row_idx, 8, df_asiento['Débito USD'].sum(), box_money_bold)
            ws3.write(row_idx, 9, df_asiento['Crédito USD'].sum(), box_money_bold)
            row_idx += 2 

            mes_txt = fecha_cierre.strftime('%b').upper() if fecha_cierre else "MES"
            anio_txt = fecha_cierre.strftime('%y') if fecha_cierre else "AA"
            texto_concepto = f"APORTE PENSIONES {mes_txt}.{anio_txt}"

            ws3.write(row_idx, 3, "(Máximo 40 posiciones...)", small_text)
            ws3.write(row_idx+1, 0, "TEXTO DEL DEBE", fmt_title_label)
            ws3.merge_range(row_idx+1, 3, row_idx+1, 5, texto_concepto, fmt_calc)
            ws3.write(row_idx+1, 7, df_asiento['Débito VES'].sum(), fmt_calc_ves) # Total VES
            ws3.write(row_idx+1, 9, df_asiento['Débito USD'].sum(), fmt_calc_usd) # Total USD con 4 decimales
            row_idx += 4

            ws3.write(row_idx, 3, "(Máximo 40 posiciones...)", small_text)
            ws3.write(row_idx+1, 0, "TEXTO DEL HABER", fmt_title_label)
            ws3.merge_range(row_idx+1, 3, row_idx+1, 5, texto_concepto, fmt_calc)
            ws3.write(row_idx+1, 7, df_asiento['Crédito VES'].sum(), fmt_calc_ves) # Total VES
            ws3.write(row_idx+1, 9, df_asiento['Crédito USD'].sum(), fmt_calc_usd) # Total USD con 4 decimales
            row_idx += 3

            top_line = workbook.add_format({'top': 1, 'font_size': 9})
            ws3.write(row_idx, 0, "Hecho por:", top_line)
            ws3.merge_range(row_idx, 3, row_idx, 4, "Aprobado por:", top_line)
            ws3.merge_range(row_idx, 6, row_idx, 7, "Procesado por:", top_line)
            ws3.merge_range(row_idx, 8, row_idx, 9, "Revisado por:", top_line)
            
            ws3.merge_range(row_idx+1, 0, row_idx+1, 2, "", fmt_input) 
            
            box_corner = workbook.add_format({'top': 1, 'left':1, 'right':1, 'font_size': 9})
            ws3.write(row_idx, 8, "Lugar y Fecha:", box_corner)
            fecha_str = fecha_cierre.strftime('%d/%m/%Y') if fecha_cierre else ""
            lugar_fecha = f"VALENCIA, {fecha_str}"
            ws3.merge_range(row_idx+1, 8, row_idx+1, 9, lugar_fecha, fmt_calc)
            
            ws3.merge_range(row_idx+3, 4, row_idx+3, 6, "ORIGINAL: CONTABILIDAD", workbook.add_format({'bold': True, 'align': 'center'}))

            ws3.set_column('A:A', 8); ws3.set_column('B:B', 15); ws3.set_column('C:C', 15)
            ws3.set_column('D:F', 15); ws3.set_column('G:J', 18)

    return output.getvalue()

def generar_cargador_asiento_pensiones(df_asiento, fecha_asiento):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Formatos
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#212121', 'font_color': 'white', 'border': 1, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1})
        num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})

        # --- PESTAÑA 1: "Asiento" ---
        ws1 = workbook.add_worksheet("Asiento")
        headers_asiento = ["Asiento", "Paquete", "Tipo Asiento", "Fecha", "Contabilidad"]
        ws1.write_row(0, 0, headers_asiento, header_fmt)
        
        ws1.write(1, 0, df_asiento['Asiento'].iloc[0], data_fmt)
        ws1.write(1, 1, "AJTB", data_fmt) # Paquete fijo según imagen
        ws1.write(1, 2, "AJTB", data_fmt) # Tipo fijo según imagen
        ws1.write_datetime(1, 3, fecha_asiento, date_fmt)
        ws1.write(1, 4, "C", data_fmt) # Contabilidad fija según imagen
        ws1.set_column('A:E', 15)

        # --- PESTAÑA 2: "ND" ---
        ws2 = workbook.add_worksheet("ND")
        headers_nd = ["Asiento", "Consecutivo", "Nit", "Centro De Costo", "Cuenta Contable", "Fuente", "Referencia", "Débito Local", "Débito Dólar", "Crédito Local", "Crédito Dólar"]
        ws2.write_row(0, 0, headers_nd, header_fmt)

        for i, row in df_asiento.iterrows():
            r = i + 1
            ws2.write(r, 0, row['Asiento'], data_fmt)
            ws2.write(r, 1, i + 1, data_fmt) # Consecutivo
            ws2.write(r, 2, row['Nit'], data_fmt)
            ws2.write(r, 3, row['Centro Costo'], data_fmt)
            ws2.write(r, 4, row['Cuenta Contable'], data_fmt)
            ws2.write(r, 5, row['Fuente'], data_fmt)
            ws2.write(r, 6, row['Referencia'], data_fmt)
            
            # Montos (Si es 0 ponemos vacío o 0.00 según prefieras, la imagen muestra celdas vacías con un guion)
            ws2.write_number(r, 7, row['Débito VES'], num_fmt)
            ws2.write_number(r, 8, row['Débito USD'], num_fmt)
            ws2.write_number(r, 9, row['Crédito VES'], num_fmt)
            ws2.write_number(r, 10, row['Crédito USD'], num_fmt)

        ws2.set_column('A:A', 15)
        ws2.set_column('C:C', 15)
        ws2.set_column('D:G', 25)
        ws2.set_column('H:K', 15)

    return output.getvalue()

def generar_reporte_ajustes_usd(df_resumen, df_bancos, df_asiento, df_balance_raw, nombre_empresa, validacion_data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- ESTILOS ---
        fmt_header_raw = workbook.add_format({'bold': False, 'font_size': 10})
        header_clean = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFFFF'})
        fmt_text = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        fmt_money = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        fmt_money_bold = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bold': True, 'bg_color': '#F2F2F2'})
        fmt_rate = workbook.add_format({'num_format': '#,##0.0000', 'border': 1})
        box_val = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bold': True, 'align': 'center'})

        ws1 = workbook.add_worksheet('1. Ajustes')
        ws1.hide_gridlines(2)
        
        # 1. ENCABEZADO (Solo Columnas A y B)
        if df_balance_raw is not None and not df_balance_raw.empty:
            for r_idx in range(min(6, len(df_balance_raw))):
                for c_idx in [0, 1]: # Solo A y B
                    val = df_balance_raw.iloc[r_idx, c_idx]
                    if pd.notna(val): ws1.write(r_idx, c_idx, val, fmt_header_raw)

        # 2. PRE-PROCESAMIENTO DE AJUSTES
        mapa_ajustes = {}
        for _, row in df_resumen.iterrows():
            cta = str(row['Cuenta']).strip()
            mapa_ajustes[cta] = mapa_ajustes.get(cta, 0.0) + row['Ajuste USD']

        # 3. ESCRITURA DE DATOS Y CÁLCULO DE TOTALES PARA EL CUADRO
        current_row = 7
        total_ajuste_1_activo = 0.0
        total_ajuste_2_pasivo = 0.0

        if df_balance_raw is not None and not df_balance_raw.empty:
            start_idx = 0
            for i, row in df_balance_raw.iterrows():
                if any('CUENTA' in str(v).upper() for v in row.values):
                    start_idx = i + 1; break
            
            for i in range(start_idx, len(df_balance_raw)):
                row = df_balance_raw.iloc[i]
                try:
                    cuenta = str(row[0]).strip()
                    if not (cuenta.startswith('1.') or cuenta.startswith('2.')): continue
                    if cuenta.endswith('.000') or cuenta.endswith('.00'): continue 
                    
                    saldo_bs = float(str(row[6]).replace('.', '').replace(',', '.')) if pd.notna(row[6]) else 0.0
                    saldo_usd = float(str(row[11]).replace('.', '').replace(',', '.')) if pd.notna(row[11]) else 0.0
                    
                    ajuste = mapa_ajustes.get(cuenta, 0.0)
                    saldo_ajustado = saldo_usd + ajuste
                    tasa = abs(saldo_bs / saldo_usd) if abs(saldo_usd) > 0.01 else 0.0

                    # Sumamos al total del cuadro resumen según el tipo de cuenta
                    if cuenta.startswith('1.'): total_ajuste_1_activo += ajuste
                    elif cuenta.startswith('2.'): total_ajuste_2_pasivo += ajuste

                    ws1.write(current_row, 0, cuenta, fmt_text)
                    ws1.write(current_row, 1, str(row[1]).strip(), fmt_text)
                    ws1.write(current_row, 2, str(row[2]).strip(), fmt_text)
                    ws1.write_number(current_row, 3, saldo_bs, fmt_money)
                    ws1.write_number(current_row, 4, saldo_usd, fmt_money)
                    ws1.write_number(current_row, 5, ajuste, fmt_money_bold if abs(ajuste) > 0.001 else fmt_money)
                    ws1.write_number(current_row, 6, saldo_ajustado, fmt_money_bold)
                    ws1.write_number(current_row, 7, tasa, fmt_rate)
                    current_row += 1
                except: continue

        # 4. ESCRIBIR CUADRO RESUMEN EN I, J, K (Ubicación Final)
        ws1.write(1, 8, "1", header_clean)    # Col I
        ws1.write(1, 9, "ACTIVO", header_clean) # Col J
        ws1.write_number(1, 10, total_ajuste_1_activo, box_val) # Col K
        
        ws1.write(2, 8, "2", header_clean) 
        ws1.write(2, 9, "PASIVO", header_clean) 
        # Mostramos valor absoluto para pasivo si así se prefiere, o neto.
        ws1.write_number(2, 10, total_ajuste_2_pasivo, box_val)
        
        ws1.write(4, 9, "DIF", workbook.add_format({'bold': True, 'align': 'right'}))
        ws1.write_formula(4, 10, '=K2-K3', box_val)

        # Ajustes de diseño
        ws1.set_column('A:A', 15); ws1.set_column('B:B', 45); ws1.set_column('D:G', 18)
        ws1.set_column('I:K', 14)

        # 5. ESCRITURA DE DATOS
        current_row = 7
        if df_balance_raw is not None and not df_balance_raw.empty:
            start_idx = 0
            for i, row in df_balance_raw.iterrows():
                if any('CUENTA' in str(v).upper() for v in row.values):
                    start_idx = i + 1; break
            
            for i in range(start_idx, len(df_balance_raw)):
                row = df_balance_raw.iloc[i]
                try:
                    cuenta = str(row[0]).strip()
                    if not (cuenta.startswith('1.') or cuenta.startswith('2.')): continue
                    if cuenta.endswith('.000') or cuenta.endswith('.00'): continue 
                    
                    def get_val(x):
                        try:
                            if isinstance(x, (int, float)): return float(x)
                            return float(str(x).replace('.', '').replace(',', '.'))
                        except: return 0.0

                    saldo_bs = get_val(row[6]); saldo_usd = get_val(row[11])
                    ajuste = mapa_ajustes.get(cuenta, 0.0)
                    saldo_ajustado = saldo_usd + ajuste
                    tasa = abs(saldo_bs / saldo_usd) if abs(saldo_usd) > 0.01 else 0.0

                    ws1.write(current_row, 0, cuenta, fmt_text)
                    ws1.write(current_row, 1, str(row[1]).strip(), fmt_text)
                    ws1.write(current_row, 2, str(row[2]).strip(), fmt_text)
                    ws1.write_number(current_row, 3, saldo_bs, fmt_money)
                    ws1.write_number(current_row, 4, saldo_usd, fmt_money)
                    ws1.write_number(current_row, 5, ajuste, fmt_money_bold if abs(ajuste) > 0.001 else fmt_money)
                    ws1.write_number(current_row, 6, saldo_ajustado, fmt_money_bold)
                    ws1.write_number(current_row, 7, tasa, fmt_rate)
                    current_row += 1
                except: continue
        
        ws1.set_column('A:A', 15); ws1.set_column('B:B', 45); ws1.set_column('D:G', 18)
        ws1.set_column('I:K', 12) # Columnas del Cuadro Resumen

        # ==========================================
        # HOJA 2: DETALLE BANCOS (SE MANTIENE IGUAL)
        # ==========================================
        ws2 = workbook.add_worksheet('2. Detalle Bancos')
        ws2.hide_gridlines(2)
        ws2.write(0, 6, "TASA BCV", workbook.add_format({'bold':True, 'align':'right'}))
        ws2.write(0, 7, validacion_data.get('tasa_bcv', 0), fmt_rate)
        ws2.write(1, 6, "TASA CORP.", workbook.add_format({'bold':True, 'align':'right'}))
        ws2.write(1, 7, validacion_data.get('tasa_corp', 0), fmt_rate)
        
        if df_bancos is not None and not df_bancos.empty:
            for c_idx, col_name in enumerate(df_bancos.columns):
                ws2.write(3, c_idx, col_name, header_clean)
            
            row_idx = 4
            for r_data in df_bancos.itertuples(index=False):
                for c_idx, value in enumerate(r_data):
                    if isinstance(value, (int, float)) and not pd.isna(value):
                        if c_idx == 12: ws2.write_number(row_idx, c_idx, value, fmt_rate)
                        else: ws2.write_number(row_idx, c_idx, value, fmt_money)
                    else:
                        ws2.write(row_idx, c_idx, str(value) if pd.notna(value) else "", fmt_text)
                row_idx += 1
            
            ws2.write(row_idx, 2, "TOTALES GENERALES", workbook.add_format({'bold':True, 'align':'right', 'border':1, 'bg_color':'#F2F2F2'}))
            for c_sum in range(3, 12):
                col_letter = chr(65 + c_sum)
                ws2.write_formula(row_idx, c_sum, f"=SUM({col_letter}5:{col_letter}{row_idx})", fmt_money_bold)

        ws2.set_column('A:A', 8); ws2.set_column('B:B', 18); ws2.set_column('C:C', 45); ws2.set_column('D:N', 18)

        # ==========================================
        # HOJA 3: ASIENTO CONTABLE
        # ==========================================
        if df_asiento is not None and not df_asiento.empty:
            ws3 = workbook.add_worksheet('3. Asiento Contable')
            ws3.hide_gridlines(2)
            ws3.merge_range('A1:F1', f"ASIENTO DE AJUSTE VALORACIÓN - {nombre_empresa}", main_title)
            ws3.write_row('A3', ['CUENTA', 'DESCRIPCIÓN', 'DEBE ($)', 'HABER ($)', 'DEBE (Bs)', 'HABER (Bs)'], header_clean)
            for r_idx, (idx, data) in enumerate(df_asiento.iterrows()):
                ws3.write(r_idx+3, 0, data['Cuenta'], fmt_text)
                ws3.write(r_idx+3, 1, data['Desc'], fmt_text)
                ws3.write_number(r_idx+3, 2, data['DebeUSD'], fmt_money)
                ws3.write_number(r_idx+3, 3, data['HaberUSD'], fmt_money)
                ws3.write_number(r_idx+3, 4, data.get('Débito VES', 0), fmt_money)
                ws3.write_number(r_idx+3, 5, data.get('Crédito VES', 0), fmt_money)
            ws3.set_column('A:A', 15); ws3.set_column('B:B', 40); ws3.set_column('C:F', 18)

        # ==========================================
        # HOJA 4: DATA (ORIGINAL)
        # ==========================================
        if df_balance_raw is not None and not df_balance_raw.empty:
            ws4 = workbook.add_worksheet('4. DATA (Original)')
            for r_idx, row in enumerate(df_balance_raw.values):
                for c_idx, val in enumerate(row):
                    if pd.notna(val): ws4.write(r_idx, c_idx, val)
            ws4.set_column('A:Z', 15)

    return output.getvalue()
    
def generar_reporte_cofersa(df_procesado):
    """
    Genera el reporte específico para Envios COFERSA con lógica basada en REFERENCIA.
    Versión Blindada: Maneja la ausencia de columna 'Tipo' y estandariza nombres.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- 0. NORMALIZACIÓN DE COLUMNAS Y DATOS ---
        if 'Descripcion Nit' in df_procesado.columns:
            df_procesado.rename(columns={'Descripcion Nit': 'Descripción Nit'}, inplace=True)
        
        if 'Descripción Nit' not in df_procesado.columns:
            df_procesado['Descripción Nit'] = 'NO DEFINIDO'

        # --- 1. DEFINICIÓN DE COLUMNAS DINÁMICAS ---
        cols_base = ['Fecha', 'Asiento', 'Fuente', 'Origen']
        if 'Tipo' in df_procesado.columns:
            cols_base.append('Tipo')
        
        cols_output = cols_base + [
            'Referencia', 'Débito Colones', 'Crédito Colones', 'Neto Colones'
            'Débito Dolar', 'Crédito Dolar', 'Neto Dólar',
            'Nit', 'Descripción Nit'
        ]
        
        # Detectar índices para formato moneda (usualmente empiezan en Débito Bolivar)
        try:
            start_num_idx = cols_output.index('Débito Bolivar')
            idx_nums = range(start_num_idx, start_num_idx + 6)
        except ValueError:
            idx_nums = range(6, 12)

        # --- 2. ESTILOS ---
        header_fmt = workbook.add_format({'bold': True, 'fg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'num_format': '#,##0.00', 'border': 1})
        label_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'right'})
        diff_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFC7CE', 'num_format': '#,##0.00', 'border': 1})
        text_fmt = workbook.add_format({'border': 1})
        date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})

        # --- 3. FUNCIONES AUXILIARES DE ESCRITURA ---
        def escribir_hoja_simple(nombre_hoja, df_data, titulo_total):
            ws = workbook.add_worksheet(nombre_hoja)
            ws.hide_gridlines(2)
            ws.write_row(0, 0, cols_output, header_fmt)
            r = 1
            for _, data in df_data.iterrows():
                for i, col in enumerate(cols_output):
                    val = data.get(col, '')
                    fmt = money_fmt if i in idx_nums else text_fmt
                    if 'Fecha' in col and pd.notna(val): ws.write_datetime(r, i, val, date_fmt)
                    elif i in idx_nums: ws.write_number(r, i, float(val) if pd.notna(val) else 0, fmt)
                    else: ws.write(r, i, str(val) if pd.notna(val) else "", fmt)
                r += 1
            ws.write(r, cols_output.index('Referencia'), titulo_total, label_fmt)
            for i in idx_nums: ws.write_number(r, i, df_data[cols_output[i]].sum(), total_fmt)
            ws.set_column(0, len(cols_output), 15)

        def escribir_hoja_agrupada(nombre_hoja, df_data, titulo_total, estilo_subtotal):
            if df_data.empty: return
            if 'Ref_Norm' not in df_data.columns:
                df_data['Ref_Norm'] = df_data['Referencia'].astype(str).str.strip().str.upper()
            
            df_data = df_data.sort_values(by=['Ref_Norm', 'Fecha'])
            ws = workbook.add_worksheet(nombre_hoja)
            ws.hide_gridlines(2)
            ws.write_row(0, 0, cols_output, header_fmt)
            
            row_idx = 1
            gran_total_bs = 0
            idx_neto_bs = cols_output.index('Neto Local')
            idx_ref = cols_output.index('Referencia')
            
            for ref, grupo in df_data.groupby('Ref_Norm'):
                subtotal_grupo = 0
                for _, data in grupo.iterrows():
                    subtotal_grupo += data.get('Neto Local', 0)
                    for i, col in enumerate(cols_output):
                        val = data.get(col, '')
                        fmt = money_fmt if i in idx_nums else text_fmt
                        if 'Fecha' in col and pd.notna(val): ws.write_datetime(row_idx, i, val, date_fmt)
                        elif i in idx_nums: ws.write_number(row_idx, i, float(val) if pd.notna(val) else 0, fmt)
                        else: ws.write(row_idx, i, str(val) if pd.notna(val) else "", fmt)
                    row_idx += 1
                
                ws.write(row_idx, idx_ref, f"SALDO {ref}:", label_fmt)
                ws.write_number(row_idx, idx_neto_bs, subtotal_grupo, estilo_subtotal)
                gran_total_bs += subtotal_grupo
                row_idx += 2
                
            ws.write(row_idx, idx_ref, titulo_total, label_fmt)
            ws.write_number(row_idx, idx_neto_bs, gran_total_bs, total_fmt)
            ws.set_column(0, len(cols_output), 15)

        # --- 4. GENERACIÓN DE HOJAS ---
        
        # Hoja 1: Pares
        df_f1 = df_procesado[df_procesado['Estado_Cofersa'] == 'PARES_1_A_1'].copy()
        escribir_hoja_simple('1. Pares 1 a 1', df_f1, "TOTAL PARES:")

        # Hoja 2: Cruce por Referencia
        df_f2 = df_procesado[df_procesado['Estado_Cofersa'] == 'CRUCE_POR_REFERENCIA'].copy()
        escribir_hoja_simple('2. Cruce por Referencia', df_f2, "TOTAL CRUCES REF:")

        # Hoja 3: Descuadres (Agrupado por Referencia)
        df_f3 = df_procesado[df_procesado['Estado_Cofersa'].isin(['REF_DESCUADRE', 'RECUPERACION_GLOBAL'])].copy()
        escribir_hoja_agrupada('3. Descuadres y Recuperados', df_f3, "TOTAL:", diff_fmt)

        # --- Lógica de Embarques Pendientes ---
        df_pendientes = df_procesado[df_procesado['Estado_Cofersa'] == 'PENDIENTE'].copy()
        
        if not df_pendientes.empty:
            # Detección segura de "EM"
            mask_ref = df_pendientes['Referencia'].astype(str).str.upper().str.contains('EM', na=False)
            mask_tipo = df_pendientes['Tipo'].astype(str).str.upper().str.contains('EM', na=False) if 'Tipo' in df_pendientes.columns else mask_ref
            mask_embarque = mask_ref | mask_tipo
            
            df_embarques = df_pendientes[mask_embarque].copy()
            df_otros = df_pendientes[~mask_embarque].copy()

            def extraer_codigo_embarque(row):
                r = str(row.get('Referencia', '')).strip().upper()
                t = str(row.get('Tipo', '')) if 'Tipo' in row else ''
                match = re.search(r'(E?M\d+)', r + t)
                return match.group(1) if match else r

            if not df_embarques.empty:
                df_embarques['Ref_Norm'] = df_embarques.apply(extraer_codigo_embarque, axis=1)
                escribir_hoja_agrupada('4. Embarques Pend.', df_embarques, "TOTAL EMBARQUES:", total_fmt)
            
            if not df_otros.empty:
                escribir_hoja_simple('5. Otros Pendientes', df_otros, "TOTAL OTROS:")
        
        # --- Hoja 6: Especificación (Resumen Ejecutivo) ---
        ws6 = workbook.add_worksheet('6. Especificación')
        ws6.hide_gridlines(2)
        ws6.merge_range('A1:F1', "RESUMEN DE SALDOS ABIERTOS (COFERSA)", workbook.add_format({'bold':True, 'font_size':14, 'align':'center'}))
        headers_spec = ['Categoría', 'Referencia', 'NIT', 'Descripción Nit', 'Fecha (Max)', 'Saldo (VES)']
        ws6.write_row('A3', headers_spec, header_fmt)
        
        # Recolectar datos para el resumen
        df_open = df_procesado[~df_procesado['Estado_Cofersa'].isin(['PARES_1_A_1', 'CRUCE_POR_REFERENCIA'])].copy()
        if not df_open.empty:
            if 'Ref_Norm' not in df_open.columns: df_open['Ref_Norm'] = df_open['Referencia'].astype(str).str.strip().str.upper()
            
            resumen = df_open.groupby('Ref_Norm').agg({
                'Neto Local': 'sum', 'Fecha': 'max', 'Nit': 'first', 'Descripción Nit': 'first'
            }).reset_index()
            
            r_idx = 3
            total_gral = 0
            for _, row in resumen.iterrows():
                if abs(row['Neto Local']) > 0.01:
                    ws6.write(r_idx, 0, "PENDIENTE", text_fmt)
                    ws6.write(r_idx, 1, row['Ref_Norm'], text_fmt)
                    ws6.write(r_idx, 2, row['Nit'], text_fmt)
                    ws6.write(r_idx, 3, row['Descripción Nit'], text_fmt)
                    ws6.write_datetime(r_idx, 4, row['Fecha'], date_fmt)
                    ws6.write_number(r_idx, 5, row['Neto Local'], money_fmt)
                    total_gral += row['Neto Local']
                    r_idx += 1
            
            ws6.write(r_idx, 4, "SALDO TOTAL:", label_fmt)
            ws6.write_number(r_idx, 5, total_gral, total_fmt)
        
        ws6.set_column('A:A', 20); ws6.set_column('B:D', 30); ws6.set_column('E:F', 20)

    return output.getvalue()

def generar_reporte_debito_fiscal(df_incidencias_raw, df_soft_raw, df_imp_raw):
    output = BytesIO()
    
    # Función de limpieza interna para evitar el error de NAN/INF
    def clean_num(val):
        try:
            if pd.isna(val) or np.isinf(val): return 0.0
            return float(val)
        except: return 0.0

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- FORMATOS EXISTENTES ---
        fmt_money = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        fmt_date = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})
        fmt_total_val = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bold': True, 'bg_color': '#E2EFDA'}) # Verde claro
        fmt_total_label = workbook.add_format({'bold': True, 'border': 1, 'align': 'right', 'bg_color': '#E2EFDA'})
        fmt_text = workbook.add_format({'border': 1})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        fmt_red_incidencia = workbook.add_format({'font_color': '#FF0000', 'num_format': '#,##0.00', 'border': 1})

        # Formatos para Totales Amarillos (Hoja 1)
        fmt_total_yellow = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bold': True, 'bg_color': '#FFFF00'})
        fmt_label_grey = workbook.add_format({'bold': True, 'border': 1, 'align': 'right', 'bg_color': '#F2F2F2'})    
        
        # Formatos para Tablas Resumen BI (Hoja 3)
        fmt_res_header = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'align': 'center', 'border': 1})
        fmt_res_total = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bold': True, 'bg_color': '#E2EFDA'})
        fmt_res_label = workbook.add_format({'bold': True, 'border': 1, 'align': 'right', 'bg_color': '#E2EFDA'})
        fmt_red_incid = workbook.add_format({'font_color': '#FF0000', 'num_format': '#,##0.00', 'border': 1})
        fmt_res_title = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left', 'bottom': 2, 'font_color': '#003366'})
        fmt_huerfanos_title = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left', 'bottom': 2, 'font_color': '#CC0000'})
        
        # Títulos de las tablas de control
        fmt_sep_casa = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'font_size': 11, 'border': 1})

        
        # ============================================================
        # HOJA 1: TRANSACCIONES SOFTLAND
        # ============================================================
        df_s_export = df_soft_raw.drop(columns=[c for c in df_soft_raw.columns if c.startswith('_')], errors='ignore')
        
        # Aseguramos que CASA sea la primera columna
        cols = ['CASA'] + [c for c in df_s_export.columns if c != 'CASA']
        df_s_export = df_s_export[cols]

        sheet_name1 = '1. Transacciones Softland'
        df_s_export.to_excel(writer, sheet_name=sheet_name1, index=False)
        ws1 = writer.sheets[sheet_name1]
        
        ws1.hide_gridlines(2) # Fondo blanco
        ws1.freeze_panes(1, 0) # Inmovilizar fila 1
        
        # Identificar columnas
        idx_date = -1; idx_deb = -1; idx_cre = -1
        for i, col in enumerate(df_s_export.columns):
            c_u = str(col).upper()
            if 'FECHA' in c_u: idx_date = i
            if any(k in c_u for k in ['DEBITO BOLIVAR', 'DEBITO LOCAL', 'DEBITO VES', 'DÉBITO VES']): idx_deb = i
            if any(k in c_u for k in ['CREDITO BOLIVAR', 'CREDITO LOCAL', 'CREDITO VES', 'CRÉDITO VES']): idx_cre = i

        # Aplicar datos con bordes y formatos
        num_rows = len(df_s_export)
        for r in range(num_rows):
            for c in range(len(df_s_export.columns)):
                val = df_s_export.iloc[r, c]
                if c == idx_date:
                    ws1.write_datetime(r + 1, c, val, fmt_date)
                elif c == idx_deb or c == idx_cre:
                    ws1.write_number(r + 1, c, float(val or 0), fmt_money)
                else:
                    ws1.write(r + 1, c, str(val) if pd.notna(val) else "", fmt_text)

        # --- FILA DE TOTALES (AMARILLA) ---
        total_row_idx = num_rows + 1
        # Llenar toda la fila con bordes vacíos primero
        for c in range(len(df_s_export.columns)):
            ws1.write(total_row_idx, c, "", fmt_text)

        ws1.write(total_row_idx, 0, "TOTALES:", fmt_total_label)
        
        if idx_deb != -1 and idx_cre != -1:
            col_deb_letter = xlsxwriter.utility.xl_col_to_name(idx_deb)
            col_cre_letter = xlsxwriter.utility.xl_col_to_name(idx_cre)
            
            # Fórmulas de Sumatoria
            ws1.write_formula(total_row_idx, idx_deb, f'=SUM({col_deb_letter}2:{col_deb_letter}{total_row_idx})', fmt_total_val)
            ws1.write_formula(total_row_idx, idx_cre, f'=SUM({col_cre_letter}2:{col_cre_letter}{total_row_idx})', fmt_total_val)
            
            # --- FILA DE SALDO (Débitos - Créditos) ---
            saldo_row_idx = total_row_idx + 1
            ws1.write(saldo_row_idx, 0, "SALDO:", fmt_total_label)
            # Escribir bordes en la fila de saldo
            for c in range(1, len(df_s_export.columns)):
                ws1.write(saldo_row_idx, c, "", fmt_text)
            
            # El saldo se coloca debajo de la columna de Créditos (o donde prefieras)
            # Requerimiento: "sume los debitos y reste creditos"
            formula_saldo = f'={col_deb_letter}{saldo_row_idx}-{col_cre_letter}{saldo_row_idx}'
            ws1.write_formula(saldo_row_idx, idx_cre, formula_saldo, fmt_total_val)

        # ============================================================
        # HOJA 2: LIBRO DE VENTAS (Copia exacta)
        # ============================================================
        # Usamos header=False e index=False para que no cree títulos nuevos
        df_imp_raw.to_excel(writer, sheet_name='2. Libro de Ventas', index=False, header=False)
        ws2 = writer.sheets['2. Libro de Ventas']

        # Fondo Blanco (Ocultar líneas de división)
        ws2.hide_gridlines(2)

        # Ajustar anchos para que se vea igual al original
        ws2.set_column('A:Z', 15)
        
        # ============================================================
        # HOJA 3: INCIDENCIAS
        # ============================================================
        ws3 = workbook.add_worksheet('3. Incidencias')
        ws3.hide_gridlines(2)
        
        # Filtramos para quitar FEBECA y registros OK
        df_audit = df_incidencias_raw[df_incidencias_raw['Estado'] != 'OK'].copy()
        df_audit = df_audit[~df_audit['_Nombre_Final'].str.upper().str.contains("FEBECA", na=False)]
        
        headers_inc = ['Casa', 'NIT', 'Nombre Proveedor', 'Tipo Doc', 'Documento', 'Monto Softland', 'Monto Imprenta', 'Estado']
        ws3.write_row(0, 0, headers_inc, fmt_header)
        
        row_i = 1
        casas_activas = sorted([c for c in df_audit['CASA'].unique() if pd.notna(c) and c != 'Libro Ventas'])
        
        # --- BLOQUE IZQUIERDO: DETALLE AGRUPADO POR CASA ---
        for casa in casas_activas:
            df_c = df_audit[df_audit['CASA'] == casa]
            if df_c.empty: continue
            
            ws3.merge_range(row_i, 0, row_i, 7, f"INCIDENCIAS DETECTADAS EN CASA: {casa}", fmt_sep_casa)
            row_i += 1
            start_block = row_i + 1
            
            for _, row in df_c.iterrows():
                ws3.write(row_i, 0, casa, fmt_text)
                ws3.write(row_i, 1, str(row.get('_NIT_Norm', '')), fmt_text)
                ws3.write(row_i, 2, str(row.get('_Nombre_Final', '')), fmt_text)
                ws3.write(row_i, 3, str(row.get('_Tipo_Final', '')), fmt_text)
                ws3.write(row_i, 4, str(row.get('_Doc_Norm', '')), fmt_text)
                ws3.write_number(row_i, 5, clean_num(row.get('_Monto_Bs_Soft')), fmt_money)
                ws3.write_number(row_i, 6, clean_num(row.get('_Monto_Imprenta')), fmt_money)
                ws3.write(row_i, 7, str(row.get('Estado', '')), fmt_text)
                row_i += 1
            
            ws3.write(row_i, 4, f"SUBTOTAL {casa}:", fmt_label_grey)
            ws3.write_formula(row_i, 5, f'=SUM(F{start_block}:F{row_i})', fmt_total_yellow)
            ws3.write_formula(row_i, 6, f'=SUM(G{start_block}:G{row_i})', fmt_total_yellow)
            row_i += 2

        # Bloque de Huérfanos
        df_h = df_audit[df_audit['_merge'] == 'right_only']
        if not df_h.empty:
            ws3.merge_range(row_i, 0, row_i, 7, "DOCUMENTOS SOLO EN IMPRENTA (Faltan por Contabilizar)", fmt_sep_casa)
            row_i += 1
            start_h = row_i + 1
            for _, row in df_h.iterrows():
                ws3.write(row_i, 0, "Imprenta", fmt_text)
                ws3.write(row_i, 1, str(row.get('_NIT_Norm', '')), fmt_text)
                ws3.write(row_i, 2, str(row.get('_Nombre_Final', '')), fmt_text)
                ws3.write(row_i, 3, str(row.get('_Tipo_Final', '')), fmt_text)
                ws3.write(row_i, 4, str(row.get('_Doc_Norm', '')), fmt_text)
                ws3.write_number(row_i, 5, 0.0, fmt_money)
                ws3.write_number(row_i, 6, clean_num(row.get('_Monto_Imprenta')), fmt_money)
                ws3.write(row_i, 7, str(row.get('Estado', '')), fmt_text)
                row_i += 1
            ws3.write(row_i, 4, "SUBTOTAL HUÉRFANOS:", fmt_label_grey)
            ws3.write_formula(row_i, 5, f'=SUM(F{start_h}:F{row_i})', fmt_total_yellow)
            ws3.write_formula(row_i, 6, f'=SUM(G{start_h}:G{row_i})', fmt_total_yellow)

        # --- BLOQUE DERECHO: TABLAS DE CONTROL (BI) ---
        col_start = 9 # Columna J
        for c_idx, casa_cod in enumerate(casas_activas):
            c_col = col_start + (c_idx * 4)
            df_c_all = df_incidencias_raw[(df_incidencias_raw['CASA'] == casa_cod) & 
                                          (~df_incidencias_raw['_Nombre_Final'].str.upper().str.contains("FEBECA", na=False))]
            
            ws3.write(0, c_col, f"ANÁLISIS: {casa_cod}", fmt_res_title)
            
            def draw_box(ws, start_r, start_c, title, m_col, source_df):
                ws.write(start_r, start_c, title, fmt_res_header)
                ws.write(start_r, start_c+1, "Cant.", fmt_res_header)
                ws.write(start_r, start_c+2, "Monto", fmt_res_header)
                tipos = [("FACTURA", "Total Débito Facturas"), ("N/C", "Total Débito N/C"), ("N/D", "Total Débito N/D")]
                curr_r, t_c, t_m = start_r + 1, 0, 0
                for code, label in tipos:
                    sub = source_df[source_df['_Tipo_Final'] == code]
                    cant = len(sub[sub[m_col] > 0.01])
                    monto = sub[m_col].sum()
                    ws.write(curr_r, start_c, label, fmt_text)
                    ws.write(curr_r, start_c+1, cant, fmt_text)
                    ws.write_number(curr_r, start_c+2, clean_num(monto), fmt_money)
                    t_c += cant; t_m += monto; curr_r += 1
                ws.write(curr_r, start_c, "Totales", fmt_res_label)
                ws.write(curr_r, start_c+1, t_c, fmt_res_total)
                ws.write_number(curr_r, start_c+2, clean_num(t_m), fmt_res_total)
                return curr_r + 2

            r_next = draw_box(ws3, 1, c_col, "Softland", "_Monto_Bs_Soft", df_c_all)
            r_next = draw_box(ws3, r_next, c_col, "Imprenta", "_Monto_Imprenta", df_c_all)
            
            # Bloque Diferencias
            ws3.write(r_next, c_col, "Diferencias", fmt_res_header)
            ws3.write(r_next, c_col+1, "Cant.", fmt_res_header)
            ws3.write(r_next, c_col+2, "Monto", fmt_res_header)
            r_next += 1
            df_dif_c = df_c_all[df_c_all['Estado'] != 'OK']
            te_c, te_m = 0, 0
            for code, label in [("FACTURA", "Total Débito Facturas"), ("N/C", "Total Débito N/C"), ("N/D", "Total Débito N/D")]:
                sub_e = df_dif_c[df_dif_c['_Tipo_Final'] == code]
                m_diff = abs(sub_e['_Monto_Bs_Soft'].sum() - sub_e['_Monto_Imprenta'].sum())
                ws3.write(r_next, c_col, label, fmt_text)
                ws3.write(r_next, c_col+1, len(sub_e), fmt_text)
                if m_diff > 0.01: ws3.write_number(r_next, c_col+2, m_diff, fmt_red_incid)
                else: ws3.write(r_next, c_col+2, "-", fmt_text)
                te_c += len(sub_e); te_m += m_diff; r_next += 1
            ws3.write(r_next, c_col, "Totales", fmt_res_label)
            ws3.write(r_next, c_col+1, te_c, fmt_res_total)
            ws3.write_number(r_next, c_col+2, clean_num(te_m), fmt_res_total)

        # Cuadro de Huérfanos BI
        df_h_all = df_incidencias_raw[(df_incidencias_raw['_merge'] == 'right_only') & 
                                      (~df_incidencias_raw['_Nombre_Final'].str.upper().str.contains("FEBECA", na=False))]
        h_col = col_start + (len(casas_activas) * 4)
        ws3.write(0, h_col, "SOLO EN IMPRENTA", fmt_huerfanos_title)
        ws3.write(1, h_col, "Documento", fmt_res_header); ws3.write(1, h_col+1, "Cant.", fmt_res_header); ws3.write(1, h_col+2, "Monto", fmt_res_header)
        tip_h = [("FACTURA", "Facturas Huérfanas"), ("N/C", "N/C Huérfanas"), ("N/D", "N/D Huérfanas")]
        rh, thc, thm = 2, 0, 0
        for code, label in tip_h:
            sub_h = df_h_all[df_h_all['_Tipo_Final'] == code]
            m_h = sub_h['_Monto_Imprenta'].sum()
            ws3.write(rh, h_col, label, fmt_text); ws3.write(rh, h_col+1, len(sub_h), fmt_text)
            ws3.write_number(rh, h_col+2, clean_num(m_h), fmt_money)
            thc += len(sub_h); thm += m_h; rh += 1
        ws3.write(rh, h_col, "Total Huérfanos", fmt_res_label)
        ws3.write(rh, h_col+1, thc, fmt_res_total); ws3.write_number(rh, h_col+2, clean_num(thm), fmt_res_total)

        # Ajustes de ancho
        ws3.set_column('A:I', 18); ws3.set_column('C:C', 35); ws3.set_column('J:Z', 15)

    return output.getvalue()
