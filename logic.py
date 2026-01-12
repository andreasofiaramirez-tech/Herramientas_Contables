# ==============================================================================
# LOGIC.PY - MOTOR DE CÁLCULO Y REGLAS DE NEGOCIO
# ==============================================================================
import pandas as pd
import numpy as np
import re
from itertools import combinations
from io import BytesIO
from difflib import SequenceMatcher
import pdfplumber

# Importación de configuraciones desde mappings.py
from mappings import (
    normalize_account,
    TOLERANCIA_MAX_BS,
    TOLERANCIA_MAX_USD,
    CUENTAS_CONOCIDAS,
    CUENTAS_BANCO,
    NOMBRES_CUENTAS_OFICIALES,
    MAPEO_CB_CG_BEVAL,
    MAPEO_CB_CG_FEBECA,
    MAPEO_CB_CG_PRISMA,
    MAPEO_CB_CG_QUINCALLA
)

# ==============================================================================
# 1. FUNCIONES AUXILIARES GLOBALES
# ==============================================================================

def _limpiar_monto(valor):
    """Convierte cualquier formato numérico a float de forma robusta."""
    if pd.isna(valor) or str(valor).strip() == '': return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    
    t = str(valor).strip().replace('Bs', '').replace('$', '').replace(' ', '').replace('\xa0', '')
    
    # Negativos entre paréntesis
    signo = 1
    if '(' in t and ')' in t:
        signo = -1
        t = t.replace('(', '').replace(')', '')
    elif '-' in t:
        signo = -1
        t = t.replace('-', '')

    # Detección de separadores
    if ',' in t and '.' in t:
        if t.rfind(',') > t.rfind('.'): # VE: 1.234,56
            t = t.replace('.', '').replace(',', '.')
        else: # US: 1,234.56
            t = t.replace(',', '')
    elif ',' in t: t = t.replace(',', '.') 
    elif '.' in t: 
         if len(t.split('.')[-1]) == 3: t = t.replace('.', '')
         
    try: return float(t) * signo
    except: return 0.0

def es_texto_numerico(texto):
    """Detecta si un string es un número válido (para PDFs)."""
    if not texto: return False
    t = texto.strip()
    if t == '-': return True
    t_clean = re.sub(r'[^\d]', '', t)
    if len(t_clean) > 0:
        if '/' in t: return False # Es fecha
        if t.count('.') > 2: return False # Es cuenta contable
        return True
    return False

def es_palabra_similiar(texto_completo, palabra_objetivo, umbral=0.80):
    """Fuzzy matching para referencias."""
    if not texto_completo: return False
    texto_limpio = re.sub(r'[^A-Z0-9]', ' ', str(texto_completo).upper())
    palabras = texto_limpio.split()
    objetivo = palabra_objetivo.upper()
    for p in palabras:
        if p == objetivo: return True
        if abs(len(p) - len(objetivo)) > 3: continue
        if SequenceMatcher(None, p, objetivo).ratio() >= umbral: return True
    return False

# ==============================================================================
# 2. MÓDULOS DE CONCILIACIÓN BANCARIA (MOTORES DE CRUCE)
# ==============================================================================

# --- NORMALIZADORES ---
def normalizar_referencia_fondos_en_transito(df):
    df_copy = df.copy()
    def clasificar(referencia_str):
        if pd.isna(referencia_str): return 'OTRO', 'OTRO', ''
        ref = str(referencia_str).upper().strip()
        ref_lit_norm = re.sub(r'[^A-Z0-9]', '', ref)
        if any(k in ref for k in ['DIFERENCIA EN CAMBIO', 'DIF. CAMBIO', 'DIFERENCIAL']): return 'DIF_CAMBIO', 'GRUPO_DIF_CAMBIO', ref_lit_norm
        if 'AJUSTE' in ref: return 'AJUSTE_GENERAL', 'GRUPO_AJUSTE', ref_lit_norm
        if 'REINTEGRO' in ref or 'SILLACA' in ref: return 'REINTEGRO_SILLACA', 'GRUPO_SILLACA', ref_lit_norm
        if 'REMESA' in ref: return 'REMESA_GENERAL', 'GRUPO_REMESA', ref_lit_norm
        if 'NOTA DE' in ref: return 'NOTA_GENERAL', 'GRUPO_NOTA', ref_lit_norm
        if 'BANCO A BANCO' in ref: return 'BANCO_A_BANCO', 'GRUPO_BANCO', ref_lit_norm
        return 'OTRO', 'OTRO', ref_lit_norm
    df_copy[['Clave_Normalizada', 'Clave_Grupo', 'Referencia_Normalizada_Literal']] = df_copy['Referencia'].apply(clasificar).apply(pd.Series)
    return df_copy

def normalizar_referencia_fondos_usd(df):
    df_copy = df.copy()
    def clasificar(ref_str):
        if pd.isna(ref_str): return 'OTRO', 'OTRO', 'OTRO'
        ref = str(ref_str).upper().strip()
        ref_lit_norm = re.sub(r'[^\w]', '', ref)
        if 'TRASPASO' in ref: return 'TRASPASO', 'GRUPO_TRASPASO', ref_lit_norm
        if 'DIFERENCIA' in ref and 'CAMBIO' in ref: return 'DIF_CAMBIO', 'GRUPO_DIF_CAMBIO', ref_lit_norm
        if 'BANCO A BANCO' in ref: return 'BANCO_A_BANCO', 'GRUPO_BANCO', 'BANCO_A_BANCO'
        if 'BANCARIZACION' in ref: return 'BANCARIZACION', 'GRUPO_BANCARIZACION', ref_lit_norm
        if 'REMESA' in ref: return 'REMESA', 'GRUPO_REMESA', ref_lit_norm
        if 'TARJETA' in ref and ('GASTOS' in ref or 'INGRESO' in ref): return 'TARJETA_GASTOS', 'GRUPO_TARJETA', 'LOTE_TARJETAS'
        if 'NOTA DE' in ref: return 'NOTA_DEBITO', 'GRUPO_NOTA', 'NOTA_DEBITO'
        return 'OTRO', 'OTRO', ref_lit_norm
    df_copy[['Clave_Normalizada', 'Clave_Grupo', 'Referencia_Normalizada_Literal']] = df_copy['Referencia'].apply(clasificar).apply(pd.Series)
    return df_copy

def normalizar_datos_deudores_empleados(df, log_messages):
    df_copy = df.copy()
    nit_col = next((c for c in df_copy.columns if c in ['NIT', 'RIF']), None)
    if nit_col:
        df_copy['Clave_Empleado'] = df_copy[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        df_copy['Clave_Empleado'] = 'SIN_NIT'
    return df_copy

def normalizar_datos_otras_cxp(df, log_messages):
    df_copy = df.copy()
    nit_col = next((c for c in df_copy.columns if c in ['NIT', 'RIF']), None)
    if nit_col: df_copy['NIT_Normalizado'] = df_copy[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else: df_copy['NIT_Normalizado'] = 'SIN_NIT'
    df_copy['Numero_Envio'] = df_copy['Referencia'].str.extract(r"ENV.*?(\d+)", expand=False, flags=re.IGNORECASE).fillna('')
    return df_copy

def normalizar_datos_cobros_viajeros(df, log_messages):
    df_copy = df.copy()
    nit_col = next((c for c in df_copy.columns if c in ['NIT', 'RIF']), None)
    if nit_col: df_copy['NIT_Normalizado'] = df_copy[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else: df_copy['NIT_Normalizado'] = 'SIN_NIT'
    
    def ext_num(t): return re.sub(r'\D', '', str(t)) if pd.notna(t) else ''
    df_copy['Referencia_Norm_Num'] = df_copy['Referencia'].apply(ext_num)
    df_copy['Fuente_Norm_Num'] = df_copy['Fuente'].apply(ext_num)
    df_copy['Es_Reverso'] = df_copy['Referencia'].str.contains('REVERSO', case=False, na=False)
    return df_copy

def normalizar_referencia_viajes(df, log_messages):
    def clasificar_tipo(ref):
        if pd.isna(ref): return 'OTRO'
        r = str(ref).upper()
        if 'TIMBRES' in r or 'FISCAL' in r: return 'IMPUESTOS'
        if 'VIAJE' in r or 'VIATICOS' in r: return 'VIATICOS'
        return 'OTRO'
    df['Tipo'] = df['Referencia'].apply(clasificar_tipo)
    nit_col = next((c for c in df.columns if c in ['NIT', 'RIF']), None)
    if nit_col: df['NIT_Normalizado'] = df[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else: df['NIT_Normalizado'] = 'SIN_NIT'
    return df

def normalizar_datos_proveedores(df, log_messages):
    df_copy = df.copy()
    nit_col = next((c for c in df_copy.columns if c in ['NIT', 'RIF']), None)
    if nit_col: df_copy['Clave_Proveedor'] = df_copy[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else: df_copy['Clave_Proveedor'] = df_copy['Nombre del Proveedor'].astype(str).str.strip().str.upper()
    
    def ext_clave(row):
        if row['Monto_USD'] > 0: return str(row['Fuente']).strip().upper()
        match = re.search(r'(COMP-\d+)', str(row['Referencia']).upper())
        return match.group(1) if match else np.nan
    df_copy['Clave_Comp'] = df_copy.apply(ext_clave, axis=1)
    return df_copy

def normalizar_datos_cdc_factoring(df, log_messages):
    df_copy = df.copy()
    nit_col = next((c for c in df_copy.columns if c in ['NIT', 'RIF']), None)
    if nit_col: df_copy['NIT_Normalizado'] = df_copy[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else: df_copy['NIT_Normalizado'] = 'SIN_NIT'

    def extraer_contrato(row):
        t = (str(row.get('Referencia', '')) + " " + str(row.get('Fuente', ''))).upper().strip()
        if not t: return 'SIN_CONTRATO'
        
        match_fq = re.search(r"(FQ-[A-Z0-9-]+)", t)
        if match_fq: return match_fq.group(1)
        match_oc = re.search(r"O/C\s*[-]?\s*([A-Z0-9]+)", t)
        if match_oc: return match_oc.group(1)
        
        if "FACTORING" in t:
            try:
                derecha = t.split("FACTORING", 1)[1].replace('$', ' ').replace(':', ' ')
                for p in derecha.split():
                    p_c = p.replace('.', '').strip()
                    if p_c in ['DE', 'DEL', 'NRO', 'NUM', 'REF', 'PAGO', 'FAC']: continue
                    if any(c.isdigit() for c in p_c) and len(p_c)>=3: return p_c
            except: pass
        
        # Nivel 3: Referencia directa
        t_ref = str(row.get('Referencia', '')).strip()
        if ' ' not in t_ref and any(c.isdigit() for c in t_ref) and len(t_ref) >= 4:
            return t_ref
            
        return 'SIN_CONTRATO'

    df_copy['Contrato'] = df_copy.apply(extraer_contrato, axis=1)
    return df_copy

# --- FUNCIONES DE CRUCE (HELPERS) ---

def conciliar_automaticos_usd(df, log_messages):
    total = 0
    grupos = [('GRUPO_DIF_CAMBIO', 'AUTOMATICO_DIF_CAMBIO'), ('GRUPO_AJUSTE', 'AUTOMATICO_AJUSTE'), ('GRUPO_TARJETA', 'AUTOMATICO_TARJETA')]
    for grupo, etiqueta in grupos:
        indices = df.loc[(df['Clave_Grupo'] == grupo) & (~df['Conciliado'])].index
        if not indices.empty:
            if abs(df.loc[indices, 'Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
                df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, etiqueta]
                total += len(indices)
    return total

def conciliar_diferencia_cambio(df, log_messages):
    indices = df[(df['Clave_Grupo'] == 'GRUPO_DIF_CAMBIO') & (~df['Conciliado'])].index
    if not indices.empty:
        df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTOMATICO_DIF_CAMBIO_SALDO']
    return len(indices)

def conciliar_ajuste_automatico(df, log_messages):
    indices = df[(df['Clave_Grupo'] == 'GRUPO_AJUSTE') & (~df['Conciliado'])].index
    if not indices.empty:
        df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTOMATICO_AJUSTE']
    return len(indices)

def conciliar_pares_exactos_cero(df, clave_grupo, fase_name, log_messages):
    pend = df[(df['Clave_Grupo'] == clave_grupo) & (~df['Conciliado'])]
    total = 0
    for _, grp in pend.groupby('Referencia_Normalizada_Literal'):
        if len(grp) < 2: continue
        debs = grp[grp['Monto_BS'] > 0].index; creds = grp[grp['Monto_BS'] < 0].index
        used_d, used_c = set(), set()
        for d in debs:
            if d in used_d: continue
            for c in creds:
                if c in used_c: continue
                if abs(df.loc[d, 'Monto_BS'] + df.loc[c, 'Monto_BS']) == 0:
                    df.loc[[d,c], ['Conciliado', 'Grupo_Conciliado']] = [True, f'PAR_EXACT_{df.loc[c,"Asiento"]}']
                    used_d.add(d); used_c.add(c); total += 2; break
    return total

def conciliar_pares_exactos_por_referencia(df, clave_grupo, fase_name, log_messages):
    pend = df[(df['Clave_Grupo'] == clave_grupo) & (~df['Conciliado'])]
    total = 0
    for _, grp in pend.groupby('Referencia_Normalizada_Literal'):
        if len(grp) < 2: continue
        debs = grp[grp['Monto_BS'] > 0].index.tolist(); creds = grp[grp['Monto_BS'] < 0].index.tolist()
        used_d, used_c = set(), set()
        for d in debs:
            if d in used_d: continue
            best_m, best_diff = None, TOLERANCIA_MAX_BS + 1
            for c in creds:
                if c in used_c: continue
                diff = abs(df.loc[d, 'Monto_BS'] + df.loc[c, 'Monto_BS'])
                if diff < best_diff: best_diff, best_m = diff, c
            if best_m is not None and best_diff <= TOLERANCIA_MAX_BS:
                df.loc[[d, best_m], ['Conciliado', 'Grupo_Conciliado']] = [True, f'PAR_REF_{df.loc[best_m,"Asiento"]}']
                used_d.add(d); used_c.add(best_m); total += 2
    return total

def cruzar_pares_simples(df, clave_norm, fase, log):
    pend = df[~df['Conciliado'] & (df['Clave_Normalizada'] == clave_norm)].copy()
    pend['M_Abs'] = pend['Monto_BS'].abs().round(0)
    total = 0
    for _, grp in pend.groupby('M_Abs'):
        debs = grp[grp['Monto_BS'] > 0].index.tolist(); creds = grp[grp['Monto_BS'] < 0].index.tolist()
        used_d, used_c = set(), set()
        for d in debs:
            if d in used_d: continue
            best_m, best_diff = None, TOLERANCIA_MAX_BS + 1
            for c in creds:
                if c in used_c: continue
                diff = abs(df.loc[d, 'Monto_BS'] + df.loc[c, 'Monto_BS'])
                if diff < best_diff: best_diff, best_m = diff, c
            if best_m is not None and best_diff <= TOLERANCIA_MAX_BS:
                df.loc[[d, best_m], ['Conciliado', 'Grupo_Conciliado']] = [True, f'PAR_BS_{df.loc[best_m,"Asiento"]}']
                used_d.add(d); used_c.add(best_m); total += 2
    return total

def cruzar_grupos_por_criterio(df, clave_norm, col, prefix, fase, log):
    pend = df[(df['Clave_Normalizada'] == clave_norm) & (~df['Conciliado'])]
    total = 0
    grouper = pend['Fecha'].dt.date.fillna('NaT') if col == 'Fecha' else pend[col]
    for crit, grp in pend.groupby(grouper):
        if len(grp) > 1 and abs(grp['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f'GRP_{prefix}_{crit}']
            total += len(grp)
    return total

def conciliar_lote_por_grupo(df, clave_grupo, fase, log):
    pend = df[~df['Conciliado'] & (df['Clave_Grupo'] == clave_grupo)]
    if len(pend) > 1 and abs(pend['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
        df.loc[pend.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f'LOTE_{clave_grupo}']
        return len(pend)
    return 0

def conciliar_grupos_globales_por_referencia(df, log):
    pend = df[~df['Conciliado'] & df['Referencia_Normalizada_Literal'].notna()]
    total = 0
    for _, grp in pend.groupby('Referencia_Normalizada_Literal'):
        if len(grp) > 1 and abs(grp['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f'GLOBAL_REF_{_}']
            total += len(grp)
    return total

def conciliar_pares_globales_remanentes(df, log):
    pend = df[~df['Conciliado']]
    if len(pend) < 2: return 0
    debs = pend[pend['Monto_BS'] > 0].index.tolist(); creds = pend[pend['Monto_BS'] < 0].index.tolist()
    total = 0; used_c = set()
    for d in debs:
        best_m, best_diff = None, TOLERANCIA_MAX_BS + 1
        for c in creds:
            if c in used_c: continue
            diff = abs(df.loc[d, 'Monto_BS'] + df.loc[c, 'Monto_BS'])
            if diff < best_diff: best_diff, best_m = diff, c
        if best_m is not None and best_diff <= TOLERANCIA_MAX_BS:
            df.loc[[d, best_m], ['Conciliado', 'Grupo_Conciliado']] = [True, f'GLOBAL_PAR_{df.loc[best_m,"Asiento"]}']
            used_c.add(best_m); total += 2
    return total

def conciliar_grupos_complejos_usd(df, log_messages, progress_bar=None):
    pendientes = df.loc[~df['Conciliado']]
    if len(pendientes) > 500: MAX_COMB = 3
    elif len(pendientes) > 200: MAX_COMB = 4
    else: MAX_COMB = 6
    
    debitos = pendientes[pendientes['Monto_USD'] > 0].copy()
    creditos = pendientes[pendientes['Monto_USD'] < 0].copy()
    
    total = 0
    for lado_a, lado_b in [(creditos, debitos), (debitos, creditos)]:
        lado_a = lado_a.sort_values('Monto_USD', key=abs)
        used = set()
        for idx_a, row_a in lado_a.iterrows():
            if idx_a in used: continue
            target = abs(row_a['Monto_USD'])
            candidatos = lado_b[~lado_b.index.isin(used) & (lado_b['Monto_USD'].abs() <= target + 0.01)]
            if len(candidatos) < 2: continue
            
            for r in range(2, MAX_COMB + 1):
                found = False
                for combo in combinations(candidatos.index, r):
                    if np.isclose(abs(lado_b.loc[list(combo), 'Monto_USD'].sum()), target, atol=0.01):
                        ids = list(combo) + [idx_a]
                        df.loc[ids, ['Conciliado', 'Grupo_Conciliado']] = [True, f'COMPLEJO_{row_a["Asiento"]}']
                        used.update(ids); total += len(ids); found = True; break
                if found: break
    return total

def conciliar_pares_globales_remanentes_usd(df, log):
    pend = df[~df['Conciliado']]
    if len(pend) < 2: return 0
    debs = pend[pend['Monto_USD'] > 0]; creds = pend[pend['Monto_USD'] < 0]
    
    pares = pd.merge(debs.reset_index(), creds.reset_index(), how='cross', suffixes=('_d', '_c'))
    pares['diff'] = (pares['Monto_USD_d'] + pares['Monto_USD_c']).abs()
    pares = pares[pares['diff'] <= TOLERANCIA_MAX_USD].sort_values('diff')
    
    total = 0; used_d, used_c = set(), set()
    for _, row in pares.iterrows():
        id_d, id_c = row['index_d'], row['index_c']
        if id_d not in used_d and id_c not in used_c:
            df.loc[[id_d, id_c], ['Conciliado', 'Grupo_Conciliado']] = [True, f'GLOBAL_USD_{df.loc[id_c,"Asiento"]}']
            used_d.add(id_d); used_c.add(id_c); total += 2
    return total

def conciliar_gran_total_final_usd(df, log):
    pend = df[~df['Conciliado']]
    if not pend.empty and abs(pend['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
        df.loc[pend.index, ['Conciliado', 'Grupo_Conciliado']] = [True, 'LOTE_FINAL']
        return len(pend)
    return 0

# --- FUNCIONES DE CONCILIACIÓN ESPECÍFICAS (RUNNERS) ---

def run_conciliation_fondos_en_transito(df, log_messages):
    df = normalizar_referencia_fondos_en_transito(df)
    conciliar_diferencia_cambio(df, log_messages)
    conciliar_ajuste_automatico(df, log_messages)
    conciliar_pares_exactos_cero(df, 'GRUPO_SILLACA', 'SILLACA 1', log_messages)
    conciliar_pares_exactos_por_referencia(df, 'GRUPO_SILLACA', 'SILLACA 2', log_messages)
    cruzar_pares_simples(df, 'REINTEGRO_SILLACA', 'SILLACA 3', log_messages)
    cruzar_grupos_por_criterio(df, 'REINTEGRO_SILLACA', 'Asiento', 'SIL_ASI', 'SILLACA 4', log_messages)
    cruzar_grupos_por_criterio(df, 'REINTEGRO_SILLACA', 'Referencia_Normalizada_Literal', 'SIL_REF', 'SILLACA 5', log_messages)
    cruzar_grupos_por_criterio(df, 'REINTEGRO_SILLACA', 'Fecha', 'SIL_FEC', 'SILLACA 6', log_messages)
    conciliar_lote_por_grupo(df, 'GRUPO_SILLACA', 'SILLACA 7', log_messages)
    # ... (Repetir lógica para NOTA y BANCO) ...
    conciliar_grupos_globales_por_referencia(df, log_messages)
    conciliar_pares_globales_remanentes(df, log_messages)
    return df

def run_conciliation_fondos_por_depositar(df, log_messages, progress_bar=None):
    df = normalizar_referencia_fondos_usd(df)
    conciliar_automaticos_usd(df, log_messages)
    
    # Agrupación Lógica (Tarjetas)
    pend = df[~df['Conciliado'] & df['Clave_Grupo'].isin(['GRUPO_TARJETA', 'GRUPO_REMESA'])]
    for grp, data in pend.groupby('Clave_Grupo'):
        if abs(data['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
            df.loc[data.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f'LOTE_{grp}']
            
    conciliar_grupos_por_referencia_usd(df, log_messages)
    conciliar_pares_banco_a_banco_usd(df, log_messages)
    conciliar_pares_globales_exactos_usd(df, log_messages)
    conciliar_pares_globales_remanentes_usd(df, log_messages)
    conciliar_grupos_complejos_usd(df, log_messages, progress_bar)
    conciliar_gran_total_final_usd(df, log_messages)
    return df

def run_conciliation_devoluciones_proveedores(df, log_messages):
    df = normalizar_datos_proveedores(df, log_messages)
    pend = df[(~df['Conciliado']) & df['Clave_Proveedor'].notna() & df['Clave_Comp'].notna()]
    for _, grp in pend.groupby(['Clave_Proveedor', 'Clave_Comp']):
        if abs(round(grp['Monto_USD'].sum(), 2)) <= TOLERANCIA_MAX_USD:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"PROV_{_}"]
    return df

def run_conciliation_viajes(df, log_messages, progress_bar=None):
    df = normalizar_referencia_viajes(df, log_messages)
    conciliar_pares_exactos_por_nit_viajes(df, log_messages)
    conciliar_grupos_por_nit_viajes(df, log_messages)
    return df

def run_conciliation_deudores_empleados_me(df, log_messages, progress_bar=None):
    df = normalizar_datos_deudores_empleados(df, log_messages)
    conciliar_grupos_por_empleado(df, log_messages)
    return df

def conciliar_grupos_por_empleado(df, log_messages):
    col_nom = next((c for c in df.columns if c in ['Descripcion NIT', 'Descripción Nit', 'Nombre']), None)
    for clv, grp in df[~df['Conciliado']].groupby('Clave_Empleado'):
        if clv in ['SIN_NIT', '']: continue
        if abs(grp['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"EMP_USD_{clv}"]
    return 0

def run_conciliation_deudores_empleados_bs(df, log_messages, progress_bar=None):
    df = normalizar_datos_deudores_empleados(df, log_messages)
    col_nom = next((c for c in df.columns if c in ['Descripcion NIT', 'Descripción Nit', 'Nombre']), None)
    
    # Fase Auto Diferencial
    idx_dif = df[df['Referencia'].str.contains('DIFF|DIFERENCIA|CAMBIO', case=False, na=False) & ~df['Conciliado']].index
    if not idx_dif.empty: df.loc[idx_dif, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTO_DIF_BS']
    
    for clv, grp in df[~df['Conciliado']].groupby('Clave_Empleado'):
        if clv in ['SIN_NIT', '']: continue
        if abs(grp['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"EMP_BS_{clv}"]
    return df

def run_conciliation_cobros_viajeros(df, log_messages, progress_bar=None):
    TOLERANCIA_ESTRICTA = 0.00
    df = normalizar_datos_cobros_viajeros(df, log_messages)
    
    idx_dif = df[df['Referencia'].str.contains('DIFF|DIFERENCIA|CAMBIO', case=False, na=False) & ~df['Conciliado']].index
    if not idx_dif.empty: df.loc[idx_dif, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTO_DIF_USD']
    
    # Reversos
    revs = df[df['Es_Reverso'] & ~df['Conciliado']]; origs = df[~df['Es_Reverso'] & ~df['Conciliado']]
    used = set(idx_dif)
    
    for ir, rr in revs.iterrows():
        if ir in used: continue
        for io, ro in origs.iterrows():
            if io in used or ro['NIT_Normalizado'] != rr['NIT_Normalizado']: continue
            match = (rr['Referencia_Norm_Num'] and rr['Referencia_Norm_Num'] in ro['Referencia_Norm_Num']) or \
                    (rr['Referencia_Norm_Num'] and rr['Referencia_Norm_Num'] in ro['Fuente_Norm_Num'])
            if match and abs(rr['Monto_USD'] + ro['Monto_USD']) <= TOLERANCIA_ESTRICTA:
                df.loc[[ir, io], ['Conciliado', 'Grupo_Conciliado']] = [True, f"REV_{rr['NIT_Normalizado']}"]
                used.update([ir, io]); break
                
    # Grupos NIT
    df['Clave_Vinculo'] = np.where(df['Asiento'].str.startswith('CC'), df['Fuente_Norm_Num'], df['Referencia_Norm_Num'])
    for (nit, c), grp in df[~df['Conciliado'] & (df['Clave_Vinculo']!='')].groupby(['NIT_Normalizado', 'Clave_Vinculo']):
        if len(grp) > 1 and abs(grp['Monto_USD'].sum()) <= TOLERANCIA_ESTRICTA:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"VIAJERO_{nit}_{c}"]
    return df

def run_conciliation_otras_cxp(df, log_messages, progress_bar=None):
    df = normalizar_datos_otras_cxp(df, log_messages)
    idx_dif = df[df['Referencia'].str.contains('DIFF|DIFERENCIA|CAMBIO', case=False, na=False) & ~df['Conciliado']].index
    if not idx_dif.empty: df.loc[idx_dif, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTO_DIF_BS']
    
    for (nit, env), grp in df[~df['Conciliado'] & (df['Numero_Envio']!='')].groupby(['NIT_Normalizado', 'Numero_Envio']):
        if len(grp) < 2: continue
        if abs(grp['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"CXP_{nit}_{env}"]
        elif len(grp) == 2 and abs(grp.iloc[0]['Monto_BS']) == abs(grp.iloc[1]['Monto_BS']):
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"CXP_MAG_{nit}_{env}"]
    return df

def run_conciliation_haberes_clientes(df, log_messages, progress_bar=None):
    df = normalizar_datos_otras_cxp(df, log_messages) # Reusamos
    
    for nit, grp in df[~df['Conciliado']].groupby('NIT_Normalizado'):
        if nit == 'SIN_NIT': continue
        if abs(grp['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"HABER_{nit}"]
            
    # Fase 2: Monto
    pend = df[~df['Conciliado']].copy(); pend['M_Abs'] = pend['Monto_BS'].abs()
    for m, grp in pend.groupby('M_Abs'):
        if len(grp) < 2 or m <= 0.01: continue
        debs = grp[grp['Monto_BS']>0].index; creds = grp[grp['Monto_BS']<0].index
        for i in range(min(len(debs), len(creds))):
            df.loc[[debs[i], creds[i]], ['Conciliado', 'Grupo_Conciliado']] = [True, f"HABER_MONTO_{int(m)}"]
    return df

def run_conciliation_cdc_factoring(df, log_messages, progress_bar=None):
    df = normalizar_datos_cdc_factoring(df, log_messages)
    
    idx_dif = df[df['Referencia'].str.contains('DIFF|DIFERENCIA|CAMBIO', case=False, na=False) & ~df['Conciliado']].index
    if not idx_dif.empty: df.loc[idx_dif, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTO_DIF_USD']
    
    for (nit, cont), grp in df[~df['Conciliado'] & (df['Contrato']!='SIN_CONTRATO')].groupby(['NIT_Normalizado', 'Contrato']):
        if abs(grp['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
            df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"FACT_{nit}_{cont}"]
    return df

def run_conciliation_asientos_por_clasificar(df, log_messages, progress_bar=None):
    TOLERANCIA_ESTRICTA = 0.00
    df_copy = df.copy()
    nit_col = next((c for c in df_copy.columns if c in ['NIT', 'RIF']), None)
    df['NIT_Normalizado'] = df_copy[nit_col].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True) if nit_col else 'SIN_NIT'

    idx_dif = df[df['Referencia'].str.contains('DIFF|DIFERENCIA|CAMBIO', case=False, na=False) & ~df['Conciliado']].index
    if not idx_dif.empty: df.loc[idx_dif, ['Conciliado', 'Grupo_Conciliado']] = [True, 'AUTO_DIF_BS']

    # Fase NIT
    for nit, grp in df[~df['Conciliado']].groupby('NIT_Normalizado'):
        if nit == 'SIN_NIT': continue
        if abs(grp['Monto_BS'].sum()) <= TOLERANCIA_ESTRICTA:
             df.loc[grp.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"CLASIF_NIT_{nit}"]
        else:
            # Pares exactos
            debs = grp[grp['Monto_BS']>0].index; creds = grp[grp['Monto_BS']<0].index
            used = set()
            for d in debs:
                for c in creds:
                    if c not in used and abs(df.loc[d,'Monto_BS'] + df.loc[c,'Monto_BS']) <= TOLERANCIA_ESTRICTA:
                        df.loc[[d,c], ['Conciliado', 'Grupo_Conciliado']] = [True, f"PAR_{nit}"]
                        used.add(c); break
                        
    # Fase Global
    pend = df[~df['Conciliado']].copy(); pend['M_Abs'] = pend['Monto_BS'].abs()
    for m, grp in pend.groupby('M_Abs'):
        if len(grp) < 2 or m <= 0.01: continue
        debs = grp[grp['Monto_BS']>0].index; creds = grp[grp['Monto_BS']<0].index
        for i in range(min(len(debs), len(creds))):
            df.loc[[debs[i], creds[i]], ['Conciliado', 'Grupo_Conciliado']] = [True, f"GLOBAL_{int(m)}"]
            
    # Fase Final Total
    rem = df[~df['Conciliado']]
    if not rem.empty:
        suma_final = round(rem['Monto_BS'].sum(), 2)
        log_messages.append(f"🔎 Suma Final Remanente: {suma_final}")
        if abs(suma_final) <= 0.01:
            df.loc[rem.index, ['Conciliado', 'Grupo_Conciliado']] = [True, 'LOTE_FINAL_REMANENTE']
            
    return df

# ==============================================================================
# 3. MÓDULO PAQUETE CC
# ==============================================================================

def _validar_asiento(asiento_group):
    grupo = asiento_group['Grupo'].iloc[0]
    monto_abs = asiento_group['Monto_USD'].abs()
    ref = str(asiento_group['Referencia'].iloc[0]).upper()

    if grupo.startswith("Grupo 6") and (monto_abs > 25).any(): return "Incidencia: Movimiento > $25."
    if grupo.startswith("Grupo 7"):
        if (monto_abs > 5).any():
            keys = ['TRASLADO', 'APLICAR', 'CRUCE', 'RECLASIFICACION', 'CORRECCION', 'TRASPASO']
            if not any(es_palabra_similiar(ref, k) for k in keys):
                return "Incidencia: Movimiento > $5 sin autorización."
        keys_err = ['DIFERENCIAL', 'DIF CAMBIARIO', 'TASA']
        if any(k in ref for k in keys_err) and "PRECIO" not in ref:
            return "Incidencia: Diferencial en cuenta de devoluciones."
            
    if grupo.startswith("Grupo 10"):
        if not np.isclose(asiento_group['Monto_USD'].sum(), 0, atol=0.01): return "Incidencia: Traspaso descuadrado."
        if not ((asiento_group['Monto_USD']>0).any() and (asiento_group['Monto_USD']<0).any()): return "Incidencia: Falta contrapartida."
        
    if grupo.startswith("Grupo 3"):
        cuentas = set(asiento_group['Cuenta Contable Norm'])
        if not (normalize_account('4.1.1.22.4.001') in cuentas and normalize_account('2.1.3.04.1.001') in cuentas):
            if "Error de Cuenta" in grupo: return "Incidencia: Diferencial en NC."
            return "Incidencia: NC Incompleta (Falta Desc o IVA)."

    if grupo.startswith("Grupo 9"):
        has_num = any(c.isdigit() for c in ref)
        has_key = any(k in ref for k in ['RET', 'IMP', 'ISLR', 'IVA'])
        if not (has_num or has_key): return "Incidencia: Ref inválida."

    if "Cuentas No Identificadas" in grupo or "No Clasificado" in grupo: return f"Incidencia: {grupo}"
    if grupo.startswith("Grupo 17"): return "Incidencia: Cuenta Transitoria. Verificar."
    
    return "Conciliado"

def _clasificar_asiento_paquete_cc(cuentas, ref, fuente, suma, max_abs):
    ref_clean = set(re.sub(r'[^\w\s]', '', ref).split())
    is_rev = 'REVERSO' in ref or 'REV' in ref_clean
    
    # Lógica Base
    def base(is_r):
        # 1. NC
        if 'N/C' in fuente or normalize_account('4.1.1.22.4.001') in cuentas:
            if is_r: return "Grupo 3: N/C"
            if ref_clean.intersection({'DIFERENCIAL', 'CAMBIO'}): return "Grupo 3: N/C - Posible Error de Cuenta"
            return "Grupo 3: N/C"
        
        # 2. Gastos Venta
        keys_merc = {'EXHIBIDOR', 'OBSEQUIO', 'MERCADEO'}
        if normalize_account('7.1.3.19.1.012') in cuentas or not keys_merc.isdisjoint(ref_clean):
            return "Grupo 4: Gastos de Ventas"
            
        # 3. Diferencial Puro
        if normalize_account('6.1.1.12.1.001') in cuentas and CUENTAS_BANCO.isdisjoint(cuentas):
            return "Grupo 2: Diferencial Cambiario"
            
        # 4. Retenciones
        if not {normalize_account('2.1.3.04.1.006'), normalize_account('2.1.3.01.1.012')}.isdisjoint(cuentas):
            return "Grupo 9: Retenciones"
            
        # 5. Traspasos/Devoluciones
        if normalize_account('4.1.1.21.4.001') in cuentas:
            if es_palabra_similiar(ref, 'TRASPASO') and abs(suma) <= 0.01: return "Grupo 10: Traspasos"
            return "Grupo 7: Devoluciones y Rebajas"
            
        # 6. Cobranzas
        if 'RECIBO' in ref or 'TEF' in fuente or 'DEPR' in fuente or not CUENTAS_BANCO.isdisjoint(cuentas):
            return "Grupo 8: Cobranzas"
            
        # 7. Ingresos Varios
        if normalize_account('6.1.1.19.1.001') in cuentas: return "Grupo 6: Ingresos Varios"
        
        # Otros
        if not {normalize_account('1.9.1.01.3.008'), normalize_account('1.9.1.01.3.009')}.isdisjoint(cuentas): return "Grupo 14: Inv. Oficinas"
        if normalize_account('7.1.3.01.1.001') in cuentas: return "Grupo 15: Incobrables"
        if normalize_account('1.1.4.01.7.044') in cuentas: return "Grupo 16: CxC Varios ME"
        if normalize_account('2.1.2.05.1.005') in cuentas: return "Grupo 17: Asientos por Clasificar"
        
        return "No Clasificado"

    res = base(is_rev)
    if is_rev and res != "No Clasificado":
        return f"{res.split(':')[0]}: Reversos - {res.split(':')[1]}"
    return res

def run_analysis_paquete_cc(df_diario, log_messages):
    log_messages.append("--- INICIANDO ANÁLISIS PAQUETE CC ---")
    df = df_diario.copy()
    
    # Normalización NIT
    rename_map = {}
    for c in df.columns:
        if c.strip().upper() in ['NIT', 'RIF']: rename_map[c] = 'NIT'
        if c.strip().upper() in ['DESCRIPCIÓN NIT', 'NOMBRE']: rename_map[c] = 'Nombre'
    df.rename(columns=rename_map, inplace=True)
    if 'NIT' not in df.columns: df['NIT'] = ''
    if 'Nombre' not in df.columns: df['Nombre'] = ''
    df['NIT'] = df['NIT'].fillna(''); df['Nombre'] = df['Nombre'].fillna('')

    df['Cuenta Contable Norm'] = df['Cuenta Contable'].astype(str).str.replace(r'\D', '', regex=True)
    df['Monto_USD'] = (df['Débito Dolar'] - df['Crédito Dolar']).round(2)
    
    # Agregación
    grp = df.groupby('Asiento')
    meta = pd.DataFrame({
        'Cuentas': grp['Cuenta Contable Norm'].apply(set),
        'Ref': grp['Referencia'].astype(str).apply(lambda x: ' '.join(x.unique()).upper()),
        'Fuente': grp['Fuente'].astype(str).apply(lambda x: ' '.join(x.unique()).upper()),
        'Suma': grp['Monto_USD'].sum(),
        'Max': grp['Monto_USD'].apply(lambda x: x.abs().max())
    })
    
    # Clasificación
    mapa = {}
    new_acc = 0
    for aid, r in meta.iterrows():
        miss = r['Cuentas'] - CUENTAS_CONOCIDAS
        if miss:
            mapa[aid] = f"Grupo 11: Cuentas No Identificadas ({','.join(miss)})"
            new_acc += 1
        else:
            mapa[aid] = _clasificar_asiento_paquete_cc(r['Cuentas'], r['Ref'], r['Fuente'], r['Suma'], r['Max'])
    df['Grupo'] = df['Asiento'].map(mapa)
    
    # Inteligencia Reversos
    is_rev = lambda r: any(k in (r['Ref']+r['Fuente']) for k in ['REVERSO', 'REV ', 'ANULA', 'CORRECCION'])
    ids_rev = set(df[df['Grupo'].str.contains('Reverso')]['Asiento']) | set(meta[meta.apply(is_rev, axis=1)].index)
    
    # Mapa Montos
    cands = df[~df['Asiento'].isin(ids_rev)].groupby('Asiento')['Monto_USD'].sum().round(2).reset_index()
    m_map = {}
    for _, r in cands.iterrows():
        m_map.setdefault(r['Monto_USD'], []).append(r['Asiento'])
        
    mapa_chg = {}; used = set()
    
    for rid in ids_rev:
        if rid in used: continue
        tgt = round(-meta.loc[rid, 'Suma'], 2)
        poss = [p for p in m_map.get(tgt, []) if p not in used]
        if not poss: continue
        
        # Match Fuerte (Ref)
        nums_r = re.findall(r'\d+', meta.loc[rid, 'Ref'] + meta.loc[rid, 'Fuente'])
        match = None
        for cid in poss:
            txt_c = meta.loc[cid, 'Ref'] + meta.loc[cid, 'Fuente']
            if any(n in txt_c for n in nums_r if len(n)>3):
                match = cid; break
        
        # Match Débil (Monto Único) - Protección Ret/Cobranza
        if not match and len(poss) == 1:
            grp_r = mapa.get(rid, ""); grp_c = mapa.get(poss[0], "")
            if not (grp_r.startswith("Grupo 9") or grp_c.startswith("Grupo 9") or grp_r.startswith("Grupo 8") or grp_c.startswith("Grupo 8")):
                match = poss[0]
                
        if match:
            mapa_chg[rid] = mapa_chg[match] = "Grupo 13: Operaciones Reversadas / Anuladas"
            used.update([rid, match])
            
    # Match ND vs NC (Barrido)
    left = [i for i in meta.index if i not in used and i not in ids_rev]
    abs_map = {}
    for i in left: abs_map.setdefault(abs(meta.loc[i,'Suma']), []).append(i)
    
    for m, lst in abs_map.items():
        if len(lst) < 2 or m <= 0.01: continue
        pos = [x for x in lst if meta.loc[x,'Suma']>0]
        neg = [x for x in lst if meta.loc[x,'Suma']<0]
        for p in pos:
            if p in used: continue
            nums_p = [n for n in re.findall(r'\d+', meta.loc[p,'Ref']+meta.loc[p,'Fuente']) if len(n)>3]
            if not nums_p: continue
            for n in neg:
                if n in used: continue
                if any(num in (meta.loc[n,'Ref']+meta.loc[n,'Fuente']) for num in nums_p):
                    mapa_chg[p] = mapa_chg[n] = "Grupo 13: Operaciones Reversadas / Anuladas"
                    used.update([p, n]); break
                    
    if mapa_chg: df['Grupo'] = df['Asiento'].map(mapa_chg).fillna(df['Grupo'])
    
    # Validación Final
    val_map = {}
    for aid, grp in df.groupby('Asiento'):
        val_map[aid] = "Conciliado (Anulado)" if grp['Grupo'].iloc[0].startswith("Grupo 13") else _validar_asiento(grp)
    df['Estado'] = df['Asiento'].map(val_map)
    
    # Orden
    df['Prio'] = df['Estado'].apply(lambda x: 1 if str(x).startswith('Conciliado') else 0)
    final = df.sort_values(['Grupo', 'Prio', 'Asiento'])
    return final.drop(columns=['Cuenta Contable Norm', 'Monto_USD', 'Ref_Str', 'Fuente_Str', 'Prio'], errors='ignore')

# ==============================================================================
# 4. MÓDULO CUADRE CB - CG
# ==============================================================================

def validar_coincidencia_empresa(file_obj, nombre_empresa_sel):
    key = "BEVAL" if "BEVAL" in nombre_empresa_sel else ("FEBECA" if "FEBECA" in nombre_empresa_sel else ("PRISMA" if "PRISMA" in nombre_empresa_sel else "SILLACA"))
    file_obj.seek(0)
    try:
        if file_obj.name.lower().endswith('.pdf'):
            with pdfplumber.open(file_obj) as pdf: text = pdf.pages[0].extract_text().upper() if pdf.pages else ""
        else:
            text = pd.read_excel(file_obj, header=None, nrows=10).to_string().upper()
    except: text = ""
    file_obj.seek(0)
    
    if key in text: return True, ""
    if key == "FEBECA" and "FEBECA" in text: return True, "" # Quincalla fix
    return False, f"Archivo no parece ser de {key}"

def extraer_saldos_cb(archivo, log_messages):
    datos = {}; name = getattr(archivo, 'name', '').lower()
    if name.endswith('.pdf'):
        try:
            with pdfplumber.open(archivo) as pdf:
                for p in pdf.pages:
                    txt = p.extract_text()
                    if not txt: continue
                    for line in txt.split('\n'):
                        parts = line.split()
                        if len(parts) < 3: continue
                        cod = parts[0].strip()
                        if len(cod)>=4 and cod[0].isdigit():
                            nums = [x for x in parts if es_texto_numerico(x)]
                            if len(nums) >= 1:
                                fin = _limpiar_monto(nums[-1])
                                ini = _limpiar_monto(nums[0]) if len(nums)>=2 else 0
                                deb = _limpiar_monto(nums[-3]) if len(nums)>=4 else 0
                                cre = _limpiar_monto(nums[-2]) if len(nums)>=4 else 0
                                nom = " ".join([x for x in parts[1:] if not es_texto_numerico(x) and not re.search(r'\d{2}/\d{2}', x)])
                                datos[cod] = {'inicial':ini, 'debitos':deb, 'creditos':cre, 'final':fin, 'nombre':nom}
        except Exception as e: log_messages.append(f"Err PDF CB: {e}")
    else:
        # Excel CB simple
        try:
            df = pd.read_excel(archivo)
            # (Lógica simplificada para excel, asume columnas estándar si existen)
            pass
        except: pass
    return datos

def extraer_saldos_cg(archivo, log_messages):
    datos = {}; name = getattr(archivo, 'name', '').lower()
    if name.endswith('.pdf'):
        try:
            with pdfplumber.open(archivo) as pdf:
                for p in pdf.pages:
                    txt = p.extract_text()
                    if not txt: continue
                    for line in txt.split('\n'):
                        parts = line.split()
                        if len(parts) < 3: continue
                        cta = parts[0].strip()
                        if not (cta.startswith('1.') and len(cta)>10): continue
                        
                        desc = NOMBRES_CUENTAS_OFICIALES.get(cta, " ".join([x for x in parts[1:] if not es_texto_numerico(x)]))
                        nums = [x for x in parts if es_texto_numerico(x)]
                        
                        ves = {'inicial':0, 'debitos':0, 'creditos':0, 'final':0}
                        usd = {'inicial':0, 'debitos':0, 'creditos':0, 'final':0}
                        
                        if len(nums) >= 8:
                            usd = {'inicial':_limpiar_monto(nums[-4]), 'debitos':_limpiar_monto(nums[-3]), 'creditos':_limpiar_monto(nums[-2]), 'final':_limpiar_monto(nums[-1])}
                            ves = {'inicial':_limpiar_monto(nums[-8]), 'debitos':_limpiar_monto(nums[-7]), 'creditos':_limpiar_monto(nums[-6]), 'final':_limpiar_monto(nums[-5])}
                        elif len(nums) >= 4:
                            ves = {'inicial':_limpiar_monto(nums[0]), 'debitos':_limpiar_monto(nums[1]), 'creditos':_limpiar_monto(nums[2]), 'final':_limpiar_monto(nums[3])}
                            
                        datos[cta] = {'VES': ves, 'USD': usd, 'descripcion': desc}
        except Exception as e: log_messages.append(f"Err PDF CG: {e}")
    return datos

def run_cuadre_cb_cg(file_cb, file_cg, nombre_empresa, log_messages):
    emp = nombre_empresa.upper()
    if "PRISMA" in emp: mapa = MAPEO_CB_CG_PRISMA
    elif "QUINCALLA" in emp or "SILLACA" in emp: mapa = MAPEO_CB_CG_QUINCALLA
    elif "FEBECA" in emp: mapa = MAPEO_CB_CG_FEBECA
    else: mapa = MAPEO_CB_CG_BEVAL
    
    d_cb = extraer_saldos_cb(file_cb, log_messages)
    d_cg = extraer_saldos_cg(file_cg, log_messages)
    
    # Agrupación N:1
    suma_cb = {}
    for cod, cfg in mapa.items():
        cta = cfg['cta']
        suma_cb[cta] = suma_cb.get(cta, 0) + d_cb.get(cod, {}).get('final', 0)
        
    res = []; huerfanos = []
    
    for cod, cfg in mapa.items():
        cta = cfg['cta']
        mon = cfg['moneda']
        
        i_cb = d_cb.get(cod, {'inicial':0, 'debitos':0, 'creditos':0, 'final':0, 'nombre':'NO ENCONTRADO'})
        i_cg_full = d_cg.get(cta, {})
        i_cg = i_cg_full.get('VES' if mon=='VES' else 'USD', {'inicial':0, 'debitos':0, 'creditos':0, 'final':0})
        desc = i_cg_full.get('descripcion', NOMBRES_CUENTAS_OFICIALES.get(cta, 'ND'))
        
        diff = round(suma_cb.get(cta, 0) - i_cg['final'], 2)
        estado = "OK" if diff == 0 else "DESCUADRE"
        s_cg_vis = i_cb['final'] if diff == 0 else i_cg['final'] # Truco visual si cuadra
        
        res.append({
            'Moneda': mon, 'Banco (Tesorería)': cod, 'Cuenta Contable': cta, 'Descripción': desc,
            'Saldo Final CB': i_cb['final'], 'Saldo Final CG': s_cg_vis, 'Diferencia': diff, 'Estado': estado,
            'CB Inicial': i_cb['inicial'], 'CB Débitos': i_cb['debitos'], 'CB Créditos': i_cb['creditos'],
            'CG Inicial': i_cg['inicial'], 'CG Débitos': i_cg['debitos'], 'CG Créditos': i_cg['creditos']
        })
        
    # Detección de Huérfanos
    mapped_cb = set(mapa.keys())
    mapped_cg = set(cfg['cta'] for cfg in mapa.values())
    
    for cod in set(d_cb.keys()) - mapped_cb:
        if d_cb[cod]['final'] != 0:
            huerfanos.append({'Origen': 'CB', 'Código/Cuenta': cod, 'Descripción/Nombre': d_cb[cod]['nombre'], 'Saldo Final': d_cb[cod]['final'], 'Mensaje': 'No mapeado'})
            
    for cta in set(d_cg.keys()) - mapped_cg:
        if (cta.startswith('1.1.1.02') or cta.startswith('1.1.1.03') or cta.startswith('1.1.1.06')) and not cta.endswith('.000'):
            s = d_cg[cta]
            if s['VES']['final'] != 0 or s['USD']['final'] != 0:
                huerfanos.append({'Origen': 'CG', 'Código/Cuenta': cta, 'Descripción/Nombre': s['descripcion'], 'Saldo Final': f"Bs:{s['VES']['final']} $:{s['USD']['final']}", 'Mensaje': 'No mapeado'})

    return pd.DataFrame(res), pd.DataFrame(huerfanos)

# ==============================================================================
# 5. MÓDULO GESTIÓN DE IMPRENTA (RETENCIONES IVA)
# ==============================================================================

# --- PARTE A: VALIDACIÓN (TXT vs TXT) ---

def parse_sales_txt(file_obj, log_messages):
    """Lee el TXT del Libro de Ventas y extrae facturas."""
    invoices_found = set()
    try:
        content = file_obj.getvalue().decode('latin-1') 
        lines = content.splitlines()
        regex_pattern = r"(?:FAC|N/C|N/D)\s*([0-9]+)"
        
        for line in lines:
            matches = re.findall(regex_pattern, line)
            for match in matches:
                clean_num = str(int(match)) 
                invoices_found.add(clean_num)
                
        log_messages.append(f"✅ Libro de Ventas procesado. {len(invoices_found)} documentos encontrados.")
        return invoices_found, lines
    except Exception as e:
        log_messages.append(f"❌ Error leyendo TXT Ventas: {str(e)}")
        return set(), []

def run_cross_check_imprenta(file_sales, file_retentions, log_messages):
    """Cruza Retenciones (TXT) vs Libro de Ventas (TXT)."""
    log_messages.append("\n--- INICIANDO CRUCE IMPRENTA (TXT) ---")
    
    valid_invoices, _ = parse_sales_txt(file_sales, log_messages)
    if not valid_invoices: return pd.DataFrame(), None

    resultados = []
    processed_invoices = set() 
    txt_original = ""
    
    try:
        content_ret = file_retentions.getvalue().decode('latin-1')
        txt_original = content_ret
        lines_ret = content_ret.splitlines()
        regex_ret = r"(FAC|N/C)\s*([0-9]+)"
        
        for line_idx, line in enumerate(lines_ret):
            match = re.search(regex_ret, line)
            if match:
                tipo = match.group(1)
                factura_raw = match.group(2)
                factura_clean = str(int(factura_raw))
                
                status = "OK"
                if factura_clean not in valid_invoices:
                    status = "ERROR: Factura no declarada en Libro de Ventas"
                if factura_clean in processed_invoices:
                    status = "ERROR: Factura duplicada en archivo de Retenciones"
                
                processed_invoices.add(factura_clean)
                
                resultados.append({
                    'Línea TXT': line_idx + 1, 'Contenido Original': line.strip(),
                    'Tipo': tipo, 'Factura Detectada': factura_raw, 'Estado': status
                })
    except Exception as e:
        log_messages.append(f"❌ Error procesando Retenciones: {str(e)}")
        return pd.DataFrame(), None

    df_res = pd.DataFrame(resultados)
    if not df_res.empty:
        errores = df_res[df_res['Estado'] != 'OK']
        if not errores.empty: log_messages.append(f"⚠️ Se encontraron {len(errores)} incidencias.")
        else: log_messages.append("✅ Validación exitosa.")
    return df_res, txt_original


# --- PARTE B: GENERACIÓN (Excel Softland -> TXT) ---

def indexar_libro_ventas(file_libro, log_messages):
    """Indexa el Excel del Libro de Ventas con búsqueda robusta de columnas."""
    db_ventas = {}
    
    def limpiar_monto_excel(valor):
        if pd.isna(valor) or str(valor).strip() == '': return 0.0
        t = str(valor).strip().replace('Bs', '').replace('$', '')
        if isinstance(valor, (int, float)): return float(valor)
        if ',' in t and '.' in t:
            if t.rfind(',') > t.rfind('.'): t = t.replace('.', '').replace(',', '.')
            else: t = t.replace(',', '')
        elif ',' in t: t = t.replace(',', '.')
        try: return float(t)
        except: return 0.0

    try:
        df_raw = pd.read_excel(file_libro, header=None)
        header_idx = None
        for i, row in df_raw.head(20).iterrows():
            row_str = row.astype(str).str.upper().values
            if any("FACTURA" in s for s in row_str) and (any("IMPUESTO" in s for s in row_str) or any("IVA" in s for s in row_str)):
                header_idx = i; break
        
        if header_idx is not None: df = pd.read_excel(file_libro, header=header_idx)
        else: df = pd.read_excel(file_libro)

        df.columns = [str(c).strip().upper().replace('\n', ' ') for c in df.columns]
        
        col_fac = next((c for c in df.columns if ('FACTURA' in c or 'NUMERO' in c) and 'AFECT' not in c and 'FECHA' not in c), None)
        col_fecha = next((c for c in df.columns if 'FECHA' in c and ('FACTURA' in c or 'EMISION' in c)), None)
        col_iva = next((c for c in df.columns if ('IMPUESTO' in c and 'IVA' in c) or 'IVA RETENIDO' in c or 'TOTAL IVA' in c), None)
        if not col_iva: col_iva = next((c for c in df.columns if 'IMPUESTO' in c and 'G' in c), None)
        col_rif = next((c for c in df.columns if 'RIF' in c and 'TERCERO' not in c), None)
        col_nom = next((c for c in df.columns if 'NOMBRE' in c or 'RAZON' in c), None)
        # Columna para validar si ya está cargada
        col_ret_libro = next((c for c in df.columns if 'IVA' in c and 'RETENIDO' in c), None)

        if not col_fac or not col_iva:
            log_messages.append("❌ Error: Faltan columnas 'Factura' o 'IVA' en Libro.")
            return {}

        for _, row in df.iterrows():
            raw_fac = str(row[col_fac])
            fac_limpia = re.sub(r'\D', '', raw_fac)
            if not fac_limpia: continue
            fac_int = str(int(fac_limpia))
            
            m_ret_libro = limpiar_monto_excel(row[col_ret_libro]) if col_ret_libro else 0.0
            
            db_ventas[fac_int] = {
                'fecha_factura': row[col_fecha] if col_fecha else None,
                'monto_iva': limpiar_monto_excel(row[col_iva]),
                'monto_ret_libro': m_ret_libro,
                'factura_original': raw_fac,
                'rif': str(row[col_rif]).strip() if col_rif else "ND",
                'nombre': str(row[col_nom]).strip() if col_nom else "ND"
            }
        log_messages.append(f"✅ Libro de Ventas indexado ({len(db_ventas)} facs).")
        return db_ventas
    except Exception as e:
        log_messages.append(f"❌ Error leyendo Libro Ventas: {str(e)}")
        return {}

def generar_txt_retenciones_galac(file_softland, file_libro, log_messages):
    """Genera TXT y Reporte Excel."""
    log_messages.append("--- INICIANDO GENERACIÓN DE TXT ---")
    db_ventas = indexar_libro_ventas(file_libro, log_messages)
    if not db_ventas: return [], None

    periodo_actual = "999999"
    try:
        fechas = [v['fecha_factura'] for k, v in db_ventas.items() if pd.notna(v['fecha_factura'])]
        if fechas:
            pers = [pd.to_datetime(f, dayfirst=True).strftime('%Y%m') for f in fechas]
            periodo_actual = max(set(pers), key=pers.count)
            log_messages.append(f"📅 Periodo Libro: {periodo_actual}")
    except: pass

    try:
        df_soft = pd.read_excel(file_softland)
        df_soft.columns = [str(c).strip().upper() for c in df_soft.columns]
        
        col_ref = next((c for c in df_soft.columns if 'REFERENCIA' in c), None)
        col_fecha = next((c for c in df_soft.columns if 'FECHA' in c), None)
        col_monto = next((c for c in df_soft.columns if 'DÉBITO' in c or 'DEBITO' in c), None)
        if not col_monto: col_monto = next((c for c in df_soft.columns if 'CRÉDITO' in c or 'CREDITO' in c), None)
        col_rif_soft = next((c for c in df_soft.columns if any(k in c for k in ['RIF', 'NIT', 'I.D', 'CEDULA'])), None)
        col_nom_soft = next((c for c in df_soft.columns if any(k in c for k in ['NOMBRE', 'CLIENTE', 'TERCERO'])), None)

        if not col_ref or not col_monto:
            log_messages.append("❌ Error: Faltan columnas Referencia o Monto en Softland.")
            return [], None
    except Exception as e:
        log_messages.append(f"❌ Error leyendo Softland: {str(e)}")
        return [], None

    filas_txt = []
    reporte_excel = []
    
    def crear_fila(estatus, mensaje, rif, nombre, comp, fac="", f_fac="", f_comp="", base=0.0, pct=0.0, m_ret=0.0, ref_orig="", m_soft=0.0):
        return {
            'Estatus': estatus, 'Mensaje': mensaje, 'RIF': rif, 'Nombre': nombre,
            'Comprobante': comp, 'Factura': fac, 'Fecha Factura': f_fac, 'Fecha Comprobante': f_comp,
            'Base IVA (Galac)': base, '% Calc': pct, 'Monto Retenido': m_ret,
            'Referencia Original': ref_orig, 'Monto Softland': m_soft
        }

    for idx, row in df_soft.iterrows():
        referencia = str(row[col_ref]).strip()
        try: monto_total_ret = float(row[col_monto])
        except: monto_total_ret = 0.0
        
        rif_c = str(row[col_rif_soft]).strip() if col_rif_soft else "ND"
        if rif_c.lower() == 'nan': rif_c = "ND"
        nom_c = str(row[col_nom_soft]).strip() if col_nom_soft else "ND"
        if nom_c.lower() == 'nan': nom_c = "ND"
        
        if monto_total_ret <= 0 or not referencia: continue
        
        parts = [p.strip() for p in referencia.split('/')]
        if len(parts) < 2:
            comp_temp = re.sub(r'\D', '', parts[0])
            msg = "Solo se detectó Comprobante" if len(comp_temp) > 10 else "Formato inválido"
            reporte_excel.append(crear_fila('FALTA INFO', msg, rif_c, nom_c, parts[0], "", "", "", 0, 0, 0, referencia, monto_total_ret))
            continue
            
        comprobante = re.sub(r'\D', '', parts[0])
        periodo_comp = comprobante[:6] if len(comprobante) >= 6 else "000000"
        es_periodo_anterior = periodo_comp < periodo_actual
        
        facturas_lista = []
        for f in parts[1:]:
            f_clean = re.sub(r'\D', '', f)
            if f_clean: facturas_lista.append(str(int(f_clean))) 
        if not facturas_lista: continue

        total_iva = 0.0; founds = []; missing = []
        for f_num in facturas_lista:
            if f_num in db_ventas:
                info = db_ventas[f_num]
                # Si ya tiene retención, la ignoramos para cálculo pero avisamos
                if info['monto_ret_libro'] > 0:
                    reporte_excel.append(crear_fila('YA CARGADA', 'Retención > 0 en Libro', rif_c, nom_c, comprobante, f_num, "", "", info['monto_iva'], 0, info['monto_ret_libro'], referencia, monto_total_ret))
                    continue
                
                total_iva += info['monto_iva']
                founds.append({'num': f_num, 'info': info, 'src': 'libro'})
            else: missing.append(f_num)

        if missing:
            if es_periodo_anterior:
                for f_miss in missing: founds.append({'num': f_miss, 'info': None, 'src': 'forced'})
            else:
                reporte_excel.append(crear_fila('FACTURA NO ENCONTRADA', f"Faltan: {','.join(missing)}", rif_c, nom_c, comprobante, "", "", "", 0, 0, 0, referencia, monto_total_ret))
                continue

        if not founds: continue # Si todas estaban 'YA CARGADA', terminamos

        has_forced = any(i['src'] == 'forced' for i in founds)
        if has_forced:
            m_ind = monto_total_ret / len(founds)
            for item in founds:
                nro = item['num'].zfill(10)
                try: f_comp = pd.to_datetime(row[col_fecha]).strftime('%d/%m/%Y')
                except: f_comp = ""
                filas_txt.append(f"FAC\t{nro}\t0\t{comprobante}\t{m_ind:.2f}\t{f_comp}\t{f_comp}")
                reporte_excel.append(crear_fila('OK - PERIODO ANTERIOR', 'Extemporáneo', rif_c, nom_c, comprobante, nro, f_comp, f_comp, 0, 0, m_ind, referencia, monto_total_ret))
        else:
            if total_iva == 0:
                reporte_excel.append(crear_fila('ERROR MATEMATICO', 'IVA 0.00', rif_c, nom_c, comprobante, "", "", "", 0, 0, 0, referencia, monto_total_ret))
                continue
            factor = monto_total_ret / total_iva
            note = "⚠️ Revisar %" if (factor > 1.01 or factor < 0.70) else "OK"
            for item in founds:
                iva_ind = item['info']['monto_iva']
                ret_ind = round(iva_ind * factor, 2)
                r_final = item['info'].get('rif', rif_c)
                n_final = item['info'].get('nombre', nom_c)
                try:
                    f_fac = pd.to_datetime(item['info']['fecha_factura'], dayfirst=True).strftime('%d/%m/%Y')
                    f_comp = pd.to_datetime(row[col_fecha]).strftime('%d/%m/%Y')
                except: f_fac=""; f_comp=""
                nro = item['num'].zfill(10)
                filas_txt.append(f"FAC\t{nro}\t0\t{comprobante}\t{ret_ind:.2f}\t{f_comp}\t{f_fac}")
                reporte_excel.append(crear_fila('GENERADO OK', note, r_final, n_final, comprobante, nro, f_fac, f_comp, iva_ind, factor, ret_ind, referencia, monto_total_ret))

    cols = ['Estatus', 'Mensaje', 'RIF', 'Nombre', 'Comprobante', 'Factura', 'Fecha Factura', 'Fecha Comprobante', 'Base IVA (Galac)', '% Calc', 'Monto Retenido', 'Referencia Original', 'Monto Softland']
    df_fin = pd.DataFrame(reporte_excel)
    if not df_fin.empty: df_fin = df_fin.reindex(columns=cols)
    return filas_txt, df_fin


# ==============================================================================
# 6. MÓDULO CÁLCULO PENSIONES (9%)
# ==============================================================================

def procesar_calculo_pensiones(file_mayor, file_nomina, tasa_cambio, nombre_empresa, log_messages):
    """
    Motor de cálculo para el impuesto del 9%.
    MEJORA FINAL: Desglose detallado de Salarios vs Tickets para validación.
    """
    log_messages.append(f"--- INICIANDO CÁLCULO DE PENSIONES (9%) - {nombre_empresa} ---")
    
    # Mapeo de nombres para búsqueda en nómina
    mapa_nombres = {
        "FEBECA, C.A": "FEBECA",
        "MAYOR BEVAL, C.A": "BEVAL",
        "PRISMA, C.A": "PRISMA",
        "FEBECA, C.A (QUINCALLA)": "QUINCALLA"
    }
    keyword_empresa = mapa_nombres.get(nombre_empresa, nombre_empresa).upper()

    mes_detectado = None
    nombres_meses = {1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'}

    # --- 1. PROCESAR MAYOR CONTABLE ---
    try:
        df_mayor = pd.read_excel(file_mayor)
        df_mayor.columns = [str(c).strip().upper() for c in df_mayor.columns]
        
        col_cta = next((c for c in df_mayor.columns if 'CUENTA' in c), None)
        col_cc = next((c for c in df_mayor.columns if 'CENTRO' in c and 'COSTO' in c), None)
        col_deb = next((c for c in df_mayor.columns if 'DÉBITO' in c or 'DEBITO' in c), None)
        col_cre = next((c for c in df_mayor.columns if 'CRÉDITO' in c or 'CREDITO' in c), None)
        col_fecha = next((c for c in df_mayor.columns if 'FECHA' in c), None)
        
        if not (col_cta and col_cc and col_deb and col_cre):
            log_messages.append("❌ Error: Faltan columnas críticas en el Mayor.")
            return None, None, None, None
            
        if col_fecha:
            try:
                fechas = pd.to_datetime(df_mayor[col_fecha], errors='coerce').dropna()
                if not fechas.empty:
                    mes_num = fechas.dt.month.mode()[0]
                    year_num = fechas.dt.year.mode()[0]
                    mes_detectado = nombres_meses[mes_num]
                    log_messages.append(f"📅 Periodo detectado: {mes_detectado} {year_num}")
            except: pass

        cuentas_base = ['7.1.1.01.1.001', '7.1.1.09.1.003']
        df_filtrado = df_mayor[df_mayor[col_cta].astype(str).str.strip().isin(cuentas_base)].copy()
        
        def clean_float(x):
            if pd.isna(x): return 0.0
            x = str(x).replace('Bs', '').replace(' ', '').replace(',', '')
            try: return float(x)
            except: return 0.0

        df_filtrado['Monto_Deb'] = df_filtrado[col_deb].apply(clean_float)
        df_filtrado['Monto_Cre'] = df_filtrado[col_cre].apply(clean_float)
        df_filtrado['Base_Neta'] = df_filtrado['Monto_Deb'] - df_filtrado['Monto_Cre']
        
        # Agrupación por 10 dígitos
        df_filtrado['CC_Agrupado'] = df_filtrado[col_cc].astype(str).str.slice(0, 10)
        df_agrupado = df_filtrado.groupby(['CC_Agrupado', col_cta]).agg({'Base_Neta': 'sum'}).reset_index()
        
        df_agrupado.rename(columns={'CC_Agrupado': 'Centro de Costo (Padre)', col_cta: 'Cuenta Contable'}, inplace=True)
        
        df_agrupado['Impuesto (9%)'] = df_agrupado['Base_Neta'] * 0.09
        
        # Totales Contables por Tipo
        base_salarios_cont = df_agrupado[df_agrupado['Cuenta Contable'].astype(str).str.contains('7.1.1.01', na=False)]['Base_Neta'].sum()
        base_tickets_cont = df_agrupado[df_agrupado['Cuenta Contable'].astype(str).str.contains('7.1.1.09', na=False)]['Base_Neta'].sum()
        total_base_contable = base_salarios_cont + base_tickets_cont

        log_messages.append(f"✅ Base Contable calculada: {total_base_contable:,.2f} Bs.")

    except Exception as e:
        log_messages.append(f"❌ Error procesando Mayor: {str(e)}")
        return None, None, None, None

    # --- 2. PROCESAR NÓMINA (VALIDACIÓN DETALLADA) ---
    val_salarios_nom = 0.0
    val_tickets_nom = 0.0
    val_impuesto_nom = 0.0
    
    try:
        if file_nomina:
            xls_nomina = pd.ExcelFile(file_nomina)
            hojas = xls_nomina.sheet_names
            hoja_objetivo = None
            if mes_detectado:
                for hoja in hojas:
                    if mes_detectado in hoja.upper():
                        hoja_objetivo = hoja; break
            if not hoja_objetivo: hoja_objetivo = hojas[0]
            
            # Buscar encabezado
            df_preview = pd.read_excel(xls_nomina, sheet_name=hoja_objetivo, header=None, nrows=15)
            header_idx = 0
            found_header = False
            for i, row in df_preview.iterrows():
                row_str = row.astype(str).str.upper().values
                if any("EMPRESA" in s for s in row_str) and any("APARTADO" in s for s in row_str):
                    header_idx = i; found_header = True; break
            
            if found_header:
                df_nom = pd.read_excel(xls_nomina, sheet_name=hoja_objetivo, header=header_idx)
            else:
                df_nom = pd.read_excel(xls_nomina, sheet_name=hoja_objetivo)

            df_nom.columns = [str(c).strip().upper() for c in df_nom.columns]
            
            col_empresa = next((c for c in df_nom.columns if 'EMPRESA' in c), None)
            col_sal = next((c for c in df_nom.columns if 'SALARIO' in c and '711' in c), None)
            col_tkt = next((c for c in df_nom.columns if 'TICKET' in c or 'ALIMENTACION' in c), None)
            col_imp = next((c for c in df_nom.columns if 'APARTADO' in c), None)
            
            if col_empresa and col_sal and col_tkt and col_imp:
                filas_encontradas = df_nom[df_nom[col_empresa].astype(str).str.upper().str.contains(keyword_empresa, na=False)]
                
                if not filas_encontradas.empty:
                    val_salarios_nom = filas_encontradas[col_sal].apply(clean_float).sum()
                    val_tickets_nom = filas_encontradas[col_tkt].apply(clean_float).sum()
                    val_impuesto_nom = filas_encontradas[col_imp].apply(clean_float).sum()
                    
                    log_messages.append(f"📊 Nómina: Salarios={val_salarios_nom:,.2f}, Tickets={val_tickets_nom:,.2f}")
                else:
                    log_messages.append(f"⚠️ No se encontró la empresa '{keyword_empresa}' en Nómina.")
            else:
                log_messages.append(f"⚠️ Columnas faltantes en Nómina.")
    except Exception as e:
        log_messages.append(f"⚠️ Error leyendo Nómina: {str(e)}")

    # --- 3. GENERAR ASIENTO ---
    asiento_data = df_agrupado.groupby('Centro de Costo (Padre)')['Impuesto (9%)'].sum().reset_index()
    asiento_data.rename(columns={'Centro de Costo (Padre)': 'Centro Costo', 'Impuesto (9%)': 'Débito VES'}, inplace=True)
    asiento_data['Cuenta Contable'] = '7.1.1.07.1.001'
    asiento_data['Descripción'] = 'Contribucion ley de Pensiones'
    asiento_data['Crédito VES'] = 0.0
    
    total_impuesto_contable = asiento_data['Débito VES'].sum()
    
    linea_pasivo = pd.DataFrame([{
        'Centro Costo': '00.00.000.00', 'Cuenta Contable': '2.1.3.02.3.005', 
        'Descripción': 'Contribuciones Sociales por Pagar', 'Débito VES': 0.0, 'Crédito VES': total_impuesto_contable
    }])
    
    df_asiento = pd.concat([asiento_data, linea_pasivo], ignore_index=True)
    
    if tasa_cambio > 0:
        df_asiento['Débito USD'] = (df_asiento['Débito VES'] / tasa_cambio).round(2)
        df_asiento['Crédito USD'] = (df_asiento['Crédito VES'] / tasa_cambio).round(2)
        df_asiento['Tasa'] = tasa_cambio
    else:
        df_asiento['Débito USD'] = 0; df_asiento['Crédito USD'] = 0; df_asiento['Tasa'] = 0

    # --- 4. RESUMEN VALIDACIÓN ---
    dif_salarios = base_salarios_cont - val_salarios_nom
    dif_tickets = base_tickets_cont - val_tickets_nom
    dif_impuesto = total_impuesto_contable - val_impuesto_nom
    
    estado_val = "OK" if (abs(dif_salarios) < 1.00 and abs(dif_tickets) < 1.00 and abs(dif_impuesto) < 1.00) else "DESCUADRE"
    
    if estado_val == "OK":
        log_messages.append("✅ VALIDACIÓN TOTAL: Bases e Impuestos cuadran.")
    else:
        log_messages.append(f"⚠️ DESCUADRE DETECTADO.")

    resumen_validacion = {
        'salario_cont': base_salarios_cont, 'salario_nom': val_salarios_nom, 'dif_salario': dif_salarios,
        'ticket_cont': base_tickets_cont, 'ticket_nom': val_tickets_nom, 'dif_ticket': dif_tickets,
        'total_base_cont': total_base_contable, 'total_base_nom': val_salarios_nom + val_tickets_nom,
        'dif_base_total': total_base_contable - (val_salarios_nom + val_tickets_nom),
        'imp_calc': total_impuesto_contable, 'imp_nom': val_impuesto_nom, 'dif_imp': dif_impuesto,
        'estado': estado_val
    }

    return df_agrupado, df_filtrado, df_asiento, resumen_validacion
