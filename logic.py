# logic.py

import pandas as pd
import numpy as np
import re
from itertools import combinations
from io import BytesIO
import unicodedata
import xlsxwriter
from utils import generar_reporte_retenciones

# --- Constantes de Tolerancia ---
TOLERANCIA_MAX_BS = 2.00
TOLERANCIA_MAX_USD = 0.50

# ==============================================================================
# LÓGICAS DE CONCILIACIÓN DETALLADAS
# ==============================================================================

# --- (A) Módulo: Fondos en Tránsito (BS) ---
def normalizar_referencia_fondos_en_transito(df):
    """Clasifica movimientos según palabras clave en la referencia para Fondos en Tránsito."""
    df_copy = df.copy()
    def clasificar(referencia_str):
        if pd.isna(referencia_str): return 'OTRO', 'OTRO', ''
        ref = str(referencia_str).upper().strip()
        ref_lit_norm = re.sub(r'[^A-Z0-9]', '', ref)
        if any(keyword in ref for keyword in ['DIFERENCIA EN CAMBIO', 'DIF. CAMBIO', 'DIFERENCIAL', 'DIFERENCIAS DE CAMBIO', 'DIFERENCIAS DE SALDOS', 'DIFERENCIA DE SALDO', 'DIF. SALDO']): return 'DIF_CAMBIO', 'GRUPO_DIF_CAMBIO', ref_lit_norm
        if 'AJUSTE' in ref: return 'AJUSTE_GENERAL', 'GRUPO_AJUSTE', ref_lit_norm
        if 'REINTEGRO' in ref or 'SILLACA' in ref: return 'REINTEGRO_SILLACA', 'GRUPO_SILLACA', ref_lit_norm
        if 'REMESA' in ref: return 'REMESA_GENERAL', 'GRUPO_REMESA', ref_lit_norm
        if 'NOTA DE DEBITO' in ref or 'NOTA DE CREDITO' in ref: return 'NOTA_GENERAL', 'GRUPO_NOTA', ref_lit_norm
        if 'BANCO A BANCO' in ref: return 'BANCO_A_BANCO', 'GRUPO_BANCO', ref_lit_norm
        return 'OTRO', 'OTRO', ref_lit_norm
    df_copy[['Clave_Normalizada', 'Clave_Grupo', 'Referencia_Normalizada_Literal']] = df_copy['Referencia'].apply(clasificar).apply(pd.Series)
    return df_copy

def conciliar_diferencia_cambio(df, log_messages):
    df_a_conciliar = df[(df['Clave_Grupo'] == 'GRUPO_DIF_CAMBIO') & (~df['Conciliado'])]
    total_conciliados = len(df_a_conciliar)
    if total_conciliados > 0:
        indices = df_a_conciliar.index
        df.loc[indices, 'Conciliado'] = True
        df.loc[indices, 'Grupo_Conciliado'] = 'AUTOMATICO_DIF_CAMBIO_SALDO'
        log_messages.append(f"✔️ Fase Auto: {total_conciliados} conciliados por ser 'Diferencia en Cambio/Saldo'.")
    return total_conciliados

def conciliar_ajuste_automatico(df, log_messages):
    df_a_conciliar = df[(df['Clave_Grupo'] == 'GRUPO_AJUSTE') & (~df['Conciliado'])]
    total_conciliados = len(df_a_conciliar)
    if total_conciliados > 0:
        indices = df_a_conciliar.index
        df.loc[indices, 'Conciliado'] = True
        df.loc[indices, 'Grupo_Conciliado'] = 'AUTOMATICO_AJUSTE'
        log_messages.append(f"✔️ Fase Auto: {total_conciliados} conciliados por ser 'AJUSTE'.")
    return total_conciliados

def conciliar_pares_exactos_cero(df, clave_grupo, fase_name, log_messages):
    TOLERANCIA_CERO = 0.0
    df_pendientes = df[(df['Clave_Grupo'] == clave_grupo) & (~df['Conciliado'])].copy()
    if df_pendientes.empty: return 0
    log_messages.append(f"\n--- {fase_name} ---")
    grupos = df_pendientes.groupby('Referencia_Normalizada_Literal')
    total_conciliados = 0
    for ref_norm, grupo in grupos:
        if len(grupo) < 2: continue
        debitos_indices = grupo[grupo['Monto_BS'] > 0].index
        creditos_indices = grupo[grupo['Monto_BS'] < 0].index
        debitos_usados = set()
        creditos_usados = set()
        for idx_d in debitos_indices:
            if idx_d in debitos_usados: continue
            monto_d = df.loc[idx_d, 'Monto_BS']
            for idx_c in creditos_indices:
                if idx_c in creditos_usados: continue
                monto_c = df.loc[idx_c, 'Monto_BS']
                if abs(monto_d + monto_c) <= TOLERANCIA_CERO:
                    asiento_d, asiento_c = df.loc[idx_d, 'Asiento'], df.loc[idx_c, 'Asiento']
                    df.loc[[idx_d, idx_c], 'Conciliado'] = True
                    df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_REF_EXACTO_{ref_norm}_{asiento_c}'
                    df.loc[idx_c, 'Grupo_Conciliado'] = f'PAR_REF_EXACTO_{ref_norm}_{asiento_d}'
                    total_conciliados += 2
                    debitos_usados.add(idx_d)
                    creditos_usados.add(idx_c)
                    break 
    if total_conciliados > 0: log_messages.append(f"✔️ {fase_name}: {total_conciliados} movimientos conciliados.")
    return total_conciliados

def conciliar_pares_exactos_por_referencia(df, clave_grupo, fase_name, log_messages):
    df_pendientes = df[(df['Clave_Grupo'] == clave_grupo) & (~df['Conciliado'])].copy()
    if df_pendientes.empty: return 0
    log_messages.append(f"\n--- {fase_name} ---")
    grupos = df_pendientes.groupby('Referencia_Normalizada_Literal')
    total_conciliados = 0
    for ref_norm, grupo in grupos:
        if len(grupo) < 2: continue
        debitos_indices = grupo[grupo['Monto_BS'] > 0].index.tolist()
        creditos_indices = grupo[grupo['Monto_BS'] < 0].index.tolist()
        debitos_usados = set()
        creditos_usados = set()
        for idx_d in debitos_indices:
            if idx_d in debitos_usados: continue
            monto_d = df.loc[idx_d, 'Monto_BS']
            mejor_match_idx, mejor_match_diff = None, TOLERANCIA_MAX_BS + 1
            for idx_c in creditos_indices:
                if idx_c in creditos_usados: continue
                diferencia = abs(monto_d + df.loc[idx_c, 'Monto_BS'])
                if diferencia < mejor_match_diff:
                    mejor_match_diff, mejor_match_idx = diferencia, idx_c
            if mejor_match_idx is not None and mejor_match_diff <= TOLERANCIA_MAX_BS:
                asiento_d, asiento_c = df.loc[idx_d, 'Asiento'], df.loc[mejor_match_idx, 'Asiento']
                df.loc[[idx_d, mejor_match_idx], 'Conciliado'] = True
                df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_REF_{ref_norm}_{asiento_c}'
                df.loc[mejor_match_idx, 'Grupo_Conciliado'] = f'PAR_REF_{ref_norm}_{asiento_d}'
                total_conciliados += 2
                debitos_usados.add(idx_d)
                creditos_usados.add(mejor_match_idx)
    if total_conciliados > 0: log_messages.append(f"✔️ {fase_name}: {total_conciliados} movimientos conciliados.")
    return total_conciliados

def cruzar_pares_simples(df, clave_normalizada, fase_name, log_messages):
    df_pendientes = df[~df['Conciliado']].copy()
    df_pendientes['Monto_BS_Abs_Redondeado'] = (df_pendientes['Monto_BS'].abs().round(0))
    df_a_cruzar = df_pendientes[df_pendientes['Clave_Normalizada'] == clave_normalizada]
    if df_a_cruzar.empty: return 0
    log_messages.append(f"\n--- {fase_name} ---")
    grupos = df_a_cruzar.groupby('Monto_BS_Abs_Redondeado')
    total_conciliados = 0
    for _, grupo in grupos:
        debitos_indices = grupo[grupo['Monto_BS'] > 0].index.tolist()
        creditos_indices = grupo[grupo['Monto_BS'] < 0].index.tolist()
        debitos_usados = set()
        creditos_usados = set()
        for idx_d in debitos_indices:
            if idx_d in debitos_usados: continue
            monto_d = df.loc[idx_d, 'Monto_BS']
            mejor_match_idx, mejor_match_diff = None, TOLERANCIA_MAX_BS + 1
            for idx_c in creditos_indices:
                if idx_c in creditos_usados: continue
                diferencia = abs(monto_d + df.loc[idx_c, 'Monto_BS'])
                if diferencia < mejor_match_diff:
                    mejor_match_diff, mejor_match_idx = diferencia, idx_c
            if mejor_match_idx is not None and mejor_match_diff <= TOLERANCIA_MAX_BS:
                asiento_d, asiento_c = df.loc[idx_d, 'Asiento'], df.loc[mejor_match_idx, 'Asiento']
                df.loc[[idx_d, mejor_match_idx], 'Conciliado'] = True
                df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_BS_{asiento_c}'
                df.loc[mejor_match_idx, 'Grupo_Conciliado'] = f'PAR_BS_{asiento_d}'
                total_conciliados += 2
                debitos_usados.add(idx_d)
                creditos_usados.add(mejor_match_idx)
    if 'Monto_BS_Abs_Redondeado' in df.columns: df.drop(columns=['Monto_BS_Abs_Redondeado'], inplace=True, errors='ignore')
    if total_conciliados > 0: log_messages.append(f"✔️ {fase_name}: {total_conciliados} movimientos conciliados.")
    return total_conciliados

def cruzar_grupos_por_criterio(df, clave_normalizada, agrupacion_col, grupo_prefix, fase_name, log_messages):
    df_pendientes = df[(df['Clave_Normalizada'] == clave_normalizada) & (~df['Conciliado'])].copy()
    if df_pendientes.empty: return 0
    log_messages.append(f"\n--- {fase_name} ---")
    indices_conciliados = set()
    if agrupacion_col == 'Fecha': grupos = df_pendientes.groupby(df_pendientes['Fecha'].dt.date.fillna('NaT'))
    else: grupos = df_pendientes.groupby(agrupacion_col)
    for criterio, grupo in grupos:
        if len(grupo) > 1 and abs(grupo['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            grupo_id = f"GRUPO_{grupo_prefix}_{criterio}"
            indices_a_conciliar = grupo.index
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = grupo_id
            indices_conciliados.update(indices_a_conciliar)
    total_conciliados = len(indices_conciliados)
    if total_conciliados > 0: log_messages.append(f"✔️ {fase_name}: {total_conciliados} movimientos conciliados.")
    return total_conciliados

def conciliar_lote_por_grupo(df, clave_grupo, fase_name, log_messages):
    log_messages.append(f"\n--- {fase_name} ---")
    df_pendientes = df[(~df['Conciliado']) & (df['Clave_Grupo'] == clave_grupo)].copy()
    if df_pendientes.empty or len(df_pendientes) < 2: return 0
    if abs(df_pendientes['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
        fecha_max = df_pendientes['Fecha'].max().strftime('%Y-%m-%d')
        grupo_id = f"LOTE_{clave_grupo.replace('GRUPO_', '')}_{fecha_max}"
        indices_a_conciliar = df_pendientes.index
        df.loc[indices_a_conciliar, 'Conciliado'] = True
        df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = grupo_id
        total_conciliados = len(indices_a_conciliar)
        log_messages.append(f"✔️ {fase_name}: {total_conciliados} movimientos conciliados como lote.")
        return total_conciliados
    return 0

def conciliar_grupos_globales_por_referencia(df, log_messages):
    log_messages.append(f"\n--- FASE GLOBAL N-a-N (Cruce por Referencia Literal) ---")
    df_pendientes = df[~df['Conciliado']].copy()
    df_pendientes = df_pendientes[df_pendientes['Referencia_Normalizada_Literal'].notna() & (df_pendientes['Referencia_Normalizada_Literal'] != '') & (df_pendientes['Referencia_Normalizada_Literal'] != 'OTRO')]
    if df_pendientes.empty: return 0
    grupos = df_pendientes.groupby('Referencia_Normalizada_Literal')
    total_conciliados = 0
    for ref_norm, grupo in grupos:
        if len(grupo) > 1 and abs(grupo['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            grupo_id = f"GRUPO_REF_GLOBAL_{ref_norm}"
            indices_a_conciliar = grupo.index
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = grupo_id
            total_conciliados += len(indices_a_conciliar)
    if total_conciliados > 0: log_messages.append(f"✔️ Fase Global N-a-N: {total_conciliados} movimientos conciliados.")
    return total_conciliados

def conciliar_pares_globales_remanentes(df, log_messages):
    log_messages.append(f"\n--- FASE GLOBAL 1-a-1 (Cruce de pares remanentes) ---")
    df_pendientes = df[~df['Conciliado']].copy()
    if df_pendientes.empty or len(df_pendientes) < 2: return 0
    debitos = df_pendientes[df_pendientes['Monto_BS'] > 0].index.tolist()
    creditos = df_pendientes[df_pendientes['Monto_BS'] < 0].index.tolist()
    total_conciliados = 0
    creditos_usados = set()
    for idx_d in debitos:
        monto_d = df.loc[idx_d, 'Monto_BS']
        mejor_match_idx, mejor_match_diff = None, TOLERANCIA_MAX_BS + 1
        for idx_c in creditos:
            if idx_c in creditos_usados: continue
            diferencia = abs(monto_d + df.loc[idx_c, 'Monto_BS'])
            if diferencia < mejor_match_diff:
                mejor_match_diff, mejor_match_idx = diferencia, idx_c
        if mejor_match_idx is not None and mejor_match_diff <= TOLERANCIA_MAX_BS:
            asiento_d, asiento_c = df.loc[idx_d, 'Asiento'], df.loc[mejor_match_idx, 'Asiento']
            df.loc[[idx_d, mejor_match_idx], 'Conciliado'] = True
            df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_GLOBAL_{asiento_c}'
            df.loc[mejor_match_idx, 'Grupo_Conciliado'] = f'PAR_GLOBAL_{asiento_d}'
            creditos_usados.add(mejor_match_idx)
            total_conciliados += 2
    if total_conciliados > 0: log_messages.append(f"✔️ Fase Global 1-a-1: {total_conciliados} movimientos conciliados.")
    return total_conciliados

# --- (B) Módulo: Fondos por Depositar (USD) ---
def normalizar_referencia_fondos_usd(df):
    df_copy = df.copy()
    def clasificar_usd(ref_str):
        if pd.isna(ref_str): return 'OTRO', 'OTRO', 'OTRO'
        ref = str(ref_str).upper().strip()
        if 'TRASPASO' in ref:
            if any(keyword in ref for keyword in ['MERCANTIL', 'ZINLI', 'BEVAL', 'SILLACA']):
                ref_lit_norm = 'TRASPASO_GENERICO_FONDOS'
                return 'TRASPASO', 'GRUPO_TRASPASO', ref_lit_norm
        ref_lit_norm = re.sub(r'[^\w]', '', ref)
        if 'DIFERENCIA' in ref and 'CAMBIO' in ref: return 'DIF_CAMBIO', 'GRUPO_DIF_CAMBIO', ref_lit_norm
        if 'BANCO A BANCO' in ref: return 'BANCO_A_BANCO', 'GRUPO_BANCO', 'BANCO_A_BANCO'
        if 'BANCARIZACION' in ref: return 'BANCARIZACION', 'GRUPO_BANCARIZACION', ref_lit_norm
        if 'REINTEGRO' in ref: return 'REINTEGRO', 'GRUPO_REINTEGRO', ref_lit_norm
        if 'REMESA' in ref: return 'REMESA', 'GRUPO_REMESA', ref_lit_norm
        if 'GASTOS POR TARJETA' in ref: return 'TARJETA_GASTOS', 'GRUPO_TARJETA', ref_lit_norm
        if 'NOTA DE DEBITO' in ref: return 'NOTA_DEBITO', 'GRUPO_NOTA', 'NOTA_DEBITO'
        if 'NOTA DE CREDITO' in ref: return 'NOTA_CREDITO', 'GRUPO_NOTA', 'NOTA_CREDITO'
        return 'OTRO', 'OTRO', ref_lit_norm

    df_copy[['Clave_Normalizada', 'Clave_Grupo', 'Referencia_Normalizada_Literal']] = df_copy['Referencia'].apply(clasificar_usd).apply(pd.Series)
    return df_copy
    
def conciliar_automaticos_usd(df, log_messages):
    total = 0
    for grupo, etiqueta in [('GRUPO_DIF_CAMBIO', 'AUTOMATICO_DIF_CAMBIO'), ('GRUPO_AJUSTE', 'AUTOMATICO_AJUSTE')]:
        indices = df.loc[(df['Clave_Grupo'] == grupo) & (~df['Conciliado'])].index
        if not indices.empty:
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, etiqueta]
            log_messages.append(f"✔️ Fase Auto (USD): {len(indices)} conciliados por ser '{etiqueta}'.")
            total += len(indices)
    return total

def conciliar_grupos_por_referencia_usd(df, log_messages):
    log_messages.append("\n--- FASE GRUPOS POR REFERENCIA EXACTA (USD) ---")
    total_conciliados = 0
    df_pendientes = df.loc[~df['Conciliado']]
    grupos = df_pendientes.groupby('Referencia_Normalizada_Literal')
    for ref_norm, grupo in grupos:
        if len(grupo) > 1 and abs(grupo['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
            indices = grupo.index
            grupo_id = f"GRUPO_REF_{ref_norm}"
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, grupo_id]
            total_conciliados += len(indices)
    if total_conciliados > 0:
        log_messages.append(f"✔️ Fase Grupos por Ref. Exacta: {total_conciliados} movimientos conciliados.")
    return total_conciliados

def conciliar_pares_globales_exactos_usd(df, log_messages):
    log_messages.append("\n--- FASE PARES GLOBALES EXACTOS (USD) ---")
    total_conciliados = 0
    df_pendientes = df.loc[~df['Conciliado']].copy()
    
    df_pendientes['Monto_Abs'] = df_pendientes['Monto_USD'].abs()
    
    grupos_por_monto = df_pendientes.groupby('Monto_Abs')
    
    for monto, grupo in grupos_por_monto:
        if len(grupo) < 2:
            continue
            
        debitos = grupo[grupo['Monto_USD'] > 0].index.to_list()
        creditos = grupo[grupo['Monto_USD'] < 0].index.to_list()
        
        pares_a_conciliar = min(len(debitos), len(creditos))
        
        if pares_a_conciliar > 0:
            for i in range(pares_a_conciliar):
                idx_d = debitos[i]
                idx_c = creditos[i]
                
                if abs(df.loc[idx_d, 'Monto_USD'] + df.loc[idx_c, 'Monto_USD']) <= 0.01:
                    asiento_d = df.loc[idx_d, 'Asiento']
                    asiento_c = df.loc[idx_c, 'Asiento']
                    df.loc[[idx_d, idx_c], 'Conciliado'] = True
                    df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_EXACTO_{asiento_c}'
                    df.loc[idx_c, 'Grupo_Conciliado'] = f'PAR_EXACTO_{asiento_d}'
                    total_conciliados += 2

    if total_conciliados > 0:
        log_messages.append(f"✔️ Fase Pares Exactos: {total_conciliados} movimientos conciliados.")
    
    return total_conciliados

def conciliar_pares_por_referencia_usd(df, clave_grupo, fase_name, log_messages):
    df_pendientes = df.loc[(df['Clave_Grupo'] == clave_grupo) & (~df['Conciliado'])].copy()
    if df_pendientes.empty: return 0
    log_messages.append(f"\n--- {fase_name} (USD) ---")
    grupos, total_conciliados = df_pendientes.groupby('Referencia_Normalizada_Literal'), 0
    for ref_norm, grupo in grupos:
        if len(grupo) < 2: continue
        debitos_idx, creditos_idx = grupo[grupo['Monto_USD'] > 0].index.tolist(), grupo[grupo['Monto_USD'] < 0].index.tolist()
        debitos_usados, creditos_usados = set(), set()
        for idx_d in debitos_idx:
            if idx_d in debitos_usados: continue
            monto_d, mejor_match, mejor_diff = df.loc[idx_d, 'Monto_USD'], None, TOLERANCIA_MAX_USD + 1
            for idx_c in creditos_idx:
                if idx_c in creditos_usados: continue
                diff = abs(monto_d + df.loc[idx_c, 'Monto_USD'])
                if diff < mejor_diff: mejor_diff, mejor_match = diff, idx_c
            if mejor_match is not None and mejor_diff <= TOLERANCIA_MAX_USD:
                asientos = (df.loc[idx_d, 'Asiento'], df.loc[mejor_match, 'Asiento'])
                df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_REF_{ref_norm[:10]}_{asientos[1]}'
                df.loc[mejor_match, 'Grupo_Conciliado'] = f'PAR_REF_{ref_norm[:10]}_{asientos[0]}'
                df.loc[[idx_d, mejor_match], 'Conciliado'] = True
                total_conciliados += 2
                debitos_usados.add(idx_d); creditos_usados.add(mejor_match)
    if total_conciliados > 0: log_messages.append(f"✔️ {fase_name}: {total_conciliados} movimientos conciliados.")
    return total_conciliados
    
def conciliar_lote_por_grupo_usd(df, clave_grupo, fase_name, log_messages):
    df_pendientes = df.loc[(~df['Conciliado']) & (df['Clave_Grupo'] == clave_grupo)].copy()
    if len(df_pendientes) > 1 and abs(df_pendientes['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
        grupo_id = f"LOTE_{clave_grupo.replace('GRUPO_', '')}_{df_pendientes['Fecha'].max().strftime('%Y%m%d')}"
        df.loc[df_pendientes.index, ['Conciliado', 'Grupo_Conciliado']] = [True, grupo_id]
        log_messages.append(f"✔️ {fase_name}: {len(df_pendientes.index)} movimientos conciliados como lote.")
        return len(df_pendientes.index)
    return 0

def conciliar_pares_banco_a_banco_usd(df, log_messages):
    log_messages.append("\n--- FASE PARES BANCO A BANCO (USD) ---")
    total_conciliados = 0
    df_pendientes = df.loc[(~df['Conciliado']) & (df['Clave_Grupo'] == 'GRUPO_BANCO')].copy()
    
    if df_pendientes.empty:
        return 0

    df_pendientes['Monto_Abs'] = df_pendientes['Monto_USD'].abs()
    
    grupos_por_monto = df_pendientes.groupby('Monto_Abs')
    
    for monto, grupo in grupos_por_monto:
        if len(grupo) < 2:
            continue
            
        debitos = grupo[grupo['Monto_USD'] > 0].index.to_list()
        creditos = grupo[grupo['Monto_USD'] < 0].index.to_list()
        
        pares_a_conciliar = min(len(debitos), len(creditos))
        
        if pares_a_conciliar > 0:
            for i in range(pares_a_conciliar):
                idx_d = debitos[i]
                idx_c = creditos[i]
                asiento_d = df.loc[idx_d, 'Asiento']
                asiento_c = df.loc[idx_c, 'Asiento']
                
                df.loc[[idx_d, idx_c], 'Conciliado'] = True
                df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_BANCO_{asiento_c}'
                df.loc[idx_c, 'Grupo_Conciliado'] = f'PAR_BANCO_{asiento_d}'
                total_conciliados += 2

    if total_conciliados > 0:
        log_messages.append(f"✔️ Fase Pares Banco a Banco: {total_conciliados} movimientos conciliados.")
    
    return total_conciliados

def conciliar_grupos_complejos_usd(df, log_messages, progress_bar=None):
    log_messages.append("\n--- FASE GRUPOS COMPLEJOS OPTIMIZADA (USD) ---")
    
    pendientes = df.loc[~df['Conciliado']]
    LIMITE_MOVIMIENTOS = 500
    if len(pendientes) > LIMITE_MOVIMIENTOS:
        log_messages.append(f"ℹ️ Se omitió la fase de grupos complejos por haber demasiados movimientos pendientes (> {LIMITE_MOVIMIENTOS}).")
        return 0

    debitos = pendientes[pendientes['Monto_USD'] > 0]
    creditos = pendientes[pendientes['Monto_USD'] < 0]

    if debitos.empty or creditos.empty or len(debitos) < 1 or len(creditos) < 1:
        return 0

    total_conciliados_fase = 0
    indices_usados = set()

    # --- CASO 1: 1 Débito vs 2 Créditos ---
    log_messages.append("Construyendo mapa de sumas de créditos...")
    sumas_pares_creditos = {}
    if len(creditos) >= 2:
        for i in range(len(creditos)):
            for j in range(i + 1, len(creditos)):
                idx1, idx2 = creditos.index[i], creditos.index[j]
                suma = creditos.loc[idx1, 'Monto_USD'] + creditos.loc[idx2, 'Monto_USD']
                if suma not in sumas_pares_creditos:
                    sumas_pares_creditos[suma] = []
                sumas_pares_creditos[suma].append((idx1, idx2))
    
    log_messages.append("Analizando 1 Débito vs 2 Créditos...")
    debitos_ordenados = debitos.sort_values('Monto_USD', ascending=False).index
    
    for idx_d in debitos_ordenados:
        if idx_d in indices_usados: continue
        monto_d = df.loc[idx_d, 'Monto_USD']
        target_sum_c = -monto_d
        
        for suma_c, pares in sumas_pares_creditos.items():
            if abs(target_sum_c - suma_c) <= TOLERANCIA_MAX_USD:
                for idx_c1, idx_c2 in pares:
                    if not indices_usados.isdisjoint([idx_d, idx_c1, idx_c2]):
                        continue
                    
                    indices_conciliados = [idx_d, idx_c1, idx_c2]
                    asiento_d = df.loc[idx_d, 'Asiento']
                    df.loc[indices_conciliados, ['Conciliado', 'Grupo_Conciliado']] = [True, f'GRUPO_1v2_{asiento_d}']
                    indices_usados.update(indices_conciliados)
                    total_conciliados_fase += 3
                    goto_next_debit = True
                    break
                if 'goto_next_debit' in locals() and goto_next_debit: break
        if 'goto_next_debit' in locals() and goto_next_debit:
            del goto_next_debit
            continue
    
    if progress_bar: progress_bar.progress(0.5)

    # --- CASO 2: 2 Débitos vs 1 Crédito ---
    debitos_disponibles = debitos.drop(indices_usados, errors='ignore')
    creditos_disponibles = creditos.drop(indices_usados, errors='ignore')
    
    log_messages.append("Construyendo mapa de sumas de débitos...")
    sumas_pares_debitos = {}
    if len(debitos_disponibles) >= 2:
        for i in range(len(debitos_disponibles)):
            for j in range(i + 1, len(debitos_disponibles)):
                idx1, idx2 = debitos_disponibles.index[i], debitos_disponibles.index[j]
                suma = debitos_disponibles.loc[idx1, 'Monto_USD'] + debitos_disponibles.loc[idx2, 'Monto_USD']
                if suma not in sumas_pares_debitos:
                    sumas_pares_debitos[suma] = []
                sumas_pares_debitos[suma].append((idx1, idx2))
    
    log_messages.append("Analizando 2 Débitos vs 1 Crédito...")
    creditos_ordenados = creditos_disponibles.sort_values('Monto_USD').index
    
    for idx_c in creditos_ordenados:
        if idx_c in indices_usados: continue
        monto_c = df.loc[idx_c, 'Monto_USD']
        target_sum_d = -monto_c
        
        for suma_d, pares in sumas_pares_debitos.items():
            if abs(target_sum_d - suma_d) <= TOLERANCIA_MAX_USD:
                for idx_d1, idx_d2 in pares:
                    if not indices_usados.isdisjoint([idx_c, idx_d1, idx_d2]):
                        continue
                        
                    indices_conciliados = [idx_c, idx_d1, idx_d2]
                    asiento_c = df.loc[idx_c, 'Asiento']
                    df.loc[indices_conciliados, ['Conciliado', 'Grupo_Conciliado']] = [True, f'GRUPO_2v1_{asiento_c}']
                    indices_usados.update(indices_conciliados)
                    total_conciliados_fase += 3
                    goto_next_credit = True
                    break
                if 'goto_next_credit' in locals() and goto_next_credit: break
        if 'goto_next_credit' in locals() and goto_next_credit:
            del goto_next_credit
            continue
            
    if total_conciliados_fase > 0:
        log_messages.append(f"✔️ {total_conciliados_fase} movimientos conciliados en grupos complejos (optimizados).")
    
    return total_conciliados_fase
    
def conciliar_pares_globales_remanentes_usd(df, log_messages):
    log_messages.append("\n--- FASE GLOBAL 1-a-1 (USD) ---")
    
    pendientes = df.loc[~df['Conciliado']].copy()
    if len(pendientes) < 2:
        return 0

    debitos = pendientes[pendientes['Monto_USD'] > 0].copy()
    creditos = pendientes[pendientes['Monto_USD'] < 0].copy()

    if debitos.empty or creditos.empty:
        return 0

    debitos.reset_index(inplace=True)
    creditos.reset_index(inplace=True)

    debitos['join_key'] = 1
    creditos['join_key'] = 1

    pares_potenciales = pd.merge(debitos, creditos, on='join_key', suffixes=('_d', '_c'))
    
    pares_potenciales['diferencia'] = abs(pares_potenciales['Monto_USD_d'] + pares_potenciales['Monto_USD_c'])
    
    pares_validos = pares_potenciales[pares_potenciales['diferencia'] <= TOLERANCIA_MAX_USD].copy()
    
    pares_validos.sort_values(by='diferencia', inplace=True)
    
    pares_finales = pares_validos.drop_duplicates(subset=['index_d'], keep='first')
    pares_finales = pares_finales.drop_duplicates(subset=['index_c'], keep='first')
    
    total_conciliados = 0
    if not pares_finales.empty:
        indices_d = pares_finales['index_d'].tolist()
        indices_c = pares_finales['index_c'].tolist()
        
        df.loc[indices_d + indices_c, 'Conciliado'] = True
        
        for _, row in pares_finales.iterrows():
            idx_d, asiento_d = row['index_d'], row['Asiento_d']
            idx_c, asiento_c = row['index_c'], row['Asiento_c']
            df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_GLOBAL_{asiento_c}'
            df.loc[idx_c, 'Grupo_Conciliado'] = f'PAR_GLOBAL_{asiento_d}'
        
        total_conciliados = len(indices_d) + len(indices_c)

    if total_conciliados > 0:
        log_messages.append(f"✔️ Fase Global Optimizada: {total_conciliados} movimientos conciliados.")
    
    return total_conciliados
    
def conciliar_gran_total_final_usd(df, log_messages):
    log_messages.append("\n--- FASE FINAL (USD) ---")
    df_pendientes = df.loc[~df['Conciliado']]
    if not df_pendientes.empty and abs(df_pendientes['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
        df.loc[df_pendientes.index, ['Conciliado', 'Grupo_Conciliado']] = [True, "LOTE_GRAN_TOTAL_FINAL"]
        log_messages.append(f"✔️ Fase Final: ¡Éxito! {len(df_pendientes.index)} remanentes conciliados.")
        return len(df_pendientes.index)
    return 0

# --- (C) Módulo: Devoluciones a Proveedores ---
def normalizar_datos_proveedores(df, log_messages):
    df_copy = df.copy()
    
    nit_col_name = None
    for col in df_copy.columns:
        if str(col).strip().upper() in ['NIT', 'RIF']:
            nit_col_name = col
            break
            
    if nit_col_name:
        log_messages.append(f"✔️ Se encontró la columna de identificador fiscal ('{nit_col_name}') y se usará como clave principal.")
        df_copy['Clave_Proveedor'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró la columna 'NIT' o 'RIF'. Se recurrirá a usar 'Nombre del Proveedor', lo cual es menos preciso.")
        df_copy['Clave_Proveedor'] = df_copy['Nombre del Proveedor'].astype(str).str.strip().str.upper()

    def extraer_clave_comp(row):
        if row['Monto_USD'] > 0:
            return str(row['Fuente']).strip().upper()
        elif row['Monto_USD'] < 0:
            match = re.search(r'(COMP-\d+)', str(row['Referencia']).upper())
            if match:
                return match.group(1)
        return np.nan
        
    df_copy['Clave_Comp'] = df_copy.apply(extraer_clave_comp, axis=1)
    return df_copy

def run_conciliation_devoluciones_proveedores(df, log_messages):
    log_messages.append("\n--- INICIANDO LÓGICA DE DEVOLUCIONES A PROVEEDORES (USD) ---")
    
    df = normalizar_datos_proveedores(df, log_messages) 
    
    total_conciliados = 0
    df_procesable = df.loc[(~df['Conciliado']) & (df['Clave_Proveedor'].notna()) & (df['Clave_Comp'].notna())]
    
    grupos = df_procesable.groupby(['Clave_Proveedor', 'Clave_Comp'])
    
    log_messages.append(f"ℹ️ Se encontraron {len(grupos)} grupos de Proveedor/COMP para analizar.")
    for (proveedor_clave, comp), grupo in grupos:
        if abs(round(grupo['Monto_USD'].sum(), 2)) <= TOLERANCIA_MAX_USD:
            indices = grupo.index
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, f"PROV_{proveedor_clave}_{comp}"]
            total_conciliados += len(indices)

    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación por Proveedor/COMP: {total_conciliados} movimientos conciliados.")
    else:
        log_messages.append("ℹ️ No se encontraron conciliaciones automáticas por Proveedor/COMP.")
        
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# --- (D) Módulo: Cuentas de Viajes ---
def normalizar_referencia_viajes(df, log_messages):
    log_messages.append("✔️ Fase de Normalización: Clasificando movimientos por tipo (Impuestos/Viáticos).")
    
    def clasificar_tipo(referencia_str):
        if pd.isna(referencia_str):
            return 'OTRO'
        ref = str(referencia_str).upper().strip()
        if 'TIMBRES' in ref or 'FISCAL' in ref:
            return 'IMPUESTOS'
        if 'VIAJE' in ref or 'VIATICOS' in ref:
            return 'VIATICOS'
        return 'OTRO'

    df['Tipo'] = df['Referencia'].apply(clasificar_tipo)
    
    nit_col_name = next((col for col in df.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df['NIT_Normalizado'] = df[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró la columna 'NIT' o 'RIF'. La conciliación puede no ser precisa.")
        df['NIT_Normalizado'] = 'SIN_NIT'
        
    return df

def conciliar_pares_exactos_por_nit_viajes(df, log_messages):
    log_messages.append("\n--- FASE 1: Búsqueda de Pares Exactos por NIT ---")
    total_conciliados = 0
    
    df_pendientes = df.loc[~df['Conciliado']].copy()
    df_pendientes['Monto_Abs'] = df_pendientes['Monto_BS'].abs()
    
    grupos = df_pendientes.groupby(['NIT_Normalizado', 'Monto_Abs'])
    
    for (nit, monto), grupo in grupos:
        if len(grupo) < 2 or nit == 'SIN_NIT':
            continue
            
        debitos = grupo[grupo['Monto_BS'] > 0].index.to_list()
        creditos = grupo[grupo['Monto_BS'] < 0].index.to_list()
        
        pares_a_conciliar = min(len(debitos), len(creditos))
        
        if pares_a_conciliar > 0:
            for i in range(pares_a_conciliar):
                idx_d, idx_c = debitos[i], creditos[i]
                
                if abs(df.loc[idx_d, 'Monto_BS'] + df.loc[idx_c, 'Monto_BS']) <= 0.01:
                    asiento_d, asiento_c = df.loc[idx_d, 'Asiento'], df.loc[idx_c, 'Asiento']
                    df.loc[[idx_d, idx_c], 'Conciliado'] = True
                    df.loc[idx_d, 'Grupo_Conciliado'] = f'PAR_NIT_{nit}_{asiento_c}'
                    df.loc[idx_c, 'Grupo_Conciliado'] = f'PAR_NIT_{nit}_{asiento_d}'
                    total_conciliados += 2

    if total_conciliados > 0:
        log_messages.append(f"✔️ Fase 1: {total_conciliados} movimientos conciliados como pares exactos por NIT.")
    return total_conciliados

def conciliar_grupos_por_nit_viajes(df, log_messages):
    log_messages.append("\n--- FASE 2: Búsqueda de Grupos por NIT ---")
    total_conciliados_fase = 0
    
    df_pendientes = df.loc[~df['Conciliado']]
    grupos_por_nit = df_pendientes.groupby('NIT_Normalizado')
    
    for nit, grupo in grupos_por_nit:
        if nit == 'SIN_NIT' or len(grupo) < 2:
            continue
            
        if abs(grupo['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            indices = grupo.index
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, f'GRUPO_TOTAL_NIT_{nit}']
            total_conciliados_fase += len(indices)
            log_messages.append(f"✔️ Conciliado grupo completo para NIT {nit} ({len(indices)} movimientos).")
            continue

        LIMITE_COMBINACION = 10
        movimientos_grupo = grupo.index.to_list()
        
        if len(movimientos_grupo) > LIMITE_COMBINACION:
            log_messages.append(f"ℹ️ Se omitió la búsqueda de sub-grupos para NIT {nit} por exceso de movimientos (> {LIMITE_COMBINACION}).")
            continue

        indices_usados_en_grupo = set()
        for i in range(2, len(movimientos_grupo) + 1):
            for combo_indices in combinations(movimientos_grupo, i):
                if not indices_usados_en_grupo.isdisjoint(combo_indices):
                    continue
                
                suma_combo = df.loc[list(combo_indices), 'Monto_BS'].sum()
                if abs(suma_combo) <= TOLERANCIA_MAX_BS:
                    grupo_id = f"GRUPO_PARCIAL_NIT_{nit}_{total_conciliados_fase}"
                    df.loc[list(combo_indices), ['Conciliado', 'Grupo_Conciliado']] = [True, grupo_id]
                    indices_usados_en_grupo.update(combo_indices)
                    total_conciliados_fase += len(combo_indices)

    if total_conciliados_fase > 0:
        log_messages.append(f"✔️ Fase 2: {total_conciliados_fase} movimientos conciliados en grupos por NIT.")
    return total_conciliados_fase

# --- (E) Módulo: Deudores Empleados - Otros (ME) ---
def normalizar_datos_deudores_empleados(df, log_messages):
    """Normaliza el NIT del empleado para usarlo como clave de agrupación."""
    df_copy = df.copy()
    
    # Busca la columna NIT/RIF dinámicamente
    nit_col_name = next((col for col in df_copy.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
            
    if nit_col_name:
        log_messages.append(f"✔️ Normalización: Usando columna '{nit_col_name}' como identificador del empleado.")
        # Limpia el NIT: quita caracteres no alfanuméricos y convierte a mayúsculas
        df_copy['Clave_Empleado'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró columna 'NIT' o 'RIF'. La conciliación puede ser imprecisa.")
        df_copy['Clave_Empleado'] = 'SIN_NIT'
        
    return df_copy

def conciliar_grupos_por_empleado(df, log_messages):
    """Concilia movimientos por empleado si la suma total en USD es cero."""
    log_messages.append("\n--- FASE 1: Conciliación de saldos totales por empleado (USD) ---")
    
    total_conciliados = 0
    
    # Agrupa por la clave de empleado normalizada, excluyendo los ya conciliados
    df_pendientes = df.loc[~df['Conciliado']]
    grupos_por_empleado = df_pendientes.groupby('Clave_Empleado')
    
    log_messages.append(f"ℹ️ Se analizarán los saldos de {len(grupos_por_empleado)} empleados.")
    
    for clave_empleado, grupo in grupos_por_empleado:
        # Omitir si la clave no es válida
        if clave_empleado == 'SIN_NIT' or pd.isna(clave_empleado) or not clave_empleado:
            continue
            
        # Suma de los movimientos en Dólares para el empleado
        suma_usd = grupo['Monto_USD'].sum()
        
        # Comprueba si la suma está dentro de la tolerancia permitida
        if abs(suma_usd) <= TOLERANCIA_MAX_USD:
            indices_a_conciliar = grupo.index
            
            # Marcar como conciliado y asignar un grupo
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"SALDO_CERO_EMP_{clave_empleado}"
            
            num_movs = len(indices_a_conciliar)
            total_conciliados += num_movs
            
            nombre_empleado = grupo['Descripción Nit'].iloc[0] if not grupo.empty else clave_empleado
            log_messages.append(f"✔️ Empleado '{nombre_empleado}' ({clave_empleado}) conciliado. Suma: ${suma_usd:.2f} ({num_movs} movimientos).")

    if total_conciliados > 0:
        log_messages.append(f"✔️ Fase 1: {total_conciliados} movimientos conciliados por saldo cero por empleado.")
    else:
        log_messages.append("ℹ️ Fase 1: No se encontraron empleados con saldo cero para conciliar automáticamente.")
        
    return total_conciliados

# --- (F) Módulo: Fondos por Depositar – Cobros Viajeros – ME ---

def normalizar_datos_cobros_viajeros(df, log_messages):
    """
    Función de normalización que ahora asegura que la columna 'Asiento'
    sea tratada siempre como texto para evitar errores de tipo.
    """
    df_copy = df.copy()
    df_copy['Asiento'] = df_copy['Asiento'].astype(str, errors='ignore').fillna('')
    
    # Normalizar NIT
    nit_col_name = next((col for col in df_copy.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df_copy['NIT_Normalizado'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró columna 'NIT' o 'RIF'.")
        df_copy['NIT_Normalizado'] = 'SIN_NIT'

    def extraer_solo_numeros(texto):
        if pd.isna(texto):
            return ''
        return re.sub(r'\D', '', str(texto))

    # Columnas normalizadas para las claves de cruce
    df_copy['Referencia_Norm_Num'] = df_copy['Referencia'].apply(extraer_solo_numeros)
    df_copy['Fuente_Norm_Num'] = df_copy['Fuente'].apply(extraer_solo_numeros)
    
    # Columna para identificar reversos
    df_copy['Es_Reverso'] = df_copy['Referencia'].str.contains('REVERSO', case=False, na=False)
    
    return df_copy


def run_conciliation_cobros_viajeros(df, log_messages, progress_bar=None):
    """
    Versión final y definitiva que maneja coincidencias parciales de claves ('endswith')
    para conciliar correctamente todos los tipos de reversos y transacciones.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE COBROS VIAJEROS (V9 - COINCIDENCIA PARCIAL) ---")
    
    df = normalizar_datos_cobros_viajeros(df, log_messages)
    if progress_bar: progress_bar.progress(0.1, text="Fase de Normalización completada.")

    total_conciliados = 0
    indices_usados = set()

    # --- FASE 1: CONCILIACIÓN DE REVERSOS (CON LÓGICA 'ENDSWITH') ---
    log_messages.append("--- Fase 1: Buscando reversos con coincidencia parcial ---")
    df_reversos = df[df['Es_Reverso']].copy()
    df_originales = df[~df['Es_Reverso']].copy()

    for idx_r, reverso_row in df_reversos.iterrows():
        if idx_r in indices_usados:
            continue
        
        clave_reverso = reverso_row['Referencia_Norm_Num']
        if not clave_reverso:
            continue

        nit_reverso = reverso_row['NIT_Normalizado']
        
        # Iterar sobre los originales para encontrar la contrapartida
        for idx_o, original_row in df_originales.iterrows():
            if idx_o in indices_usados or original_row['NIT_Normalizado'] != nit_reverso:
                continue

            clave_orig_ref = original_row['Referencia_Norm_Num']
            clave_orig_fuente = original_row['Fuente_Norm_Num']
            
            # --- LA LÓGICA DE CRUCE CLAVE ---
            # Comprobar si una clave termina con la otra, en ambas direcciones
            match_en_referencia = (clave_reverso and clave_orig_ref and (clave_reverso.endswith(clave_orig_ref) or clave_orig_ref.endswith(clave_reverso)))
            match_en_fuente = (clave_reverso and clave_orig_fuente and (clave_reverso.endswith(clave_orig_fuente) or clave_orig_fuente.endswith(clave_reverso)))

            # Si hay un match de clave Y los montos se anulan
            if (match_en_referencia or match_en_fuente) and np.isclose(reverso_row['Monto_USD'] + original_row['Monto_USD'], 0, atol=TOLERANCIA_MAX_USD):
                indices_a_conciliar = [idx_r, idx_o]
                df.loc[indices_a_conciliar, 'Conciliado'] = True
                df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"REVERSO_{nit_reverso}_{clave_reverso}"
                indices_usados.update(indices_a_conciliar)
                total_conciliados += 2
                log_messages.append(f"✔️ Reverso (parcial) conciliado para NIT {nit_reverso} con clave {clave_reverso}.")
                break # Salir del bucle de originales y pasar al siguiente reverso

    if progress_bar: progress_bar.progress(0.5, text="Fase de Reversos completada.")

    # --- FASE 2: CONCILIACIÓN ESTÁNDAR N-a-N (Movimientos Restantes) ---
    # (Esta fase no necesita cambios)
    log_messages.append("--- Fase 2: Buscando grupos de conciliación estándar N-a-N ---")
    
    df['Clave_Vinculo'] = ''
    df_restante = df[~df.index.isin(indices_usados) & ~df['Conciliado']]
    
    for index, row in df_restante.iterrows():
        if row['Asiento'].startswith('CC'):
            df.loc[index, 'Clave_Vinculo'] = row['Fuente_Norm_Num']
        elif row['Asiento'].startswith('CB'):
            df.loc[index, 'Clave_Vinculo'] = row['Referencia_Norm_Num']

    df_procesable = df[(~df['Conciliado']) & (df['Clave_Vinculo'] != '')]
    grupos = df_procesable.groupby(['NIT_Normalizado', 'Clave_Vinculo'])
    
    for (nit, clave), grupo in grupos:
        if len(grupo) < 2 or not ((grupo['Monto_USD'] > 0).any() and (grupo['Monto_USD'] < 0).any()):
            continue
        if np.isclose(grupo['Monto_USD'].sum(), 0, atol=TOLERANCIA_MAX_USD):
            indices_a_conciliar = grupo.index
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"VIAJERO_{nit}_{clave}"
            indices_usados.update(indices_a_conciliar)
            total_conciliados += len(indices_a_conciliar)
            log_messages.append(f"✔️ Grupo estándar conciliado para NIT {nit} con clave {clave} ({len(indices_a_conciliar)} movimientos).")
            
    if progress_bar: progress_bar.progress(1.0, text="Conciliación completada.")

    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación finalizada: Se conciliaron un total de {total_conciliados} movimientos en ambas fases.")
    else:
        log_messages.append("ℹ️ No se encontraron movimientos para conciliar en ninguna de las fases.")
        
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# --- (G) Módulo: Otras Cuentas por Pagar (VES) ---

def normalizar_datos_otras_cxp(df, log_messages):
    """
    Prepara el DataFrame extrayendo el número de envío de la Referencia.
    """
    df_copy = df.copy()
    
    # Normalizar NIT para usarlo como clave de agrupación
    nit_col_name = next((col for col in df_copy.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df_copy['NIT_Normalizado'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró columna 'NIT' o 'RIF'.")
        df_copy['NIT_Normalizado'] = 'SIN_NIT'

    # Extraer el número de envío usando una expresión regular
    # Busca el patrón "ENV:" seguido de uno o más dígitos (\d+)
    df_copy['Numero_Envio'] = df_copy['Referencia'].str.extract(r"ENV:(\d+)", expand=False, flags=re.IGNORECASE)
    
    return df_copy


def run_conciliation_otras_cxp(df, log_messages, progress_bar=None):
    """
    Orquesta la conciliación para Otras Cuentas por Pagar, agrupando por NIT
    y cruzando por el número de envío extraído. La conciliación es en VES.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE OTRAS CUENTAS POR PAGAR (VES) ---")
    
    df = normalizar_datos_otras_cxp(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase de Normalización completada.")

    total_conciliados = 0
    
    # Filtrar solo las filas donde se pudo extraer un número de envío
    df_procesable = df[(~df['Conciliado']) & (df['Numero_Envio'].notna())]
    
    # Agrupar por NIT y luego por Número de Envío
    grupos = df_procesable.groupby(['NIT_Normalizado', 'Numero_Envio'])
    log_messages.append(f"ℹ️ Se encontraron {len(grupos)} combinaciones de NIT/Envío para analizar.")

    for (nit, envio), grupo in grupos:
        if len(grupo) < 2: # Se necesita al menos un débito y un crédito
            continue

        # Verificar si la suma de los movimientos en Bolívares es cero
        if np.isclose(grupo['Monto_BS'].sum(), 0, atol=TOLERANCIA_MAX_BS):
            indices_a_conciliar = grupo.index
            
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"OTRAS_CXP_{nit}_{envio}"
            
            total_conciliados += len(indices_a_conciliar)

    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación finalizada: Se conciliaron {total_conciliados} movimientos.")
    else:
        log_messages.append("ℹ️ No se encontraron grupos por NIT/Envío que sumen cero.")
        
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# ==============================================================================
# FUNCIONES MAESTRAS DE ESTRATEGIA
# ==============================================================================

def run_conciliation_fondos_en_transito (df, log_messages):
    df = normalizar_referencia_fondos_en_transito(df)
    log_messages.append("\n--- INICIANDO LÓGICA DE FONDOS EN TRÁNSITO ---")
    conciliar_diferencia_cambio(df, log_messages)
    conciliar_ajuste_automatico(df, log_messages)
    conciliar_pares_exactos_cero(df, 'GRUPO_SILLACA', 'FASE SILLACA 1/7 (Cruce CERO)', log_messages)
    conciliar_pares_exactos_por_referencia(df, 'GRUPO_SILLACA', 'FASE SILLACA 2/7 (Pares por Referencia)', log_messages)
    cruzar_pares_simples(df, 'REINTEGRO_SILLACA', 'FASE SILLACA 3/7 (Pares por Monto)', log_messages)
    cruzar_grupos_por_criterio(df, 'REINTEGRO_SILLACA', 'Asiento', 'SILLACA_ASIENTO', 'FASE SILLACA 4/7 (Grupos por Asiento)', log_messages)
    cruzar_grupos_por_criterio(df, 'REINTEGRO_SILLACA', 'Referencia_Normalizada_Literal', 'SILLACA_REF', 'FASE SILLACA 5/7 (Grupos por Ref. Literal)', log_messages)
    cruzar_grupos_por_criterio(df, 'REINTEGRO_SILLACA', 'Fecha', 'SILLACA_FECHA', 'FASE SILLACA 6/7 (Grupos por Fecha)', log_messages)
    conciliar_lote_por_grupo(df, 'GRUPO_SILLACA', 'FASE SILLACA 7/7 (CRUCE POR LOTE)', log_messages)
    conciliar_pares_exactos_cero(df, 'GRUPO_NOTA', 'FASE NOTAS 1/6 (Cruce CERO)', log_messages)
    conciliar_pares_exactos_por_referencia(df, 'GRUPO_NOTA', 'FASE NOTAS 2/6 (Pares por Referencia)', log_messages)
    cruzar_pares_simples(df, 'NOTA_GENERAL', 'FASE NOTAS 3/6 (Pares por Monto)', log_messages)
    cruzar_grupos_por_criterio(df, 'NOTA_GENERAL', 'Referencia_Normalizada_Literal', 'NOTA_REF', 'FASE NOTAS 4/6 (Grupos por Ref. Literal)', log_messages)
    cruzar_grupos_por_criterio(df, 'NOTA_GENERAL', 'Fecha', 'NOTA_FECHA', 'FASE NOTAS 5/6 (Grupos por Fecha)', log_messages)
    conciliar_lote_por_grupo(df, 'GRUPO_NOTA', 'FASE NOTAS 6/6 (CRUCE POR LOTE)', log_messages)
    conciliar_pares_exactos_cero(df, 'GRUPO_BANCO', 'FASE BANCO 1/5 (Cruce CERO)', log_messages)
    conciliar_pares_exactos_por_referencia(df, 'GRUPO_BANCO', 'FASE BANCO 2/5 (Pares por Referencia)', log_messages)
    cruzar_pares_simples(df, 'BANCO_A_BANCO', 'FASE BANCO 3/5 (Pares por Monto)', log_messages)
    cruzar_grupos_por_criterio(df, 'BANCO_A_BANCO', 'Referencia_Normalizada_Literal', 'BANCO_REF', 'FASE BANCO 4/5 (Grupos por Ref. Literal)', log_messages)
    cruzar_grupos_por_criterio(df, 'BANCO_A_BANCO', 'Fecha', 'BANCO_FECHA', 'FASE BANCO 5/5 (Grupos por Fecha)', log_messages)
    conciliar_pares_exactos_cero(df, 'GRUPO_REMESA', 'FASE REMESA 1/3 (Cruce CERO)', log_messages)
    cruzar_pares_simples(df, 'REMESA_GENERAL', 'FASE REMESA 2/3 (Pares por Monto)', log_messages)
    cruzar_grupos_por_criterio(df, 'REMESA_GENERAL', 'Referencia_Normalizada_Literal', 'REMESA_REF', 'FASE REMESA 3/3 (Grupos por Ref. Literal)', log_messages)
    conciliar_grupos_globales_por_referencia(df, log_messages)
    conciliar_pares_globales_remanentes(df, log_messages)
    conciliar_grupos_complejos_usd(df, log_messages)
    conciliar_pares_globales_remanentes(df, log_messages)
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

def run_conciliation_fondos_por_depositar(df, log_messages, progress_bar=None):
    log_messages.append("\n--- INICIANDO LÓGICA DE FONDOS POR DEPOSITAR (USD) ---")
    df = normalizar_referencia_fondos_usd(df)
    
    conciliar_automaticos_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.1, text="Fase 1/6: Conciliaciones automáticas completada.")
    
    conciliar_grupos_por_referencia_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase 2/6: Grupos por referencia específica completada.")
    
    conciliar_pares_banco_a_banco_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.35, text="Fase 3/6: Pares 'Banco a Banco' completada.")
    
    conciliar_pares_globales_exactos_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.5, text="Fase 4/6: Pares globales exactos completada.")
    
    conciliar_pares_globales_remanentes_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.65, text="Fase 5/6: Búsqueda de pares con tolerancia completada.")

    conciliar_grupos_complejos_usd(df, log_messages, progress_bar)
    if progress_bar: progress_bar.progress(0.9, text="Fase 6/6: Búsqueda de grupos complejos completada.")
    
    conciliar_gran_total_final_usd(df, log_messages)
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

def run_conciliation_viajes(df, log_messages, progress_bar=None):
    log_messages.append("\n--- INICIANDO LÓGICA DE CUENTAS DE VIAJES (BS) ---")
    
    df = normalizar_referencia_viajes(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase de Normalización completada.")
    
    conciliar_pares_exactos_por_nit_viajes(df, log_messages)
    if progress_bar: progress_bar.progress(0.5, text="Fase 1/2: Búsqueda de pares exactos completada.")
    
    conciliar_grupos_por_nit_viajes(df, log_messages)
    if progress_bar: progress_bar.progress(0.9, text="Fase 2/2: Búsqueda de grupos complejos completada.")
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

def run_conciliation_deudores_empleados_me(df, log_messages, progress_bar=None):
    """
    Orquesta el proceso de conciliación para la cuenta Deudores Empleados en ME.
    La lógica principal es conciliar por empleado si su saldo total en USD es cero.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE DEUDORES EMPLEADOS (ME) ---")
    
    # Paso 1: Normalizar los datos para obtener una clave de empleado fiable
    df = normalizar_datos_deudores_empleados(df, log_messages)
    if progress_bar: progress_bar.progress(0.3, text="Fase de Normalización completada.")
    
    # Paso 2: Ejecutar la lógica de conciliación principal
    conciliar_grupos_por_empleado(df, log_messages)
    if progress_bar: progress_bar.progress(0.8, text="Fase de Conciliación por Empleado completada.")
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# ==============================================================================
# LÓGICAS PARA LA HERRAMIENTA DE RELACIONES DE RETENCIONES
# ==============================================================================

# --- NUEVAS FUNCIONES DE NORMALIZACIÓN (Paso 4) ---

def _normalizar_rif(valor):
    """Normaliza RIF, eliminando espacios y caracteres no alfanuméricos."""
    if pd.isna(valor): return ''
    val_str = str(valor).strip().upper()
    val_str = re.sub(r'[^A-Z0-9]', '', val_str)
    if val_str.startswith(('J', 'V', 'E', 'G')) and len(val_str) > 8:
        return val_str[1:]
    return val_str

def _normalizar_numerico(valor):
    """
    (Versión Definitiva) Normaliza un valor numérico como texto,
    eliminando cualquier caracter no-numérico Y los ceros a la izquierda.
    """
    if pd.isna(valor):
        return ''
    
    # 1. Quita espacios al principio y al final
    val_str = str(valor).strip()
    
    # 2. Quita todo lo que no sea un dígito
    solo_digitos = re.sub(r'[^0-9]', '', val_str)
    
    if not solo_digitos:
        return ''
        
    # 3. Convierte a entero para eliminar ceros a la izquierda y luego de nuevo a string
    return str(int(solo_digitos))

def _extraer_factura_cp(aplicacion):
    """
    (Versión Final y Robusta) Extrae el número de factura buscando el
    último "bloque" de texto alfanumérico en la cadena de aplicación.
    """
    if pd.isna(aplicacion):
        return ''
    
    # Esta expresión regular busca uno o más caracteres de palabra (letras, números, _)
    # o guiones, que se encuentren al final de la línea ($).
    match = re.search(r'([\w-]+)$', str(aplicacion).strip())
    
    if match:
        # El texto extraído (ej. 'B-00010' o 'A-7000095120') se guarda
        numero_texto_extraido = match.group(1)
        
        # Lo pasamos a nuestra función de normalización para limpiar letras y ceros
        return _normalizar_numerico(numero_texto_extraido)
        
    return ''

# --- NUEVAS FUNCIONES DE CARGA Y PREPARACIÓN ---

def preparar_df_cp(file_cp):
    df = pd.read_excel(file_cp, header=4, dtype=str).rename(columns={
        'Asiento Contable': 'Asiento', 'Proveedor': 'RIF', 'Tipo': 'Tipo', 
        'Fecha': 'Fecha', 'Número': 'Comprobante', 'Monto': 'Monto',
        'Aplicación': 'Aplicacion', 'Subtipo': 'Subtipo'
    })
    df['RIF_norm'] = df['RIF'].apply(_normalizar_rif)
    df['Comprobante_norm'] = df['Comprobante'].apply(_normalizar_numerico)
    df['Factura_norm'] = df['Aplicacion'].apply(_extraer_factura_cp)
    df['Monto'] = df['Monto'].str.replace(',', '.', regex=False).astype(float)
    return df

def preparar_df_iva(file_iva):
    df = pd.read_excel(file_iva, header=4, dtype=str).rename(columns={
        'Rif Prov.': 'RIF', 'Nombre o Razón Social': 'Nombre_Proveedor', 
        'Nº Documento': 'Factura', 'No. Comprobante': 'Comprobante', 
        'IVA Retenido': 'Monto'
    })
    
    df['RIF_norm'] = df['RIF'].apply(_normalizar_rif)
    df['Comprobante_norm'] = df['Comprobante'].apply(_normalizar_numerico)
    df['Factura_norm'] = df['Factura'].apply(_normalizar_numerico)
    df['Monto'] = df['Monto'].str.replace(',', '.', regex=False).astype(float)
    return df

def preparar_df_municipal(file_path):
    """
    (Versión Final y Robusta) Carga y prepara el archivo de Retenciones Municipales,
    buscando dinámicamente las columnas clave, incluyendo el Nombre del Proveedor.
    """
    df = pd.read_excel(file_path, header=8, dtype=str)
    
    # --- Búsqueda Robusta de Columnas ---
    column_map = {}
    for col in df.columns:
        # Normalizamos el nombre de la columna para la comparación
        col_normalized = col.strip().lower().replace(" ", "")
        
        if col_normalized == 'númerorif':
            column_map[col] = 'RIF'
        
        # --- ¡AQUÍ SE AÑADE LA LÓGICA DEL PROVEEDOR! ---
        elif col_normalized == 'razonsocialdelsujetoretenido':
            column_map[col] = 'Nombre_Proveedor'
            
        elif col_normalized == 'númerodefactura':
            column_map[col] = 'Factura'
        elif col_normalized == 'valor':
            column_map[col] = 'Monto'
            
    # Renombramos usando el mapa que creamos
    df.rename(columns=column_map, inplace=True)

    # Verificamos que las columnas clave ahora existen para evitar errores posteriores
    if 'RIF' not in df.columns: df['RIF'] = ''
    if 'Nombre_Proveedor' not in df.columns: df['Nombre_Proveedor'] = '' # Aseguramos que la columna exista
    if 'Factura' not in df.columns: df['Factura'] = ''
    if 'Monto' not in df.columns: df['Monto'] = 0
    
    # --- Normalización de Datos (sin cambios) ---
    df['RIF_norm'] = df['RIF'].apply(_normalizar_rif)
    df['Factura_norm'] = df['Factura'].apply(_normalizar_numerico)
    df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    
    # Añadimos una columna de Comprobante vacía para consistencia
    df['Comprobante_norm'] = ''
    
    return df

def preparar_df_islr(file_path):
    """
    (Versión Definitiva y Corregida) Carga y prepara el archivo de ISLR, combinando
    la búsqueda robusta de columnas por nombre y la extracción posicional para la factura.
    """
    df = pd.read_excel(file_path, header=8, dtype=str)

    # --- 1. Búsqueda Robusta de Columnas por Nombre (para RIF y Proveedor) ---
    # Usamos el mismo método que en la función de Municipal para máxima fiabilidad.
    column_map = {}
    for col in df.columns:
        col_normalized = col.strip().upper().replace(" ", "").replace(".", "")
        if col_normalized == 'RIFPROVEEDOR':
            column_map[col] = 'RIF'
        elif col_normalized == 'RAZÓNSOCIALDELSUJETORETENIDO':
            column_map[col] = 'Nombre_Proveedor'
    df.rename(columns=column_map, inplace=True)

    # --- 2. Extracción Posicional para la Factura (la columna sin nombre) ---
    try:
        col_anclaje_idx = df.columns.get_loc('Nº Documento')
        col_factura_idx = col_anclaje_idx + 1
        df['Factura'] = df.iloc[:, col_factura_idx]
    except KeyError:
        df['Factura'] = ''

    # --- 3. Renombrar Columnas Restantes con Nombres Fijos ---
    df.rename(columns={
        'Nº Referencia': 'Comprobante',
        'Monto Retenido': 'Monto'
    }, inplace=True)

    # --- 4. Asegurar que las columnas clave existan para evitar errores ---
    if 'RIF' not in df.columns: df['RIF'] = ''
    if 'Nombre_Proveedor' not in df.columns: df['Nombre_Proveedor'] = ''
    if 'Factura' not in df.columns: df['Factura'] = ''
    if 'Monto' not in df.columns: df['Monto'] = 0

    # --- 5. Normalizar TODOS los datos ---
    df['RIF_norm'] = df['RIF'].apply(_normalizar_rif)
    df['Comprobante_norm'] = df['Comprobante'].apply(_normalizar_numerico)
    df['Factura_norm'] = df['Factura'].apply(_normalizar_numerico)
    df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)

    return df
    
# --- NUEVAS FUNCIONES DE LÓGICA DE CONCILIACIÓN ---

def _conciliar_iva(cp_row, df_iva):
    """
    (Versión con manejo de Notas de Crédito) Lógica de conciliación que compara
    el valor absoluto de los montos si detecta una Nota de Crédito.
    """
    rif_cp = cp_row['RIF_norm']
    comprobante_cp_norm = cp_row['Comprobante_norm']
    
    # 1. Búsqueda principal (sin cambios)
    match_encontrado = df_iva[
        (df_iva['RIF_norm'] == rif_cp) & 
        (df_iva['Comprobante_norm'] == comprobante_cp_norm)
    ]
    
    # 2. Lógica de error si no se encuentra (sin cambios)
    if match_encontrado.empty:
        # Tu lógica de error inteligente para sugerir comprobantes se mantiene aquí...
        probable_match = df_iva[
            (df_iva['RIF_norm'] == rif_cp) &
            (df_iva['Factura_norm'] == cp_row['Factura_norm']) &
            (np.isclose(df_iva['Monto'].abs(), abs(cp_row['Monto']))) # Usamos abs aquí también por si acaso
        ]
        if not probable_match.empty and len(probable_match) == 1:
            comprobante_sugerido = probable_match.iloc[0]['Comprobante']
            mensaje_error = f"Comprobante no coincide. CP: {cp_row['Comprobante']}, GALAC sugiere: {comprobante_sugerido}"
            return 'No Conciliado', mensaje_error
        else:
            registros_del_rif_en_galac = df_iva[df_iva['RIF_norm'] == rif_cp]
            if registros_del_rif_en_galac.empty:
                return 'No Conciliado', 'RIF no se encuentra en GALAC'
            else:
                comprobantes_disponibles_en_galac = registros_del_rif_en_galac['Comprobante'].unique().tolist()
                comprobantes_str = ", ".join(map(str, comprobantes_disponibles_en_galac))
                mensaje_error = (f"El Comprobante de CP ({cp_row['Comprobante']}) no se encontró. "
                                 f"Para ese RIF, GALAC tiene estos: [{comprobantes_str}]")
                return 'No Conciliado', mensaje_error

    # 3. Si se encuentra la coincidencia, procedemos a las validaciones.
    match_row = match_encontrado.iloc[0]
    errores = []
    
    # Validación de Factura (sin cambios)
    if cp_row['Factura_norm'] != match_row['Factura_norm']:
        msg = f"Numero de factura no coincide. CP: {cp_row['Factura_norm']}, GALAC: {match_row['Factura_norm']}"
        errores.append(msg)
    
    # --- INICIO DE LA NUEVA LÓGICA PARA NOTAS DE CRÉDITO ---
    
    # Convertimos el texto de "Aplicacion" a mayúsculas para una comparación robusta.
    aplicacion_text = str(cp_row.get('Aplicacion', '')).upper()
    
    # Verificamos si alguna de las palabras clave de Nota de Crédito está presente.
    # "N/C" se convierte en "NC" al quitar caracteres no alfanuméricos, por eso solo buscar "NC" es suficiente.
    is_credit_note = 'NC' in aplicacion_text or 'NOTA CREDITO' in aplicacion_text

    if is_credit_note:
        # Si es una Nota de Crédito, comparamos los valores absolutos.
        if not np.isclose(abs(cp_row['Monto']), abs(match_row['Monto'])):
            msg = f"Monto (NC) no coincide. CP: {cp_row['Monto']:.2f}, GALAC: {match_row['Monto']:.2f}"
            errores.append(msg)
    else:
        # Si NO es una Nota de Crédito, usamos la comparación normal.
        if not np.isclose(cp_row['Monto'], match_row['Monto']):
            msg = f"Monto no coincide. CP: {cp_row['Monto']:.2f}, GALAC: {match_row['Monto']:.2f}"
            errores.append(msg)
            
    # --- FIN DE LA NUEVA LÓGICA ---
        
    return ('Conciliado', 'OK') if not errores else ('Parcialmente Conciliado', ' | '.join(errores))

def _conciliar_islr(cp_row, df_islr):
    """
    (Versión Definitiva con Suma Inteligente y Mensajes Corregidos) Maneja
    comprobantes con múltiples facturas Y múltiples retenciones para una misma factura.
    """
    rif_cp = cp_row['RIF_norm']
    comprobante_cp_norm = cp_row['Comprobante_norm']
    factura_cp_norm = cp_row['Factura_norm']
    
    # 1. Encontrar el grupo completo del comprobante
    comprobante_group = df_islr[
        (df_islr['RIF_norm'] == rif_cp) & 
        (df_islr['Comprobante_norm'] == comprobante_cp_norm)
    ]
    
    if comprobante_group.empty:
        # Lógica de error si el comprobante no existe (se mantiene)
        if rif_cp not in df_islr['RIF_norm'].values: return 'No Conciliado', 'RIF no se encuentra en el reporte de ISLR'
        return 'No Conciliado', f"Comprobante de CP ({cp_row['Comprobante']}) no encontrado."

    # 2. Dentro de ese grupo, encontrar TODAS las retenciones para nuestra factura específica
    specific_invoice_matches = comprobante_group[comprobante_group['Factura_norm'] == factura_cp_norm]
    
    if specific_invoice_matches.empty:
        # --- LÍNEA CORREGIDA ---
        # Se cambió cp_row['Factura'] por cp_row['Factura_norm'] para que coincida
        # con el nombre de columna que sí existe en el DataFrame de CP.
        all_invoices_in_group = comprobante_group['Factura'].unique().tolist()
        msg = (f"Factura de CP ({cp_row['Factura_norm']}) no encontrada para el Comprobante {cp_row['Comprobante']}. "
               f"Este comprobante en GALAC contiene estas facturas: {all_invoices_in_group}")
        return 'No Conciliado', msg

    # 3. Sumar los montos de TODAS las retenciones encontradas para ESA factura
    monto_islr_sumado = specific_invoice_matches['Monto'].sum()
    errores = []
    
    # 4. Comparar el monto de CP con el monto SUMADO de ISLR
    if not np.isclose(cp_row['Monto'], monto_islr_sumado):
        msg = f"Monto no coincide. CP: {cp_row['Monto']:.2f}, ISLR (suma de retenciones para factura): {monto_islr_sumado:.2f}"
        errores.append(msg)
        
    # 5. Añadir mensaje informativo si el comprobante contenía otras facturas
    if len(comprobante_group['Factura_norm'].unique()) > 1:
        all_invoices_in_group = comprobante_group['Factura'].unique().tolist()
        info_msg = f"INFO: Comprobante {cp_row['Comprobante']} incluye otras facturas en GALAC: {all_invoices_in_group}"
        errores.append(info_msg)
        
    return ('Conciliado', 'OK') if not errores else ('Parcialmente Conciliado', ' | '.join(errores))
    
def _conciliar_municipal(cp_row, df_municipal):
    """
    (Versión Robusta y Multi-paso) Aplica la lógica de conciliación Municipal.
    - Usa np.isclose para una comparación segura de montos.
    - Busca la factura exacta dentro de un grupo de posibles candidatos.
    """
    rif_cp = cp_row['RIF_norm']
    monto_cp = cp_row['Monto']
    factura_cp_norm = cp_row['Factura_norm']
    
    candidatos_por_rif = df_municipal[df_municipal['RIF_norm'] == rif_cp]
    
    if candidatos_por_rif.empty:
        return 'No Conciliado', 'RIF no se encuentra en GALAC'
        
    posibles_matches = candidatos_por_rif[np.isclose(candidatos_por_rif['Monto'], monto_cp)]
    
    if posibles_matches.empty:
        return 'No Conciliado', f"Monto de retencion no encontrado en GALAC. Monto CP: {monto_cp:.2f}"

    match_perfecto = posibles_matches[posibles_matches['Factura_norm'] == factura_cp_norm]
    
    if not match_perfecto.empty:
        return 'Conciliado', 'OK'
    else:
        factura_sugerida_galac = posibles_matches.iloc[0]['Factura_norm']
        msg = f"Numero de factura no coincide. CP: {factura_cp_norm}, GALAC sugiere: {factura_sugerida_galac}"
        return 'Parcialmente Conciliado', msg
        
def _traducir_resultados_para_reporte(row, asientos_en_cg_set, df_cg):
    """
    (Versión Final con Doble Lógica Jerárquica) La validación de CG ahora respeta
    tanto los errores de subtipo como los errores de monto de la conciliación de GALAC.
    """
    # El resultado de CP vs GALAC no cambia
    estado_galac = row['Estado_Conciliacion']
    detalle_galac = row['Detalle']
    if estado_galac == 'Anulado': cp_vs_galac = 'Anulado'
    elif estado_galac == 'Conciliado': cp_vs_galac = 'Sí'
    else: cp_vs_galac = detalle_galac
    
    # --- INICIO DE LA LÓGICA DE VALIDACIÓN DE CG ---
    asiento_cp = row.get('Asiento', None)
    
    if not asiento_cp or df_cg.empty:
        return cp_vs_galac, 'No Aplica'

    if asiento_cp not in asientos_en_cg_set:
        return cp_vs_galac, 'Asiento no encontrado en CG'

    errores_cg = []
    transacciones_asiento = df_cg[df_cg['ASIENTO'] == asiento_cp]
    
    # --- 1. VALIDACIÓN DE CUENTA CONTABLE (con comprobación jerárquica) ---
    if estado_galac == 'Error de Subtipo':
        errores_cg.append('Cuenta Contable no coincide (debido a Error de Subtipo en GALAC)')
    else:
        mapa_cuentas = {'IVA': '2.1.3.05.1.001', 'ISLR': '2.1.3.02.1.002', 'MUNICIPAL': '2.1.3.02.4.002'}
        subtipo_cp = str(row.get('Subtipo', '')).upper()
        cuenta_esperada = None
        for key, value in mapa_cuentas.items():
            if key in subtipo_cp:
                cuenta_esperada = value
                break
        
        if cuenta_esperada and 'CUENTACONTABLE' in transacciones_asiento.columns:
            cuentas_en_asiento = set(transacciones_asiento['CUENTACONTABLE'].str.strip())
            if cuenta_esperada not in cuentas_en_asiento:
                errores_cg.append('Cuenta Contable no coincide con Tipo de retencion')

    # --- 2. VALIDACIÓN DE MONTO (con comprobación jerárquica) ---
    monto_cp = row.get('Monto', 0)
    
    # ¡NUEVA COMPROBACIÓN PRIORITARIA PARA MONTO!
    # Si el detalle de GALAC ya indica un error de monto, heredamos ese error.
    if 'Monto no coincide' in detalle_galac:
        errores_cg.append('Monto no coincide (debido a discrepancia con GALAC)')
    
    # Solo si no hay error de monto en GALAC, procedemos a comparar con CG.
    else:
        aplicacion_text = str(row.get('Aplicacion', '')).upper()
        es_nota_credito = 'NC' in aplicacion_text or 'NOTA CREDITO' in aplicacion_text
        columna_monto_cg = 'CREDITO_NORM' if es_nota_credito else 'DEBITO_NORM'
        
        if columna_monto_cg in transacciones_asiento.columns:
            montos_cg = pd.to_numeric(transacciones_asiento[columna_monto_cg], errors='coerce').fillna(0)
            suma_monto_cg = montos_cg.sum()
            if not np.isclose(monto_cp, suma_monto_cg):
                errores_cg.append(f'Monto no coincide (CP: {monto_cp:.2f}, CG: {suma_monto_cg:.2f})')
        else:
            errores_cg.append(f'Columna de monto ({columna_monto_cg}) no encontrada en CG')

    # --- 3. GENERAR RESULTADO FINAL ---
    if not errores_cg:
        resultado_final_cg = 'Conciliado en CG'
    else:
        resultado_final_cg = ' | '.join(errores_cg)
        
    return cp_vs_galac, resultado_final_cg

def run_conciliation_retenciones(file_cp, file_cg, file_iva, file_islr, file_mun, log_messages):
    """
    Función principal que orquesta todo el proceso de conciliación de retenciones.
    (Versión final y completa, con todas las correcciones).
    """
    try:
        log_messages.append("--- INICIANDO PROCESO DE CONCILIACIÓN DE RETENCIONES ---")
        
        # --- 1. CARGA Y PREPARACIÓN DE ARCHIVOS DE DATOS ---
        df_cp = preparar_df_cp(file_cp)
        df_iva = preparar_df_iva(file_iva)
        df_islr = preparar_df_islr(file_islr)
        df_municipal = preparar_df_municipal(file_mun)
        
        # --- 2. CREACIÓN DEL MAPA DE PROVEEDORES ---
        provider_dfs = []
        if all(col in df_iva.columns for col in ['RIF_norm', 'Nombre_Proveedor']):
            provider_dfs.append(df_iva[['RIF_norm', 'Nombre_Proveedor']])
        if all(col in df_islr.columns for col in ['RIF_norm', 'Nombre_Proveedor']):
            provider_dfs.append(df_islr[['RIF_norm', 'Nombre_Proveedor']])
        if all(col in df_municipal.columns for col in ['RIF_norm', 'Nombre_Proveedor']):
            provider_dfs.append(df_municipal[['RIF_norm', 'Nombre_Proveedor']])
        if provider_dfs:
            provider_map_df = pd.concat(provider_dfs).dropna(subset=['RIF_norm', 'Nombre_Proveedor']).drop_duplicates(subset=['RIF_norm'])
        else:
            provider_map_df = pd.DataFrame(columns=['RIF_norm', 'Nombre_Proveedor'])

        # --- 3. PREPARACIÓN DE DATOS DE CONTABILIDAD GENERAL (CG) ---
        if file_cg:
            df_cg_dummy = pd.read_excel(file_cg, header=0, dtype=str)
            df_cg_dummy.columns = [col.strip().upper().replace(' ', '') for col in df_cg_dummy.columns]
            debit_names = ['DEBITOVES', 'DÉBITOVES', 'DEBITO', 'DÉBITO', 'DEBEVESDÉBITO', 'MONEDALOCAL', 'DÉBITOBOLIVAR']
            credit_names = ['CREDITOVES', 'CRÉDITOVES', 'CREDITO', 'CRÉDITO', 'CREDITOVESMCREDITOLOCAL', 'CRÉDITOBOLIVAR']
            for col_name in df_cg_dummy.columns:
                if col_name in debit_names:
                    df_cg_dummy.rename(columns={col_name: 'DEBITO_NORM'}, inplace=True); break
            for col_name in df_cg_dummy.columns:
                if col_name in credit_names:
                    df_cg_dummy.rename(columns={col_name: 'CREDITO_NORM'}, inplace=True); break
            if 'ASIENTO' in df_cg_dummy.columns:
                asientos_en_cg_set = set(df_cg_dummy['ASIENTO'].dropna().unique())
            else:
                log_messages.append("Advertencia: No se encontró la columna 'ASIENTO' en el archivo de CG.")
                asientos_en_cg_set = set()
        else:
            df_cg_dummy = pd.DataFrame()
            asientos_en_cg_set = set()
        
        log_messages.append("Iniciando conciliación por tipo de impuesto...")
        
        # --- ¡LÍNEA CORREGIDA! ---
        # Aquí se inicializa la lista para guardar los resultados.
        resultados = []
        
        # --- 4. BUCLE PRINCIPAL DE CONCILIACIÓN ---
        for index, row in df_cp.iterrows():
            if 'ANULADO' in str(row.get('Aplicacion', '')).upper():
                resultados.append({'Estado_Conciliacion': 'Anulado', 'Detalle': 'Movimiento Anulado en CP'})
                continue

            subtipo = str(row.get('Subtipo', '')).upper()
            
            if 'IVA' in subtipo:
                tipo_primario = 'IVA'; busqueda_primaria = lambda r: _conciliar_iva(r, df_iva)
                busquedas_cruzadas = [('ISLR', lambda r: _conciliar_islr(r, df_islr)), ('Municipal', lambda r: _conciliar_municipal(r, df_municipal))]
            elif 'ISLR' in subtipo:
                tipo_primario = 'ISLR'; busqueda_primaria = lambda r: _conciliar_islr(r, df_islr)
                busquedas_cruzadas = [('IVA', lambda r: _conciliar_iva(r, df_iva)), ('Municipal', lambda r: _conciliar_municipal(r, df_municipal))]
            elif 'MUNICIPAL' in subtipo:
                tipo_primario = 'Municipal'; busqueda_primaria = lambda r: _conciliar_municipal(r, df_municipal)
                busquedas_cruzadas = [('IVA', lambda r: _conciliar_iva(r, df_iva)), ('ISLR', lambda r: _conciliar_islr(r, df_islr))]
            else:
                resultados.append({'Estado_Conciliacion': 'No Conciliado', 'Detalle': 'Subtipo no reconocido'})
                continue

            estado, mensaje = busqueda_primaria(row)

            if estado == 'No Conciliado':
                for nombre_otro_tipo, busqueda_otro_tipo in busquedas_cruzadas:
                    estado_otro, _ = busqueda_otro_tipo(row)
                    if estado_otro in ['Conciliado', 'Parcialmente Conciliado']:
                        estado = 'Error de Subtipo'; mensaje = f'Declarado como {tipo_primario}, pero encontrado en {nombre_otro_tipo}'; break
            
            resultados.append({'Estado_Conciliacion': estado, 'Detalle': mensaje})

        # --- 5. POST-PROCESAMIENTO Y GENERACIÓN DEL REPORTE FINAL ---
        df_resultados = pd.DataFrame(resultados)
        df_cp_temp = pd.concat([df_cp.reset_index(drop=True), df_resultados], axis=1)

        df_cp_temp[['CP_Vs_Galac', 'Validacion_CG']] = df_cp_temp.apply(
            lambda row: _traducir_resultados_para_reporte(row, asientos_en_cg_set, df_cg_dummy), 
            axis=1, 
            result_type='expand'
        )
        
        df_cp_final = pd.merge(df_cp_temp, provider_map_df, on='RIF_norm', how='left')
        df_cp_final.rename(columns={'Nombre_Proveedor': 'Nombre Proveedor'}, inplace=True)
        
        log_messages.append("¡Proceso de conciliación completado con éxito!")
        
        df_galac_no_cp = pd.DataFrame(); cuentas_map_dummy = {}

        return generar_reporte_retenciones(df_cp_final, df_galac_no_cp, df_cg_dummy, cuentas_map_dummy)

    except Exception as e:
        log_messages.append(f"❌ ERROR CRÍTICO: {e}")
        import traceback
        log_messages.append(traceback.format_exc())
        return None

# ==============================================================================
# LÓGICA PARA LA HERRAMIENTA DE ANÁLISIS DE PAQUETE CC (VERSIÓN ACTUALIZADA)
# ==============================================================================

def normalize_account(acc):
    """Función auxiliar que limpia un número de cuenta, eliminando todo lo que no sea un dígito."""
    return re.sub(r'\D', '', str(acc))

# --- Directorio de Cuentas (Versión Normalizada) ---
CUENTAS_CONOCIDAS = {normalize_account(acc) for acc in [
    '1.1.3.01.1.001', '1.1.3.01.1.901', '7.1.3.45.1.997', '6.1.1.12.1.001',
    '4.1.1.22.4.001', '2.1.3.04.1.001', '7.1.3.19.1.012', '2.1.2.05.1.108',
    '6.1.1.19.1.001', '4.1.1.21.4.001', '2.1.3.04.1.006', '2.1.3.01.1.012',
    '7.1.3.04.1.004', '7.1.3.06.1.998', '1.1.1.04.6.003', '1.1.4.01.7.020',
    '1.1.4.01.7.021', '1.1.1.02.1.004', '1.1.1.02.1.007', '1.1.1.02.1.009',
    '1.1.1.02.1.016', '1.1.1.02.1.112', '1.1.1.02.1.124', '1.1.1.02.1.132',
    '1.1.1.02.6.002', '1.1.1.02.6.003', '1.1.1.02.6.005', '1.1.1.02.6.010',
    '1.1.1.03.6.012', '1.1.1.03.6.024', '1.1.1.03.6.026', '1.1.1.03.6.031',
]}

CUENTAS_BANCO = {normalize_account(acc) for acc in [
    '1.1.4.01.7.020', '1.1.4.01.7.021', '1.1.1.02.1.004', '1.1.1.02.1.007',
    '1.1.1.02.1.009', '1.1.1.02.1.016', '1.1.1.02.1.112', '1.1.1.02.1.124',
    '1.1.1.02.1.132', '1.1.1.02.6.002', '1.1.1.02.6.003', '1.1.1.02.6.005',
    '1.1.1.02.6.010', '1.1.1.03.6.012', '1.1.1.03.6.024', '1.1.1.03.6.026',
    '1.1.1.03.6.031',
]}

def _get_base_classification(asiento_group, cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, is_reverso_check=False):
    """
    Función auxiliar que contiene la lógica de clasificación base, ordenada por prioridad.
    El parámetro 'is_reverso_check' relaja ciertas reglas para identificar el tipo de un reverso.
    """
    # PRIORIDAD 1: Notas de Crédito
    # Si estamos chequeando un reverso, no exigimos que la Fuente contenga 'N/C', solo las cuentas.
    if (is_reverso_check or 'N/C' in fuente_completa):
        if {normalize_account('4.1.1.22.4.001'), normalize_account('2.1.3.04.1.001')}.issubset(cuentas_del_asiento):
            if is_reverso_check: return "Grupo 3: N/C" # Para reversos, solo identificamos el tipo base
            if 'AVISOS DE CREDITO' in referencia_completa: return "Grupo 3: N/C - Avisos de Crédito"
            if referencia_limpia_palabras.intersection({'ESTRATEGIA', 'ESTRATEGIAS'}): return "Grupo 3: N/C - Estrategias"
            if referencia_limpia_palabras.intersection({'INCENTIVO', 'INCENTIVOS'}): return "Grupo 3: N/C - Incentivos"
            if referencia_limpia_palabras.intersection({'BONIFICACION', 'BONIFICACIONES', 'BONIF', 'BONF'}): return "Grupo 3: N/C - Bonificaciones"
            if referencia_limpia_palabras.intersection({'DESCUENTO', 'DESCUENTOS', 'DSCTO', 'DESC', 'DESTO'}): return "Grupo 3: N/C - Descuentos"
            return "Grupo 3: N/C - Otros"

    # PRIORIDAD 2: Retenciones
    if normalize_account('2.1.3.04.1.006') in cuentas_del_asiento: return "Grupo 9: Retenciones - IVA"
    if normalize_account('2.1.3.01.1.012') in cuentas_del_asiento: return "Grupo 9: Retenciones - ISLR"
    if normalize_account('7.1.3.04.1.004') in cuentas_del_asiento: return "Grupo 9: Retenciones - Municipal"
    
    # PRIORIDAD 3: Cobranzas
    is_cobranza = 'RECIBO DE COBRANZA' in referencia_completa or 'TEF' in fuente_completa
    if is_cobranza:
        if is_reverso_check: return "Grupo 8: Cobranzas"
        if normalize_account('6.1.1.12.1.001') in cuentas_del_asiento: return "Grupo 8: Cobranzas - Con Diferencial Cambiario"
        if normalize_account('1.1.1.04.6.003') in cuentas_del_asiento: return "Grupo 8: Cobranzas - Fondos por Depositar"
        if not CUENTAS_BANCO.isdisjoint(cuentas_del_asiento):
            if 'TEF' in fuente_completa: return "Grupo 8: Cobranzas - TEF (Bancos)"
            else: return "Grupo 8: Cobranzas - Recibos (Bancos)"
        return "Grupo 8: Cobranzas - Otros"

    # PRIORIDAD 4: Ingresos Varios (Grupo 6)
    if normalize_account('6.1.1.19.1.001') in cuentas_del_asiento:
        if is_reverso_check: return "Grupo 6: Ingresos Varios"
        keywords_limpieza = {'LIMPIEZA', 'LIMPIEZAS', 'SALDO', 'SALDOS', 'HISTORICO'}
        if not keywords_limpieza.isdisjoint(referencia_limpia_palabras):
            if (asiento_group['Monto_USD'].abs() <= 5).all(): return "Grupo 6: Ingresos Varios - Limpieza (<= $5)"
            else: return "Grupo 6: Ingresos Varios - Limpieza (> $5)"
        else: return "Grupo 6: Ingresos Varios - Otros"

    # PRIORIDAD 5: Traspasos vs. Devoluciones (Grupo 10 y 7)
    if normalize_account('4.1.1.21.4.001') in cuentas_del_asiento:
        if 'TRASPASO' in referencia_completa and abs(asiento_group['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD: return "Grupo 10: Traspasos"
        if is_reverso_check: return "Grupo 7: Devoluciones y Rebajas"
        keywords_limpieza_dev = {'LIMPIEZA', 'LIMPIEZAS', 'SALDO', 'SALDOS', 'HISTORICO', 'AJUSTE'}
        if not keywords_limpieza_dev.isdisjoint(referencia_limpia_palabras):
            if (asiento_group['Monto_USD'].abs() <= 5).all(): return "Grupo 7: Devoluciones y Rebajas - Limpieza (<= $5)"
            else: return "Grupo 7: Devoluciones y Rebajas - Limpieza (> $5)"
        else: return "Grupo 7: Devoluciones y Rebajas - Otros Ajustes"
            
    # Resto de prioridades
    if normalize_account('7.1.3.06.1.998') in cuentas_del_asiento: return "Grupo 12: Perdida p/Venta o Retiro Activo ND"
    if normalize_account('7.1.3.45.1.997') in cuentas_del_asiento: return "Grupo 1: Acarreos y Fletes Recuperados"
    if normalize_account('6.1.1.12.1.001') in cuentas_del_asiento: return "Grupo 2: Diferencial Cambiario"
    if normalize_account('7.1.3.19.1.012') in cuentas_del_asiento: return "Grupo 4: Gastos de Ventas"
    if normalize_account('2.1.2.05.1.108') in cuentas_del_asiento: return "Grupo 5: Haberes de Clientes"

    return "No Clasificado"

def _clasificar_asiento_paquete_cc(asiento_group):
    """
    Función principal de clasificación que implementa una capa global para manejar Reversos.
    """
    cuentas_del_asiento = set(asiento_group['Cuenta Contable Norm'])
    referencia_completa = ' '.join(asiento_group['Referencia'].astype(str).unique()).upper()
    fuente_completa = ' '.join(asiento_group['Fuente'].astype(str).unique()).upper()
    referencia_limpia_palabras = set(re.sub(r'[^\w\s]', '', referencia_completa).split())

    # CAPA 1: Detección de Reversos
    if 'REVERSO' in referencia_completa or 'REV' in referencia_limpia_palabras:
        # Si es un reverso, llamamos a la lógica base en modo "reverso" para saber de qué tipo es
        base_group = _get_base_classification(asiento_group, cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, is_reverso_check=True)
        
        if base_group != "No Clasificado":
            # Formateamos el resultado como un subgrupo de Reverso
            parts = base_group.split(':', 1)
            group_number = parts[0].strip()
            description = parts[1].split('-')[0].strip() # Tomamos la descripción principal (ej: "N/C", "Acarreos y Fletes")
            return f"{group_number}: Reversos - {description}"
        else:
            return "Grupo 11: Reversos No Identificados"
            
    # CAPA 2: Si no es un reverso, aplicar la lógica de clasificación estándar
    return _get_base_classification(asiento_group, cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, is_reverso_check=False)


def _validar_asiento(asiento_group):
    """
    Recibe un asiento completo (ya clasificado) y aplica las reglas de negocio
    para determinar si está Conciliado o tiene una Incidencia.
    """
    grupo = asiento_group['Grupo'].iloc[0]
    
    if grupo.startswith("Grupo 1:"):
        fletes_lines = asiento_group[asiento_group['Cuenta Contable Norm'] == normalize_account('7.1.3.45.1.997')]
        if not fletes_lines['Referencia'].str.contains('FLETE', case=False, na=False).all():
            return "Incidencia: Referencia sin 'FLETE' encontrada."
            
    elif grupo.startswith("Grupo 2:"):
        diff_lines = asiento_group[asiento_group['Cuenta Contable Norm'] == normalize_account('6.1.1.12.1.001')]
        keywords = ['DIFERENCIAL', 'DIFERENCIA EN CAMBIO', 'DIF CAMBIARIO']
        if not diff_lines['Referencia'].str.contains('|'.join(keywords), case=False, na=False).all():
            return "Incidencia: Referencia sin palabra clave de diferencial."
            
    elif grupo.startswith("Grupo 6:"):
        if (asiento_group['Monto_USD'].abs() > 25).any():
            return "Incidencia: Movimiento mayor a $25 encontrado."
            
    elif grupo.startswith("Grupo 7:"):
        if (asiento_group['Monto_USD'].abs() > 5).any():
            return "Incidencia: Movimiento mayor a $5 encontrado."

    elif grupo.startswith("Grupo 9:"):
        referencia_str = asiento_group['Referencia'].iloc[0]
        if not re.fullmatch(r'\d+', str(referencia_str).strip()):
            return "Incidencia: Referencia no es un número de comprobante válido."

    elif grupo.startswith("Grupo 10:"):
        if not np.isclose(asiento_group['Monto_USD'].sum(), 0, atol=TOLERANCIA_MAX_USD):
            return "Incidencia: El traspaso no suma cero."
    
    return "Conciliado"

def run_analysis_paquete_cc(df_diario, log_messages):
    """
    Función principal que orquesta la clasificación Y la validación de asientos.
    """
    log_messages.append("--- INICIANDO ANÁLISIS Y VALIDACIÓN DE PAQUETE CC ---")
    df = df_diario.copy()
    df['Cuenta Contable Norm'] = df['Cuenta Contable'].apply(normalize_account)
    df['Monto_USD'] = (df['Débito Dolar'] - df['Crédito Dolar']).round(2)
    
    resultados_clasificacion = {}
    grupos_de_asientos = df.groupby('Asiento')
    asientos_con_cuentas_nuevas = 0
    for asiento_id, asiento_group in grupos_de_asientos:
        cuentas_del_asiento_norm = set(asiento_group['Cuenta Contable Norm'])
        if not cuentas_del_asiento_norm.issubset(CUENTAS_CONOCIDAS):
            grupo_asignado = "Grupo 11: Cuentas No Identificadas"
            asientos_con_cuentas_nuevas += 1
        else:
            grupo_asignado = _clasificar_asiento_paquete_cc(asiento_group)
        resultados_clasificacion[asiento_id] = grupo_asignado
        
    df['Grupo'] = df['Asiento'].map(resultados_clasificacion)
    log_messages.append("✔️ Clasificación de asientos completada.")
    
    resultados_validacion = {}
    grupos_de_asientos_clasificados = df.groupby('Asiento')
    for asiento_id, asiento_group in grupos_de_asientos_clasificados:
        estado_validacion = _validar_asiento(asiento_group)
        resultados_validacion[asiento_id] = estado_validacion
        
    df['Estado'] = df['Asiento'].map(resultados_validacion)
    log_messages.append("✔️ Validación de reglas de negocio completada.")

    if asientos_con_cuentas_nuevas > 0:
        log_messages.append(f"⚠️ Se encontraron {asientos_con_cuentas_nuevas} asientos con cuentas no registradas.")
    
    df_final = df.drop(columns=['Cuenta Contable Norm']).sort_values(by=['Grupo', 'Asiento', 'Monto_USD'], ascending=[True, True, False])
    log_messages.append("--- ANÁLISIS FINALIZADO CON ÉXITO ---")
    return df_final
