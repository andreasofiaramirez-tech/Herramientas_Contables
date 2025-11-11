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

# ==============================================================================
# LÓGICAS PARA LA HERRAMIENTA DE RELACIONES DE RETENCIONES
# ==============================================================================

# --- Funciones Auxiliares ---
def _limpiar_nombre_columna_retenciones(col_name):
    s = str(col_name).strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^A-Z0-9]', '', s.upper())

def _normalizar_valor(valor):
    if pd.isna(valor): return ''
    val_str = str(valor).strip().upper()
    if 'E+' in val_str:
        try: val_str = f"{int(float(val_str)):d}"
        except (ValueError, TypeError): pass
    val_str = re.sub(r'[^A-Z0-9]', '', val_str)
    if val_str.startswith('J') and len(val_str) in [9, 10]: val_str = val_str[1:]
    val_str = re.sub(r'^0+', '', val_str)
    return val_str

# --- LÓGICA DE CONCILIACIÓN REFACTORIZADA ---

def run_conciliation_retenciones(file_cp, file_cg, file_iva, file_islr, file_mun, log_messages):
    log_messages.append("--- INICIANDO PROCESO DE CONCILIACIÓN DE RETENCIONES (VERSIÓN OPTIMIZADA) ---")
    try:
        # --- 1. CARGA Y PREPARACIÓN INICIAL ---
        log_messages.append("Paso 1: Cargando y estandarizando archivos...")
        df_cp = pd.read_excel(file_cp, header=4, dtype=str)
        df_cg = pd.read_excel(file_cg, header=0, dtype=str)
        df_galac_iva = pd.read_excel(file_iva, header=4, dtype=str)
        df_galac_islr = pd.read_excel(file_islr, header=8, dtype=str)
        df_galac_mun = pd.read_excel(file_mun, header=8, dtype=str)
        if not df_galac_islr.empty and df_galac_islr.columns[0].upper().strip() == 'NRO':
            df_galac_islr = df_galac_islr[pd.to_numeric(df_galac_islr.iloc[:, 0], errors='coerce').notna()]
            
        CUENTAS_MAP = {'IVA': '2111101004', 'ISLR': '2111101005', 'MUNICIPAL': '2111101006'}

        # Estandarización de columnas
        dfs = {'cp': df_cp, 'cg': df_cg, 'iva': df_galac_iva, 'islr': df_galac_islr, 'mun': df_galac_mun}
        for name, df in dfs.items():
            df.columns = [_limpiar_nombre_columna_retenciones(c) for c in df.columns]

        synonyms_map = {
            'MONTO': ['MONTOTOTAL', 'MONTOBS', 'MONTO', 'IVARETENIDO', 'MONTORETENIDO', 'VALOR'], 'RIF': ['RIF', 'PROVEEDOR', 'RIFPROV', 'RIFPROVEEDOR', 'NUMERORIF', 'NIT'], 'COMPROBANTE': ['COMPROBANTE', 'NOCOMPROBANTE', 'NREFERENCIA', 'NUMERO'], 'FACTURA': ['FACTURA', 'NDOCUMENTO', 'NUMERODEFACTURA'], 'FECHA': ['FECHA', 'FECHARET', 'OPERACION'], 'CREDITOVES': ['CREDITOVES', 'CREDITO', 'CREDITOBS'], 'ASIENTO': ['ASIENTO', 'ASIENTOCONTABLE'], 'CUENTACONTABLE': ['CUENTACONTABLE', 'CUENTA']
        }
        def estandarizar_columnas(df):
            for standard_name, synonyms in synonyms_map.items():
                col_encontrada = next((s for s in synonyms if s in df.columns), None)
                if col_encontrada and standard_name != col_encontrada:
                    df.rename(columns={col_encontrada: standard_name}, inplace=True)

        for df in dfs.values(): estandarizar_columnas(df)

        # Consolidar GALAC
        df_galac_iva['TIPO'] = 'IVA'; df_galac_islr['TIPO'] = 'ISLR'; df_galac_mun['TIPO'] = 'MUNICIPAL'
        df_galac_full = pd.concat([df_galac_iva, df_galac_islr, df_galac_mun], ignore_index=True)
        df_galac_full['galac_unique_id'] = df_galac_full.index
        columnas_criticas_galac = ['RIF', 'COMPROBANTE', 'MONTO']
        df_galac_full.dropna(subset=[col for col in columnas_criticas_galac if col in df_galac_full.columns], inplace=True)

        # --- 2. NORMALIZACIÓN PROFUNDA Y CREACIÓN DE CLAVES ---
        log_messages.append("Paso 2: Normalizando datos y creando claves de conciliación...")
        
        # Normalizar claves de texto
        for df in [df_cp, df_galac_full]:
            df['RIF_norm'] = df['RIF'].apply(_normalizar_valor)
            df['COMPROBANTE_norm'] = df.get('COMPROBANTE', pd.Series(dtype=str)).apply(_normalizar_valor)
            df['FACTURA_norm'] = df.get('FACTURA', pd.Series(dtype=str)).apply(_normalizar_valor)
            
            if 'APLICACION' in df.columns:
                mask_factura_vacia = (df['FACTURA_norm'] == '') | (df['FACTURA_norm'].isna())
                df.loc[mask_factura_vacia, 'FACTURA_norm'] = df.loc[mask_factura_vacia, 'APLICACION'].str.upper().str.extract(r'FACT\s*N?[°º]?\s*(\S+)')[0].apply(_normalizar_valor)


        # Normalizar y convertir montos y fechas
        for col in ['MONTO', 'CREDITOVES']:
            for df in [df_cp, df_cg, df_galac_full]:
                if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Normalizar fechas
        for df in [df_cp, df_galac_full]:
            if 'FECHA' in df.columns: df['FECHA_norm'] = pd.to_datetime(df['FECHA'], errors='coerce')

        # Determinar el tipo de retención declarado en CP
        df_cp['SUBTIPO_DECLARADO'] = 'OTRO'
        df_cp.loc[df_cp['SUBTIPO'].str.contains('IVA', na=False), 'SUBTIPO_DECLARADO'] = 'IVA'
        df_cp.loc[df_cp['SUBTIPO'].str.contains('ISLR', na=False), 'SUBTIPO_DECLARADO'] = 'ISLR'
        df_cp.loc[df_cp['SUBTIPO'].str.contains('MUNICIPAL', na=False), 'SUBTIPO_DECLARADO'] = 'MUNICIPAL'
        
        # Guardar un índice único para CP para poder reconstruir el orden original
        df_cp['CP_INDEX'] = df_cp.index
        
        # --- 3. PROCESO DE CONCILIACIÓN VECTORIZADO ---
        log_messages.append("Paso 3: Ejecutando conciliación principal (CP vs GALAC)...")
        
        # --- 3a. Conciliación para IVA (Lógica por Comprobante) ---
        cp_iva = df_cp[df_cp['SUBTIPO_DECLARADO'] == 'IVA'].copy()
        galac_iva = df_galac_full[df_galac_full['TIPO'] == 'IVA'].copy()
        
        # Se cruza por RIF y Comprobante, que es la clave correcta según los datos.
        conciliado_iva = pd.merge(cp_iva, galac_iva.add_suffix('_galac'),
                                      left_on=['RIF_norm', 'COMPROBANTE_norm'],
                                      right_on=['RIF_norm_galac', 'COMPROBANTE_norm_galac'],
                                      how='left')
        
        monto_match_iva = np.isclose(conciliado_iva['MONTO'], conciliado_iva['MONTO_galac'])
        date_match_iva = (conciliado_iva['FECHA_norm'] - conciliado_iva['FECHA_norm_galac']).abs() <= pd.Timedelta(days=30)
        conciliado_iva['CP_Vs_Galac'] = np.where(monto_match_iva & date_match_iva, 'Sí', 'No Encontrado en GALAC')
        

        # --- 3b. Conciliación para MUNICIPAL (Lógica por Factura) ---
        cp_mun = df_cp[df_cp['SUBTIPO_DECLARADO'] == 'MUNICIPAL'].copy()
        galac_mun = df_galac_full[df_galac_full['TIPO'] == 'MUNICIPAL'].copy()
        
        # Se cruza por RIF y Factura, que es la clave correcta para este tipo.
        conciliado_mun = pd.merge(cp_mun, galac_mun.add_suffix('_galac'),
                                      left_on=['RIF_norm', 'FACTURA_norm'],
                                      right_on=['RIF_norm_galac', 'FACTURA_norm_galac'],
                                      how='left')
        
        monto_match_mun = np.isclose(conciliado_mun['MONTO'], conciliado_mun['MONTO_galac'])
        date_match_mun = (conciliado_mun['FECHA_norm'] - conciliado_mun['FECHA_norm_galac']).abs() <= pd.Timedelta(days=30)
        conciliado_mun['CP_Vs_Galac'] = np.where(monto_match_mun & date_match_mun, 'Sí', 'No Encontrado en GALAC')


        # --- 3c. Conciliación para ISLR (Lógica de Agrupación por Comprobante) ---
        cp_islr = df_cp[df_cp['SUBTIPO_DECLARADO'] == 'ISLR'].copy()
        
        # Agrupa el reporte de GALAC para sumar los montos por RIF y Comprobante.
        galac_islr_grouped = df_galac_full[df_galac_full['TIPO'] == 'ISLR'].groupby(['RIF_norm', 'COMPROBANTE_norm'], as_index=False).agg(
            MONTO_galac_sum=('MONTO', 'sum'),
            FECHA_galac_max=('FECHA_norm', 'max')
        )
        
        conciliado_islr = pd.merge(cp_islr, galac_islr_grouped, on=['RIF_norm', 'COMPROBANTE_norm'], how='left')
        
        monto_match_islr = np.isclose(conciliado_islr['MONTO'], conciliado_islr['MONTO_galac_sum'])
        date_match_islr = (conciliado_islr['FECHA_norm'] - conciliado_islr['FECHA_galac_max']).abs() <= pd.Timedelta(days=30)
        conciliado_islr['CP_Vs_Galac'] = np.where(monto_match_islr & date_match_islr, 'Sí', 'No Encontrado en GALAC')

        # --- 4. CONSOLIDACIÓN Y VERIFICACIÓN FINAL ---
        log_messages.append("Paso 4: Consolidando resultados y verificando contra CG...")
        
        # Se concatenan los tres DataFrames correctos.
        df_cp_results = pd.concat([conciliado_iva, conciliado_mun, conciliado_islr], ignore_index=True)
        
        df_cp_results.loc[df_cp_results['APLICACION'].str.contains('ANULADO', case=False, na=False), 'CP_Vs_Galac'] = 'No Aplica (Anulado)'
        
        df_cg_grouped = df_cg.groupby(['ASIENTO', 'CUENTACONTABLE'])['CREDITOVES'].sum().reset_index()
        df_cp_results['CUENTA_ESPERADA'] = df_cp_results['SUBTIPO_DECLARADO'].map(CUENTAS_MAP)
        df_cp_final = pd.merge(df_cp_results, df_cg_grouped, left_on=['ASIENTO', 'CUENTA_ESPERADA'], right_on=['ASIENTO', 'CUENTACONTABLE'], how='left')
        
        df_cp_final['Asiento_en_CG'] = np.where(df_cp_final['CUENTACONTABLE'].notna(), 'Sí', 'No')
        monto_coincide_cg = np.isclose(df_cp_final['MONTO'].abs(), df_cp_final['CREDITOVES'].fillna(0))
        df_cp_final['Monto_coincide_CG'] = np.select([df_cp_final['Asiento_en_CG'] == 'No', monto_coincide_cg], ['No Aplica', 'Sí'], default='No')
        
        # --- 5. VERIFICACIÓN CONTRA CONTABILIDAD GENERAL (CG) ---
        log_messages.append("Paso 5: Verificando contra Contabilidad General...")

        # Preparar CG: agrupar por asiento y cuenta para tener un total por partida
        df_cg_grouped = df_cg.groupby(['ASIENTO', 'CUENTACONTABLE'])['CREDITOVES'].sum().reset_index()

        # Añadir la cuenta contable esperada a los resultados de CP
        df_cp_results['CUENTA_ESPERADA'] = df_cp_results['SUBTIPO_DECLARADO'].map(CUENTAS_MAP)
        
        # Fusionar con los datos de CG
        df_cp_final = pd.merge(
            df_cp_results,
            df_cg_grouped,
            left_on=['ASIENTO', 'CUENTA_ESPERADA'],
            right_on=['ASIENTO', 'CUENTACONTABLE'],
            how='left',
            suffixes=('', '_cg')
        )
        
        # Evaluar resultados de la verificación con CG
        df_cp_final['Asiento_en_CG'] = np.where(df_cp_final['CUENTACONTABLE'].notna(), 'Sí', 'No')
        monto_coincide_cg = np.isclose(df_cp_final['MONTO'].abs(), df_cp_final['CREDITOVES'].fillna(0))
        df_cp_final['Monto_coincide_CG'] = np.select(
            [df_cp_final['Asiento_en_CG'] == 'No', monto_coincide_cg],
            ['No Aplica', 'Sí'],
            default='No'
        )
        
        # --- 6. PREPARACIÓN DEL REPORTE FINAL ---
        log_messages.append("Paso 6: Generando reporte final...")
        
        # Reordenar para que se parezca al original
        df_cp_final.sort_values('CP_INDEX', inplace=True)
        # Limpiar columnas auxiliares antes de generar el reporte
        
        # --- Cálculo de registros de GALAC no encontrados en CP ---
        # Obtenemos los IDs únicos de las filas de Galac que sí se encontraron
        # La columna puede tener el sufijo '_galac' o '_islr' etc. Buscamos cualquiera que termine en 'galac_unique_id'
        id_col_name = next((col for col in df_cp_final.columns if 'galac_unique_id' in col), None)
        if id_col_name:
            ids_encontrados = df_cp_final[id_col_name].dropna().unique()
            # Filtramos el df_galac_full original para quedarnos con las filas que NO están en ids_encontrados
            df_galac_no_cp = df_galac_full[~df_galac_full['galac_unique_id'].isin(ids_encontrados)].copy()
        else:
            df_galac_no_cp = pd.DataFrame() # Creamos un DF vacío si no se encontró ninguna coincidencia

        reporte_bytes = generar_reporte_retenciones(df_cp_final, df_galac_no_cp, df_cg, CUENTAS_MAP)
        
        log_messages.append("¡Proceso de conciliación de retenciones completado con éxito!")
        return reporte_bytes

    except Exception as e:
        log_messages.append(f"❌ ERROR CRÍTICO en la conciliación de retenciones: {e}")
        import traceback
        log_messages.append(traceback.format_exc())
        return None
