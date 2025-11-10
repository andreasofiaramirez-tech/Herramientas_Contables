# logic.py

import pandas as pd
import numpy as np
import re
import unicodedata
from itertools import combinations

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
    """
    (Versión CORREGIDA) Busca y concilia pares 1-a-1 de débito/crédito cuyo monto absoluto es EXACTAMENTE el mismo.
    Ahora maneja múltiples pares del mismo monto.
    """
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
    """
    Busca y concilia pares 1-a-1 de débito/crédito DENTRO del grupo 'BANCO A BANCO'.
    Maneja múltiples pares del mismo monto de forma robusta.
    """
    log_messages.append("\n--- FASE PARES BANCO A BANCO (USD) ---")
    total_conciliados = 0
    # Seleccionamos solo los movimientos pendientes que pertenecen al grupo de banco
    df_pendientes = df.loc[(~df['Conciliado']) & (df['Clave_Grupo'] == 'GRUPO_BANCO')].copy()
    
    if df_pendientes.empty:
        return 0

    df_pendientes['Monto_Abs'] = df_pendientes['Monto_USD'].abs()
    
    # Agrupamos por el monto absoluto para encontrar pares potenciales
    grupos_por_monto = df_pendientes.groupby('Monto_Abs')
    
    for monto, grupo in grupos_por_monto:
        if len(grupo) < 2:
            continue
            
        debitos = grupo[grupo['Monto_USD'] > 0].index.to_list()
        creditos = grupo[grupo['Monto_USD'] < 0].index.to_list()
        
        # Determinamos cuántos pares podemos formar
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
    """
    Versión OPTIMIZADA que busca grupos complejos (1-vs-2 y 2-vs-1) utilizando
    diccionarios para búsquedas O(1) en lugar de fuerza bruta combinatoria.
    """
    log_messages.append("\n--- FASE GRUPOS COMPLEJOS OPTIMIZADA (USD) ---")
    
    pendientes = df.loc[~df['Conciliado']]
    LIMITE_MOVIMIENTOS = 500  # Podemos aumentar el límite con el algoritmo optimizado
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
    # Creamos un diccionario con la suma de cada par de créditos
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
        
        # Buscamos una suma cercana en nuestro diccionario
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
    # Actualizamos los movimientos disponibles
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
    """
    Versión OPTIMIZADA que busca pares 1-a-1 usando un self-merge de pandas,
    evitando bucles anidados de Python.
    """
    log_messages.append("\n--- FASE GLOBAL 1-a-1 (USD) ---")
    
    pendientes = df.loc[~df['Conciliado']].copy()
    if len(pendientes) < 2:
        return 0

    debitos = pendientes[pendientes['Monto_USD'] > 0].copy()
    creditos = pendientes[pendientes['Monto_USD'] < 0].copy()

    if debitos.empty or creditos.empty:
        return 0

    # Convertimos el índice en una columna para que el merge lo preserve.
    # Esto creará las columnas 'index_d' e 'index_c' que necesitamos.
    debitos.reset_index(inplace=True)
    creditos.reset_index(inplace=True)

    debitos['join_key'] = 1
    creditos['join_key'] = 1

    # Cruzamos todos los débitos con todos los créditos
    pares_potenciales = pd.merge(debitos, creditos, on='join_key', suffixes=('_d', '_c'))
    
    # Calculamos la diferencia absoluta entre los montos
    pares_potenciales['diferencia'] = abs(pares_potenciales['Monto_USD_d'] + pares_potenciales['Monto_USD_c'])
    
    # Filtramos solo aquellos pares que están dentro de la tolerancia
    pares_validos = pares_potenciales[pares_potenciales['diferencia'] <= TOLERANCIA_MAX_USD].copy()
    
    # Ordenamos por la menor diferencia para encontrar los mejores matches primero
    pares_validos.sort_values(by='diferencia', inplace=True)
    
    # Eliminamos duplicados para que cada movimiento se use solo una vez
    # Ahora 'index_d' e 'index_c' existirán y esta línea funcionará.
    pares_finales = pares_validos.drop_duplicates(subset=['index_d'], keep='first')
    pares_finales = pares_finales.drop_duplicates(subset=['index_c'], keep='first')
    
    total_conciliados = 0
    if not pares_finales.empty:
        indices_d = pares_finales['index_d'].tolist()
        indices_c = pares_finales['index_c'].tolist()
        
        # Aplicamos la conciliación en lote
        df.loc[indices_d + indices_c, 'Conciliado'] = True
        
        # Actualizamos los grupos de conciliación
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
    """
    (Versión MEJORADA) Normaliza los datos para la conciliación de proveedores,
    utilizando el NIT/RIF como la clave principal para una mayor precisión.
    """
    df_copy = df.copy()
    
    # --- LÓGICA MEJORADA PARA ENCONTRAR Y USAR EL NIT/RIF ---
    
    # 1. Buscamos la columna del identificador único (NIT o RIF)
    nit_col_name = None
    for col in df_copy.columns:
        if str(col).strip().upper() in ['NIT', 'RIF']:
            nit_col_name = col
            break
            
    if nit_col_name:
        log_messages.append(f"✔️ Se encontró la columna de identificador fiscal ('{nit_col_name}') y se usará como clave principal.")
        # 2. Creamos la clave usando el NIT/RIF normalizado (sin guiones, puntos, etc.)
        df_copy['Clave_Proveedor'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró la columna 'NIT' o 'RIF'. Se recurrirá a usar 'Nombre del Proveedor', lo cual es menos preciso.")
        df_copy['Clave_Proveedor'] = df_copy['Nombre del Proveedor'].astype(str).str.strip().str.upper()

    # La lógica para extraer la clave de la factura/documento no cambia
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
    """Orquesta el proceso completo de conciliación para Devoluciones a Proveedores."""
    log_messages.append("\n--- INICIANDO LÓGICA DE DEVOLUCIONES A PROVEEDORES (USD) ---")
    
    # Esta llamada ahora usará la lógica mejorada con NIT/RIF
    df = normalizar_datos_proveedores(df, log_messages) 
    
    total_conciliados = 0
    df_procesable = df.loc[(~df['Conciliado']) & (df['Clave_Proveedor'].notna()) & (df['Clave_Comp'].notna())]
    
    # El groupby ahora usará la clave basada en NIT, que es mucho más precisa
    grupos = df_procesable.groupby(['Clave_Proveedor', 'Clave_Comp'])
    
    log_messages.append(f"ℹ️ Se encontraron {len(grupos)} grupos de Proveedor/COMP para analizar.")
    for (proveedor_clave, comp), grupo in grupos:
        if abs(round(grupo['Monto_USD'].sum(), 2)) <= TOLERANCIA_MAX_USD:
            indices = grupo.index
            # Usamos la clave del proveedor (el NIT) en el ID del grupo para mayor claridad
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
    """
    Clasifica los movimientos de la cuenta de viajes en 'IMPUESTOS' o 'VIATICOS'
    basado en palabras clave en la columna 'Referencia'.
    """
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

    # Creamos la nueva columna 'Tipo' que se usará en el reporte
    df['Tipo'] = df['Referencia'].apply(clasificar_tipo)
    
    # También necesitamos normalizar el NIT para usarlo como clave
    nit_col_name = next((col for col in df.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df['NIT_Normalizado'] = df[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró la columna 'NIT' o 'RIF'. La conciliación puede no ser precisa.")
        df['NIT_Normalizado'] = 'SIN_NIT'
        
    return df

def conciliar_pares_exactos_por_nit_viajes(df, log_messages):
    """
    Busca y concilia pares 1-a-1 de débito/crédito que se anulan mutuamente (suma cero)
    y que pertenecen al MISMO NIT_Normalizado.
    """
    log_messages.append("\n--- FASE 1: Búsqueda de Pares Exactos por NIT ---")
    total_conciliados = 0
    
    # Agrupamos por NIT y por el valor absoluto del monto para encontrar pares potenciales
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
                
                # Doble chequeo de que la suma es cero (o muy cercana)
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
    """
    Para cada NIT, busca combinaciones de movimientos pendientes que sumen cero.
    """
    log_messages.append("\n--- FASE 2: Búsqueda de Grupos por NIT ---")
    total_conciliados_fase = 0
    
    df_pendientes = df.loc[~df['Conciliado']]
    grupos_por_nit = df_pendientes.groupby('NIT_Normalizado')
    
    for nit, grupo in grupos_por_nit:
        if nit == 'SIN_NIT' or len(grupo) < 2:
            continue
            
        # Si todo el grupo de un NIT ya suma cero, lo conciliamos
        if abs(grupo['Monto_BS'].sum()) <= TOLERANCIA_MAX_BS:
            indices = grupo.index
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, f'GRUPO_TOTAL_NIT_{nit}']
            total_conciliados_fase += len(indices)
            log_messages.append(f"✔️ Conciliado grupo completo para NIT {nit} ({len(indices)} movimientos).")
            continue # Pasamos al siguiente NIT

        # Si no, buscamos sub-combinaciones (lógica similar a la de USD)
        LIMITE_COMBINACION = 10 # Límite de seguridad para evitar congelamiento
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
    
    # PASO 1: Conciliaciones automáticas
    conciliar_automaticos_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.1, text="Fase 1/6: Conciliaciones automáticas completada.")
    
    # PASO 2: Grupos por referencia ESPECÍFICA (Ignora 'BANCO A BANCO')
    conciliar_grupos_por_referencia_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase 2/6: Grupos por referencia específica completada.")
    
    # --- NUEVO PASO ESTRATÉGICO ---
    # PASO 3: Lógica dedicada para encontrar pares dentro de 'BANCO A BANCO'
    conciliar_pares_banco_a_banco_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.35, text="Fase 3/6: Pares 'Banco a Banco' completada.")
    
    # PASO 4: Pares globales con monto EXACTO (Ahora más robusto)
    conciliar_pares_globales_exactos_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.5, text="Fase 4/6: Pares globales exactos completada.")
    
    # PASO 5: Búsqueda de pares globales con TOLERANCIA (versión optimizada)
    conciliar_pares_globales_remanentes_usd(df, log_messages)
    if progress_bar: progress_bar.progress(0.65, text="Fase 5/6: Búsqueda de pares con tolerancia completada.")

    # PASO 6: Búsqueda de grupos complejos (versión optimizada)
    conciliar_grupos_complejos_usd(df, log_messages, progress_bar)
    if progress_bar: progress_bar.progress(0.9, text="Fase 6/6: Búsqueda de grupos complejos completada.")
    
    conciliar_gran_total_final_usd(df, log_messages)
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

def run_conciliation_devoluciones_proveedores(df, log_messages):
    """Orquesta el proceso completo de conciliación para Devoluciones a Proveedores."""
    log_messages.append("\n--- INICIANDO LÓGICA DE DEVOLUCIONES A PROVEEDORES (USD) ---")
    df = normalizar_datos_proveedores(df)
    log_messages.append("✔️ Datos normalizados: Claves de Proveedor y COMP generadas.")
    total_conciliados = 0
    df_procesable = df.loc[(~df['Conciliado']) & (df['Clave_Proveedor'].notna()) & (df['Clave_Comp'].notna())]
    grupos = df_procesable.groupby(['Clave_Proveedor', 'Clave_Comp'])
    log_messages.append(f"ℹ️ Se encontraron {len(grupos)} grupos de Proveedor/COMP para analizar.")
    for (proveedor, comp), grupo in grupos:
        if abs(round(grupo['Monto_USD'].sum(), 2)) <= TOLERANCIA_MAX_USD:
            indices = grupo.index
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, f"PROV_{proveedor}_{comp}"]
            total_conciliados += len(indices)
    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación por Proveedor/COMP: {total_conciliados} movimientos conciliados.")
    else:
        log_messages.append("ℹ️ No se encontraron conciliaciones automáticas por Proveedor/COMP.")
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

def run_conciliation_viajes(df, log_messages, progress_bar=None):
    """
    Orquesta el proceso completo de conciliación para la cuenta de Viajes.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE CUENTAS DE VIAJES (BS) ---")
    
    # Paso 0: Clasificar y preparar datos
    df = normalizar_referencia_viajes(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase de Normalización completada.")
    
    # Paso 1: Buscar pares exactos por NIT (alta precisión)
    conciliar_pares_exactos_por_nit_viajes(df, log_messages)
    if progress_bar: progress_bar.progress(0.5, text="Fase 1/2: Búsqueda de pares exactos completada.")
    
    # Paso 2: Buscar grupos complejos por NIT (baja precisión)
    conciliar_grupos_por_nit_viajes(df, log_messages)
    if progress_bar: progress_bar.progress(0.9, text="Fase 2/2: Búsqueda de grupos complejos completada.")
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# ==============================================================================
# LÓGICAS PARA LA HERRAMIENTA DE RELACIONES DE RETENCIONES
# ==============================================================================

# --- Funciones Auxiliares Específicas de Retenciones ---

def _limpiar_nombre_columna_retenciones(col_name):
    """Limpia y normaliza un nombre de columna, manejando acentos."""
    s = str(col_name).strip()
    # Normaliza para remover acentos (ej: 'Crédito' -> 'Credito')
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    # Convierte a mayúsculas y elimina caracteres no alfanuméricos
    return re.sub(r'[^A-Z0-9]', '', s.upper())
    
def _normalizar_valor(valor):
    """Normaliza RIF, comprobantes y facturas para una comparación precisa."""
    if pd.isna(valor):
        return ''
    # Convierte a string, elimina espacios, guiones, puntos, y quita ceros a la izquierda
    val_str = str(valor).strip().upper().replace('.', '').replace('-', '')
    val_str = re.sub(r'^0+', '', val_str) # Elimina ceros iniciales
    if val_str.startswith('J'): # Quita la 'J' de los RIFs si existe
        val_str = val_str[1:]
    return val_str

# --- Función Maestra de Conciliación de Retenciones ---

# EN logic.py, REEMPLAZA LA FUNCIÓN COMPLETA CON ESTA VERSIÓN FINAL Y DEFINITIVA

def run_conciliation_retenciones(file_cp, file_cg, file_iva, file_islr, file_mun, log_messages):
    """
    Función principal que encapsula toda la lógica de conciliación de retenciones
    para ser ejecutada desde la interfaz de Streamlit.
    """
    log_messages.append("--- INICIANDO PROCESO DE CONCILIACIÓN DE RETENCIONES ---")

    try:
        # --- 1. CARGA Y PREPARACIÓN DE DATOS ---
        log_messages.append("Cargando archivos de entrada...")
        
        # --- CORRECCIÓN CLAVE: Ajuste de los encabezados según las imágenes ---
        df_cp = pd.read_excel(file_cp, header=4)
        df_cg = pd.read_excel(file_cg, header=0)
        df_galac_iva = pd.read_excel(file_iva, header=4)    # Fila 5 en Excel
        df_galac_islr = pd.read_excel(file_islr, header=8)   # Fila 9 en Excel
        df_galac_mun = pd.read_excel(file_mun, header=8)     # Fila 9 en Excel

        CUENTAS_MAP = {'IVA': '2111101004', 'ISLR': '2111101005', 'MUNICIPAL': '2111101006'}

        log_messages.append("Limpiando y estandarizando datos...")
        # Limpiamos los nombres de todas las columnas primero
        df_cp.columns = [_limpiar_nombre_columna_retenciones(c) for c in df_cp.columns]
        df_cg.columns = [_limpiar_nombre_columna_retenciones(c) for c in df_cg.columns]
        df_galac_iva.columns = [_limpiar_nombre_columna_retenciones(c) for c in df_galac_iva.columns]
        df_galac_islr.columns = [_limpiar_nombre_columna_retenciones(c) for c in df_galac_islr.columns]
        df_galac_mun.columns = [_limpiar_nombre_columna_retenciones(c) for c in df_galac_mun.columns]
        
        # Búsqueda y renombrado robusto para CP y CG
        monto_synonyms_cp = ['MONTOTOTAL', 'MONTOBS', 'MONTO']
        if not any(col in df_cp.columns for col in monto_synonyms_cp): raise KeyError("No se pudo encontrar una columna de Monto en el archivo CP.")
        for col in monto_synonyms_cp:
            if col in df_cp.columns: df_cp.rename(columns={col: 'MONTO'}, inplace=True)
        
        credito_synonyms_cg = ['CREDITOVES', 'CREDITO', 'CREDITOBS']
        if not any(col in df_cg.columns for col in credito_synonyms_cg): raise KeyError("No se pudo encontrar una columna de Crédito en el archivo CG.")
        for col in credito_synonyms_cg:
            if col in df_cg.columns: df_cg.rename(columns={col: 'CREDITOVES'}, inplace=True)

        # --- CORRECCIÓN ROBUSTA PARA GALAC CON SINÓNIMOS EXACTOS DE LAS IMÁGENES ---
        galac_synonyms = {
            'MONTO': ['MONTO', 'IVARETENIDO', 'MONTORETENIDO', 'VALOR'],
            'RIF': ['RIF', 'RIFPROV', 'RIFPROVEEDOR', 'NUMERORIF'],
            'COMPROBANTE': ['COMPROBANTE', 'NOCOMPROBANTE', 'NREFERENCIA'],
            'FACTURA': ['FACTURA', 'NDOCUMENTO', 'NUMERODEFACTURA'],
            'FECHA': ['FECHA', 'FECHARET', 'FECHAOPERACION', 'FECHARETENCION']
        }

        for df_galac, nombre_archivo in [(df_galac_iva, 'IVA'), (df_galac_islr, 'ISLR'), (df_galac_mun, 'Municipal')]:
            for col_estandar, sinonimos in galac_synonyms.items():
                for sinonimo in sinonimos:
                    if sinonimo in df_galac.columns:
                        df_galac.rename(columns={sinonimo: col_estandar}, inplace=True)
                        break
        
        # Columnas que pueden no existir en todos los archivos
        if 'COMPROBANTE' not in df_galac_mun.columns: df_galac_mun['COMPROBANTE'] = ''
        if 'FACTURA' not in df_galac_iva.columns: df_galac_iva['FACTURA'] = ''
        
        df_galac_iva['TIPO'] = 'IVA'
        df_galac_islr['TIPO'] = 'ISLR'
        df_galac_mun['TIPO'] = 'MUNICIPAL'
        
        df_galac_full = pd.concat([df_galac_iva, df_galac_islr, df_galac_mun], ignore_index=True)
        
        # Normalización de valores clave
        for df in [df_cp, df_cg, df_galac_full]:
            if 'PROVEEDOR' in df.columns and 'RIF' not in df.columns: df.rename(columns={'PROVEEDOR': 'RIF'}, inplace=True)
            if 'RIF' in df.columns: df['RIF_norm'] = df['RIF'].apply(_normalizar_valor)
            if 'NIT' in df.columns: df['RIF_norm'] = df['NIT'].apply(_normalizar_valor)
            if 'NUMERO' in df.columns: df['COMPROBANTE_norm'] = df['NUMERO'].apply(_normalizar_valor)
            if 'COMPROBANTE' in df.columns: df['COMPROBANTE_norm'] = df['COMPROBANTE'].apply(_normalizar_valor)
            if 'FACTURA' in df.columns: df['FACTURA_norm'] = df['FACTURA'].apply(_normalizar_valor)

        # Conversión a numérico (ahora debería funcionar)
        df_cp['MONTO'] = pd.to_numeric(df_cp['MONTO'], errors='coerce').fillna(0)
        df_cg['CREDITOVES'] = pd.to_numeric(df_cg['CREDITOVES'], errors='coerce').fillna(0)
        df_galac_full['MONTO'] = pd.to_numeric(df_galac_full['MONTO'], errors='coerce').fillna(0)
        
        # (El resto de la función permanece exactamente igual...)

        # --- 2. LÓGICA DE CONCILIACIÓN ---
        log_messages.append("Iniciando auditoría en cascada por registro...")
        results = []
        indices_galac_encontrados = set()

        for index, row_cp in df_cp.iterrows():
            subtipo = str(row_cp.get('SUBTIPO', '')).upper()
            # Corrección para "Retención IVA" -> "IVA"
            if 'IVA' in subtipo: subtipo = 'IVA'
                
            rif_cp = row_cp.get('RIF_norm', '')
            comprobante_cp = row_cp.get('COMPROBANTE_norm', '')
            factura_cp = row_cp.get('FACTURA_norm', '')
            monto_cp = row_cp.get('MONTO', 0)
            
            resultado = {
                'CP_Vs_Galac': 'No Encontrado en GALAC',
                'Asiento_en_CG': 'No',
                'Monto_coincide_CG': 'No Aplica'
            }

            if "ANULADO" in str(row_cp.get('APLICACION', '')).upper():
                resultado['CP_Vs_Galac'] = 'No Aplica (Anulado)'
            else:
                df_galac_target = df_galac_full[df_galac_full['TIPO'] == subtipo]
                
                match = pd.Series(False, index=df_galac_target.index)
                if subtipo == 'IVA':
                    match = (df_galac_target['RIF_norm'] == rif_cp) & (df_galac_target['COMPROBANTE_norm'].str.endswith(comprobante_cp[-6:]))
                elif subtipo == 'ISLR':
                    match = (df_galac_target['RIF_norm'] == rif_cp) & (df_galac_target['COMPROBANTE_norm'] == comprobante_cp) & (df_galac_target['FACTURA_norm'] == factura_cp)
                elif subtipo == 'MUNICIPAL':
                    match = (df_galac_target['RIF_norm'] == rif_cp) & (df_galac_target['FACTURA_norm'] == factura_cp)
                
                found_df = df_galac_target[match]
                
                if not found_df.empty:
                    resultado['CP_Vs_Galac'] = 'Sí'
                    indices_galac_encontrados.update(found_df.index)
                else:
                    for otro_tipo in [t for t in ['IVA', 'ISLR', 'MUNICIPAL'] if t != subtipo]:
                        df_otro_galac = df_galac_full[df_galac_full['TIPO'] == otro_tipo]
                        if not df_otro_galac.empty:
                            if otro_tipo == 'IVA':
                                match_otro = (df_otro_galac['RIF_norm'] == rif_cp) & (df_otro_galac['COMPROBANTE_norm'].str.endswith(comprobante_cp[-6:]))
                            elif otro_tipo == 'ISLR':
                                match_otro = (df_otro_galac['RIF_norm'] == rif_cp) & (df_otro_galac['COMPROBANTE_norm'] == comprobante_cp) & (df_otro_galac['FACTURA_norm'] == factura_cp)
                            elif otro_tipo == 'MUNICIPAL':
                                match_otro = (df_otro_galac['RIF_norm'] == rif_cp) & (df_otro_galac['FACTURA_norm'] == factura_cp)
                            
                            if match_otro.any():
                                resultado['CP_Vs_Galac'] = f'Error: Subtipo {subtipo}, Encontrado en {otro_tipo}'
                                break
                    
                    if resultado['CP_Vs_Galac'] == 'No Encontrado en GALAC':
                        match_doc_errado = (df_galac_target['RIF_norm'] == rif_cp) & (np.isclose(df_galac_target['MONTO'].abs(), abs(monto_cp)))
                        if match_doc_errado.sum() == 1:
                            resultado['CP_Vs_Galac'] = 'Error: Documento No Coincide'

            asiento_cp = row_cp.get('ASIENTOCONTABLE', '')
            if asiento_cp:
                df_asiento_cg = df_cg[df_cg['ASIENTO'] == asiento_cp]
                if not df_asiento_cg.empty:
                    resultado['Asiento_en_CG'] = 'Sí'
                    monto_cg = df_asiento_cg[df_asiento_cg['CUENTACONTABLE'] == CUENTAS_MAP.get(subtipo, '')]['CREDITOVES'].sum()
                    if np.isclose(monto_cg, abs(monto_cp)):
                        resultado['Monto_coincide_CG'] = 'Sí'
                    else:
                        resultado['Monto_coincide_CG'] = 'No'

            results.append(resultado)

        df_cp_results = df_cp.join(pd.DataFrame(results))
        df_galac_no_cp = df_galac_full.drop(indices_galac_encontrados)

        # --- 3. GENERACIÓN DE REPORTES ---
        log_messages.append("Generando reporte final en formato Excel...")
        
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df_galac_no_cp.to_excel(writer, sheet_name='GALAC', index=False)
            df_cp_results.to_excel(writer, sheet_name='Relacion CP', index=False)
            df_errores = df_cp_results[(df_cp_results['CP_Vs_Galac'] != 'Sí') | (df_cp_results['Asiento_en_CG'] != 'Sí') | (df_cp_results['Monto_coincide_CG'] != 'Sí')]
            asientos_con_error = df_errores['ASIENTOCONTABLE'].unique()
            df_cg_errores = df_cg[df_cg['ASIENTO'].isin(asientos_con_error)]
            df_cg_errores.to_excel(writer, sheet_name='Diario CG', index=False)
            
            workbook = writer.book
            formato_titulo = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
            worksheet_galac = writer.sheets['GALAC']
            worksheet_galac.set_column('A:Z', 15)
            worksheet_galac.merge_range('A1:G1', 'Retenciones en GALAC No Encontradas en CP (Omisiones)', formato_titulo)
            
            worksheet_cp = writer.sheets['Relacion CP']
            worksheet_cp.set_column('A:Z', 15)
            worksheet_cp.merge_range('A1:J1', 'Panel de Control de Conciliación - Relación CP', formato_titulo)

            worksheet_cg = writer.sheets['Diario CG']
            worksheet_cg.set_column('A:Z', 15)
            worksheet_cg.merge_range('A1:H1', 'Detalle de Asientos con Discrepancias en Diario', formato_titulo)

        log_messages.append("¡Proceso de conciliación de retenciones completado con éxito!")
        return output_buffer.getvalue()

    except Exception as e:
        log_messages.append(f"❌ ERROR CRÍTICO en la conciliación de retenciones: {e}")
        import traceback
        log_messages.append(traceback.format_exc())
        return None
