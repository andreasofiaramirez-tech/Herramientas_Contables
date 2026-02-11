# logic.py

import pandas as pd
import numpy as np
import re
import itertools
from itertools import combinations
from io import BytesIO
import unicodedata
import xlsxwriter
from difflib import SequenceMatcher  # Necesario para la detección de errores de tipeo
from utils import generar_reporte_retenciones
import bisect

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
        if 'TARJETA' in ref and ('GASTOS' in ref or 'INGRESO' in ref): return 'TARJETA_GASTOS', 'GRUPO_TARJETA', 'LOTE_TARJETAS'
        if 'NOTA DE DEBITO' in ref: return 'NOTA_DEBITO', 'GRUPO_NOTA', 'NOTA_DEBITO'
        if 'NOTA DE CREDITO' in ref: return 'NOTA_CREDITO', 'GRUPO_NOTA', 'NOTA_CREDITO'
        return 'OTRO', 'OTRO', ref_lit_norm

    df_copy[['Clave_Normalizada', 'Clave_Grupo', 'Referencia_Normalizada_Literal']] = df_copy['Referencia'].apply(clasificar_usd).apply(pd.Series)
    return df_copy
    
def conciliar_automaticos_usd(df, log_messages):
    total = 0
    grupos_a_revisar = [
        ('GRUPO_DIF_CAMBIO', 'AUTOMATICO_DIF_CAMBIO'), 
        ('GRUPO_AJUSTE', 'AUTOMATICO_AJUSTE'),
        ('GRUPO_TARJETA', 'AUTOMATICO_TARJETA') 
    ]
    
    for grupo, etiqueta in grupos_a_revisar:
        # Filtramos los movimientos de ese grupo que no estén conciliados
        indices = df.loc[(df['Clave_Grupo'] == grupo) & (~df['Conciliado'])].index
        
        if not indices.empty:
            # Verificamos si la suma total del grupo es Cero
            suma_grupo = df.loc[indices, 'Monto_USD'].sum()
            
            if abs(suma_grupo) <= TOLERANCIA_MAX_USD:
                df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, etiqueta]
                log_messages.append(f"✔️ Fase Auto (USD): {len(indices)} conciliados por ser '{etiqueta}'.")
                total += len(indices)
            else:
                # Si no suman cero todos juntos, intentamos agrupar por la referencia literal forzada
                # Esto ayuda si hay varios meses de tarjetas en el mismo archivo
                subgrupos = df.loc[indices].groupby('Referencia_Normalizada_Literal')
                for ref_lit, subgrupo in subgrupos:
                    if abs(subgrupo['Monto_USD'].sum()) <= TOLERANCIA_MAX_USD:
                         df.loc[subgrupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"{etiqueta}_{ref_lit}"]
                         total += len(subgrupo.index)
                         log_messages.append(f"✔️ Fase Auto (USD): {len(subgrupo)} conciliados en '{etiqueta}' (Subgrupo).")

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
    """
    Fase Avanzada: Busca combinaciones de N movimientos contra 1 (N:1).
    MEJORA: Capacidad ampliada hasta 6 combinaciones si el volumen de datos lo permite.
    """
    log_messages.append("\n--- FASE GRUPOS COMPLEJOS (N vs 1) (USD) ---")
    
    pendientes = df.loc[~df['Conciliado']]
    num_pendientes = len(pendientes)
    
    # --- AJUSTE DINÁMICO DE POTENCIA ---
    # Mientras menos datos, más combinaciones podemos probar sin colgar el sistema.
    if num_pendientes < 50:
        MAX_COMBINACIONES = 6 # ¡Potencia máxima! (Busca hasta 6 facturas contra 1 pago)
    elif num_pendientes < 200:
        MAX_COMBINACIONES = 5
    elif num_pendientes < 1000:
        MAX_COMBINACIONES = 4
    else:
        MAX_COMBINACIONES = 3 # Modo seguro para archivos grandes
        
    log_messages.append(f"ℹ️ Analizando combinaciones de hasta {MAX_COMBINACIONES} elementos contra 1...")
    
    debitos = pendientes[pendientes['Monto_USD'] > 0].copy()
    creditos = pendientes[pendientes['Monto_USD'] < 0].copy()

    if debitos.empty or creditos.empty: return 0

    total_conciliados_fase = 0
    indices_usados = set()

    # --- CASO 1: N Débitos vs 1 Crédito (Ej: Varias Facturas vs 1 Pago) ---
    creditos_ordenados = creditos.sort_values('Monto_USD', ascending=True) 

    for idx_c, row_c in creditos_ordenados.iterrows():
        if idx_c in indices_usados: continue
        
        target = abs(row_c['Monto_USD'])
        
        # Filtramos candidatos para reducir la carga computacional
        candidatos = debitos[
            (debitos['Monto_USD'] <= target + TOLERANCIA_MAX_USD) & 
            (~debitos.index.isin(indices_usados))
        ]
        
        # Si hay demasiados candidatos (>30), limitamos la profundidad para este caso específico
        limit_local = MAX_COMBINACIONES
        if len(candidatos) > 40 and limit_local > 4:
            limit_local = 4 # Bajamos a 4 si hay demasiados candidatos para este número específico
        
        if len(candidatos) < 2: continue
        
        encontrado = False
        for r in range(2, limit_local + 1):
            # Itertools es rápido, pero con r=6 puede tardar
            for combo_idx in combinations(candidatos.index, r):
                suma_combo = df.loc[list(combo_idx), 'Monto_USD'].sum()
                
                if np.isclose(suma_combo, target, atol=TOLERANCIA_MAX_USD):
                    indices_todos = list(combo_idx) + [idx_c]
                    asiento_c = row_c['Asiento']
                    
                    df.loc[indices_todos, 'Conciliado'] = True
                    df.loc[indices_todos, 'Grupo_Conciliado'] = f'GRUPO_{r}v1_{asiento_c}'
                    
                    indices_usados.update(indices_todos)
                    total_conciliados_fase += len(indices_todos)
                    encontrado = True
                    log_messages.append(f"   ⚡ Match Complejo: {r} Débitos suman {target:.2f}")
                    break
            if encontrado: break

    if progress_bar: progress_bar.progress(0.7, text="Buscando N créditos contra 1 débito...")

    # --- CASO 2: 1 Débito vs N Créditos (Ej: 1 Depósito vs Varias CxC) ---
    debitos_ordenados = debitos.sort_values('Monto_USD', ascending=False)

    for idx_d, row_d in debitos_ordenados.iterrows():
        if idx_d in indices_usados: continue
        
        target = row_d['Monto_USD']
        
        candidatos = creditos[
            (creditos['Monto_USD'].abs() <= target + TOLERANCIA_MAX_USD) & 
            (~creditos.index.isin(indices_usados))
        ]
        
        limit_local = MAX_COMBINACIONES
        if len(candidatos) > 40 and limit_local > 4: limit_local = 4

        if len(candidatos) < 2: continue
        
        encontrado = False
        for r in range(2, limit_local + 1):
            for combo_idx in combinations(candidatos.index, r):
                suma_combo = df.loc[list(combo_idx), 'Monto_USD'].sum()
                
                if np.isclose(target + suma_combo, 0, atol=TOLERANCIA_MAX_USD):
                    indices_todos = list(combo_idx) + [idx_d]
                    asiento_d = row_d['Asiento']
                    
                    df.loc[indices_todos, 'Conciliado'] = True
                    df.loc[indices_todos, 'Grupo_Conciliado'] = f'GRUPO_1v{r}_{asiento_d}'
                    
                    indices_usados.update(indices_todos)
                    total_conciliados_fase += len(indices_todos)
                    encontrado = True
                    log_messages.append(f"   ⚡ Match Complejo: 1 Débito cruza con {r} Créditos")
                    break
            if encontrado: break

    if total_conciliados_fase > 0:
        log_messages.append(f"✔️ Fase Grupos Complejos: {total_conciliados_fase} movimientos conciliados.")
    
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
    df_pendientes = df.loc[~df['Conciliado']]
    grupos_por_empleado = df_pendientes.groupby('Clave_Empleado')
    
    log_messages.append(f"ℹ️ Se analizarán los saldos de {len(grupos_por_empleado)} empleados.")
    
    # --- BÚSQUEDA INTELIGENTE DE LA COLUMNA DE NOMBRE ---
    col_nombre = None
    # Lista de posibles nombres que puede tener la columna en el Excel
    candidatos = ['Descripcion NIT', 'Descripción Nit', 'Nombre del Proveedor', 'Nombre', 'Cliente']
    
    for col in candidatos:
        if col in df.columns:
            col_nombre = col
            break # Encontramos una válida, dejamos de buscar
    # ----------------------------------------------------

    for clave_empleado, grupo in grupos_por_empleado:
        if clave_empleado == 'SIN_NIT' or pd.isna(clave_empleado) or not clave_empleado:
            continue
            
        suma_usd = grupo['Monto_USD'].sum()
        
        if abs(suma_usd) <= TOLERANCIA_MAX_USD:
            indices_a_conciliar = grupo.index
            
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"SALDO_CERO_EMP_{clave_empleado}"
            
            num_movs = len(indices_a_conciliar)
            total_conciliados += num_movs
            
            # Extracción segura del nombre
            if col_nombre and not grupo.empty:
                nombre_empleado = grupo[col_nombre].iloc[0]
            else:
                nombre_empleado = clave_empleado # Si no hay columna de nombre, usamos el NIT
                
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
    Versión final que maneja coincidencias parciales y limpieza automática de Diferencial Cambiario.
    USA TOLERANCIA EXACTA (0.00) para contabilidad precisa.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE COBROS VIAJEROS (V12 - TOLERANCIA CERO) ---")
    
    # --- DEFINICIÓN DE TOLERANCIA LOCAL ---
    # Cambio solicitado: Tolerancia CERO. No se permiten diferencias.
    TOLERANCIA_ESTRICTA_USD = 0.00
    # --------------------------------------

    df = normalizar_datos_cobros_viajeros(df, log_messages)
    if progress_bar: progress_bar.progress(0.1, text="Fase de Normalización completada.")

    total_conciliados = 0
    indices_usados = set()

    # --- FASE 0: CONCILIACIÓN AUTOMÁTICA (DIFERENCIAL CAMBIARIO) ---
    def es_ajuste_cambiario(texto):
        t = str(texto).upper()
        if 'DIFF' in t: return True
        if 'CAMBIO' in t and ('DIFERENCIA' in t or 'DIF' in t or 'AJUSTE' in t): return True
        return False

    indices_dif = df[df['Referencia'].apply(es_ajuste_cambiario) & (~df['Conciliado'])].index

    if not indices_dif.empty:
        df.loc[indices_dif, 'Conciliado'] = True
        df.loc[indices_dif, 'Grupo_Conciliado'] = 'AUTOMATICO_DIF_CAMBIO'
        indices_usados.update(indices_dif)
        count_dif = len(indices_dif)
        total_conciliados += count_dif
        log_messages.append(f"✔️ Fase Auto: {count_dif} movimientos conciliados por ser 'Diferencia en Cambio'.")

    # --- FASE 1: CONCILIACIÓN DE REVERSOS ---
    log_messages.append("--- Fase 1: Buscando reversos con coincidencia parcial ---")
    df_reversos = df[df['Es_Reverso'] & (~df['Conciliado'])].copy()
    df_originales = df[~df['Es_Reverso'] & (~df['Conciliado'])].copy()

    for idx_r, reverso_row in df_reversos.iterrows():
        if idx_r in indices_usados: continue
        
        clave_reverso = reverso_row['Referencia_Norm_Num']
        if not clave_reverso: continue

        nit_reverso = reverso_row['NIT_Normalizado']
        
        for idx_o, original_row in df_originales.iterrows():
            if idx_o in indices_usados or original_row['NIT_Normalizado'] != nit_reverso:
                continue

            clave_orig_ref = original_row['Referencia_Norm_Num']
            clave_orig_fuente = original_row['Fuente_Norm_Num']
            
            match_en_referencia = (clave_reverso and clave_orig_ref and (clave_reverso.endswith(clave_orig_ref) or clave_orig_ref.endswith(clave_reverso)))
            match_en_fuente = (clave_reverso and clave_orig_fuente and (clave_reverso.endswith(clave_orig_fuente) or clave_orig_fuente.endswith(clave_reverso)))

            # Comparación con tolerancia 0.00
            if (match_en_referencia or match_en_fuente) and np.isclose(reverso_row['Monto_USD'] + original_row['Monto_USD'], 0, atol=TOLERANCIA_ESTRICTA_USD):
                indices_a_conciliar = [idx_r, idx_o]
                df.loc[indices_a_conciliar, 'Conciliado'] = True
                df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"REVERSO_{nit_reverso}_{clave_reverso}"
                indices_usados.update(indices_a_conciliar)
                total_conciliados += 2
                log_messages.append(f"✔️ Reverso conciliado para NIT {nit_reverso}.")
                break 

    if progress_bar: progress_bar.progress(0.5, text="Fase de Reversos completada.")

    # --- FASE 2: CONCILIACIÓN ESTÁNDAR N-a-N (Búsqueda por Recibo/Depósito) ---
    log_messages.append("--- Fase 2: Buscando grupos de conciliación estándar N-a-N ---")
    
    df['Clave_Vinculo'] = ''
    for index, row in df[~df['Conciliado']].iterrows():
        asiento = str(row['Asiento']).upper()
        f_num = str(row.get('Fuente_Norm_Num', ''))
        r_num = str(row.get('Referencia_Norm_Num', ''))
        
        # MEJORA: Si ambos números existen, tomamos el más largo para mayor precisión
        # Esto resuelve el problema de los números truncados en JIANLONG MO
        if asiento.startswith(('CC', 'CG')):
            df.loc[index, 'Clave_Vinculo'] = f_num if f_num != '' else r_num
        elif asiento.startswith('CB'):
            df.loc[index, 'Clave_Vinculo'] = r_num if r_num != '' else f_num
        else:
            df.loc[index, 'Clave_Vinculo'] = f_num if len(f_num) > len(r_num) else r_num

    df_procesable = df[(~df['Conciliado']) & (df['Clave_Vinculo'] != '')]
    grupos = df_procesable.groupby(['NIT_Normalizado', 'Clave_Vinculo'])
    
    for (nit, clave), grupo in grupos:
        if len(grupo) >= 2 and np.isclose(grupo['Monto_USD'].sum(), 0, atol=TOLERANCIA_ESTRICTA_USD):
            df.loc[grupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"VIAJERO_{nit}_{clave}"]
            indices_usados.update(grupo.index)
            total_conciliados += len(grupo)

    # --- FASE 3: EL BARRIDO DEFINITIVO POR NIT (Solución JIANLONG MO) ---
    # Si después de todo, el saldo de un NIT es CERO, se cierra.
    log_messages.append("--- Fase 3: Ejecutando Barrido de Saldo Neto por NIT ---")
    
    df_pendientes_final = df[~df['Conciliado']]
    # Agrupamos por NIT y sumamos. Usamos filter para quedarnos con los que dan CERO.
    resumen_nit = df_pendientes_final.groupby('NIT_Normalizado')['Monto_USD'].sum().round(2)
    nits_a_cerrar = resumen_nit[abs(resumen_nit) <= TOLERANCIA_ESTRICTA_USD].index

    for nit in nits_a_cerrar:
        if nit == 'SIN_NIT': continue
        indices = df_pendientes_final[df_pendientes_final['NIT_Normalizado'] == nit].index
        if len(indices) > 0:
            df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, f"BARRIDO_NETO_NIT_{nit}"]
            total_conciliados += len(indices)
            log_messages.append(f"✔️ NIT {nit}: Conciliado por saldo neto cero en el barrido final.")

    if progress_bar: progress_bar.progress(1.0)
    log_messages.append(f"✔️ Proceso finalizado. Conciliados: {total_conciliados}")
    return df

# --- (G) Módulo: Otras Cuentas por Pagar (VES) ---

def normalizar_datos_otras_cxp(df, log_messages):
    """
    Prepara el DataFrame extrayendo el número de envío de la Referencia.
    Versión 'Nuclear': Captura ENV seguido de cualquier cosa hasta encontrar números.
    """
    df_copy = df.copy()
    
    # Normalizar NIT
    nit_col_name = next((col for col in df_copy.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df_copy['NIT_Normalizado'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        log_messages.append("⚠️ ADVERTENCIA: No se encontró columna 'NIT' o 'RIF'.")
        df_copy['NIT_Normalizado'] = 'SIN_NIT'

    # --- REGEX NUCLEAR ---
    # 1. Busca 'ENV' (insensible a mayúsculas)
    # 2. .*? salta cualquier caracter (puntos, espacios, guiones) de forma no codiciosa
    # 3. (\d+) captura el primer grupo de números que encuentre después
    df_copy['Numero_Envio'] = df_copy['Referencia'].str.extract(r"ENV.*?(\d+)", expand=False, flags=re.IGNORECASE)
    
    # Rellenar vacíos para evitar errores al ordenar
    df_copy['Numero_Envio'] = df_copy['Numero_Envio'].fillna('')
    
    return df_copy

def run_conciliation_otras_cxp(df, log_messages, progress_bar=None):
    """
    Orquesta la conciliación para Otras Cuentas por Pagar (VES).
    1. Concilia automáticamente Diferencias en Cambio.
    2. Busca grupos por NIT y ENV que sumen CERO o coincidan en magnitud.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE OTRAS CUENTAS POR PAGAR (VES) ---")
    
    df = normalizar_datos_otras_cxp(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase de Normalización completada.")

    total_conciliados = 0

    # --- FASE 0: CONCILIACIÓN AUTOMÁTICA (DIFERENCIA EN CAMBIO) ---
    # Detectamos referencias que contengan 'DIFF' o la combinación 'DIFERENCIA' y 'CAMBIO'
    # Esto cubre "Diferencias de cambio al...", "DIFF000182", "Ajuste Diferencial", etc.
    def es_diferencial_cambiario(texto):
        t = str(texto).upper()
        if 'DIFF' in t: return True
        if 'CAMBIO' in t and ('DIFERENCIA' in t or 'DIF.' in t or 'AJUSTE' in t): return True
        return False

    indices_dif = df[df['Referencia'].apply(es_diferencial_cambiario) & (~df['Conciliado'])].index

    if not indices_dif.empty:
        df.loc[indices_dif, 'Conciliado'] = True
        df.loc[indices_dif, 'Grupo_Conciliado'] = 'AUTOMATICO_DIF_CAMBIO'
        count_dif = len(indices_dif)
        total_conciliados += count_dif
        log_messages.append(f"✔️ Fase Auto: {count_dif} movimientos conciliados por ser 'Diferencia en Cambio'.")
    # ---------------------------------------------------------------
    
    # Filtrar solo las filas pendientes donde se pudo extraer un número de envío válido
    df_procesable = df[(~df['Conciliado']) & (df['Numero_Envio'] != '')]
    
    # Agrupar por NIT y Número de Envío
    grupos = df_procesable.groupby(['NIT_Normalizado', 'Numero_Envio'])
    if not df_procesable.empty:
        log_messages.append(f"ℹ️ Se encontraron {len(grupos)} combinaciones de NIT/Envío para analizar.")

    for (nit, envio), grupo in grupos:
        if len(grupo) < 2: 
            continue

        # CASO 1: Conciliación Estándar (Suma Cero)
        if np.isclose(grupo['Monto_BS'].sum(), 0, atol=TOLERANCIA_MAX_BS):
            indices_a_conciliar = grupo.index
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"OTRAS_CXP_{nit}_{envio}"
            total_conciliados += len(indices_a_conciliar)
            
        # CASO 2: Conciliación por Magnitud (Corrección de Signos para Débitos vs Débitos)
        elif len(grupo) == 2:
            vals = grupo['Monto_BS'].abs().values
            # Si el valor absoluto es igual (con tolerancia)
            if np.isclose(vals[0], vals[1], atol=TOLERANCIA_MAX_BS):
                indices_a_conciliar = grupo.index
                df.loc[indices_a_conciliar, 'Conciliado'] = True
                df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"OTRAS_CXP_MAGNITUD_{nit}_{envio}"
                total_conciliados += len(indices_a_conciliar)

    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación finalizada: Se conciliaron {total_conciliados} movimientos en total.")
    else:
        log_messages.append("ℹ️ No se encontraron movimientos para conciliar.")
        
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# --- (H) Módulo: Haberes de Clientes (VES) ---

def run_conciliation_haberes_clientes(df, log_messages, progress_bar=None):
    """
    Conciliación de Haberes de Clientes (BS).
    Fase 1: Por NIT (Suma Cero).
    Fase 2: Por Monto Exacto (Recuperación de NITs perdidos).
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE HABERES DE CLIENTES (BS) ---")
    
    # 1. Normalización (Usamos la misma lógica de limpieza de NIT)
    # Reutilizamos normalizar_datos_otras_cxp que ya limpia NITs y Referencias
    df = normalizar_datos_otras_cxp(df, log_messages) 
    if progress_bar: progress_bar.progress(0.2, text="Fase de Normalización completada.")

    total_conciliados = 0

    # --- FASE 1: CONCILIACIÓN POR NIT (ESTÁNDAR) ---
    # Agrupamos por NIT. Si la suma de todos sus movimientos es 0, se cierran.
    df_pendientes = df[~df['Conciliado']]
    grupos_nit = df_pendientes.groupby('NIT_Normalizado')
    
    for nit, grupo in grupos_nit:
        if nit == 'SIN_NIT': continue # Saltamos los vacíos para la fase 2
        
        if np.isclose(grupo['Monto_BS'].sum(), 0, atol=TOLERANCIA_MAX_BS):
            indices = grupo.index
            df.loc[indices, 'Conciliado'] = True
            df.loc[indices, 'Grupo_Conciliado'] = f"HABER_NIT_{nit}"
            total_conciliados += len(indices)
            
    log_messages.append(f"✔️ Fase 1 (Por NIT): {total_conciliados} movimientos conciliados.")
    if progress_bar: progress_bar.progress(0.5, text="Fase 1 completada.")

    # --- FASE 2: CONCILIACIÓN POR MONTO EXACTO (RECUPERACIÓN) ---
    # Buscamos pares (Débito vs Crédito) que tengan exactamente el mismo monto absoluto
    # Esto cruza filas con NIT vs filas SIN NIT (o NIT errado).
    
    df_pendientes = df[~df['Conciliado']].copy()
    df_pendientes['Monto_Abs'] = df_pendientes['Monto_BS'].abs()
    
    # Agrupamos por monto absoluto
    grupos_monto = df_pendientes.groupby('Monto_Abs')
    count_fase2 = 0
    
    for monto, grupo in grupos_monto:
        if len(grupo) < 2 or monto <= TOLERANCIA_MAX_BS: continue
        
        debitos = grupo[grupo['Monto_BS'] > 0].index.tolist()
        creditos = grupo[grupo['Monto_BS'] < 0].index.tolist()
        
        # Emparejamos 1 a 1
        pares = min(len(debitos), len(creditos))
        for i in range(pares):
            idx_d = debitos[i]
            idx_c = creditos[i]
            
            # Validamos suma cero estricta
            if np.isclose(df.loc[idx_d, 'Monto_BS'] + df.loc[idx_c, 'Monto_BS'], 0, atol=TOLERANCIA_MAX_BS):
                # Intentamos rescatar el NIT del que sí lo tenga para la etiqueta
                nit_d = df.loc[idx_d, 'NIT_Normalizado']
                nit_c = df.loc[idx_c, 'NIT_Normalizado']
                nit_ref = nit_d if nit_d != 'SIN_NIT' else (nit_c if nit_c != 'SIN_NIT' else 'GENERICO')
                
                df.loc[[idx_d, idx_c], 'Conciliado'] = True
                df.loc[[idx_d, idx_c], 'Grupo_Conciliado'] = f"HABER_MONTO_{nit_ref}_{int(monto)}"
                count_fase2 += 2

    total_conciliados += count_fase2
    log_messages.append(f"✔️ Fase 2 (Por Monto/Sin NIT): {count_fase2} movimientos conciliados.")
    if progress_bar: progress_bar.progress(1.0, text="Proceso Finalizado.")
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# --- (I) Módulo: CDC - Factoring (USD) ---

def normalizar_datos_cdc_factoring(df, log_messages):
    """
    Extrae el código del contrato.
    Nivel 1: Patrones FQ u O/C.
    Nivel 2: Barrido después de palabra 'FACTORING'.
    Nivel 3: Detección directa (si la referencia es solo el número).
    """
    df_copy = df.copy()
    
    # 1. Normalizar NIT
    nit_col_name = next((col for col in df_copy.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df_copy['NIT_Normalizado'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        df_copy['NIT_Normalizado'] = 'SIN_NIT'

    def limpiar_y_extraer(texto_crudo):
        if not isinstance(texto_crudo, str): return None
        texto = texto_crudo.upper().strip()
        if not texto: return None
        
        # --- NIVEL 1: Patrones de Alta Precisión (FQ y O/C) ---
        match_fq = re.search(r"(FQ-[A-Z0-9-]+)", texto)
        if match_fq: return match_fq.group(1)
        
        match_oc = re.search(r"O/C\s*[-]?\s*([A-Z0-9]+)", texto)
        if match_oc: return match_oc.group(1)

        # --- NIVEL 2: Barrido Inteligente después de 'FACTORING' ---
        if "FACTORING" in texto:
            try:
                parte_derecha = texto.split("FACTORING", 1)[1]
                parte_derecha = parte_derecha.replace('$', ' ').replace(':', ' ').replace(',', ' ')
                palabras = parte_derecha.split()
                ignorar = ['DE', 'DEL', 'AL', 'NRO', 'NUM', 'NUMERO', 'REF', 'NO', 'PAGO', 'ABONO', 'CANCELACION', 'FAC', 'FACT']
                
                for palabra in palabras:
                    p_clean = palabra.replace('.', '').strip()
                    if p_clean in ignorar: continue
                    if any(char.isdigit() for char in p_clean):
                        if len(p_clean) >= 3: return p_clean
            except IndexError:
                pass

        # --- NIVEL 3: Referencia Directa (Caso Febeca) ---
        # Si la referencia NO tiene espacios y contiene números (ej: "6016301")
        # Verificamos que sea un bloque sólido alfanumérico
        if ' ' not in texto:
            # Debe tener al menos un número
            if any(char.isdigit() for char in texto):
                # Longitud mínima 4 para evitar capturar "1", "2023" (años), etc.
                if len(texto) >= 4:
                    return texto

        return None

    def extraer_contrato_row(row):
        # Buscamos primero en Referencia, luego en Fuente
        contrato_ref = limpiar_y_extraer(row.get('Referencia', ''))
        if contrato_ref: return contrato_ref
        
        contrato_fuente = limpiar_y_extraer(row.get('Fuente', ''))
        if contrato_fuente: return contrato_fuente
        
        return 'SIN_CONTRATO'

    df_copy['Contrato'] = df_copy.apply(extraer_contrato_row, axis=1)
    
    return df_copy
    
def run_conciliation_cdc_factoring(df, log_messages, progress_bar=None):
    """
    Conciliación de Factoring (USD).
    Incluye limpieza automática de Diferencial Cambiario y cruce por Contrato.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE CDC - FACTORING (USD) ---")
    
    df = normalizar_datos_cdc_factoring(df, log_messages)
    if progress_bar: progress_bar.progress(0.2, text="Fase de Normalización completada.")

    total_conciliados = 0

    # --- FASE 0: CONCILIACIÓN AUTOMÁTICA (DIFERENCIAL CAMBIARIO) ---
    # Detectamos ajustes contables por valoración de moneda.
    # Buscamos en Referencia, Fuente y Descripción si existe.
    def es_ajuste_cambiario(row):
        texto_completo = (str(row.get('Referencia', '')) + " " + 
                          str(row.get('Fuente', '')) + " " + 
                          str(row.get('Descripción', ''))).upper()
        
        palabras_clave = ['DIFERENCIA DE CAMBIO', 'DIFERENCIAS DE CAMBIO', 'DIFERENCIAL CAMBIARIO', 'AJUSTE CAMBIARIO', 'DIFF']
        
        for palabra in palabras_clave:
            if palabra in texto_completo:
                return True
        return False

    # Aplicamos el filtro
    indices_dif = df[df.apply(es_ajuste_cambiario, axis=1) & (~df['Conciliado'])].index

    if not indices_dif.empty:
        df.loc[indices_dif, 'Conciliado'] = True
        df.loc[indices_dif, 'Grupo_Conciliado'] = 'AUTOMATICO_DIF_CAMBIO'
        count_dif = len(indices_dif)
        total_conciliados += count_dif
        log_messages.append(f"✔️ Fase Auto: {count_dif} movimientos conciliados por ser 'Diferencia en Cambio'.")
    # ------------------------------------------------------------------
    
    # Filtramos pendientes reales para la lógica de contratos
    df_pendientes = df[~df['Conciliado']]
    
    # Agrupamos por NIT y Contrato
    # Ignoramos los que no tienen contrato identificado para esta fase
    grupos = df_pendientes[df_pendientes['Contrato'] != 'SIN_CONTRATO'].groupby(['NIT_Normalizado', 'Contrato'])
    
    log_messages.append(f"ℹ️ Se encontraron {len(grupos)} contratos para analizar.")

    for (nit, contrato), grupo in grupos:
        if len(grupo) < 2: continue
        
        # Validar suma cero en Dólares
        if np.isclose(grupo['Monto_USD'].sum(), 0, atol=TOLERANCIA_MAX_USD):
            indices = grupo.index
            df.loc[indices, 'Conciliado'] = True
            df.loc[indices, 'Grupo_Conciliado'] = f"FACT_{nit}_{contrato}"
            total_conciliados += len(indices)

    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación Factoring: {total_conciliados} movimientos conciliados.")
    
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    return df

# --- (J) Módulo: Asientos por Clasificar (VES) ---
def run_conciliation_asientos_por_clasificar(df, log_messages, progress_bar=None):
    """
    Conciliación de Asientos por Clasificar (BS).
    Incluye diagnóstico de saldo final y redondeo forzado.
    """
    log_messages.append("\n--- INICIANDO LÓGICA DE ASIENTOS POR CLASIFICAR (BS) ---")
    TOLERANCIA_ESTRICTA_BS = 0.00
    
    # 1. Normalización
    df_copy = df.copy()
    nit_col_name = next((col for col in df_copy.columns if str(col).strip().upper() in ['NIT', 'RIF']), None)
    if nit_col_name:
        df['NIT_Normalizado'] = df_copy[nit_col_name].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    else:
        df['NIT_Normalizado'] = 'SIN_NIT'

    if progress_bar: progress_bar.progress(0.1, text="Fase de Normalización completada.")
    
    total_conciliados = 0

    # --- FASE 0: LIMPIEZA AUTOMÁTICA (DIFERENCIAL CAMBIARIO) ---
    def es_ajuste_cambiario(texto):
        t = str(texto).upper()
        if 'DIFF' in t: return True
        if 'CAMBIO' in t and ('DIFERENCIA' in t or 'DIF' in t or 'AJUSTE' in t): return True
        return False

    mask_dif = df['Referencia'].apply(es_ajuste_cambiario) | df['Fuente'].apply(es_ajuste_cambiario)
    indices_dif = df[mask_dif & (~df['Conciliado'])].index

    if not indices_dif.empty:
        df.loc[indices_dif, 'Conciliado'] = True
        df.loc[indices_dif, 'Grupo_Conciliado'] = 'AUTOMATICO_DIF_CAMBIO_BS'
        total_conciliados += len(indices_dif)
        log_messages.append(f"✔️ Fase Auto: {len(indices_dif)} movimientos de Diferencial Cambiario conciliados.")

    # --- FASE 1: CRUCE POR NIT (1 a 1 y N a N) ---
    df_pendientes = df[~df['Conciliado']]
    grupos_nit = df_pendientes.groupby('NIT_Normalizado')
    
    for nit, grupo in grupos_nit:
        if nit == 'SIN_NIT' or len(grupo) < 2: continue
        
        # A. Pares Exactos (1 a 1)
        debitos = grupo[grupo['Monto_BS'] > 0].index.tolist()
        creditos = grupo[grupo['Monto_BS'] < 0].index.tolist()
        usados_local = set()
        
        for idx_d in debitos:
            if idx_d in usados_local: continue
            monto_d = df.loc[idx_d, 'Monto_BS']
            for idx_c in creditos:
                if idx_c in usados_local: continue
                if np.isclose(monto_d + df.loc[idx_c, 'Monto_BS'], 0, atol=TOLERANCIA_ESTRICTA_BS):
                    df.loc[[idx_d, idx_c], 'Conciliado'] = True
                    df.loc[[idx_d, idx_c], 'Grupo_Conciliado'] = f"PAR_NIT_{nit}"
                    total_conciliados += 2
                    usados_local.add(idx_d); usados_local.add(idx_c)
                    break
        
        # B. Grupo Completo (N a N)
        remanente = grupo[~grupo.index.isin(usados_local)]
        if len(remanente) > 1:
            # Usamos round para evitar errores de flotante
            if round(remanente['Monto_BS'].sum(), 2) == 0.00:
                indices = remanente.index
                df.loc[indices, 'Conciliado'] = True
                df.loc[indices, 'Grupo_Conciliado'] = f"GRUPO_NIT_{nit}"
                total_conciliados += len(indices)

    if progress_bar: progress_bar.progress(0.6, text="Fase por NIT completada.")

    # --- FASE 2: CRUCE GLOBAL POR MONTO ---
    df_pendientes_final = df[~df['Conciliado']].copy()
    df_pendientes_final['Monto_Abs'] = df_pendientes_final['Monto_BS'].abs()
    
    for monto, grupo in df_pendientes_final.groupby('Monto_Abs'):
        if len(grupo) < 2 or monto <= 0.01: continue
        
        debitos = grupo[grupo['Monto_BS'] > 0].index.tolist()
        creditos = grupo[grupo['Monto_BS'] < 0].index.tolist()
        
        pares = min(len(debitos), len(creditos))
        for i in range(pares):
            idx_d, idx_c = debitos[i], creditos[i]
            if np.isclose(df.loc[idx_d, 'Monto_BS'] + df.loc[idx_c, 'Monto_BS'], 0, atol=TOLERANCIA_ESTRICTA_BS):
                df.loc[[idx_d, idx_c], 'Conciliado'] = True
                df.loc[[idx_d, idx_c], 'Grupo_Conciliado'] = f"GLOBAL_MONTO_{int(monto)}"
                total_conciliados += 2

    # --- FASE 3: BARRIDO FINAL (DIAGNÓSTICO Y CORRECCIÓN) ---
    df_remanente = df[~df['Conciliado']]
    
    if not df_remanente.empty:
        suma_final_real = df_remanente['Monto_BS'].sum()
        suma_final_redondeada = round(suma_final_real, 2)
        
        # MENSAJE DE DIAGNÓSTICO (Aparecerá en el Log de la web)
        log_messages.append(f"🔎 DIAGNÓSTICO FASE FINAL:")
        log_messages.append(f"   > Movimientos pendientes: {len(df_remanente)}")
        log_messages.append(f"   > Suma Real (Python): {suma_final_real:.10f}")
        log_messages.append(f"   > Suma Redondeada (2 dec): {suma_final_redondeada}")

        # Si la suma redondeada es cero (o casi cero, permitiendo 1 centimo de basura)
        if abs(suma_final_redondeada) <= 0.01:
            indices = df_remanente.index
            df.loc[indices, 'Conciliado'] = True
            df.loc[indices, 'Grupo_Conciliado'] = 'LOTE_FINAL_REMANENTE'
            
            cantidad_final = len(indices)
            total_conciliados += cantidad_final
            log_messages.append(f"✔️ Fase 3: ÉXITO. Se conciliaron {cantidad_final} movimientos finales.")
        else:
            log_messages.append("⚠️ Fase 3: NO se concilió el remanente porque la suma no es 0.00.")

    if progress_bar: progress_bar.progress(1.0, text="Proceso finalizado.")
    log_messages.append(f"✔️ Conciliación finalizada: {total_conciliados} movimientos cerrados.")
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
    
    return df

# --- (K) Módulo: Proveedores d/Mcia -Costos Causados ---

def run_conciliation_proveedores_costos(df, log_messages, progress_bar=None):
    """
    Conciliación 212.07.1012 - VERSIÓN DEFINITIVA V7.
    Incluye: Herencia NIT, Vínculo Factura, Rescate de Huérfanos y Cruce entre Embarques.
    """
    log_messages.append("\n--- INICIANDO CONCILIACIÓN PROVEEDORES COSTOS (212.07.1012) ---")
    
    # --- 1. NORMALIZACIÓN Y LIMPIEZA ---
    nit_col = next((c for c in df.columns if c.upper() in ['NIT', 'RIF']), None)
    df['NIT_Norm'] = df[nit_col].astype(str).str.strip().str.upper().replace(['NAN', 'NONE', 'NAT', '0', ''], 'ND') if nit_col else 'ND'
    
    def extraer_y_normalizar_emb(referencia):
        if pd.isna(referencia): return 'NO_EMB'
        ref = str(referencia).upper()
        match = re.search(r'([EM])[.:\-\s]*(\d{3,})', ref)
        if not match:
            match = re.search(r'(?:EMB|EEM|EMBARQUE|EMBARUE|EMB:)[.:\-\s]*(\d+)', ref)
        if match:
            try:
                letra = match.group(1) if match.group(1).isalpha() else 'E'
                num = re.sub(r'\D', '', match.group(1 if not match.lastindex or match.lastindex < 2 else 2))
                return f"{letra}{int(num)}" if num else 'NO_EMB'
            except: return 'NO_EMB'
        return 'NO_EMB'

    df['Numero_Embarque'] = df['Referencia'].apply(extraer_y_normalizar_emb)

    # --- 2. VÍNCULO POR FACTURA (Backfill) ---
    def extraer_factura_clean(texto):
        if pd.isna(texto): return None
        match = re.search(r'(?:FAC|S/F|FACT|N[R°]O|FACTURA)[.:\-\s]*([A-Z]*\d+)', str(texto).upper())
        if match:
            val = match.group(1)
            num_only = re.sub(r'\D', '', val)
            return str(int(num_only)) if num_only else val
        return None

    df['Factura_Norm'] = df['Fuente'].apply(extraer_factura_clean).fillna(df['Referencia'].apply(extraer_factura_clean))
    df_con_ambos = df[(df['Numero_Embarque'] != 'NO_EMB') & (df['Factura_Norm'].notna())]
    mapa_fac_emb = df_con_ambos.groupby('Factura_Norm')['Numero_Embarque'].first().to_dict()

    def backfill_embarque(row):
        if row['Numero_Embarque'] == 'NO_EMB' and row['Factura_Norm'] in mapa_fac_emb:
            return mapa_fac_emb[row['Factura_Norm']]
        return row['Numero_Embarque']

    df['Numero_Embarque'] = df.apply(backfill_embarque, axis=1)

    # --- 3. HERENCIA DE NIT ---
    shipments_with_nit = df[(df['NIT_Norm'] != 'ND') & (df['Numero_Embarque'] != 'NO_EMB')]
    mapa_emb_nit = shipments_with_nit.groupby('Numero_Embarque')['NIT_Norm'].first().to_dict()
    df['NIT_Reporte'] = df['Numero_Embarque'].map(mapa_emb_nit).fillna(df['NIT_Norm'])

    # Inicialización
    df['Conciliado'] = False
    df['Grupo_Conciliado'] = ""
    total_conciliados = 0 

    # --- FASE 1: POR EMBARQUE INDIVIDUAL (DOBLE LLAVE USD/BS) ---
    df_p1 = df[~df['Conciliado']]
    grupos_emb = df_p1[df_p1['Numero_Embarque'] != 'NO_EMB'].groupby('Numero_Embarque')
    
    for emb, grupo in grupos_emb:
        if len(grupo) < 2: continue
        
        suma_usd_abs = abs(round(grupo['Monto_USD'].sum(), 2))
        suma_bs_abs = abs(round(grupo['Monto_BS'].sum(), 2))
        
        # REGLA: Si cuadra en USD pero tiene diferencia en BS (> 1.00), va a AJUSTE
        if suma_usd_abs <= 0.01 and suma_bs_abs <= 1.00:
            df.loc[grupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"EMBARQUE_{emb}"]
            total_conciliados += len(grupo)
        elif suma_usd_abs <= 1.00:
            # Aquí entra: diferencia en USD de hasta $1.00 
            # O USD en cero pero con diferencia en BS
            df.loc[grupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"REQUIERE_AJUSTE_{emb}"]
            total_conciliados += len(grupo)

    # --- FASE 1.7: RESCATE DE HUÉRFANOS VS EMBARQUE (CON TOLERANCIA USD/BS) ---
    df_p1_7 = df[~df['Conciliado']]
    for nit, grupo_nit in df_p1_7.groupby('NIT_Reporte'):
        if nit == 'ND': continue
        abiertos = grupo_nit[grupo_nit['Numero_Embarque'] != 'NO_EMB'].copy()
        huerfanos = grupo_nit[grupo_nit['Numero_Embarque'] == 'NO_EMB'].copy()
        if abiertos.empty or huerfanos.empty: continue
        
        for emb, grupo_emb in abiertos.groupby('Numero_Embarque'):
            saldo_usd_emb = grupo_emb['Monto_USD'].sum()
            
            # Buscamos en los huérfanos alguno que cuadre en USD (Tolerancia $1.00)
            match_huerfano = huerfanos[ (huerfanos['Monto_USD'] + saldo_usd_emb).abs() <= 1.00 ]
            
            if not match_huerfano.empty:
                idx_huerfano = (match_huerfano['Monto_USD'] + saldo_usd_emb).abs().idxmin()
                indices = list(grupo_emb.index) + [idx_huerfano]
                
                # Verificamos saldos combinados para la etiqueta
                res_usd = abs(round(df.loc[indices, 'Monto_USD'].sum(), 2))
                res_bs = abs(round(df.loc[indices, 'Monto_BS'].sum(), 2))
                
                # Doble validación para decidir a qué pestaña va
                if res_usd <= 0.01 and res_bs <= 1.00:
                    etiqueta = f"RESCATE_HUERF_{emb}"
                else:
                    etiqueta = f"REQUIERE_AJUSTE_HUERF_{emb}"
                
                df.loc[indices, ['Conciliado', 'Grupo_Conciliado']] = [True, etiqueta]
                total_conciliados += len(indices)
                huerfanos = huerfanos.drop(idx_huerfano)

    # --- FASE 1.8: CRUCE POR COMBINATORIA DE EMBARQUES (BLINDADO USD/BS) ---
    df_p1_8 = df[~df['Conciliado']]
    for nit, grupo_nit in df_p1_8.groupby('NIT_Reporte'):
        if nit == 'ND': continue
        solo_emb = grupo_nit[grupo_nit['Numero_Embarque'] != 'NO_EMB']
        if solo_emb.empty: continue
        
        # Saldos pendientes por cada moneda
        saldos_usd_emb = solo_emb.groupby('Numero_Embarque')['Monto_USD'].sum().round(2)
        saldos_bs_emb = solo_emb.groupby('Numero_Embarque')['Monto_BS'].sum().round(2)
        
        embarques_lista = saldos_usd_emb[saldos_usd_emb.abs() > 0.01].index.tolist()
        embarques_usados = set()

        for r in range(2, min(len(embarques_lista) + 1, 5)):
            for combo in combinations(embarques_lista, r):
                if any(e in embarques_usados for e in combo): continue
                
                res_usd = abs(round(sum(saldos_usd_emb[list(combo)]), 2))
                res_bs = abs(round(sum(saldos_bs_emb[list(combo)]), 2))
                
                if res_usd <= 1.00:
                    indices_a_cerrar = solo_emb[solo_emb['Numero_Embarque'].isin(combo)].index
                    
                    # Decidimos etiqueta según doble llave
                    if res_usd <= 0.01 and res_bs <= 1.00:
                        etiqueta = f"COMB_EMB_{nit}"
                    else:
                        etiqueta = f"REQUIERE_AJUSTE_COMB_{nit}"
                    
                    df.loc[indices_a_cerrar, ['Conciliado', 'Grupo_Conciliado']] = [True, etiqueta]
                    total_conciliados += len(indices_a_cerrar)
                    embarques_usados.update(combo)

    # --- FASES FINALES DE SEGURIDAD (FUENTE, REFERENCIA, GLOBAL) ---
    # Fase 1.5: Match Fuente
    df_p1_5 = df[~df['Conciliado']]
    grupos_fuente = df_p1_5[df_p1_5['Fuente'].notna() & (df_p1_5['Fuente'] != '')].groupby(['NIT_Reporte', 'Fuente'])
    for (nit, fuente), grupo in grupos_fuente:
        if len(grupo) >= 2 and abs(round(grupo['Monto_USD'].sum(), 2)) <= 0.01:
            df.loc[grupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"FUENTE_{fuente[:15]}"]
            total_conciliados += len(grupo)

    # Fase 2: Match Referencia
    df_p2 = df[~df['Conciliado']]
    for (nit, ref), grupo in df_p2.groupby(['NIT_Reporte', 'Referencia']):
        if len(grupo) >= 2 and abs(round(grupo['Monto_USD'].sum(), 2)) <= 0.01:
            df.loc[grupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"REF_{ref[:15]}"]
            total_conciliados += len(grupo)

    # Fase 3: Saldo Global por NIT
    df_p3 = df[~df['Conciliado']]
    for nit, grupo in df_p3.groupby('NIT_Reporte'):
        if nit != 'ND' and len(grupo) >= 2 and abs(round(grupo['Monto_USD'].sum(), 2)) <= 0.01:
            df.loc[grupo.index, ['Conciliado', 'Grupo_Conciliado']] = [True, f"SALDO_NIT_{nit}"]
            total_conciliados += len(grupo)
            
    log_messages.append(f"✔️ Conciliación finalizada. Total: {total_conciliados} movimientos.")
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
    """
    Carga y prepara el archivo CP.
    CORREGIDO: Mapea 'Proveedor' como RIF y 'Nombre' como Nombre_Proveedor.
    """
    # 1. Leemos el archivo asumiendo encabezado en fila 5 (índice 4)
    df = pd.read_excel(file_cp, header=4, dtype=str)
    
    # 2. Limpieza de nombres de columnas
    df.columns = [str(col).strip() for col in df.columns]
    
    # 3. Mapeo específico según tu aclaratoria
    rename_map = {}
    
    for col in df.columns:
        c_upper = col.upper()
        
        # --- LÓGICA DE IDENTIFICACIÓN ---
        
        # Si la columna se llama explícitamente PROVEEDOR, el usuario indicó que ese es el RIF.
        if c_upper == 'PROVEEDOR':
            rename_map[col] = 'RIF'
            
        # Si la columna se llama RIF o NIT, obviamente es el RIF.
        elif c_upper in ['RIF', 'NIT', 'R.I.F.']:
            rename_map[col] = 'RIF'
            
        # Si la columna es NOMBRE o BENEFICIARIO, la guardamos aparte, NO como RIF.
        elif c_upper in ['NOMBRE', 'BENEFICIARIO', 'NOMBRE PROVEEDOR']:
            rename_map[col] = 'Nombre_Proveedor'
            
        # --- OTROS CAMPOS ---
        elif 'ASIENTO' in c_upper:
            rename_map[col] = 'Asiento'
        elif c_upper in ['TIPO']:
            rename_map[col] = 'Tipo'
        elif 'FECHA' in c_upper:
            rename_map[col] = 'Fecha'
        elif c_upper in ['NÚMERO', 'NUMERO', 'COMPROBANTE', 'NO. COMPROBANTE']:
            rename_map[col] = 'Comprobante'
        elif 'MONTO' in c_upper or 'IMPORTE' in c_upper:
            rename_map[col] = 'Monto'
        elif 'APLICACIÓN' in c_upper or 'APLICACION' in c_upper:
            rename_map[col] = 'Aplicacion'
        elif 'SUBTIPO' in c_upper:
            rename_map[col] = 'Subtipo'

    # Aplicar renombrado
    df.rename(columns=rename_map, inplace=True)

    # --- LIMPIEZA DE DUPLICADOS (SEGURIDAD) ---
    # Si por casualidad el archivo traía "Proveedor" Y "RIF", nos quedamos con la primera 'RIF' que aparezca
    df = df.loc[:, ~df.columns.duplicated()]

    # 4. Verificación
    if 'RIF' not in df.columns:
        cols_encontradas = ", ".join(df.columns)
        raise KeyError(f"No se encontró la columna 'Proveedor' (o RIF) en el archivo CP. Columnas detectadas: [{cols_encontradas}]")

    # 5. Normalización
    df['RIF_norm'] = df['RIF'].apply(_normalizar_rif)
    
    if 'Comprobante' in df.columns:
        df['Comprobante_norm'] = df['Comprobante'].apply(_normalizar_numerico)
    else:
        df['Comprobante_norm'] = ''
        
    if 'Aplicacion' in df.columns:
        df['Factura_norm'] = df['Aplicacion'].apply(_extraer_factura_cp)
    else:
        df['Factura_norm'] = ''

    # Limpieza de montos
    if 'Monto' in df.columns:
        df['Monto'] = df['Monto'].astype(str).str.replace(',', '.', regex=False)
        df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    else:
        df['Monto'] = 0.0

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

## --- Directorio de Cuentas (Versión Normalizada) ---
CUENTAS_CONOCIDAS = {normalize_account(acc) for acc in [
    '1.1.3.01.1.001', '1.1.3.01.1.901', '7.1.3.45.1.997', '6.1.1.12.1.001',
    '4.1.1.22.4.001', '2.1.3.04.1.001', '7.1.3.19.1.012', '2.1.2.05.1.108',
    '6.1.1.19.1.001', '4.1.1.21.4.001', '2.1.3.04.1.006', '2.1.3.01.1.012',
    '7.1.3.04.1.004', '7.1.3.06.1.998', '1.1.1.04.6.003', '1.1.4.01.7.020',
    '1.1.4.01.7.021', '1.1.1.02.1.004', '1.1.1.02.1.007', '1.1.1.02.1.009',
    '1.1.1.02.1.016', '1.1.1.02.1.112', '1.1.1.02.1.124', '1.1.1.02.1.132',
    '1.1.1.02.6.002', '1.1.1.02.6.003', '1.1.1.02.6.005', '1.1.1.02.6.010',
    '1.1.1.03.6.012', '1.1.1.03.6.024', '1.1.1.03.6.026', '1.1.1.03.6.031',
    # --- BANCOS ADICIONALES ---
    '1.1.1.02.1.002', '1.1.1.02.1.005', '1.1.1.02.6.001', '1.1.1.02.1.003',
    '4.1.1.21.4.001', '2.1.3.04.1.001', '4.1.1.22.4.001', '1.1.1.03.6.002', 
    '1.1.1.06.6.003', '1.1.1.02.1.018', '1.1.1.02.6.013', '1.1.1.03.6.028',
    # --- CUENTAS GRUPOS NUEVOS ---
    '1.9.1.01.3.008', # Inv. Oficinas
    '1.9.1.01.3.009', # Inv. Oficinas
    '7.1.3.01.1.001', # Deudores Incobrables
    '1.1.4.01.7.044', # CxC Varios ME
    '2.1.2.05.1.005'  # Asientos por Clasificar (NUEVA)
]}

CUENTAS_BANCO = {normalize_account(acc) for acc in [
    '1.1.4.01.7.020', '1.1.4.01.7.021', '1.1.1.02.1.004', '1.1.1.02.1.007',
    '1.1.1.02.1.009', '1.1.1.02.1.016', '1.1.1.02.1.112', '1.1.1.02.1.124',
    '1.1.1.02.1.132', '1.1.1.02.6.002', '1.1.1.02.6.003', '1.1.1.02.6.005',
    '1.1.1.02.6.010', '1.1.1.03.6.012', '1.1.1.03.6.024', '1.1.1.03.6.026',
    '1.1.1.03.6.031',
    '1.1.1.02.1.002', # Banco Venezolano de Credito
    '1.1.1.02.1.005', # Banesco
    '1.1.1.02.6.001', # Banco Mercantil
    '1.1.1.02.1.003', # Banco de Venezuela
    '1.1.1.02.1.018','1.1.1.02.6.013', '1.1.1.06.6.003', '1.1.1.03.6.002',
    '1.1.1.03.6.028', '1.1.1.03.6.002', '1.1.1.06.6.003', '1.1.1.02.1.018',
    '1.1.1.02.6.013', '1.1.1.03.6.028'
]}

def es_palabra_similiar(texto_completo, palabra_objetivo, umbral=0.80):
    """
    Busca si 'palabra_objetivo' está en 'texto_completo' permitiendo errores de tipeo.
    MEJORA: Separa palabras por cualquier símbolo para evitar uniones accidentales.
    """
    if not texto_completo: return False
    
    # 1. Reemplazar cualquier carácter que NO sea letra o número por ESPACIO
    # Ej: "TRANSPASO-SALDO" -> "TRANSPASO SALDO"
    texto_limpio = re.sub(r'[^A-Z0-9]', ' ', str(texto_completo).upper())
    
    # 2. Dividir en palabras
    palabras = texto_limpio.split()
    
    objetivo = palabra_objetivo.upper()
    
    for palabra in palabras:
        # Coincidencia Exacta
        if palabra == objetivo:
            return True
            
        # Optimización: Si la longitud varía mucho, no es la misma palabra (ahorra proceso)
        if abs(len(palabra) - len(objetivo)) > 3:
            continue
        
        # Coincidencia Difusa (Fuzzy)
        ratio = SequenceMatcher(None, palabra, objetivo).ratio()
        if ratio >= umbral:
            return True
            
    return False

def _get_base_classification(cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, monto_suma, monto_max_abs, is_reverso_check=False):
    """
    Clasifica el asiento basándose en reglas de jerarquía.
    """
    
    # --- PRIORIDAD 1: Notas de Crédito (Grupo 3) ---
    es_fuente_nc = 'N/C' in fuente_completa
    tiene_cuenta_descuento = normalize_account('4.1.1.22.4.001') in cuentas_del_asiento
    tiene_cuenta_iva = normalize_account('2.1.3.04.1.001') in cuentas_del_asiento
    
    if is_reverso_check or es_fuente_nc or tiene_cuenta_descuento:
        if tiene_cuenta_descuento or (es_fuente_nc and tiene_cuenta_iva):
            if is_reverso_check: return "Grupo 3: N/C" 
            if 'AVISOS DE CREDITO' in referencia_completa: return "Grupo 3: N/C - Avisos de Crédito"
            if referencia_limpia_palabras.intersection({'DIFERENCIAL', 'DIFERENCIA', 'CAMBIO', 'DIF'}): return "Grupo 3: N/C - Posible Error de Cuenta (Ref. Diferencial)"
            if referencia_limpia_palabras.intersection({'ESTRATEGIA', 'ESTRATEGIAS'}): return "Grupo 3: N/C - Estrategias"
            if referencia_limpia_palabras.intersection({'INCENTIVO', 'INCENTIVOS'}): return "Grupo 3: N/C - Incentivos"
            if referencia_limpia_palabras.intersection({'BONIFICACION', 'BONIFICACIONES', 'BONIF', 'BONF'}): return "Grupo 3: N/C - Bonificaciones"
            if referencia_limpia_palabras.intersection({'DESCUENTO', 'DESCUENTOS', 'DSCTO', 'DESC', 'DESTO'}): return "Grupo 3: N/C - Descuentos"
            return "Grupo 3: N/C - Otros"

    # --- PRIORIDAD 2: Gastos de Ventas (Grupo 4) ---
    keywords_mercadeo = {'EXHIBIDOR', 'EXHIBIDORES', 'OBSEQUIO', 'OBSEQUIOS', 'MERCADEO', 'PUBLICIDAD', 'PROPAGANDA'}
    tiene_cuenta_gasto = normalize_account('7.1.3.19.1.012') in cuentas_del_asiento
    tiene_texto_gasto = not keywords_mercadeo.isdisjoint(referencia_limpia_palabras)
    if tiene_cuenta_gasto or tiene_texto_gasto: 
        return "Grupo 4: Gastos de Ventas"

    # --- PRIORIDAD 3: Diferencial Cambiario PURO (Grupo 2) ---
    tiene_diferencial = normalize_account('6.1.1.12.1.001') in cuentas_del_asiento
    tiene_banco = not CUENTAS_BANCO.isdisjoint(cuentas_del_asiento)
    if tiene_diferencial and not tiene_banco:
        return "Grupo 2: Diferencial Cambiario"

    # --- PRIORIDAD 4: Retenciones (Grupo 9) ---
    if normalize_account('2.1.3.04.1.006') in cuentas_del_asiento: return "Grupo 9: Retenciones - IVA"
    if normalize_account('2.1.3.01.1.012') in cuentas_del_asiento: return "Grupo 9: Retenciones - ISLR"
    if normalize_account('7.1.3.04.1.004') in cuentas_del_asiento: return "Grupo 9: Retenciones - Municipal"

    # --- PRIORIDAD 5: Traspasos vs. Devoluciones (Grupo 10 y 7) ---
    if normalize_account('4.1.1.21.4.001') in cuentas_del_asiento:
        
        # CAMBIO INTELIGENTE: Detecta 'TRASPASO', 'TRANSPASO', 'TRAPASO', etc.
        es_traspaso = es_palabra_similiar(referencia_completa, 'TRASPASO')
        
        if es_traspaso and abs(monto_suma) <= TOLERANCIA_MAX_USD: 
            return "Grupo 10: Traspasos"
        
        if is_reverso_check: return "Grupo 7: Devoluciones y Rebajas"
        
        keywords_limpieza_dev = {'LIMPIEZA', 'LIMPIEZAS', 'SALDO', 'SALDOS', 'HISTORICO', 'AJUSTE', 'APLICAR', 'CRUCE', 'FAVOR', 'TRASLADO'}
        if not keywords_limpieza_dev.isdisjoint(referencia_limpia_palabras):
            if monto_max_abs <= 5: 
                return "Grupo 7: Devoluciones y Rebajas - Limpieza (<= $5)"
            else: 
                # Reutilizamos la lógica inteligente aquí también para traslados
                keywords_autorizadas = ['TRASLADO', 'APLICAR', 'CRUCE', 'RECLASIFICACION', 'CORRECCION', 'TRASPASO']
                
                # Verificamos si ALGUNA palabra clave está presente (con tolerancia a errores)
                es_autorizado = False
                for key in keywords_autorizadas:
                    if es_palabra_similiar(referencia_completa, key):
                        es_autorizado = True
                        break
                
                if es_autorizado:
                     return "Grupo 7: Devoluciones y Rebajas - Traslados/Cruce"
                return "Grupo 7: Devoluciones y Rebajas - Limpieza (> $5)"
        else: 
            return "Grupo 7: Devoluciones y Rebajas - Otros Ajustes"

    # --- PRIORIDAD 6: Cobranzas (Grupo 8) ---
    # MODIFICACIÓN: Ahora consideramos Cobranza si:
    # 1. Tiene texto 'RECIBO', 'TEF', 'DEPR'
    # 2. O tiene cuenta de Banco Real
    # 3. O tiene cuenta de Fondos por Depositar (1.1.1.04.6.003) <--- NUEVO
    
    is_cobranza_texto = 'RECIBO DE COBRANZA' in referencia_completa or 'TEF' in fuente_completa or 'DEPR' in fuente_completa
    tiene_cuenta_fondos = normalize_account('1.1.1.04.6.003') in cuentas_del_asiento
    
    if is_cobranza_texto or tiene_banco or tiene_cuenta_fondos:
        if is_reverso_check: return "Grupo 8: Cobranzas"
        
        if normalize_account('6.1.1.12.1.001') in cuentas_del_asiento: return "Grupo 8: Cobranzas - Con Diferencial Cambiario"
        if tiene_cuenta_fondos: return "Grupo 8: Cobranzas - Fondos por Depositar"
        
        if tiene_banco:
            if 'TEF' in fuente_completa: return "Grupo 8: Cobranzas - TEF (Bancos)"
            return "Grupo 8: Cobranzas - Recibos (Bancos)"
        return "Grupo 8: Cobranzas - Otros"

    # --- PRIORIDAD 7: Ingresos Varios (Grupo 6) ---
    if normalize_account('6.1.1.19.1.001') in cuentas_del_asiento:
        if is_reverso_check: return "Grupo 6: Ingresos Varios"
        keywords_limpieza = {'LIMPIEZA', 'LIMPIEZAS', 'SALDO', 'SALDOS', 'HISTORICO', 'INGRESOS', 'INGRESO', 'AJUSTE'}
        if not keywords_limpieza.isdisjoint(referencia_limpia_palabras):
            if monto_max_abs <= 25: return "Grupo 6: Ingresos Varios - Limpieza (<= $25)"
            else: return "Grupo 6: Ingresos Varios - Limpieza (> $25)"
        else: return "Grupo 6: Ingresos Varios - Otros"
            
    # --- RESTO DE PRIORIDADES (Grupos Específicos) ---
    ctas_inversion = {normalize_account('1.9.1.01.3.008'), normalize_account('1.9.1.01.3.009')}
    if not ctas_inversion.isdisjoint(cuentas_del_asiento):
        return "Grupo 14: Inv. entre Oficinas"

    if normalize_account('7.1.3.01.1.001') in cuentas_del_asiento:
        return "Grupo 15: Deudores Incobrables"

    if normalize_account('1.1.4.01.7.044') in cuentas_del_asiento:
        return "Grupo 16: Cuentas por Cobrar - Varios en ME"

    # --- NUEVO GRUPO 17 ---
    if normalize_account('2.1.2.05.1.005') in cuentas_del_asiento:
        return "Grupo 17: Asientos por Clasificar"
    # ----------------------

    if normalize_account('7.1.3.06.1.998') in cuentas_del_asiento: return "Grupo 12: Perdida p/Venta o Retiro Activo ND"
    if normalize_account('7.1.3.45.1.997') in cuentas_del_asiento: return "Grupo 1: Acarreos y Fletes Recuperados"
    if normalize_account('2.1.2.05.1.108') in cuentas_del_asiento: return "Grupo 5: Haberes de Clientes"

    return "No Clasificado"

def _clasificar_asiento_paquete_cc(cuentas_del_asiento, referencia_completa, fuente_completa, monto_suma, monto_max_abs):
    """
    Función adaptada para recibir parámetros pre-calculados.
    """
    referencia_limpia_palabras = set(re.sub(r'[^\w\s]', '', referencia_completa).split())

    # CAPA 1: Detección de Reversos
    if 'REVERSO' in referencia_completa or 'REV' in referencia_limpia_palabras:
        base_group = _get_base_classification(cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, monto_suma, monto_max_abs, is_reverso_check=True)
        
        if base_group != "No Clasificado":
            parts = base_group.split(':', 1)
            group_number = parts[0].strip()
            description = parts[1].split('-')[0].strip()
            return f"{group_number}: Reversos - {description}"
        else:
            return "Grupo 11: Reversos No Identificados"
            
    # CAPA 2: Estándar
    return _get_base_classification(cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, monto_suma, monto_max_abs, is_reverso_check=False)


def _validar_asiento(asiento_group):
    """
    Recibe un asiento completo (ya clasificado) y aplica las reglas de negocio.
    """
    grupo = asiento_group['Grupo'].iloc[0]
    
    # GRUPO 1: FLETES
    if grupo.startswith("Grupo 1:"):
        fletes_lines = asiento_group[asiento_group['Cuenta Contable Norm'] == normalize_account('7.1.3.45.1.997')]
        if not fletes_lines['Referencia'].str.contains('FLETE', case=False, na=False).all():
            return "Incidencia: Referencia sin 'FLETE' encontrada."
            
    # GRUPO 2: DIFERENCIAL CAMBIARIO
    elif grupo.startswith("Grupo 2:"):
        diff_lines = asiento_group[asiento_group['Cuenta Contable Norm'] == normalize_account('6.1.1.12.1.001')]
        keywords_estrictas = ['DIFERENCIAL', 'DIFERENCIA', 'DIF CAMBIARIO', 'DIF.', 'TASA', 'AJUSTE', 'IVA', 'DC']
        patron_regex = '|'.join([re.escape(k) if '.' in k else k for k in keywords_estrictas])
        
        for ref in diff_lines['Referencia']:
            ref_str = str(ref).upper()
            if re.search(patron_regex, ref_str): continue
            
            # Fuzzy match
            palabras_referencia = ref_str.split()
            objetivos = ['DIFERENCIAL', 'CAMBIARIO', 'DIFERENCIA', 'AJUSTE']
            es_typo_aceptable = False
            for palabra in palabras_referencia:
                p_clean = re.sub(r'[^A-Z]', '', palabra)
                for objetivo in objetivos:
                    ratio = SequenceMatcher(None, p_clean, objetivo).ratio()
                    if ratio > 0.80: 
                        es_typo_aceptable = True; break
                if es_typo_aceptable: break
            if not es_typo_aceptable: return f"Incidencia: Referencia '{ref}' no parece indicar Diferencial Cambiario."
            
    # GRUPO 6: INGRESOS VARIOS
    elif grupo.startswith("Grupo 6:"):
        if (asiento_group['Monto_USD'].abs() > 25).any():
            return "Incidencia: Movimiento mayor al límite permitido ($25)."
            
    # --- GRUPO 7: DEVOLUCIONES ---
    elif grupo.startswith("Grupo 7:"):
        # Regla Semántica (Diferencial)
        referencia_upper = str(asiento_group['Referencia'].iloc[0]).upper()
        keywords_error_cuenta = ['DIFERENCIAL', 'DIF. CAMBIARIO', 'DIF CAMBIARIO', 'TASA', 'DIFF', 'CAMBIO']
        if any(k in referencia_upper for k in keywords_error_cuenta) and "PRECIO" not in referencia_upper:
            return "Incidencia: Referencia indica 'Diferencial/Tasa' pero usa cuenta de Devoluciones."

        # Regla de Monto con Inteligencia Artificial
        if (asiento_group['Monto_USD'].abs() > 5).any():
            keywords_autorizadas = ['TRASLADO', 'APLICAR', 'CRUCE', 'RECLASIFICACION', 'CORRECCION', 'TRASPASO']
            
            # Verificamos si alguna de las palabras autorizadas está en la referencia (con fuzzy match)
            es_valido = False
            for key in keywords_autorizadas:
                if es_palabra_similiar(referencia_upper, key):
                    es_valido = True
                    break
            
            if not es_valido:
                return "Incidencia: Movimiento mayor a $5 (y no indica ser Traslado/Cruce)."

    # GRUPO 9: RETENCIONES
    elif grupo.startswith("Grupo 9:"):
        referencia_str = str(asiento_group['Referencia'].iloc[0]).upper().strip()
        tiene_numeros = any(char.isdigit() for char in referencia_str)
        tiene_keywords = any(k in referencia_str for k in ['RET', 'IMP', 'ISLR', 'IVA', 'MUNICIPAL'])
        if not (tiene_numeros or tiene_keywords):
            return f"Incidencia: Referencia '{referencia_str}' inválida."

    # GRUPO 3: NOTAS DE CRÉDITO
    elif grupo.startswith("Grupo 3:"):
        if "Error de Cuenta" in grupo:
            return "Incidencia: Diferencial Cambiario registrado en cuenta de Descuentos/NC."
        cuentas_presentes = set(asiento_group['Cuenta Contable Norm'])
        tiene_descuento = normalize_account('4.1.1.22.4.001') in cuentas_presentes
        tiene_iva = normalize_account('2.1.3.04.1.001') in cuentas_presentes
        if not (tiene_descuento and tiene_iva):
            return "Incidencia: Asiento de N/C incompleto (Falta Descuento o IVA)."

    # GRUPO 10: TRASPASOS
    elif grupo.startswith("Grupo 10:"):
        if not np.isclose(asiento_group['Monto_USD'].sum(), 0, atol=TOLERANCIA_MAX_USD):
            return "Incidencia: El traspaso no suma cero."
        if not ((asiento_group['Monto_USD'] > 0).any() and (asiento_group['Monto_USD'] < 0).any()):
             return "Incidencia: Traspaso incompleto (Falta contrapartida)."

    # GRUPO 17
    elif grupo.startswith("Grupo 17:"):
        return "Incidencia: Cuenta Transitoria. Verificar cruce en Mayor antes de mayorizar."
    
    # GRUPOS AUTOMÁTICOS (14, 15, 16)
    elif grupo.startswith("Grupo 14:") or grupo.startswith("Grupo 15:") or grupo.startswith("Grupo 16:"):
        return "Conciliado"

    # NO CLASIFICADOS
    elif grupo.startswith("Grupo 11") or grupo == "No Clasificado":
        return f"Incidencia: Revisión requerida. {grupo}"

    return "Conciliado"

def run_analysis_paquete_cc(df_diario, log_messages):
    """
    Función principal optimizada.
    Fase 3 BLINDADA: Protege las Cobranzas (Grupo 8) de ser marcadas como reversos falsos.
    """
    log_messages.append("--- INICIANDO ANÁLISIS Y VALIDACIÓN DE PAQUETE CC (ULTRA RÁPIDO) ---")
    
    df = df_diario.copy()
    
    # --- PASO 0: NORMALIZACIÓN ---
    rename_map = {}
    for col in df.columns:
        c_upper = col.strip().upper()
        if c_upper in ['NIT', 'RIF', 'R.I.F.', 'CEDULA']:
            rename_map[col] = 'NIT'
        elif c_upper in ['DESCRIPCIÓN NIT', 'DESCRIPCION NIT', 'NOMBRE', 'CLIENTE', 'NOMBRE DEL PROVEEDOR']:
            rename_map[col] = 'Nombre'
            
    df.rename(columns=rename_map, inplace=True)
    if 'NIT' not in df.columns: df['NIT'] = ''
    if 'Nombre' not in df.columns: df['Nombre'] = ''
    df['NIT'] = df['NIT'].fillna('')
    df['Nombre'] = df['Nombre'].fillna('')

    # Limpieza
    df['Cuenta Contable Norm'] = df['Cuenta Contable'].astype(str).str.replace(r'\D', '', regex=True)
    df['Monto_USD'] = (df['Débito Dolar'] - df['Crédito Dolar']).round(2)
    df['Ref_Str'] = df['Referencia'].astype(str).fillna('').str.upper()
    df['Fuente_Str'] = df['Fuente'].astype(str).fillna('').str.upper()
    
    log_messages.append("⚙️ Pre-calculando metadatos por asiento...")
    
    # --- FASE 1: AGREGACIÓN MASIVA ---
    df_grouped = df.groupby('Asiento')
    
    s_cuentas = df_grouped['Cuenta Contable Norm'].apply(set)
    s_ref = df_grouped['Ref_Str'].apply(lambda x: ' '.join(x.unique()))
    s_fuente = df_grouped['Fuente_Str'].apply(lambda x: ' '.join(x.unique()))
    s_suma = df_grouped['Monto_USD'].sum()
    s_max_abs = df_grouped['Monto_USD'].apply(lambda x: x.abs().max())
    
    df_meta = pd.DataFrame({
        'Cuentas': s_cuentas, 'Ref': s_ref, 'Fuente': s_fuente, 'Suma': s_suma, 'Max_Abs': s_max_abs
    })
    
    # --- FASE 2: CLASIFICACIÓN ---
    mapa_grupos = {}
    asientos_con_cuentas_nuevas = 0
    
    for asiento_id, row in df_meta.iterrows():
        cuentas_del_asiento = row['Cuentas']
        cuentas_desconocidas = cuentas_del_asiento - CUENTAS_CONOCIDAS
        if cuentas_desconocidas:
            lista_faltantes = ", ".join(sorted(cuentas_desconocidas))
            grupo_asignado = f"Grupo 11: Cuentas No Identificadas ({lista_faltantes})"
            asientos_con_cuentas_nuevas += 1
        else:
            grupo_asignado = _clasificar_asiento_paquete_cc(
                cuentas_del_asiento, row['Ref'], row['Fuente'], row['Suma'], row['Max_Abs']
            )
        mapa_grupos[asiento_id] = grupo_asignado

    df['Grupo'] = df['Asiento'].map(mapa_grupos)
    
    # --- FASE 3: INTELIGENCIA DE REVERSOS (BLINDADA) ---
    log_messages.append("🧠 Ejecutando cruce inteligente de reversos...")
    
    mapa_cambio_grupo = {}
    procesados = set()
    
    # 1. Definición de Reverso
    def es_potencial_reverso(row_meta):
        texto_full = (row_meta['Ref'] + " " + row_meta['Fuente']).upper()
        # Palabras clave más estrictas (quitamos 'ERROR' solo para evitar falsos en descripciones)
        keywords = ['REVERSO', 'REV ', 'ANULA', 'NO CORRESPONDE', 'CORRECCION', 'DEVOLUCION DEL ASIENTO']
        return any(k in texto_full for k in keywords)

    # Identificación inicial
    ids_por_grupo = df[df['Grupo'].astype(str).str.contains("Reverso", case=False, na=False)]['Asiento'].unique()
    ids_por_texto = df_meta[df_meta.apply(es_potencial_reverso, axis=1)].index.tolist()
    ids_candidatos_reverso = set(list(ids_por_grupo) + ids_por_texto)
    
    # --- FILTRO DE INMUNIDAD (La corrección clave) ---
    ids_reversos_final = set()
    for aid in ids_candidatos_reverso:
        grupo_actual = mapa_grupos.get(aid, "")
        ref_actual = df_meta.loc[aid, 'Ref']
        
        # Si es Cobranza (Grupo 8), SOLO es reverso si dice explícitamente "REVERSO"
        if grupo_actual.startswith("Grupo 8"):
            if "REVERSO" in ref_actual:
                ids_reversos_final.add(aid)
        else:
            # Para otros grupos, aceptamos la detección normal
            ids_reversos_final.add(aid)
    # -------------------------------------------------

    # Mapa de montos global
    df_todos = df.groupby(['Asiento'])['Monto_USD'].sum().round(2).reset_index()
    mapa_montos_global = {}
    for _, row in df_todos.iterrows():
        m = row['Monto_USD']
        if m not in mapa_montos_global: mapa_montos_global[m] = []
        mapa_montos_global[m].append(row['Asiento'])

    # A. Procesar Reversos
    for id_rev in ids_reversos_final:
        if id_rev in procesados: continue
        
        monto_rev = df_meta.loc[id_rev, 'Suma']
        monto_target = round(-monto_rev, 2)
        posibles = [p for p in mapa_montos_global.get(monto_target, []) if p not in procesados]
        
        if not posibles: continue
        
        ref_rev = df_meta.loc[id_rev, 'Ref'] + " " + df_meta.loc[id_rev, 'Fuente']
        numeros_clave = re.findall(r'\d+', ref_rev)
        match_final = None
        
        # Estrategia Fuerte
        for cand_id in posibles:
            ref_cand = df_meta.loc[cand_id, 'Ref'] + " " + df_meta.loc[cand_id, 'Fuente']
            for num in numeros_clave:
                if len(num) > 3 and num in ref_cand:
                    match_final = cand_id; break
            if match_final: break
            
        # Estrategia Débil
        if not match_final and len(posibles) == 1:
            cand_unico = posibles[0]
            grp_rev = mapa_grupos.get(id_rev, "")
            grp_cand = mapa_grupos.get(cand_unico, "")
            # Protección Retenciones y Cobranzas
            if not (grp_rev.startswith("Grupo 9") or grp_cand.startswith("Grupo 9") or 
                    grp_rev.startswith("Grupo 8") or grp_cand.startswith("Grupo 8")):
                match_final = cand_unico
            
        if match_final:
            mapa_cambio_grupo[id_rev] = "Grupo 13: Operaciones Reversadas / Anuladas"
            mapa_cambio_grupo[match_final] = "Grupo 13: Operaciones Reversadas / Anuladas"
            procesados.update([id_rev, match_final])

    # B. Barrido de Emparejamiento (Protegido)
    ids_restantes = [i for i in df_meta.index if i not in procesados and i not in ids_reversos_final]
    mapa_abs = {}
    for aid in ids_restantes:
        m = abs(df_meta.loc[aid, 'Suma'])
        if m > 0.01:
            if m not in mapa_abs: mapa_abs[m] = []
            mapa_abs[m].append(aid)
            
    for monto, candidatos in mapa_abs.items():
        if len(candidatos) < 2: continue
        pos = [c for c in candidatos if df_meta.loc[c, 'Suma'] > 0]
        neg = [c for c in candidatos if df_meta.loc[c, 'Suma'] < 0]
        
        for p_id in pos:
            if p_id in procesados: continue
            
            # Protección extra: No cruzar Cobranzas (Grupo 8) automáticamente aquí
            if mapa_grupos.get(p_id, "").startswith("Grupo 8"): continue

            ref_p = df_meta.loc[p_id, 'Ref'] + " " + df_meta.loc[p_id, 'Fuente']
            nums_p = set([n for n in re.findall(r'\d+', ref_p) if len(n) > 3])
            if not nums_p: continue 
            
            for n_id in neg:
                if n_id in procesados: continue
                # Protección extra Grupo 8
                if mapa_grupos.get(n_id, "").startswith("Grupo 8"): continue

                ref_n = df_meta.loc[n_id, 'Ref'] + " " + df_meta.loc[n_id, 'Fuente']
                for num in nums_p:
                    if num in ref_n:
                        mapa_cambio_grupo[p_id] = "Grupo 13: Operaciones Reversadas / Anuladas"
                        mapa_cambio_grupo[n_id] = "Grupo 13: Operaciones Reversadas / Anuladas"
                        procesados.update([p_id, n_id])
                        break
                if p_id in procesados: break

    if mapa_cambio_grupo:
        df['Grupo'] = df['Asiento'].map(mapa_cambio_grupo).fillna(df['Grupo'])

    # --- FASE 4: VALIDACIÓN Y ORDEN ---
    resultados_validacion = {}
    for asiento_id, asiento_group in df.groupby('Asiento'):
        if asiento_group['Grupo'].iloc[0].startswith("Grupo 13"):
            val = "Conciliado (Anulado)"
        else:
            val = _validar_asiento(asiento_group)
        resultados_validacion[asiento_id] = val
        
    df['Estado'] = df['Asiento'].map(resultados_validacion)
    
    if asientos_con_cuentas_nuevas > 0:
        log_messages.append(f"⚠️ Se encontraron {asientos_con_cuentas_nuevas} asientos con cuentas no registradas.")

    df['Orden_Prioridad'] = df['Estado'].apply(lambda x: 1 if str(x).startswith('Conciliado') else 0)
    
    # ORDENAR PRIMERO
    df_sorted = df.sort_values(by=['Grupo', 'Orden_Prioridad', 'Asiento'], ascending=[True, True, True])
    
    cols_drop = ['Ref_Str', 'Fuente_Str', 'Cuenta Contable Norm', 'Monto_USD', 'Orden_Prioridad']
    df_final = df_sorted.drop(columns=cols_drop, errors='ignore')
    
    log_messages.append("--- ANÁLISIS FINALIZADO CON ÉXITO ---")
    return df_final

# ==============================================================================
# LÓGICA PARA CUADRE CB - CG (TESORERÍA VS CONTABILIDAD) - VERSIÓN FINAL BLINDADA
# ==============================================================================
import pdfplumber

# 1. DICCIONARIO MAESTRO DE NOMBRES
NOMBRES_CUENTAS_OFICIALES = {
    '1.1.1.02.1.000': 'Bancos del País',
    '1.1.1.02.1.002': 'Banco Venezolano de Credito, S.A.',
    '1.1.1.02.1.003': 'Banco de Venezuela, S.A. Banco Universal',
    '1.1.1.02.1.004': 'Banco Provincial, S.A. Banco Universal',
    '1.1.1.02.1.005': 'Banesco, C.A. Banco Universal',
    '1.1.1.02.1.006': 'Banco Provincial S.A. Banco Universal',
    '1.1.1.02.1.008': 'Banco Bicentenario, Banco Universal',
    '1.1.1.02.1.009': 'Banco Mercantil C.A. Banco Universal',
    '1.1.1.02.1.010': 'Banco del Caribe C.A. Banco Universal',
    '1.1.1.02.1.015': 'Banco Exterior S.A. Banco Universal',
    '1.1.1.02.1.018': 'Banco Nacional de Cdto.C.A. Bco.Univer.',
    '1.1.1.02.1.019': 'Banco Caroní, C.A. Banco Universal',
    '1.1.1.02.1.021': 'Bancamiga Banco Universal, C.A.',
    '1.1.1.02.1.022': 'Banco Sofitasa Banco Universal, C.A.',
    '1.1.1.02.1.111': 'Banco Exterior C.A. Banco Universal',
    '1.1.1.02.1.112': 'Venezolano de Credito S.A. Bco.Universal',
    '1.1.1.02.1.115': 'Bicentenario Banco Universal, C.A.',
    '1.1.1.02.1.116': 'Venezolano de Credito S.A. Bco.Universal',
    '1.1.1.02.1.124': 'Banco Nacional de Cdto.C.A. Bco.Univer.',
    '1.1.1.02.1.126': 'Del Sur Banco Universal, C.A.',
    '1.1.1.02.1.132': 'Banplus, C.A Banco Universal',
    '1.1.1.02.6.000': 'Bancos del Pais en Monedas Extranjeras',
    '1.1.1.02.6.001': 'Banco Mercantil Banco Universal',
    '1.1.1.02.6.002': 'Banco Nacional de Crédito C.A.',
    '1.1.1.02.6.003': 'Banco de Venezuela S.A. Banco Universal',
    '1.1.1.02.6.005': 'Banesco, C.A. Banco Universal',
    '1.1.1.02.6.006': 'Bancaribe',
    '1.1.1.02.6.010': 'Banplus (US$)',
    '1.1.1.02.6.011': 'Bancamiga Banco Universal',
    '1.1.1.02.6.013': 'Banco Provincial, s.a.Banco Universal',
    '1.1.1.02.6.015': 'Banco Sofitasa, Banco Universal, C.A.',
    '1.1.1.02.6.017': 'Banco Banesco (US$) - Cta. Custodia',
    '1.1.1.02.6.210': 'Banplus (EUR)',
    '1.1.1.02.6.213': 'Banco de Venezuela (EUR)',
    '1.1.1.02.6.214': 'Banco Sofitasa, Banco Universal C.A(COP)',
    '1.1.1.02.6.995': 'Bancos del País - Dif.en Cambio No Reali',
    '1.1.1.03.0.000': 'Bancos del Exterior',
    '1.1.1.03.6.000': 'Bancos del Exterior',
    '1.1.1.03.6.002': 'Amerant Bank, N.A.',
    '1.1.1.03.6.012': 'Banesco, S.A. Panamá',
    '1.1.1.03.6.015': 'Santander Private Banking',
    '1.1.1.03.6.024': 'Venezolano de Crédito Cayman Branch',
    '1.1.1.03.6.026': 'Banco Mercantil Panamá',
    '1.1.1.03.6.028': 'FACEBANK International',
    '1.1.1.03.6.031': 'Banesco USA',
    '1.1.1.06.6.000': 'Monedero Electrónico - Moneda Extranjera',
    '1.1.1.06.6.001': 'PayPal',
    '1.1.1.06.6.002': 'Creska',
    '1.1.1.06.6.003': 'Zinli',
    '1.1.4.01.7.020': 'Servicios de Administración de Fondos -Z',
    '1.1.4.01.7.021': 'Servicios de Administración de Fondos - USDT',
    '1.1.1.01.6.001': 'Cuenta Dolares',
    '1.1.1.01.6.002': 'Cuenta Euros'
}

# 2. MAPEO DE CÓDIGOS
MAPEO_CB_CG_BEVAL = {
    "0102E":  {"cta": "1.1.1.02.6.003", "moneda": "USD"},
    "0102EU": {"cta": "1.1.1.02.6.213", "moneda": "EUR"},
    "0102L":  {"cta": "1.1.1.02.1.003", "moneda": "VES"},
    "0104L":  {"cta": "1.1.1.02.1.002", "moneda": "VES"},
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105L":  {"cta": "1.1.1.02.1.009", "moneda": "VES"},
    "0108E":  {"cta": "1.1.1.02.6.013", "moneda": "USD"},
    "0108L":  {"cta": "1.1.1.02.1.004", "moneda": "VES"},
    "0114E":  {"cta": "1.1.1.02.6.006", "moneda": "USD"},
    "0114L":  {"cta": "1.1.1.02.1.010", "moneda": "VES"},
    "0115L":  {"cta": "1.1.1.02.1.015", "moneda": "VES"},
    "0134E":  {"cta": "1.1.1.02.6.017", "moneda": "USD"},
    "0134EC": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0134L":  {"cta": "1.1.1.02.1.005", "moneda": "VES"},
    "0137CP": {"cta": "1.1.1.02.6.214", "moneda": "COP"},
    "0137E":  {"cta": "1.1.1.02.6.015", "moneda": "USD"},
    "0137L":  {"cta": "1.1.1.02.1.022", "moneda": "VES"},
    "0172E":  {"cta": "1.1.1.02.6.011", "moneda": "USD"},
    "0172L":  {"cta": "1.1.1.02.1.021", "moneda": "VES"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174EU": {"cta": "1.1.1.02.6.210", "moneda": "EUR"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.008", "moneda": "VES"},
    "0191E":  {"cta": "1.1.1.02.6.002", "moneda": "USD"},
    "0191L":  {"cta": "1.1.1.02.1.018", "moneda": "VES"},
    "0201E":  {"cta": "1.1.1.03.6.012", "moneda": "USD"},
    "0202E":  {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0203E":  {"cta": "1.1.4.01.7.020", "moneda": "USD"},
    "0204E":  {"cta": "1.1.1.03.6.028", "moneda": "USD"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0206E":  {"cta": "1.1.1.06.6.001", "moneda": "USD"},
    "0207E":  {"cta": "1.1.4.01.7.021", "moneda": "USD"},
    "0209E":  {"cta": "1.1.1.01.6.001", "moneda": "USD"},
    "0210EU": {"cta": "1.1.1.01.6.002", "moneda": "EUR"},
    "0211E":  {"cta": "1.1.1.03.6.015", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "2407E":  {"cta": "1.1.1.06.6.003", "moneda": "USD"},
}

MAPEO_CB_CG_FEBECA = {
    "0102EU": {"cta": "1.1.1.02.6.213", "moneda": "EUR"},
    "0102E":  {"cta": "1.1.1.02.6.003", "moneda": "USD"},
    "0102L":  {"cta": "1.1.1.02.1.016", "moneda": "VES"},
    "0104L":  {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0104LP": {"cta": "1.1.1.02.1.116", "moneda": "VES"}, 
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105L":  {"cta": "1.1.1.02.1.009", "moneda": "VES"},
    "0108E":  {"cta": "1.1.1.02.6.013", "moneda": "USD"},
    "0108L":  {"cta": "1.1.1.02.1.004", "moneda": "VES"},
    "0114E":  {"cta": "1.1.1.02.6.006", "moneda": "USD"},
    "0114L":  {"cta": "1.1.1.02.1.011", "moneda": "VES"},
    "0115L":  {"cta": "1.1.1.02.1.111", "moneda": "VES"},
    "0134E2": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0134EC": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0134L":  {"cta": "1.1.1.02.1.007", "moneda": "VES"},
    "0134L2": {"cta": "1.1.1.02.1.007", "moneda": "VES"}, 
    "0137CP": {"cta": "1.1.1.02.6.214", "moneda": "COP"},
    "0137E":  {"cta": "1.1.1.02.6.015", "moneda": "USD"},
    "0137L":  {"cta": "1.1.1.02.1.022", "moneda": "VES"},
    "0172E":  {"cta": "1.1.1.02.6.011", "moneda": "USD"},
    "0172L":  {"cta": "1.1.1.02.1.021", "moneda": "VES"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174EU": {"cta": "1.1.1.02.6.210", "moneda": "EUR"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.115", "moneda": "VES"},
    "0191E":  {"cta": "1.1.1.02.6.002", "moneda": "USD"},
    "0191L":  {"cta": "1.1.1.02.1.124", "moneda": "VES"},
    "0201E":  {"cta": "1.1.1.03.6.012", "moneda": "USD"},
    "0202E":  {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0202E2": {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0202E3": {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0203E":  {"cta": "1.1.4.01.7.020", "moneda": "USD"},
    "0204E":  {"cta": "1.1.1.03.6.028", "moneda": "USD"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0206E":  {"cta": "1.1.1.06.6.001", "moneda": "USD"},
    "0207E":  {"cta": "1.1.4.01.7.021", "moneda": "USD"},
    "0211E":  {"cta": "1.1.1.03.6.015", "moneda": "USD"},
    "0212E":  {"cta": "1.1.1.03.6.031", "moneda": "USD"},
    "0212E2":  {"cta": "1.1.1.03.6.031", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "2407E":  {"cta": "1.1.1.06.6.003", "moneda": "USD"},
}

MAPEO_CB_CG_PRISMA = {
    "0104L":  {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105L":  {"cta": "1.1.1.02.1.003", "moneda": "VES"},
    "0114L":  {"cta": "1.1.1.02.1.005", "moneda": "VES"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.115", "moneda": "VES"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0209E":  {"cta": "1.1.1.01.6.001", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
}

# 5. MAPEO SILLACA - VERSIÓN LIMPIA Y SIN DUPLICADOS
MAPEO_CB_CG_SILLACA = {
    # --- VES ---
    "0102L":  {"cta": "1.1.1.02.1.016", "moneda": "VES"},
    "0104L":  {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0104L2": {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0105L":  {"cta": "1.1.1.02.1.009", "moneda": "VES"},
    "0105L2": {"cta": "1.1.1.02.1.009", "moneda": "VES"},
    "0108L":  {"cta": "1.1.1.02.1.004", "moneda": "VES"},
    "0114L":  {"cta": "1.1.1.02.1.011", "moneda": "VES"},
    "0115L":  {"cta": "1.1.1.02.1.015", "moneda": "VES"},
    "0134L":  {"cta": "1.1.1.02.1.007", "moneda": "VES"},
    "0137L":  {"cta": "1.1.1.02.1.022", "moneda": "VES"},
    "0172L":  {"cta": "1.1.1.02.1.021", "moneda": "VES"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.115", "moneda": "VES"},
    "0191L":  {"cta": "1.1.1.02.1.124", "moneda": "VES"},
    # --- USD ---
    "0102E":  {"cta": "1.1.1.02.6.003", "moneda": "USD"},
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105E2": {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0108E":  {"cta": "1.1.1.02.6.013", "moneda": "USD"},
    "0114E":  {"cta": "1.1.1.02.6.006", "moneda": "USD"},
    "0134EC": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0137E":  {"cta": "1.1.1.02.6.015", "moneda": "USD"},
    "0172E":  {"cta": "1.1.1.02.6.011", "moneda": "USD"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174E2": {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0191E":  {"cta": "1.1.1.02.6.002", "moneda": "USD"},
    "0201E":  {"cta": "1.1.1.03.6.012", "moneda": "USD"},
    "0202E":  {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0203E":  {"cta": "1.1.4.01.7.020", "moneda": "USD"},
    "0204E":  {"cta": "1.1.1.03.6.028", "moneda": "USD"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0205E2": {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0206E":  {"cta": "1.1.1.06.6.001", "moneda": "USD"},
    "0207E":  {"cta": "1.1.4.01.7.021", "moneda": "USD"},
    "0209E":  {"cta": "1.1.1.01.6.001", "moneda": "USD"},
    "0211E":  {"cta": "1.1.1.03.6.015", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "0501E2": {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "2407E":  {"cta": "1.1.1.06.6.003", "moneda": "USD"},
    # --- OTROS ---
    "0102EU": {"cta": "1.1.1.02.6.213", "moneda": "EUR"},
    "0174EU": {"cta": "1.1.1.02.6.210", "moneda": "EUR"},
    "0210EU": {"cta": "1.1.1.01.6.002", "moneda": "EUR"},
    "0137CP": {"cta": "1.1.1.02.6.214", "moneda": "COP"},
}
def limpiar_monto_pdf(texto):
    """
    Convierte texto a float. Maneja formatos US/VE, paréntesis y guiones (-).
    """
    if not texto: return 0.0
    t = str(texto).strip()
    
    # CASO: Guion solo (Cero contable)
    if t == '-': return 0.0
    
    # Limpieza general
    t = t.replace(' ', '').replace('$', '').replace('Bs', '')
    
    # Manejo de negativos (1.00) -> -1.00
    signo = 1
    if '(' in t and ')' in t:
        signo = -1
        t = t.replace('(', '').replace(')', '')
    elif '-' in t:
        signo = -1
        t = t.replace('-', '') # Quitamos el menos para procesar el número
        
    # Detección de formato decimal (Inteligente)
    idx_punto = t.rfind('.')
    idx_coma = t.rfind(',')

    if idx_punto > idx_coma:
        t = t.replace(',', '') # US: 1,234.56 -> 1234.56
    elif idx_coma > idx_punto:
        t = t.replace('.', '').replace(',', '.') # VE: 1.234,56 -> 1234.56
    elif idx_coma != -1 and idx_punto == -1:
        t = t.replace(',', '.') # Solo comas
        
    try:
        return float(t) * signo
    except ValueError:
        return 0.0

def es_texto_numerico(texto):
    """
    Detecta si un string es un número válido en un reporte contable.
    Acepta: '1.000,00', '-500', '(200.00)', '0.00', '-'.
    """
    if not texto: return False
    t = texto.strip()
    
    # El guion solo es un número válido (cero)
    if t == '-': return True
    
    # Limpiamos para verificar contenido numérico
    # Permitimos dígitos, comas, puntos, paréntesis y signo menos
    t_clean = re.sub(r'[^\d]', '', t)
    
    # Debe tener al menos un dígito para ser considerado número (si no fue guion)
    if len(t_clean) > 0:
        # Verificamos que no sea una fecha (ej: 01/01/2025)
        if '/' in t: return False
        # Verificamos que no sea un código de cuenta (ej: 1.1.1.02)
        # Los montos no suelen tener más de 2 puntos, las cuentas sí.
        if t.count('.') > 2: return False
        return True
        
    return False

def extraer_saldos_cb(archivo, log_messages):
    """
    Extrae saldos de CB con detección mejorada de números negativos/parentesis.
    """
    datos = {} 
    nombre_archivo = getattr(archivo, 'name', '').lower()
    
    if nombre_archivo.endswith('.pdf'):
        import pdfplumber
        
        log_messages.append("📄 Procesando Reporte CB como PDF...")
        try:
            with pdfplumber.open(archivo) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text: continue
                    
                    for line in text.split('\n'):
                        parts = line.split()
                        if len(parts) < 3: continue
                        
                        codigo = parts[0].strip()
                        
                        if len(codigo) >= 4 and codigo[0].isdigit():
                            try:
                                numeros_encontrados = []
                                indices_numeros = []
                                
                                for i, p in enumerate(parts):
                                    # USAMOS LA NUEVA FUNCIÓN DE DETECCIÓN
                                    if es_texto_numerico(p):
                                        numeros_encontrados.append(p)
                                        indices_numeros.append(i)
                                
                                # Lógica flexible de asignación
                                s_ini = 0.0; s_deb = 0.0; s_cre = 0.0; s_fin = 0.0
                                cant = len(numeros_encontrados)
                                
                                if cant >= 1: s_fin = limpiar_monto_pdf(numeros_encontrados[-1])
                                if cant >= 2: s_ini = limpiar_monto_pdf(numeros_encontrados[0])
                                
                                if cant >= 4:
                                    # Si hay 4 o más, tomamos los últimos 4 (lo estándar)
                                    s_ini = limpiar_monto_pdf(numeros_encontrados[-4])
                                    s_deb = limpiar_monto_pdf(numeros_encontrados[-3])
                                    s_cre = limpiar_monto_pdf(numeros_encontrados[-2])
                                    s_fin = limpiar_monto_pdf(numeros_encontrados[-1])
                                    
                                    # Extracción de Nombre
                                    if indices_numeros:
                                        # El nombre está antes del primer número del bloque de saldos
                                        idx_corte = indices_numeros[-4]
                                        nombre_parts = parts[1:idx_corte]
                                        # Limpieza
                                        nombre_limpio = []
                                        for p in nombre_parts:
                                            # Filtramos fechas dd/mm/yyyy
                                            if not re.search(r'\d{2}/\d{2}/\d{4}', p) and not (p.isdigit() and len(p)==4):
                                                nombre_limpio.append(p)
                                        nombre_banco = " ".join(nombre_limpio)
                                    else: nombre_banco = "SIN NOMBRE"
                                else:
                                    # Caso raro: hay código pero no 4 números. Asumimos lo que haya.
                                    nombre_banco = "DETECTADO (Saldo Parcial)"

                                datos[codigo] = {
                                    'inicial': s_ini, 'debitos': s_deb, 
                                    'creditos': s_cre, 'final': s_fin, 
                                    'nombre': nombre_banco
                                }
                            except: continue
        except Exception as e:
            log_messages.append(f"❌ Error leyendo PDF CB: {str(e)}")
    
    # --- MODO EXCEL ---
    else:
        log_messages.append("📗 Procesando Reporte CB como Excel...")
        try:
            df = pd.read_excel(archivo)
            df.columns = [str(c).strip().upper() for c in df.columns]
            col_cta = next((c for c in df.columns if 'CUENTA' in c), None)
            col_fin = next((c for c in df.columns if 'FINAL' in c), None)
            if col_cta and col_fin:
                for _, row in df.iterrows():
                    codigo = str(row[col_cta]).strip()
                    try: s_fin = float(row[col_fin])
                    except: s_fin = 0.0
                    datos[codigo] = {'inicial':0, 'debitos':0, 'creditos':0, 'final':s_fin, 'nombre':"Excel Import"}
        except: pass

    return datos

def extraer_saldos_cg(archivo, log_messages):
    """
    Extrae saldos de CG usando REGEX para mayor precisión en números grandes.
    Soluciona problemas donde el PDF separa los miles con espacios.
    """
    datos_cg = {}
    nombre_archivo = getattr(archivo, 'name', '').lower()
    
    if nombre_archivo.endswith('.pdf'):
        import pdfplumber
        
        log_messages.append("📄 Procesando Balance CG como PDF (Modo Regex)...")
        try:
            with pdfplumber.open(archivo) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text: continue
                    
                    for line in text.split('\n'):
                        # Limpieza inicial de la línea
                        line_clean = line.strip()
                        
                        # 1. IDENTIFICAR LA CUENTA (Al inicio de la línea)
                        # Busca patrón: Empieza con 1, puntos y dígitos, longitud min 10
                        match_cuenta = re.match(r'^(1\.[\d\.]+)', line_clean)
                        if not match_cuenta:
                            continue
                            
                        cuenta = match_cuenta.group(1)
                        if len(cuenta) < 10: continue # Falso positivo corto
                        
                        # 2. IDENTIFICAR MONTOS
                        # Este Regex busca números financieros:
                        # - Opcional: Signo menos o paréntesis
                        # - Dígitos seguidos de (puntos/comas y más dígitos)
                        # - Debe terminar en un decimal de 2 dígitos
                        
                        # Patrón: ( -?  DIGITOS  ([.,] DIGITOS)*  [.,]  DIGITOS_2 )
                        patron_monto = r'(?:\(?\-?[\d]{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{2})\)?)'
                        
                        # Encontramos todos los patrones que parezcan dinero en la línea
                        montos_encontrados_raw = re.findall(patron_monto, line_clean)
                        
                        # Filtramos: A veces el Regex atrapa la cuenta contable si termina en .XX
                        # Eliminamos cualquier match que sea idéntico a la cuenta
                        montos_validos = []
                        for m in montos_encontrados_raw:
                            # Limpiamos para comparar
                            m_val = limpiar_monto_pdf(m)
                            # Si no es la cuenta (comparando longitud o valor numérico) y parece un monto válido
                            if m != cuenta and len(m) < 20: 
                                montos_validos.append(m_val)
                            # Caso especial: Si el monto es "0.00" o "-" (que el regex de arriba no atrapa bien a veces)
                        
                        # Búsqueda auxiliar para ceros explícitos "0.00" o guiones aislados "-"
                        # Si el regex complejo falló en ceros, hacemos un split tradicional SOLO para rellenar huecos?
                        # Mejor estrategia: El regex complejo atrapa 0.00.
                        # Si hay guiones "-", el split los ve.
                        
                        # ESTRATEGIA HÍBRIDA:
                        # Si el regex encontró pocos números (menos de 8), intentamos completar con la lógica posicional
                        # porque a veces hay columnas vacías o guiones que el regex numérico salta.
                        
                        numeros_finales = montos_validos
                        
                        # Si detectamos menos de 4 números, es peligroso asignar.
                        # Volvemos a mirar la línea buscando guiones "-" que representan ceros.
                        if len(numeros_finales) < 8:
                            parts = line.split()
                            numeros_con_guiones = []
                            for p in parts:
                                if p == '-' or p == '0.00' or re.match(patron_monto, p):
                                    val = 0.0 if p == '-' else limpiar_monto_pdf(p)
                                    # Evitamos agregar la cuenta contable como dinero
                                    if p != cuenta:
                                        numeros_con_guiones.append(val)
                            
                            # Si esta estrategia dio más resultados, la usamos
                            if len(numeros_con_guiones) > len(numeros_finales):
                                numeros_finales = numeros_con_guiones

                        # 3. ASIGNACIÓN DE SALDOS
                        vals_ves = {'inicial':0.0, 'debitos':0.0, 'creditos':0.0, 'final':0.0}
                        vals_usd = {'inicial':0.0, 'debitos':0.0, 'creditos':0.0, 'final':0.0}
                        
                        cant = len(numeros_finales)
                        
                        # Asumimos estructura estándar de 8 columnas de montos
                        if cant >= 8:
                            # Últimos 4 -> Dólar
                            vals_usd = {
                                'inicial': numeros_finales[-4],
                                'debitos': numeros_finales[-3],
                                'creditos': numeros_finales[-2],
                                'final': numeros_finales[-1]
                            }
                            # Antepenúltimos 4 -> Local
                            vals_ves = {
                                'inicial': numeros_finales[-8],
                                'debitos': numeros_finales[-7],
                                'creditos': numeros_finales[-6],
                                'final': numeros_finales[-5]
                            }
                        elif cant >= 4:
                            # Solo Local detectado
                            vals_ves = {
                                'inicial': numeros_finales[-4],
                                'debitos': numeros_finales[-3],
                                'creditos': numeros_finales[-2],
                                'final': numeros_finales[-1]
                            }

                        # 4. OBTENER NOMBRE
                        if cuenta in NOMBRES_CUENTAS_OFICIALES:
                            descripcion = NOMBRES_CUENTAS_OFICIALES[cuenta]
                        else:
                            # Fallback: Cortar string entre cuenta y primer número/palabra clave
                            # Buscamos índice donde termina la cuenta
                            idx_cuenta = line.find(cuenta) + len(cuenta)
                            
                            # Buscamos índice donde empieza "Deudor" o el primer número
                            match_deudor = re.search(r'\s(DEUDOR|ACREEDOR)\s', line, re.IGNORECASE)
                            
                            if match_deudor:
                                idx_fin = match_deudor.start()
                                descripcion = line[idx_cuenta:idx_fin].strip()
                            else:
                                descripcion = "NOMBRE NO DETECTADO"

                        datos_cg[cuenta] = {'VES': vals_ves, 'USD': vals_usd, 'descripcion': descripcion}

        except Exception as e:
            log_messages.append(f"❌ Error leyendo PDF CG: {str(e)}")

    # --- MODO EXCEL ---
    else:
        pass
            
    return datos_cg

def validar_coincidencia_empresa(file_obj, nombre_empresa_sel):
    """
    Verifica si el archivo subido pertenece a la empresa seleccionada.
    Lee la primera página (PDF) o las primeras filas (Excel) buscando el nombre.
    """
    # 1. Definir palabra clave según la selección del menú
    empresa_sel_upper = str(nombre_empresa_sel).upper()
    keyword = ""
    
    if "BEVAL" in empresa_sel_upper: keyword = "BEVAL"
    elif "FEBECA" in empresa_sel_upper: keyword = "FEBECA"
    elif "PRISMA" in empresa_sel_upper: keyword = "PRISMA"
    elif "SILLACA" in empresa_sel_upper: keyword = "SILLACA"
    
    # Si no hay keyword definida (caso raro), pasamos la validación
    if not keyword: return True, ""

    # 2. Leer el encabezado del archivo
    file_obj.seek(0) # Importante: Rebobinar el archivo al inicio para leerlo
    texto_cabecera = ""
    
    try:
        nombre_archivo = getattr(file_obj, 'name', '').lower()
        
        if nombre_archivo.endswith('.pdf'):
            import pdfplumber
            
            with pdfplumber.open(file_obj) as pdf:
                if pdf.pages:
                    # Leemos solo la primera página
                    texto_cabecera = pdf.pages[0].extract_text() or ""
                    texto_cabecera = texto_cabecera.upper()
        else:
            # Excel: Leemos las primeras 10 filas
            df = pd.read_excel(file_obj, header=None, nrows=10)
            # Convertimos todo el dataframe pequeño a texto mayúscula
            texto_cabecera = df.to_string().upper()
            
    except Exception:
        # Si falla la lectura técnica (archivo dañado), dejamos que el proceso principal maneje el error
        file_obj.seek(0)
        return True, ""

    file_obj.seek(0) # Importante: Rebobinar de nuevo para que el proceso principal lo lea desde el principio

    # 3. Comparación
    if keyword in texto_cabecera:
        return True, ""
    else:
        # Caso especial: A veces el archivo de "Febeca Quincalla" dice solo "FEBECA"
        if keyword == "FEBECA" and "FEBECA" in texto_cabecera:
             return True, ""
             
        return False, f"El archivo **'{file_obj.name}'** no parece corresponder a **{keyword}**. No se encontró el nombre de la empresa en el encabezado."
        
def run_cuadre_cb_cg(file_cb, file_cg, nombre_empresa, log_messages):
    """
    Función Principal: Cruza Tesorería vs Contabilidad.
    Soporta: BEVAL, FEBECA, PRISMA, SILLACA.
    """
    # 1. Configuración (Selección de Diccionario)
    empresa_upper = str(nombre_empresa).upper()
    
    if "PRISMA" in empresa_upper:
        mapeo_actual = MAPEO_CB_CG_PRISMA
        log_messages.append(f"🏢 Configuración activa: PRISMA")
    elif "SILLACA" in empresa_upper or "QUINCALLA" in empresa_upper:
        mapeo_actual = MAPEO_CB_CG_SILLACA
        log_messages.append(f"🏢 Configuración activa: SILLACA / QUINCALLA")
    elif "FEBECA" in empresa_upper:
        # Aplica tanto para Febeca C.A. como Febeca Quincalla
        mapeo_actual = MAPEO_CB_CG_FEBECA
        log_messages.append(f"🏢 Configuración activa: FEBECA")
    else:
        mapeo_actual = MAPEO_CB_CG_BEVAL
        log_messages.append(f"🏢 Configuración activa: BEVAL")

    # 2. Extracción
    raw_cb = extraer_saldos_cb(file_cb, log_messages)
    data_cb = {}
    for k, v in raw_cb.items():
        key_limpia = str(k).strip().replace(".", "").replace(",", "").replace("\t", "")
        data_cb[key_limpia] = v
    data_cg = extraer_saldos_cg(file_cg, log_messages)
    
    cb_encontrados = set(data_cb.keys())
    cg_encontrados = set(data_cg.keys())
    cb_mapeados = set()
    cg_mapeados = set()
    
    # 3. Pre-cálculo agrupación
    suma_cb_por_cuenta = {}
    for codigo_cb, config in mapeo_actual.items():
        cod_key = str(codigo_cb).strip()
        cb_mapeados.add(cod_key) # Registramos que este código ya lo procesamos
        
        info_banco = data_cb.get(cod_key, {})
        saldo_individual = info_banco.get('final', 0.0)
        
        cuenta_cg = config['cta']
        if cuenta_cg not in suma_cb_por_cuenta: suma_cb_por_cuenta[cuenta_cg] = 0.0
        suma_cb_por_cuenta[cuenta_cg] += saldo_individual

    resultados = []

    # 4. Cruce
    for codigo_cb, config in mapeo_actual.items():
        cb_mapeados.add(codigo_cb)
        cg_mapeados.add(config['cta'])
        
        cuenta_cg = config['cta']
        moneda = config['moneda']
        
        info_cb = data_cb.get(codigo_cb, {'inicial':0, 'debitos':0, 'creditos':0, 'final':0, 'nombre':'NO ENCONTRADO'})
        saldo_cb_individual = info_cb.get('final', 0)
        
        clave_cg = 'VES' if moneda == 'VES' else 'USD'
        info_cg_full = data_cg.get(cuenta_cg, {})
        info_cg = info_cg_full.get(clave_cg, {'inicial':0, 'debitos':0, 'creditos':0, 'final':0})
        saldo_cg_total_real = info_cg.get('final', 0)
        desc_cg = info_cg_full.get('descripcion', NOMBRES_CUENTAS_OFICIALES.get(cuenta_cg, 'NO DEFINIDO'))
        
        saldo_cb_grupo_total = suma_cb_por_cuenta.get(cuenta_cg, 0.0)
        diferencia_grupo = round(saldo_cb_grupo_total - saldo_cg_total_real, 2)
        
        if diferencia_grupo == 0:
            estado = "OK"
            diferencia_visual = 0.0
            saldo_cg_visual = saldo_cb_individual 
        else:
            estado = "DESCUADRE"
            diferencia_visual = diferencia_grupo 
            saldo_cg_visual = saldo_cg_total_real
            
        #if saldo_cb_individual == 0 and saldo_cg_total_real == 0 and info_cb.get('debitos', 0) == 0 and diferencia_grupo == 0:
            #continue

        resultados.append({
            'Moneda': moneda,
            'Banco (Tesorería)': codigo_cb, 
            'Cuenta Contable': cuenta_cg,
            'Descripción': desc_cg,
            'Saldo Final CB': saldo_cb_individual,
            'Saldo Final CG': saldo_cg_visual,
            'Diferencia': diferencia_visual,
            'Estado': estado,
            'CB Inicial': info_cb.get('inicial', 0), 'CB Débitos': info_cb.get('debitos', 0), 'CB Créditos': info_cb.get('creditos', 0),
            'CG Inicial': info_cg.get('inicial', 0), 'CG Débitos': info_cg.get('debitos', 0), 'CG Créditos': info_cg.get('creditos', 0)
        })

    # 5. HUÉRFANOS
    huerfanos = []
    
    sobrantes_cb = cb_encontrados - cb_mapeados
    for cod in sobrantes_cb:
        info = data_cb[cod]
        if info['final'] != 0 or info['debitos'] != 0:
            huerfanos.append({
                'Origen': 'TESORERÍA (CB)',
                'Código/Cuenta': cod,
                'Descripción/Nombre': info.get('nombre', 'Desconocido'),
                'Saldo Final': info['final'],
                'Mensaje': 'Código en reporte CB pero no en diccionario.'
            })
            
    sobrantes_cg = cg_encontrados - cg_mapeados
    for cta in sobrantes_cg:
        es_banco = (cta.startswith('1.1.1.02') or cta.startswith('1.1.1.03') or cta.startswith('1.1.1.06'))
        es_agrupadora = cta.endswith('.000')
        
        if es_banco and not es_agrupadora:
            info = data_cg[cta]
            s_ves = info['VES']['final']
            s_usd = info['USD']['final']
            if s_ves != 0 or s_usd != 0:
                huerfanos.append({
                    'Origen': 'CONTABILIDAD (CG)',
                    'Código/Cuenta': cta,
                    'Descripción/Nombre': info.get('descripcion', 'Desconocido'),
                    'Saldo Final': f"Bs: {s_ves} | $: {s_usd}",
                    'Mensaje': 'Cuenta contable con saldo, no mapeada.'
                })

    return pd.DataFrame(resultados), pd.DataFrame(huerfanos)

# ==============================================================================
# LÓGICA PARA GESTIÓN DE IMPRENTA (VALIDACIÓN Y GENERACIÓN) - BLOQUE MAESTRO
# ==============================================================================
import re # Asegúrate de que esto esté importado al inicio del archivo también

# --- PARTE A: VALIDACIÓN (Lectura de TXT) ---

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

def limpiar_string_factura(txt):
    """Limpia letras y ceros a la izquierda para match perfecto entre sistemas."""
    if not txt or pd.isna(txt): return ""
    # Quitar cualquier letra o caracter no numérico y luego ceros iniciales
    num = re.sub(r'\D', '', str(txt))
    return str(int(num)) if num else ""
    
def indexar_libro_ventas(file_libro, log_messages):
    """Radar Mejorado: Busca la base imponible del IVA para evitar los 0.00."""
    db_ventas = {}
    periodo_final = "000000"
    
    try:
        file_libro.seek(0)
        df_raw = pd.read_excel(file_libro, header=None)
        
        # 1. DETECCIÓN DE PERIODO
        for i, row in df_raw.head(15).iterrows():
            fila_txt = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
            if "DEL" in fila_txt and "AL" in fila_txt:
                match_fecha = re.findall(r'(\d{2}/\d{2}/\d{4})', fila_txt)
                if match_fecha:
                    periodo_final = pd.to_datetime(match_fecha[0], dayfirst=True).strftime('%Y%m')
                    break

        # 2. LOCALIZAR TABLA
        header_idx = None
        for i, row in df_raw.head(20).iterrows():
            fila_vals = [str(x).upper() for x in row.values]
            if any("FACTURA" in s for s in fila_vals) and any("RIF" in s for s in fila_vals):
                header_idx = i
                break
        
        if header_idx is None: return {}, "000000"

        df = pd.read_excel(file_libro, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # --- RADAR DE COLUMNAS MEJORADO ---
        col_fac = next((c for c in df.columns if 'N DE FACTURA' in c or 'Nº FACTURA' in c), None)
        
        # Priorizamos 'IMPUESTO IVA G' que es donde Galac guarda el IVA de la factura
        col_iva = next((c for c in df.columns if 'IMPUESTO IVA G' in c), None)
        # Si no existe, buscamos alternativas
        if not col_iva:
            col_iva = next((c for c in df.columns if 'IVA RETENIDO' in c or 'TOTAL IVA' in c), None)
            
        col_fecha = next((c for c in df.columns if 'FECHA FACTURA' in c), None)
        col_nom = next((c for c in df.columns if 'NOMBRE' in c or 'RAZON' in c), None)
        col_comp_existente = next((c for c in df.columns if 'NUMERO COMPROBANTE RETENCION' in c or 'Nº COMPROBANTE' in c), None)

        def safe_float(valor):
            if pd.isna(valor) or str(valor).strip() == "": return 0.0
            if isinstance(valor, (int, float)): return float(valor)
            t = str(valor).strip()
            if ',' in t and '.' in t:
                if t.rfind(',') > t.rfind('.'): t = t.replace('.', '').replace(',', '.')
                else: t = t.replace(',', '')
            elif ',' in t: t = t.replace(',', '.')
            try: return float(t)
            except: return 0.0

        # 3. LLENADO DE MEMORIA
        for _, row in df.iterrows():
            f_raw = row.get(col_fac)
            if pd.isna(f_raw): continue
            
            f_key = re.sub(r'\D', '', str(f_raw))
            if f_key:
                f_key = str(int(f_key))
                
                c_reg = row.get(col_comp_existente)
                comp_val = str(c_reg).strip() if pd.notna(c_reg) and re.sub(r'\D', '', str(c_reg)) != "" else None

                db_ventas[f_key] = {
                    'fecha': pd.to_datetime(row.get(col_fecha), dayfirst=True, errors='coerce'),
                    'iva': safe_float(row.get(col_iva)), # <-- Aquí cargamos la base detectada
                    'nombre': str(row.get(col_nom, "ND")).strip(),
                    'comp_ya_registrado': comp_val
                }
        
        return db_ventas, periodo_final

    except Exception as e:
        log_messages.append(f"❌ Error Indexando Galac: {str(e)}")
        return {}, "000000"

def generar_txt_retenciones_galac(file_softland, file_libro, log_messages):
    """
    Versión Blindada V5: Si una factura del grupo existe en el libro, 
    se usa la base real y se marca error en las faltantes, incluso si es periodo anterior.
    """
    db_ventas, periodo_libro = indexar_libro_ventas(file_libro, log_messages)
    
    try:
        df_soft = pd.read_excel(file_softland, dtype={'Nit': str, 'NIT': str, 'Referencia': str})
        df_soft.columns = [str(c).strip().upper() for c in df_soft.columns]
        col_ref = next((c for c in df_soft.columns if 'REFERENCIA' in c), None)
        col_fecha = next((c for c in df_soft.columns if 'FECHA' in c), None)
        col_monto = next((c for c in df_soft.columns if 'DÉBITO' in c or 'DEBITO' in c or 'LOCAL' in c), None)
        col_rif_s = next((c for c in df_soft.columns if any(k in c for k in ['RIF', 'NIT', 'I.D', 'CEDULA'])), None)
        col_nom_s = next((c for c in df_soft.columns if any(k in c for k in ['NOMBRE', 'CLIENTE', 'TERCERO', 'DESCRIPCION NIT'])), None)
    except: return [], None

    def safe_numeric_parsing(val):
        if pd.isna(val) or str(val).strip() == "": return 0.0
        if isinstance(val, (int, float)): return float(val)
        t = str(val).strip().replace('Bs', '').replace('$', '')
        if ',' in t and '.' in t:
            if t.rfind(',') > t.rfind('.'): t = t.replace('.', '').replace(',', '.')
            else: t = t.replace(',', '')
        elif ',' in t: t = t.replace(',', '.')
        try: return float(t)
        except: return 0.0

    filas_txt = []
    audit = []

    for idx, row in df_soft.iterrows():
        ref = str(row.get(col_ref, "")).strip()
        if "/" not in ref: continue
        
        m_soft_total = safe_numeric_parsing(row.get(col_monto))
        if m_soft_total <= 0: continue
        
        rif_s = str(row.get(col_rif_s, "ND")).strip()
        comprobante = re.sub(r'\D', '', ref.split('/')[0])
        p_voucher = comprobante[:6]
        es_anterior = (p_voucher < periodo_libro) if periodo_libro != "000000" else False
        
        f_nums = [str(int(re.sub(r'\D', '', f))) for f in ref.split('/')[1:] if re.sub(r'\D', '', f)]
        
        facturas_data = []
        existe_alguna_en_libro = False
        todas_existen_en_libro = True
        total_iva_galac = 0.0
        
        for fn in f_nums:
            info = db_ventas.get(fn)
            if info:
                existe_alguna_en_libro = True
                total_iva_galac += info['iva']
            else:
                todas_existen_en_libro = False
            facturas_data.append({'nro': fn, 'info': info})

        # --- CÁLCULO DE ASIGNACIÓN ---
        for f_item in facturas_data:
            f_n = f_item['nro']
            info_g = f_item['info']
            f_txt = f_n.zfill(10)
            f_c = pd.to_datetime(row[col_fecha]).strftime('%d/%m/%Y')
            
            monto_final = 0.0
            iva_base_g = 0.0
            pct_aplicado = 0.0
            estatus = ""
            f_f = f_c 
            nombre_f = info_g['nombre'] if info_g and info_g['nombre'] != "ND" else str(row.get(col_nom_s, "ND"))

            # 1. ¿YA ESTÁ REGISTRADA?
            if info_g and info_g.get('comp_ya_registrado'):
                estatus = "RETENCION REGISTRADA"
                iva_base_g = info_g['iva']
                f_f = info_g['fecha'].strftime('%d/%m/%Y')

            # 2. ¿EXISTE EN EL LIBRO? (Cálculo real de 75/100)
            elif info_g:
                iva_base_g = info_g['iva']
                f_f = info_g['fecha'].strftime('%d/%m/%Y')
                
                # Intentar cuadre matemático 75/100 si todas están, sino prorrateo
                if todas_existen_en_libro:
                    factor = m_soft_total / total_iva_galac if total_iva_galac > 0 else 0
                    # Si el factor es muy cercano a 0.75 o 1.0, lo redondeamos para que el reporte sea limpio
                    if abs(factor - 0.75) < 0.01: factor = 0.75
                    elif abs(factor - 1.0) < 0.01: factor = 1.0
                    
                    monto_final = round(iva_base_g * factor, 2)
                    pct_aplicado = factor
                    estatus = "GENERADO OK"
                else:
                    # Si algunas existen y otras no, no podemos prorratear con seguridad
                    estatus = "⚠️ NO ENCONTRADA EN LIBRO"

            # 3. NO EXISTE EN EL LIBRO
            else:
                # Si es anterior Y ninguna del grupo existe, hacemos el prorrateo de rescate
                if es_anterior and not existe_alguna_en_libro:
                    monto_final = round(m_soft_total / len(f_nums), 2)
                    estatus = "OK - PERIODO ANTERIOR"
                else:
                    # Si el periodo es actual O si algunas del grupo sí existen en el libro,
                    # la que falta se marca como error y monto 0.
                    monto_final = 0.0
                    estatus = "⚠️ NO ENCONTRADA EN LIBRO"

            # Generar TXT
            if monto_final > 0 and estatus != "RETENCION REGISTRADA":
                filas_txt.append(f"FAC\t{f_txt}\t0\t{comprobante}\t{monto_final:.2f}\t{f_c}\t{f_f}")
            
            audit.append({
                'Estatus': estatus,
                'Rif Origen Softland': rif_s,
                'Nombre proveedor Origen Softland': nombre_f,
                'Comprobante': comprobante,
                'Factura': f_txt,
                'IVA Origen Softland': m_soft_total,
                'IVA GALAC (Base)': iva_base_g,
                '% Retención': pct_aplicado,
                'Monto Retenido GALAC': monto_final,
                'Referencia Original': ref
            })

    return filas_txt, pd.DataFrame(audit)

# ==============================================================================
# LÓGICA CÁLCULO LEY PROTECCIÓN PENSIONES
# ==============================================================================

def procesar_calculo_pensiones(file_mayor, file_nomina, tasa_cambio, nombre_empresa, log_messages, num_asiento):
    """
    Motor de cálculo para el impuesto del 9%.
    """
    log_messages.append(f"--- INICIANDO CÁLCULO DE PENSIONES (9%) - {nombre_empresa} ---")
    
    # 0. HERRAMIENTAS INTERNAS
    def limpiar_monto_inteligente(valor):
        if pd.isna(valor) or str(valor).strip() == '': return 0.0
        if isinstance(valor, (int, float)): return float(valor)
        t = str(valor).strip().replace('Bs', '').replace(' ', '').replace('\xa0', '')
        if ',' in t and '.' in t:
            if t.rfind(',') > t.rfind('.'): t = t.replace('.', '').replace(',', '.')
            else: t = t.replace(',', '')
        elif ',' in t: t = t.replace(',', '.')
        elif '.' in t: 
             if len(t.split('.')[-1]) == 3: t = t.replace('.', '')
        try: return float(t)
        except: return 0.0

    mapa_nombres = { "FEBECA, C.A": "FEBECA", "MAYOR BEVAL, C.A": "BEVAL", "PRISMA, C.A": "PRISMA", "FEBECA, C.A (QUINCALLA)": "QUINCALLA" }
    keyword_empresa = mapa_nombres.get(nombre_empresa, nombre_empresa).upper()
    
    mes_detectado = None
    anio_detectado = None
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
                    anio_detectado = str(year_num)
                    log_messages.append(f"📅 Periodo detectado en Mayor: {mes_detectado} {anio_detectado}")
            except: pass

        cuentas_base = ['7.1.1.01.1.001', '7.1.1.09.1.003']
        df_filtrado = df_mayor[df_mayor[col_cta].astype(str).str.strip().isin(cuentas_base)].copy()
        
        df_filtrado['Monto_Deb'] = df_filtrado[col_deb].apply(limpiar_monto_inteligente)
        df_filtrado['Monto_Cre'] = df_filtrado[col_cre].apply(limpiar_monto_inteligente)
        df_filtrado['Base_Neta'] = df_filtrado['Monto_Deb'] - df_filtrado['Monto_Cre']
        df_filtrado['CC_Agrupado'] = df_filtrado[col_cc].astype(str).str.slice(0, 10)
        
        df_agrupado = df_filtrado.groupby(['CC_Agrupado', col_cta]).agg({'Base_Neta': 'sum'}).reset_index()
        df_agrupado.rename(columns={'CC_Agrupado': 'Centro de Costo (Padre)', col_cta: 'Cuenta Contable'}, inplace=True)
        df_agrupado['Impuesto (9%)'] = df_agrupado['Base_Neta'] * 0.09
        
        base_salarios_cont = df_agrupado[df_agrupado['Cuenta Contable'].astype(str).str.contains('7.1.1.01', na=False)]['Base_Neta'].sum()
        base_tickets_cont = df_agrupado[df_agrupado['Cuenta Contable'].astype(str).str.contains('7.1.1.09', na=False)]['Base_Neta'].sum()
        total_base_contable = base_salarios_cont + base_tickets_cont

    except Exception as e:
        log_messages.append(f"❌ Error procesando Mayor: {str(e)}")
        return None, None, None, None

    # --- 2. PROCESAR NÓMINA (SUMA GLOBAL PRISMA) ---
    val_salarios_nom = 0.0
    val_tickets_nom = 0.0
    val_impuesto_nom = 0.0
    
    try:
        if file_nomina:
            xls_nomina = pd.ExcelFile(file_nomina)
            hojas = xls_nomina.sheet_names
            hoja_objetivo = None
            
            # Buscar Hoja por Mes + Año
            if mes_detectado and anio_detectado:
                anio_corto = anio_detectado[-2:]
                for h in hojas:
                    h_upper = h.upper()
                    if mes_detectado in h_upper and (anio_detectado in h_upper or anio_corto in h_upper):
                        hoja_objetivo = h; break
            
            if not hoja_objetivo and mes_detectado:
                for h in hojas:
                    if mes_detectado in h.upper():
                        hoja_objetivo = h; log_messages.append(f"⚠️ Aviso: Se usó hoja '{h}' por mes (sin validar año)."); break
            
            if not hoja_objetivo: 
                hoja_objetivo = hojas[0]
                log_messages.append(f"⚠️ Usando primera hoja: '{hoja_objetivo}'")
            
            # Leer encabezado
            df_raw = pd.read_excel(xls_nomina, sheet_name=hoja_objetivo, header=None, nrows=20)
            header_idx = 0
            for i, row in df_raw.iterrows():
                s = [str(x).upper().replace('\n', ' ').strip() for x in row.values]
                if any("EMPRESA" in x for x in s) and (any("SALARIO" in x for x in s) or any("TOTAL" in x for x in s)):
                    header_idx = i; break
            
            df_nom = pd.read_excel(xls_nomina, sheet_name=hoja_objetivo, header=header_idx)
            df_nom.columns = [str(c).strip().upper().replace('\n', ' ') for c in df_nom.columns]
            
            col_emp = next((c for c in df_nom.columns if 'EMPRESA' in c), None)
            col_sal = next((c for c in df_nom.columns if 'SALARIO' in c), None)
            col_tkt = next((c for c in df_nom.columns if 'TICKET' in c or 'ALIMENTACION' in c), None)
            col_imp = next((c for c in df_nom.columns if 'APARTADO' in c), None)
            
            if col_emp:
                # Filtrar filas que contengan la palabra clave (ej: PRISMA)
                mask = df_nom[col_emp].astype(str).str.upper().str.contains(keyword_empresa, na=False)
                filas_encontradas = df_nom[mask]
                
                if not filas_encontradas.empty:
                    log_messages.append(f"🔎 Filas encontradas para '{keyword_empresa}': {len(filas_encontradas)}")
                    
                    # Sumar todas las filas encontradas (Prisma 01 + Prisma 99)
                    for idx, row in filas_encontradas.iterrows():
                        v_sal = cleaning_sal = limpiar_monto_inteligente(row[col_sal]) if col_sal else 0
                        v_tkt = cleaning_tkt = limpiar_monto_inteligente(row[col_tkt]) if col_tkt else 0
                        
                        val_salarios_nom += v_sal
                        val_tickets_nom += v_tkt
                        if col_imp: val_impuesto_nom += limpiar_monto_inteligente(row[col_imp])

                        # Log para verificar que sumó ambas
                        log_messages.append(f"   + Sumando: {row[col_emp]} (Salario: {v_sal:,.2f})")
                    
                    log_messages.append(f"📊 Total Nómina Global: {val_salarios_nom:,.2f}")
                else:
                    log_messages.append(f"⚠️ No se encontró '{keyword_empresa}' en Nómina.")
            else:
                log_messages.append("❌ Columna EMPRESA no encontrada.")

    except Exception as e:
        log_messages.append(f"⚠️ Error leyendo Nómina: {str(e)}")

    # --- 3. GENERAR ASIENTO ---
    # 1. Agrupamos y redondeamos el Débito en BS
    asiento_data = df_agrupado.groupby('Centro de Costo (Padre)')['Impuesto (9%)'].sum().reset_index()
    asiento_data['Débito VES'] = asiento_data['Impuesto (9%)'].round(2)
    asiento_data.rename(columns={'Centro de Costo (Padre)': 'Centro Costo'}, inplace=True)

    # 2. Convertimos a USD línea por línea
    asiento_data['Débito USD'] = (asiento_data['Débito VES'] / tasa_cambio).round(4)

    # 3. CUADRE MATEMÁTICO: Forzamos que la suma de USD sea exacta a la Tasa del Total
    total_ves_general = asiento_data['Débito VES'].sum()
    total_usd_objetivo = round(total_ves_general / tasa_cambio, 2)
    diferencia_centavos = round(total_usd_objetivo - asiento_data['Débito USD'].sum(), 2)

    if diferencia_centavos != 0 and not asiento_data.empty:
        # Aplicamos el centavo de ajuste a la fila con el monto más alto para que sea imperceptible
        idx_max = asiento_data['Débito USD'].idxmax()
        asiento_data.loc[idx_max, 'Débito USD'] += diferencia_centavos

    # 4. Completamos el resto de las columnas del asiento
    asiento_data['Cuenta Contable'] = '7.1.1.07.1.001'
    asiento_data['Descripción'] = 'Contribucion ley de Pensiones'
    asiento_data['Crédito VES'] = 0.0
    asiento_data['Crédito USD'] = 0.0
    asiento_data['Tasa'] = tasa_cambio
    asiento_data['Asiento'] = num_asiento
    asiento_data['Nit'] = "ND" 
    asiento_data['Fuente'] = "PENSIONES"
    asiento_data['Referencia'] = f"APORTE PENSIONES {mes_detectado[:3]}.{anio_detectado[-2:]}"

    # 5. Calcular Totales para la línea del Pasivo (Crédito)
    total_impuesto_ves = asiento_data['Débito VES'].sum().round(2)
    total_impuesto_usd = asiento_data['Débito USD'].sum().round(4) # Ahora coincide con el total de débitos
    
    linea_pasivo = pd.DataFrame([{
        'Asiento': num_asiento,
        'Nit': "ND",
        'Centro Costo': '00.00.000.00', 
        'Cuenta Contable': '2.1.3.02.3.005', 
        'Descripción': 'Contribuciones Sociales por Pagar', 
        'Fuente': "PENSIONES",
        'Referencia': f"APORTE PENSIONES {mes_detectado[:3]}.{anio_detectado[-2:]}",
        'Débito VES': 0.0, 
        'Crédito VES': total_impuesto_ves,
        'Débito USD': 0.0,
        'Crédito USD': total_impuesto_usd,
        'Tasa': tasa_cambio
    }])
    
    df_asiento = pd.concat([asiento_data, linea_pasivo], ignore_index=True)

    # --- 4. RESUMEN Y VALIDACIÓN ---
    dif_salarios = round(base_salarios_cont - val_salarios_nom, 2)
    dif_tickets = round(base_tickets_cont - val_tickets_nom, 2)
    dif_impuesto = round(total_impuesto_ves - val_impuesto_nom, 2)
    
    total_base_nomina = val_salarios_nom + val_tickets_nom
    
    estado_val = "OK" if (abs(dif_salarios) < 1.00 and abs(dif_tickets) < 1.00) else "DESCUADRE"

    resumen_validacion = {
        'salario_cont': base_salarios_cont, 'salario_nom': val_salarios_nom, 'dif_salario': dif_salarios,
        'ticket_cont': base_tickets_cont, 'ticket_nom': val_tickets_nom, 'dif_ticket': dif_tickets,
        'total_base_cont': total_base_contable, 'total_base_nom': total_base_nomina,
        'dif_base_total': round(total_base_contable - total_base_nomina, 2),
        'imp_calc': total_impuesto_ves, 'imp_nom': val_impuesto_nom, 'dif_imp': dif_impuesto,
        'estado': estado_val
    }

    return df_agrupado, df_filtrado, df_asiento, resumen_validacion

# ==============================================================================
# LÓGICA AJUSTES AL BALANCE EN USD
# ==============================================================================

# MAPEO DE RECLASIFICACIÓN (Saldos Contrarios)
# Estructura: 'Cuenta_Origen': 'Contrapartida'
MAPEO_SALDOS_CONTRARIOS = {
    # --- CLIENTES Y HABERES ---
    '1.1.3.01.1.001': '2.1.2.05.1.108', # Deudores Comerciales <-> Haberes Clientes
    '2.1.2.05.1.108': '1.1.3.01.1.001', # Haberes Clientes <-> Deudores Comerciales
    '1.1.1.04.1.003': '1.1.3.01.1.001', # Fondos por Depositar (Cobros) <-> Deudores Comerciales

    # --- CUENTAS POR COBRAR VARIOS ---
    '1.1.4.01.1.044': '2.1.2.05.1.005', # CxC Varios <-> Asientos por Clasificar
    '2.1.2.05.1.005': '1.1.4.01.1.044', # Asientos por Clasificar <-> CxC Varios

    # --- SEGUROS Y SERVICIOS ---
    '1.1.4.01.1.015': '2.1.2.05.1.015', # HCM Póliza Básica <-> Primas por Pagar
    '1.1.4.01.1.016': '2.1.2.05.1.015', # HCM Póliza Exceso <-> Primas por Pagar
    '1.1.4.01.1.019': '2.1.2.05.1.015', # Servicios Funerarios <-> Primas por Pagar
    '2.1.2.05.1.015': '1.1.4.01.1.015', # Primas por Pagar <-> HCM Póliza Básica (Default)

    # --- PROVEEDORES Y ADELANTOS ---
    '1.1.4.01.1.006': '2.1.2.05.1.012', # Adelanto a Proveedores <-> Prov. Compra Muebles
    '2.1.2.05.1.012': '1.1.4.01.1.006', # Prov. Compra Muebles <-> Adelanto a Proveedores
    
    '1.1.4.05.1.001': '2.1.2.07.1.001', # Avances Compras Locales <-> Proveedores Locales
    '2.1.2.07.1.001': '1.1.4.05.1.001', # Proveedores Locales <-> Avances Compras Locales
    
    '1.1.4.05.1.002': '2.1.2.07.1.011', # Avances Compras ME <-> Proveedores ME
    '2.1.2.07.1.011': '1.1.4.05.1.002', # Proveedores ME <-> Avances Compras ME

    # --- EMPLEADOS Y UTILIDADES ---
    '1.1.4.02.1.006': '2.1.2.05.3.004', # Deudores Empleados Otros <-> Liquidaciones Pendientes
    '2.1.2.05.3.004': '1.1.4.02.1.006', # Liquidaciones Pendientes <-> Deudores Empleados Otros
    
    '1.1.4.02.1.001': '2.1.2.05.1.019', # Deudores Empleados <-> Otras CxP
    '2.1.2.05.1.019': '1.1.4.02.1.001', # Otras CxP <-> Deudores Empleados

    '2.1.2.09.1.001': '2.1.2.05.3.001', # Apartado Utilidades <-> Utilidades Legales por Pagar
    '2.1.2.05.3.001': '2.1.2.09.1.001', # Utilidades Legales por Pagar <-> Apartado Utilidades

    # --- VIAJES ---
    '1.1.4.03.1.002': '2.1.2.09.1.900', # Anticipo Viajes <-> Gastos Estimados por Pagar
    '2.1.2.09.1.900': '1.1.4.03.1.002', # Gastos Estimados por Pagar <-> Anticipo Viajes

    # --- IMPUESTOS (IVA / ISLR / MUNICIPAL) ---
    '1.1.4.04.1.007': '2.1.3.02.5.004', # Deudores Fiscales <-> Ley del Deporte
    '2.1.3.02.5.004': '1.1.4.04.1.007', # Ley del Deporte <-> Deudores Fiscales

    '1.1.4.04.1.003': '2.1.3.04.1.006', # IVA Retenciones <-> IVA Retenido Terceros
    '2.1.3.04.1.006': '1.1.4.04.1.003', # IVA Retenido Terceros <-> IVA Retenciones

    '1.1.4.04.1.004': '2.1.3.04.1.005', # IVA Créditos Pendientes <-> IVA Compensación
    '2.1.3.04.1.005': '1.1.4.04.1.004', # IVA Compensación <-> IVA Créditos Pendientes

    '1.1.4.04.1.001': '2.1.3.01.1.015', # ISLR Pagado Exceso <-> ISLR Anticipo
    '2.1.3.01.1.015': '1.1.4.04.1.001', # ISLR Anticipo <-> ISLR Pagado Exceso

    # --- INTERCOMPAÑÍAS ---
    '1.1.4.07.1.002': '2.1.2.08.1.002', # Beconsult Cta Cte (Activo) <-> Beconsult Cta Cte (Pasivo)
    '2.1.2.08.1.002': '1.1.4.07.1.002', 
    
    '1.1.4.07.1.004': '2.1.2.08.1.004', # Febeca Cta Cte (Activo) <-> Febeca Cta Cte (Pasivo)
    '2.1.2.08.1.004': '1.1.4.07.1.004',
    
    '1.1.4.07.1.071': '2.1.2.08.1.071', # Dist. Sillas California (Activo) <-> (Pasivo)
    '2.1.2.08.1.071': '1.1.4.07.1.071',

    '1.1.4.01.1.508': '2.1.2.05.1.086', # Prisma Cta Cte (Activo) <-> Prisma Cta Cte (Pasivo)
    '2.1.2.05.1.086': '1.1.4.01.1.508',

    # --- OTROS ---
    '1.1.5.07.1.002': '2.1.2.07.1.002', # Envíos en Tránsito Exterior <-> Pasivo Relacionado
    '2.1.2.07.1.002': '1.1.5.07.1.002'
}

# 2. FUNCIONES DE LECTURA AUXILIAR
def leer_saldo_haberes_negativos(file_haberes):
    """Busca la fila 'Total de Saldos Negativos:' y extrae el monto."""
    try:
        df = pd.read_excel(file_haberes)
        for col in df.columns:
            fila_match = df[df[col].astype(str).str.contains("Total de Saldos Negativos", na=False, case=False)]
            if not fila_match.empty:
                val_crudo = fila_match.iloc[0, -1] 
                if isinstance(val_crudo, (int, float)): return abs(float(val_crudo))
                val_limpio = str(val_crudo).replace('.', '').replace(',', '.')
                return abs(float(val_limpio))
    except: pass
    return 0.0

def leer_saldo_viajes(file_obj, columna_busqueda):
    """Busca el total en la columna especificada (SALDO $ o SALDO BS)."""
    try:
        df = pd.read_excel(file_obj)
        df.columns = [str(c).strip().upper() for c in df.columns]
        col_target = next((c for c in df.columns if columna_busqueda in c), None)
        if col_target:
            # Intento 1: Buscar fila TOTAL
            for c_desc in df.columns:
                fila_total = df[df[c_desc].astype(str).str.upper() == 'TOTAL']
                if not fila_total.empty:
                    val = fila_total.iloc[0][col_target]
                    if isinstance(val, (int, float)): return float(val)
            # Intento 2: Sumar columna
            return pd.to_numeric(df[col_target], errors='coerce').sum()
    except: pass
    return 0.0

def extraer_saldos_cg_ajustes(archivo, log_messages):
    """
    Función EXCLUSIVA para el módulo de Ajustes al Balance.
    Lee el Excel del Balance y extrae solo el SALDO FINAL (VES y USD).
    Soporta doble encabezado y columnas repetidas.
    """
    datos_cg = {}
    nombre_archivo = getattr(archivo, 'name', '').lower()
    
    # --- PROCESAMIENTO EXCEL (PRIORIDAD) ---
    if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xls'):
        log_messages.append("📗 [Ajustes] Leyendo Balance CG como Excel...")
        try:
            # 1. Buscar fila de encabezados
            df_raw = pd.read_excel(archivo, header=None)
            header_idx = None
            for i, row in df_raw.head(15).iterrows():
                row_str = [str(x).upper() for x in row.values]
                if any("CUENTA" in s for s in row_str) and any("DESCRIPCI" in s for s in row_str):
                    header_idx = i; break
            
            if header_idx is None:
                log_messages.append("❌ No se encontró encabezado en Balance Excel.")
                return {}

            # 2. Cargar datos
            df = pd.read_excel(archivo, header=header_idx)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # 3. Identificar columnas por posición (Estrategia Balance Fiscal)
            # Buscamos columnas que se llamen "BALANCE FINAL"
            cols_balance = [i for i, c in enumerate(df.columns) if "BALANCE" in c and "FINAL" in c]
            
            # Si la estructura es fija según tu imagen:
            # Col G (index 6) = Balance Final Local
            # Col L (index 11) = Balance Final Dólar
            idx_ves = 6
            idx_usd = 11
            
            # Validación dinámica por si acaso
            if len(cols_balance) >= 2:
                idx_ves = cols_balance[0] # El primero
                idx_usd = cols_balance[-1] # El último
            
            col_cta = next((c for c in df.columns if 'CUENTA' in c), None)
            col_desc = next((c for c in df.columns if 'DESCRIPCI' in c), None)

            if col_cta:
                for _, row in df.iterrows():
                    cuenta = str(row[col_cta]).strip()
                    if not (cuenta.startswith('1.') or cuenta.startswith('2.')): continue
                    
                    desc = str(row[col_desc]).strip() if col_desc else "Sin Descripción"
                    
                    # Helpers
                    def get_val(val):
                        if pd.isna(val): return 0.0
                        if isinstance(val, (int, float)): return float(val)
                        t = str(val).replace('.', '').replace(',', '.')
                        try: return float(t)
                        except: return 0.0

                    # Extraer por índice numérico de columna (iloc) para evitar confusión de nombres
                    saldo_ves = get_val(row.iloc[idx_ves])
                    saldo_usd = get_val(row.iloc[idx_usd])
                    
                    datos_cg[cuenta] = {'VES': saldo_ves, 'USD': saldo_usd, 'descripcion': desc}
        except Exception as e:
            log_messages.append(f"❌ Error leyendo Excel CG para Ajustes: {str(e)}")

    # --- PROCESAMIENTO PDF (Respaldo) ---
    elif nombre_archivo.endswith('.pdf'):
        # Reutilizamos la lógica de extracción de PDF existente pero simplificando la salida
        # para que coincida con la estructura {'VES': float, 'USD': float}
        raw_data = extraer_saldos_cg(archivo, log_messages) # Llamamos a la función vieja
        for cta, info in raw_data.items():
            # La función vieja devuelve diccionarios complejos {'inicial':x, 'final':y}
            # Aplanamos la estructura para este módulo
            s_ves = info['VES']['final'] if isinstance(info['VES'], dict) else info['VES']
            s_usd = info['USD']['final'] if isinstance(info['USD'], dict) else info['USD']
            datos_cg[cta] = {'VES': s_ves, 'USD': s_usd, 'descripcion': info['descripcion']}

    return datos_cg

def procesar_ajustes_balance_usd(f_bancos, f_balance, f_viajes_me, f_viajes_bs, f_haberes, tasa_bcv, tasa_corp, log):
    """
    Motor principal de Ajustes USD.
    Versión blindada con detección de Fila 4 y mapeo de contrapartidas funerarias.
    """
    log.append("--- INICIANDO CÁLCULO DE AJUSTES (USD) ---")
    
    asientos = [] 
    resumen_ajustes = [] 
    df_balance_raw = pd.DataFrame()
    val_activo_ajuste = 0.0
    val_pasivo_ajuste = 0.0

    # --- PASO 0: CAPTURA DE BALANCE ORIGINAL ---
    if f_balance:
        try:
            f_balance.seek(0)
            df_balance_raw = pd.read_excel(f_balance, header=None)
            f_balance.seek(0)
        except: pass

    # --- PASO 1: EXTRAER SALDOS DEL BALANCE (Usando Helper) ---
    datos_cg = extraer_saldos_cg_ajustes(f_balance, log)
    
    # --- PASO 2: PROCESAR BANCOS (Fila 4 + Comas) ---
    lista_bancos_reporte = []
    if f_bancos:
        try:
            df_raw_b = pd.read_excel(f_bancos, header=None)
            header_idx_b = 0
            for i, row in df_raw_b.head(15).iterrows():
                row_str = [str(x).upper() for x in row.values]
                if any("CUENTA CONTABLE" in s for s in row_str):
                    header_idx_b = i
                    break
            
            df_b = pd.read_excel(f_bancos, header=header_idx_b)
            df_b.columns = [str(c).strip().upper().replace('\n', ' ') for c in df_b.columns]
            
            col_no_conc = next((c for c in df_b.columns if "BANCO" in c and "NO" in c and "CONCILIADO" in c), None)
            col_cta = next((c for c in df_b.columns if "CUENTA CONTABLE" in c), None)
            col_tipo = next((c for c in df_b.columns if "TIPO" in c), None)
            col_desc = next((c for c in df_b.columns if "DESCRIPCI" in c), None)
            col_sal_lib = next((c for c in df_b.columns if "SALDO" in c and "LIBROS" in c), None)
            col_sal_bco = next((c for c in df_b.columns if "SALDO" in c and "BANCOS" in c), None)

            def limpiar_val(v):
                if pd.isna(v) or str(v).strip() in ['', '-']: return 0.0
                if isinstance(v, (int, float)): return float(v)
                return float(str(v).replace('.', '').replace(',', '.'))

            if col_no_conc and col_cta:
                for _, row in df_b.iterrows():
                    cta = str(row[col_cta]).strip()
                    if cta == 'nan' or not cta: continue
                    
                    tipo = str(row[col_tipo]).strip().upper() if col_tipo else "L"
                    desc_banco = str(row[col_desc]) if col_desc else "Banco"
                    monto_nc = limpiar_val(row[col_no_conc])
                    s_lib_orig = limpiar_val(row[col_sal_lib])
                    s_bco_orig = limpiar_val(row[col_sal_bco])

                    # Cálculos
                    if tipo in ['E', 'C']:
                        tasa_ref = tasa_bcv
                        ajuste_usd = monto_nc
                    else:
                        tasa_ref = tasa_corp
                        ajuste_usd = monto_nc / tasa_corp if tasa_corp else 0
                    
                    lista_bancos_reporte.append({
                        'TIPO': tipo, 'Cuenta Contable': cta, 'Descripción': desc_banco,
                        'Saldo Libros': s_lib_orig, 'Saldo Bancos': s_bco_orig, 'MOV NC': monto_nc,
                        'SALDO LIB BS': 0, 'SALDO BCO BS': 0, 'SALDO LIB $': 0, 'SALDO BCO $': 0,
                        'AJUSTE BS': 0, 'AJUSTE $': 0, 'TASA': tasa_ref, 'VERIF': 0
                    })
                    
                    if abs(ajuste_usd) > 0.001:
                        val_activo_ajuste += ajuste_usd
                        resumen_ajustes.append({
                            'Cuenta': cta, 'Descripción': desc_banco, 'Origen': 'Bancos',
                            'Saldo Actual USD': datos_cg.get(cta, {}).get('USD', 0.0), 
                            'Ajuste USD': ajuste_usd, 'Saldo Final USD': datos_cg.get(cta, {}).get('USD', 0.0) + ajuste_usd
                        })
                        # Asiento Bancos
                        if ajuste_usd > 0:
                            asientos.append({'Cuenta': cta, 'Desc': desc_banco, 'DebeUSD': ajuste_usd, 'HaberUSD': 0})
                            asientos.append({'Cuenta': '1.1.3.01.1.001', 'Desc': 'Deudores (Ajuste Bco)', 'DebeUSD': 0, 'HaberUSD': ajuste_usd})
                        else:
                            m = abs(ajuste_usd)
                            asientos.append({'Cuenta': '1.1.3.01.1.001', 'Desc': 'Deudores (Ajuste Bco)', 'DebeUSD': m, 'HaberUSD': 0})
                            asientos.append({'Cuenta': cta, 'Desc': desc_banco, 'DebeUSD': 0, 'HaberUSD': m})
        except Exception as e: log.append(f"❌ Error Bancos: {e}")

    # --- PASO 3: PROCESAR VIAJES ---
    if f_viajes_me:
        try:
            val_real = leer_saldo_viajes(f_viajes_me, "SALDO $")
            cta = '1.1.4.03.6.002'
            s_cg = datos_cg.get(cta, {}).get('USD', 0.0)
            adj = val_real - s_cg
            if abs(adj) > 0.01:
                val_activo_ajuste += adj
                resumen_ajustes.append({'Cuenta': cta, 'Descripción': 'Viajes ME', 'Origen': 'Viajes', 'Saldo Actual USD': s_cg, 'Ajuste USD': adj, 'Saldo Final USD': val_real})
                asientos.append({'Cuenta': cta, 'Desc': 'Viajes ME', 'DebeUSD': adj if adj > 0 else 0, 'HaberUSD': abs(adj) if adj < 0 else 0})
                asientos.append({'Cuenta': '2.1.2.09.6.900', 'Desc': 'Gastos Est. ME', 'DebeUSD': abs(adj) if adj < 0 else 0, 'HaberUSD': adj if adj > 0 else 0})
        except: pass

    # --- PASO 4: PROCESAR HABERES ---
    if f_haberes:
        try:
            m_hab = leer_saldo_haberes_negativos(f_haberes)
            if m_hab > 0:
                val_pasivo_ajuste += m_hab
                asientos.append({'Cuenta': '1.1.3.01.1.001', 'Desc': 'Deudores Comerciales', 'DebeUSD': m_hab, 'HaberUSD': 0})
                asientos.append({'Cuenta': '2.1.2.05.1.108', 'Desc': 'Haberes Clientes', 'DebeUSD': 0, 'HaberUSD': m_hab})
                resumen_ajustes.append({'Cuenta': '2.1.2.05.1.108', 'Descripción': 'Haberes Clientes', 'Origen': 'Haberes', 'Saldo Actual USD': 0, 'Ajuste USD': m_hab, 'Saldo Final USD': 0})
        except: pass

    # --- PASO 5: SALDOS CONTRARIOS (INCLUYE FUNERARIOS) ---
    log.append("🔄 Analizando saldos contrarios...")
    for cta, data in datos_cg.items():
        s_usd = data['USD']
        if s_usd < -0.01:
            adj = abs(s_usd)
            contra = MAPEO_SALDOS_CONTRARIOS.get(cta)
            if contra:
                if cta.startswith('1.'): val_activo_ajuste += adj
                elif cta.startswith('2.'): val_pasivo_ajuste += adj
                
                asientos.append({'Cuenta': cta, 'Desc': data['descripcion'], 'DebeUSD': adj, 'HaberUSD': 0})
                asientos.append({'Cuenta': contra, 'Desc': datos_cg.get(contra, {}).get('descripcion', 'Contrapartida'), 'DebeUSD': 0, 'HaberUSD': adj})
                resumen_ajustes.append({'Cuenta': cta, 'Descripción': data['descripcion'], 'Origen': 'Saldo Contrario', 'Saldo Actual USD': s_usd, 'Ajuste USD': adj, 'Saldo Final USD': 0.00})

    # --- PASO 6: COMPILAR RESULTADOS ---
    df_asiento = pd.DataFrame(asientos)
    if not df_asiento.empty:
        df_asiento['Débito VES'] = (df_asiento['DebeUSD'] * tasa_bcv).round(2)
        df_asiento['Crédito VES'] = (df_asiento['HaberUSD'] * tasa_bcv).round(2)
            
    val_data = {
        'total_ajuste_activo': val_activo_ajuste,
        'total_ajuste_pasivo': val_pasivo_ajuste,
        'tasa_bcv': tasa_bcv, 'tasa_corp': tasa_corp
    }

    return pd.DataFrame(resumen_ajustes), pd.DataFrame(lista_bancos_reporte), df_asiento, df_balance_raw, val_data
    
# ==============================================================================
# LÓGICA ENVIOS EN TRANSITO COFERSA (101050200)
# ==============================================================================
def run_conciliation_envios_cofersa(df, log_messages, progress_bar=None):
    """
    Conciliación COFERSA (V16).
    Busca pares exactos DENTRO de cada TIPO antes de evaluar el saldo total del grupo.
    """
    log_messages.append("\n--- INICIANDO CONCILIACIÓN COFERSA (V16 - PARES INTERNOS POR TIPO) ---")
    
    df['Conciliado'] = False
    df['Estado_Cofersa'] = 'PENDIENTE'
    TOLERANCIA_ESTRICTA = 0.01 
    
# Eliminar movimientos que no tienen impacto monetario (Neto 0 en ambas monedas)
    df = df[(df['Neto Local'].abs() > 0.001) | (df['Neto Dólar'].abs() > 0.001)].copy()
    
    # 1. Normalización de la llave TIPO
    df['Ref_Norm'] = (
        df['Tipo'].astype(str)
        .str.strip()
        .str.upper()
        .replace(['NAN', 'NONE', '', '0', '0.0'], 'SIN_TIPO')
    )
    
    # 2. Sincronización de montos
    df['Neto Local'] = (df['Débito Colones'].fillna(0) - df['Crédito Colones'].fillna(0)).round(2)
    df['Neto Dólar'] = (df['Débito Dolar'].fillna(0) - df['Crédito Dolar'].fillna(0)).round(2)
    
    # Verificación de seguridad para el log
    log_messages.append("✅ Montos calculados automáticamente: [Neto = Débito - Crédito]")
    
    total_conciliados = 0
    indices_usados = set()

    # --- FASE ÚNICA: ANÁLISIS DE GRUPOS POR TIPO ---
    df_procesable = df[df['Ref_Norm'] != 'SIN_TIPO'].copy()
    
    if not df_procesable.empty:
        for tipo_val, grupo in df_procesable.groupby('Ref_Norm'):
            # --- MEJORA: VALIDACIÓN GLOBAL DEL GRUPO (BI-MONEDA) ---
            # Si TODO el grupo suma cero en Colones Y en Dólares, cerramos todo de una vez
            suma_local = round(grupo['Neto Local'].sum(), 2)
            suma_usd = round(grupo['Neto Dólar'].sum(), 2)
            
            if abs(suma_local) <= TOLERANCIA_ESTRICTA and abs(suma_usd) <= TOLERANCIA_ESTRICTA:
                idx_grupo = grupo.index
                df.loc[idx_grupo, 'Conciliado'] = True
                df.loc[idx_grupo, 'Estado_Cofersa'] = f'GRUPO_CERRADO_{tipo_val}'
                indices_usados.update(idx_grupo)
                continue # Pasa al siguiente grupo
            
            # --- SUB-FASE A: BUSCAR PARES EXACTOS DENTRO DEL TIPO ---
            # Esto resuelve el caso de Débito y Crédito iguales que se quedaban abiertos
            debitos = grupo[grupo['Neto Local'] > 0].copy()
            creditos = grupo[grupo['Neto Local'] < 0].copy()
            
            for idx_d, row_d in debitos.iterrows():
                monto_buscar = abs(row_d['Neto Local'])
                # Buscamos en los créditos uno que tenga el mismo monto y no haya sido usado
                match_credito = creditos[
                    (creditos['Neto Local'].abs() == monto_buscar) & 
                    (~creditos.index.isin(indices_usados))
                ]
                
                if not match_credito.empty:
                    idx_c = match_credito.index[0]
                    # Validamos que el par también sume cero en USD o sea despreciable
                    # Si no suma cero en USD, lo dejamos para la sub-fase B
                    if abs(row_d['Neto Dólar'] + df.loc[idx_c, 'Neto Dólar']) <= TOLERANCIA:
                        indices_pareja = [idx_d, idx_c]
                        df.loc[indices_pareja, 'Conciliado'] = True
                        df.loc[indices_pareja, 'Estado_Cofersa'] = f'PAR_BI_MONEDA_{tipo_val}'
                        indices_usados.update(indices_pareja)

            # --- SUB-FASE B: VERIFICAR SI EL RESTO DEL GRUPO SUMA CERO ---
            # Lo que no se concilió como par exacto, vemos si suma cero como bloque
            remanente_grupo = grupo[~grupo.index.isin(indices_usados)]
            
            if len(remanente_grupo) >= 2:
                suma_remanente = round(remanente_grupo['Neto Local'].sum(), 2)
                if abs(suma_remanente) <= TOLERANCIA_ESTRICTA:
                    indices_rem = remanente_grupo.index
                    df.loc[indices_rem, 'Conciliado'] = True
                    df.loc[indices_rem, 'Estado_Cofersa'] = f'GRUPO_NETO_{tipo_val}'
                    indices_usados.update(indices_rem)
                    total_conciliados += len(indices_rem)

    if progress_bar:
        progress_bar.progress(1.0)

    # 1. Calculamos el total de filas que quedaron marcadas como Conciliado = True
    total_movimientos_cerrados = len(df[df['Conciliado'] == True])

    # 2. Preparamos los contadores para la UI (opcional si vas a usar el retorno múltiple)
    count_pares = len(df[df['Estado_Cofersa'].str.contains('PAR_', na=False)])
    count_grupos = len(df[df['Estado_Cofersa'].str.contains('GRUPO_|AJUSTE_', na=False)])
    count_pendientes = len(df[df['Estado_Cofersa'] == 'PENDIENTE'])

    # 3. Corregimos el log (usando el nombre de variable correcto)
    log_messages.append(f"🏁 Proceso finalizado. Total movimientos cerrados: {total_movimientos_cerrados}")
    
    if progress_bar:
        progress_bar.progress(1.0)

    # 4. Retornamos según lo que espera tu app.py
    # Si tu app.py espera solo el DF, usa: return df
    # Si aplicaste mi consejo anterior de retorno múltiple, usa:
    return df, count_pares, count_grupos, count_pendientes

# ==============================================================================
# LÓGICA FONDOS EN TRANSITO (101010300)
# ==============================================================================
def normalizar_fondos_transito_cofersa(df):
    """Extrae números de referencia y fuente, y normaliza texto para cruces."""
    def extraer_id(texto):
        if pd.isna(texto): return ""
        # Extraemos solo los números de más de 4 dígitos (posibles depósitos)
        nums = re.findall(r'\d{4,}', str(texto))
        return nums[0] if nums else ""

    # --- CORRECCIÓN: Definimos Ref_Norm para evitar el KeyError ---
    df['Ref_Norm'] = df['Referencia'].astype(str).str.strip().str.upper()
    # --------------------------------------------------------------

    df['Ref_Num'] = df['Referencia'].apply(extraer_id)
    df['Fuente_Num'] = df['Fuente'].apply(extraer_id)
    return df

def run_conciliation_fondos_transito_cofersa(df, log_messages, progress_bar=None):
    """
    Lógica: 
    1. Pares 1-1 por monto exacto y misma Referencia.
    2. Cruce de Referencia (Débito) contra Fuente (Crédito) usando ID de depósito.
    """
    log_messages.append("\n--- INICIANDO FONDOS EN TRÁNSITO COFERSA (101.01.03.00) ---")
    
    # Aplicamos la normalización (que ahora ya incluye Ref_Norm)
    df = normalizar_fondos_transito_cofersa(df)
    total_conciliados = 0
    indices_usados = set()

    # --- FASE 1: PARES 1-1 POR MONTO EXACTO (MISMA REFERENCIA TEXTUAL) ---
    df_pendientes = df[~df['Conciliado']]
    # Agrupamos por el valor absoluto del monto para encontrar parejas
    grupos_monto = df_pendientes.groupby(df_pendientes['Monto_BS'].abs())
    
    for monto_abs, grupo in grupos_monto:
        if len(grupo) < 2: continue
        
        debitos = grupo[grupo['Monto_BS'] > 0].index.tolist()
        creditos = grupo[grupo['Monto_BS'] < 0].index.tolist()
        
        for idx_d in debitos:
            if idx_d in indices_usados: continue
            for idx_c in creditos:
                if idx_c in indices_usados: continue
                
                # Comparamos la Referencia Normalizada (Texto)
                if df.loc[idx_d, 'Ref_Norm'] == df.loc[idx_c, 'Ref_Norm']:
                    df.loc[[idx_d, idx_c], 'Conciliado'] = True
                    df.loc[[idx_d, idx_c], 'Grupo_Conciliado'] = "PAR_MONTO_REF"
                    indices_usados.update([idx_d, idx_c])
                    total_conciliados += 2
                    break

    # --- FASE 2: CRUCE REFERENCIA VS FUENTE (NÚMERO DE DEPÓSITO) ---
    # Solo procesamos lo que no se concilió en la Fase 1
    df_restante = df[~df['Conciliado']]
    debitos_res = df_restante[df_restante['Monto_BS'] > 0]
    creditos_res = df_restante[df_restante['Monto_BS'] < 0]

    for idx_d, row_d in debitos_res.iterrows():
        id_deposito = row_d['Ref_Num']
        if not id_deposito: continue
        
        # Buscamos en los créditos alguien que tenga el mismo número en Fuente o Referencia
        match = creditos_res[
            (creditos_res['Fuente_Num'] == id_deposito) | 
            (creditos_res['Ref_Num'] == id_deposito)
        ]
        
        # Validamos que el monto sea el mismo
        match_monto = match[np.isclose(match['Monto_BS'] + row_d['Monto_BS'], 0, atol=0.01)]
        
        if not match_monto.empty:
            idx_c = match_monto.index[0]
            df.loc[[idx_d, idx_c], 'Conciliado'] = True
            df.loc[[idx_d, idx_c], 'Grupo_Conciliado'] = f"DEPOSITO_{id_deposito}"
            total_conciliados += 2
            # Eliminamos de la lista local para no duplicar el cruce
            creditos_res = creditos_res.drop(idx_c)

    log_messages.append(f"✔️ Conciliación finalizada. Total: {total_conciliados} movimientos.")
    return df

def run_conciliation_dev_proveedores_cofersa(df, log_messages, moneda_base='CRC'):
    """
    Lógica para Devoluciones a Proveedores COFERSA.
    Cruce por NIT + Número de EMB (extraído de Referencia).
    moneda_base: 'CRC' para Colones, 'USD' para Dólares.
    """
    log_messages.append(f"\n--- INICIANDO CONCILIACIÓN DEV. PROVEEDORES ({moneda_base}) ---")
    
    # 1. Parámetros según moneda
    col_monto = 'Neto Local' if moneda_base == 'CRC' else 'Neto Dólar'
    TOLERANCIA = 0.01

    # 2. Extracción de la Llave de Embarque (Regex)
    def extraer_emb(texto):
        if pd.isna(texto): return 'SIN_EMB'
        # Busca EM o M seguido de dígitos (ej: EM25010 o M3500)
        match = re.search(r'([EM]\d+)', str(texto).upper())
        return match.group(1) if match else 'SIN_EMB'

    df['EMB_Key'] = df['Referencia'].apply(extraer_emb)
    
    # 3. Limpieza de NIT y Estados
    df['NIT'] = df['NIT'].astype(str).str.strip().replace(['NAN', 'NONE', '0', '0.0'], 'SIN_NIT')
    df['Conciliado'] = False
    df['Estado_Cofersa'] = 'PENDIENTE'
    
    total_conciliados = 0

    # 4. Cruce Maestro: Grupo por NIT + EMB_Key
    # Ignoramos los que no tienen EMB o NIT para evitar cruces masivos erróneos
    df_procesable = df[(df['EMB_Key'] != 'SIN_EMB') & (df['NIT'] != 'SIN_NIT')].copy()
    
    if not df_procesable.empty:
        grupos = df_procesable.groupby(['NIT', 'EMB_Key'])
        for (nit, emb), grupo in grupos:
            if len(grupo) < 2: continue
            
            suma_grupo = round(grupo[col_monto].sum(), 2)
            
            if abs(suma_grupo) <= TOLERANCIA:
                indices = grupo.index
                df.loc[indices, 'Conciliado'] = True
                df.loc[indices, 'Estado_Cofersa'] = f'DEV_PROV_{emb}'
                total_conciliados += len(indices)

    log_messages.append(f"✔️ Proceso finalizado. Se conciliaron {total_conciliados} movimientos por NIT/EMBARQUE.")
    return df

# ==============================================================================
# LÓGICA VERIFICACIÓN DE DÉBITO FISCAL (BS)
# ==============================================================================

def normalizar_doc_fiscal(texto):
    """Extrae el número de documento limpiando letras y ceros a la izquierda."""
    if pd.isna(texto) or str(texto).strip() == "": return ""
    nums = re.findall(r'\d+', str(texto))
    if nums:
        # Retorna el último bloque numérico como entero para quitar ceros (ej: 000501 -> 501)
        return str(int(nums[-1]))
    return ""

def preparar_datos_softland_debito(df_diario, df_mayor, tag_casa):
    """Mantiene columnas originales y agrega metadatos. NIT normalizado solo a números."""
    df_soft = pd.concat([df_diario, df_mayor], ignore_index=True)
    
    def normalizar_header(t):
        import unicodedata
        return ''.join(c for c in unicodedata.normalize('NFD', str(t))
                      if unicodedata.category(c) != 'Mn').upper().strip()

    col_deb, col_cre, col_rif, col_ref, col_fue, col_nom = None, None, None, None, None, None
    for c in df_soft.columns:
        c_norm = normalizar_header(c)
        if any(k in c_norm for k in ['DEBITO BOLIVAR', 'DEBITO LOCAL', 'DEBITO VES']): col_deb = c
        elif any(k in c_norm for k in ['CREDITO BOLIVAR', 'CREDITO LOCAL', 'CREDITO VES']): col_cre = c
        elif any(k in c_norm for k in ['NIT', 'RIF']): col_rif = c
        elif 'REFERENCIA' in c_norm: col_ref = c
        elif 'FUENTE' in c_norm: col_fue = c
        elif any(k in c_norm for k in ['NOMBRE', 'RAZON SOCIAL', 'DESCRIPCION NIT', 'CLIENTE']): col_nom = c

    def extraer_doc_softland(row):
        doc_fuente = normalizar_doc_fiscal(row.get(col_fue, ""))
        if doc_fuente != "": return doc_fuente
        return normalizar_doc_fiscal(row.get(col_ref, ""))

    def detectar_tipo_softland(row):
        texto = (str(row.get(col_fue, "")) + " " + str(row.get(col_ref, ""))).upper()
        if "N/C" in texto or "NC" in texto: return "N/C"
        if "N/D" in texto or "ND" in texto: return "N/D"
        return "FACTURA"

    df_soft['CASA'] = tag_casa 
    df_soft['_Doc_Norm'] = df_soft.apply(extraer_doc_softland, axis=1)
    df_soft['_Tipo'] = df_soft.apply(detectar_tipo_softland, axis=1)
    
    # --- CAMBIO CLAVE: NIT SOLO NÚMEROS ---
    df_soft['_NIT_Norm'] = df_soft[col_rif].astype(str).str.replace(r'[^0-9]', '', regex=True) if col_rif else "0"
    
    df_soft['_Nombre_Soft'] = df_soft[col_nom].fillna("SIN NOMBRE") if col_nom else "NOMBRE NO DETECTADO"
    val_deb = pd.to_numeric(df_soft[col_deb], errors='coerce').fillna(0) if col_deb else 0
    val_cre = pd.to_numeric(df_soft[col_cre], errors='coerce').fillna(0) if col_cre else 0
    df_soft['_Monto_Bs_Soft'] = abs(val_deb - val_cre)
    
    return df_soft

def run_conciliation_debito_fiscal(df_soft_total, df_imprenta_logica, tolerancia_bs, log_messages):
    """Cruce N-a-N con NIT numérico, filtro exentos y exclusión de FEBECA/Totales."""
    log_messages.append("\n--- INICIANDO AUDITORÍA DE DÉBITO FISCAL ---")
    
    df_imp = df_imprenta_logica.copy()
    def find_col(keywords, df):
        for c in df.columns:
            if any(k in str(c).upper() for k in keywords): return c
        return None

    col_rif = find_col(['RIF'], df_imp)
    col_fact = find_col(['N DE FACTURA'], df_imp)
    col_nd = find_col(['NOTA DE DEBITO'], df_imp)
    col_nc = find_col(['NOTA DE CREDITO'], df_imp)
    col_iva = find_col(['IMPUESTO IVA G'], df_imp)
    col_nom_imp = find_col(['NOMBRE O RAZON SOCIAL', 'NOMBRE', 'RAZON SOCIAL'], df_imp)

    # Filtro Anti-Totales/Resúmenes
    if col_rif:
        df_imp = df_imp[df_imp[col_rif].notna()]
        mask_totales = df_imp.astype(str).apply(lambda x: x.str.contains('TOTALES', case=False, na=False)).any(axis=1)
        df_imp = df_imp[~mask_totales]

    def identificar_tipo_y_doc_imp(row):
        if pd.notna(row.get(col_nc)) and str(row.get(col_nc)).strip() != "":
            return normalizar_doc_fiscal(row.get(col_nc)), "N/C"
        if pd.notna(row.get(col_nd)) and str(row.get(col_nd)).strip() != "":
            return normalizar_doc_fiscal(row.get(col_nd)), "N/D"
        val_f = str(row.get(col_fact, "")).strip()
        if val_f == "" or val_f == "nan": return "", "RESUMEN"
        return normalizar_doc_fiscal(val_f), "FACTURA"

    df_imp[['_Doc_Norm', '_Tipo']] = df_imp.apply(identificar_tipo_y_doc_imp, axis=1, result_type='expand')
    df_imp = df_imp[df_imp['_Doc_Norm'] != ""]
    
    # --- CAMBIO CLAVE: NIT SOLO NÚMEROS ---
    df_imp['_NIT_Norm'] = df_imp[col_rif].astype(str).str.replace(r'[^0-9]', '', regex=True) if col_rif else "0"
    
    df_imp['_Monto_Imprenta'] = pd.to_numeric(df_imp[col_iva], errors='coerce').fillna(0).abs()
    df_imp['_Nombre_Imp'] = df_imp[col_nom_imp].fillna("SIN NOMBRE") if col_nom_imp else "NOMBRE NO DETECTADO"

    # Exclusión FEBECA (Terceros)
    df_imp = df_imp[~df_imp['_Nombre_Imp'].str.upper().str.contains("FEBECA", na=False)]

    soft_grouped = df_soft_total.groupby(['_NIT_Norm', '_Doc_Norm', 'CASA', '_Tipo'], as_index=False).agg({
        '_Monto_Bs_Soft': 'sum',
        '_Nombre_Soft': 'first'
    })
    soft_grouped = soft_grouped[~soft_grouped['_Nombre_Soft'].str.upper().str.contains("FEBECA", na=False)]

    # Cruce por NIT numérico + Documento numérico
    merged = pd.merge(
        soft_grouped, 
        df_imp[['_NIT_Norm', '_Doc_Norm', '_Monto_Imprenta', '_Tipo', '_Nombre_Imp']], 
        on=['_NIT_Norm', '_Doc_Norm'], 
        how='outer', 
        indicator=True,
        suffixes=('_soft', '_imp')
    )

    def clasificar(row):
        tipo_final = row['_Tipo_imp'] if pd.notna(row['_Tipo_imp']) else row['_Tipo_soft']
        nombre_final = row['_Nombre_Imp'] if pd.notna(row['_Nombre_Imp']) else row['_Nombre_Soft']
        m_s = float(row['_Monto_Bs_Soft']) if pd.notna(row['_Monto_Bs_Soft']) else 0.0
        m_i = float(row['_Monto_Imprenta']) if pd.notna(row['_Monto_Imprenta']) else 0.0
        
        if m_i <= 0.001 and m_s <= 0.001: return "OK", tipo_final, nombre_final
        if row['_merge'] == 'left_only': return "NO APARECE EN LIBRO DE VENTAS", tipo_final, nombre_final
        if row['_merge'] == 'right_only': return "NO APARECE EN CONTABILIDAD", tipo_final, nombre_final
        
        dif = abs(m_s - m_i)
        if dif > tolerancia_bs: return f"DIFERENCIA DE MONTO (Bs. {dif:,.2f})", tipo_final, nombre_final
        return "OK", tipo_final, nombre_final

    merged[['Estado', '_Tipo_Final', '_Nombre_Final']] = merged.apply(clasificar, axis=1, result_type='expand')
    return merged
