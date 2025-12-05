# logic.py

import pandas as pd
import numpy as np
import re
from itertools import combinations
from io import BytesIO
import unicodedata
import xlsxwriter
from difflib import SequenceMatcher
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

    # --- FASE 2: CONCILIACIÓN ESTÁNDAR N-a-N ---
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
            
        # Comparación con tolerancia 0.00
        if np.isclose(grupo['Monto_USD'].sum(), 0, atol=TOLERANCIA_ESTRICTA_USD):
            indices_a_conciliar = grupo.index
            df.loc[indices_a_conciliar, 'Conciliado'] = True
            df.loc[indices_a_conciliar, 'Grupo_Conciliado'] = f"VIAJERO_{nit}_{clave}"
            indices_usados.update(indices_a_conciliar)
            total_conciliados += len(indices_a_conciliar)
            
    if progress_bar: progress_bar.progress(1.0, text="Conciliación completada.")

    if total_conciliados > 0:
        log_messages.append(f"✔️ Conciliación finalizada: Se conciliaron un total de {total_conciliados} movimientos.")
    else:
        log_messages.append("ℹ️ No se encontraron movimientos para conciliar.")
        
    log_messages.append("\n--- PROCESO DE CONCILIACIÓN FINALIZADO ---")
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
    '1.1.1.02.1.002', '1.1.1.02.1.005', '1.1.1.02.6.001', '1.1.1.02.1.003',
    '1.1.1.02.1.018', '1.1.1.02.6.013', '1.1.1.06.6.003', '1.1.1.03.6.002',
    '1.1.1.03.6.028', '1.9.1.01.3.008', # Inversión entre oficinas
    '1.9.1.01.3.009', # Inversión entre oficinas
    '7.1.3.01.1.001',  # Deudores Incobrables
    '1.1.4.01.7.044'  # Cuentas por Cobrar - Varios en ME
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
    '1.1.1.03.6.028'
]}

def _get_base_classification(cuentas_del_asiento, referencia_completa, fuente_completa, referencia_limpia_palabras, monto_suma, monto_max_abs, is_reverso_check=False):
    """
    Versión Optimizada: Recibe valores pre-calculados en lugar de DataFrames.
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

    # --- PRIORIDAD 2: Diferencial Cambiario PURO (Grupo 2) ---
    tiene_diferencial = normalize_account('6.1.1.12.1.001') in cuentas_del_asiento
    tiene_banco = not CUENTAS_BANCO.isdisjoint(cuentas_del_asiento)
    
    if tiene_diferencial and not tiene_banco:
        return "Grupo 2: Diferencial Cambiario"

    # --- PRIORIDAD 3: Retenciones (Grupo 9) ---
    if normalize_account('2.1.3.04.1.006') in cuentas_del_asiento: return "Grupo 9: Retenciones - IVA"
    if normalize_account('2.1.3.01.1.012') in cuentas_del_asiento: return "Grupo 9: Retenciones - ISLR"
    if normalize_account('7.1.3.04.1.004') in cuentas_del_asiento: return "Grupo 9: Retenciones - Municipal"

    # --- PRIORIDAD 4: Traspasos vs. Devoluciones (CORREGIDO LÍMITE $5) ---
    if normalize_account('4.1.1.21.4.001') in cuentas_del_asiento:
        if 'TRASPASO' in referencia_completa and abs(monto_suma) <= TOLERANCIA_MAX_USD: 
            return "Grupo 10: Traspasos"
        
        if is_reverso_check: return "Grupo 7: Devoluciones y Rebajas"
        
        keywords_limpieza_dev = {'LIMPIEZA', 'LIMPIEZAS', 'SALDO', 'SALDOS', 'HISTORICO', 'AJUSTE', 'APLICAR', 'CRUCE', 'FAVOR', 'TRASLADO'}
        if not keywords_limpieza_dev.isdisjoint(referencia_limpia_palabras):
            # LÍMITE CORREGIDO A $5
            if monto_max_abs <= 5: 
                return "Grupo 7: Devoluciones y Rebajas - Limpieza (<= $5)"
            else: 
                keywords_autorizadas = ['TRASLADO', 'APLICAR', 'CRUCE', 'RECLASIFICACION', 'CORRECCION']
                if any(k in referencia_completa for k in keywords_autorizadas):
                     return "Grupo 7: Devoluciones y Rebajas - Traslados/Cruce"
                return "Grupo 7: Devoluciones y Rebajas - Limpieza (> $5)"
        else: 
            return "Grupo 7: Devoluciones y Rebajas - Otros Ajustes"

    # --- PRIORIDAD 5: Gastos de Ventas (Grupo 4) ---
    if normalize_account('7.1.3.19.1.012') in cuentas_del_asiento: 
        return "Grupo 4: Gastos de Ventas"

    # --- PRIORIDAD 6: Cobranzas (Grupo 8) ---
    is_cobranza_texto = 'RECIBO DE COBRANZA' in referencia_completa or 'TEF' in fuente_completa or 'DEPR' in fuente_completa
    
    if is_cobranza_texto or tiene_banco:
        if is_reverso_check: return "Grupo 8: Cobranzas"
        
        if normalize_account('6.1.1.12.1.001') in cuentas_del_asiento: return "Grupo 8: Cobranzas - Con Diferencial Cambiario"
        if normalize_account('1.1.1.04.6.003') in cuentas_del_asiento: return "Grupo 8: Cobranzas - Fondos por Depositar"
        
        if tiene_banco:
            if 'TEF' in fuente_completa: return "Grupo 8: Cobranzas - TEF (Bancos)"
            return "Grupo 8: Cobranzas - Recibos (Bancos)"
        return "Grupo 8: Cobranzas - Otros"

    # --- PRIORIDAD 7: Ingresos Varios (CORREGIDO PALABRAS CLAVE) ---
    if normalize_account('6.1.1.19.1.001') in cuentas_del_asiento:
        if is_reverso_check: return "Grupo 6: Ingresos Varios"
        keywords_limpieza = {'LIMPIEZA', 'LIMPIEZAS', 'SALDO', 'SALDOS', 'HISTORICO', 'INGRESOS', 'INGRESO', 'AJUSTE'}
        
        if not keywords_limpieza.isdisjoint(referencia_limpia_palabras):
            if monto_max_abs <= 25: 
                return "Grupo 6: Ingresos Varios - Limpieza (<= $25)"
            else: 
                return "Grupo 6: Ingresos Varios - Limpieza (> $25)"
        else: 
            return "Grupo 6: Ingresos Varios - Otros"

    # --- PRIORIDAD 8: Inversión entre Oficinas (Grupo 14) ---
    # Cuentas 1.9.1.01.3.008 y 1.9.1.01.3.009
    ctas_inversion = {normalize_account('1.9.1.01.3.008'), normalize_account('1.9.1.01.3.009')}
    if not ctas_inversion.isdisjoint(cuentas_del_asiento):
        return "Grupo 14: Inv. entre Oficinas"

    # --- PRIORIDAD 9: Deudores Incobrables (Grupo 15) ---
    # Cuenta 7.1.3.01.1.001
    if normalize_account('7.1.3.01.1.001') in cuentas_del_asiento:
        return "Grupo 15: Deudores Incobrables"

    # --- NUEVO: PRIORIDAD 10: CxC Varios ME (Grupo 16) ---
    if normalize_account('1.1.4.01.7.044') in cuentas_del_asiento:
        return "Grupo 16: Cuentas por Cobrar - Varios en ME"
            
    # --- RESTO DE PRIORIDADES ---
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
    Recibe un asiento completo (ya clasificado) y aplica las reglas de negocio
    para determinar si está Conciliado o tiene una Incidencia.
    """
    grupo = asiento_group['Grupo'].iloc[0]
    
    # --- GRUPO 1: FLETES ---
    if grupo.startswith("Grupo 1:"):
        fletes_lines = asiento_group[asiento_group['Cuenta Contable Norm'] == normalize_account('7.1.3.45.1.997')]
        if not fletes_lines['Referencia'].str.contains('FLETE', case=False, na=False).all():
            return "Incidencia: Referencia sin 'FLETE' encontrada."
            
    # --- GRUPO 2: DIFERENCIAL CAMBIARIO (CORREGIDO) ---
    elif grupo.startswith("Grupo 2:"):
        diff_lines = asiento_group[asiento_group['Cuenta Contable Norm'] == normalize_account('6.1.1.12.1.001')]
        
        # 1. Validación Estricta (Regex)
        keywords_estrictas = [
            'DIFERENCIAL', 'DIFERENCIA', 'DIF CAMBIARIO', 'DIF.', 
            'TASA', 'AJUSTE', 'IVA', 'DC'  # <--- NUEVO: Acepta abreviatura DC
        ]
        
        # Nota: re.escape es útil para puntos, pero para palabras normales no afecta.
        # Aseguramos que busque la palabra completa o parte significativa.
        patron_regex = '|'.join([re.escape(k) if '.' in k else k for k in keywords_estrictas])
        
        # Revisamos línea por línea
        for ref in diff_lines['Referencia']:
            ref_str = str(ref).upper()
            
            # Si pasa la prueba estricta (tiene DIF, DC, TASA...), continuamos
            if re.search(patron_regex, ref_str):
                continue
                
            # 2. Validación de Similitud (Detector de Typos)
            palabras_referencia = ref_str.split()
            objetivos = ['DIFERENCIAL', 'CAMBIARIO', 'DIFERENCIA', 'AJUSTE']
            es_typo_aceptable = False
            
            for palabra in palabras_referencia:
                p_clean = re.sub(r'[^A-Z]', '', palabra)
                for objetivo in objetivos:
                    ratio = SequenceMatcher(None, p_clean, objetivo).ratio()
                    if ratio > 0.80: 
                        es_typo_aceptable = True
                        break
                if es_typo_aceptable: break
            
            if not es_typo_aceptable:
                return f"Incidencia: Referencia '{ref}' no parece indicar Diferencial Cambiario."
            
    # --- GRUPO 6: INGRESOS VARIOS ---
    elif grupo.startswith("Grupo 6:"):
        if (asiento_group['Monto_USD'].abs() > 25).any():
            return "Incidencia: Movimiento mayor al límite permitido ($25)."
            
    # --- GRUPO 7: DEVOLUCIONES Y REBAJAS (Logica Inteligente) ---
    elif grupo.startswith("Grupo 7:"):
        # Límite corregido a $5
        if (asiento_group['Monto_USD'].abs() > 5).any():
            
            # Excepción para traslados autorizados
            referencia_upper = str(asiento_group['Referencia'].iloc[0]).upper()
            keywords_autorizadas = ['TRASLADO', 'APLICAR', 'CRUCE', 'RECLASIFICACION', 'CORRECCION']
            
            if not any(k in referencia_upper for k in keywords_autorizadas):
                return "Incidencia: Movimiento mayor a $5 (y no indica ser Traslado/Cruce)."

    # --- GRUPO 9: RETENCIONES ---
    elif grupo.startswith("Grupo 9:"):
        # Tomamos la referencia como texto y mayúsculas
        referencia_str = str(asiento_group['Referencia'].iloc[0]).upper().strip()
        
        # Validacion 1: ¿Tiene algún número? (Ej: "00000072", "2025...", "123")
        tiene_numeros = any(char.isdigit() for char in referencia_str)
        
        # Validacion 2: ¿Tiene palabras clave?
        tiene_keywords = any(k in referencia_str for k in ['RET', 'IMP', 'ISLR', 'IVA', 'MUNICIPAL'])
        
        # Si cumple CUALQUIERA de las dos, es válido.
        if tiene_numeros or tiene_keywords:
            pass # Está correcto, pasará al return "Conciliado" final.
        else:
            return f"Incidencia: Referencia '{referencia_str}' inválida (Se requiere Nro Comprobante o RET/IMP)."

    # --- GRUPO 10: TRASPASOS ---
    elif grupo.startswith("Grupo 10:"):
        if not np.isclose(asiento_group['Monto_USD'].sum(), 0, atol=TOLERANCIA_MAX_USD):
            return "Incidencia: El traspaso no suma cero."
            
    # --- GRUPO 3: N/C ---
    elif grupo.startswith("Grupo 3:"):
        # Regla 1: Si la referencia habla de Diferencial Cambiario, es un error de cuenta.
        if "Error de Cuenta" in grupo:
            return "Incidencia: Diferencial Cambiario registrado en cuenta de Descuentos/NC."
            
        # Regla 2: Auditoría de Cuentas Cruzadas (Debe tener Descuento + IVA)
        # Obtenemos las cuentas presentes en este asiento específico
        cuentas_presentes = set(asiento_group['Cuenta Contable Norm'])
        tiene_descuento = normalize_account('4.1.1.22.4.001') in cuentas_presentes
        tiene_iva = normalize_account('2.1.3.04.1.001') in cuentas_presentes
        
        # Si es una Bonificación/Estrategia/Descuento, usualmente esperamos que afecte el IVA.
        # Si falta alguna de las dos, avisamos.
        if not (tiene_descuento and tiene_iva):
            faltante = []
            if not tiene_descuento: faltante.append("Cta Descuentos")
            if not tiene_iva: faltante.append("Cta IVA")
            return f"Incidencia: Asiento de N/C incompleto. Falta: {', '.join(faltante)}."

    # --- Validaciones para Grupos Nuevos ---
    
    elif grupo.startswith("Grupo 14:"):
        # Regla: "Estas cuentas están conciliadas desde que se cargan"
        return "Conciliado"

    elif grupo.startswith("Grupo 15:"):
        # Regla: Asumimos conciliado por defecto al ser un gasto/pérdida directa
        return "Conciliado"

    elif grupo.startswith("Grupo 16:"):
        return "Conciliado"

    # --- GRUPO 11: No identificados ---
    elif grupo.startswith("Grupo 11") or grupo == "No Clasificado":
        # Si falta la cuenta contable en el sistema o no encajó en ninguna regla,
        # es IMPOSIBLE que esté conciliado automáticamente.
        return f"Incidencia: Revisión requerida. {grupo}"
    
    # Si pasó todas las validaciones (o es un grupo sin reglas específicas como Cobranzas)
    return "Conciliado"

def run_analysis_paquete_cc(df_diario, log_messages):
    """
    Función principal optimizada con VECTORIZACIÓN y AGREGACIÓN PREVIA.
    Incluye normalización de columnas de Cliente/NIT y CORRECCIÓN DE ORDENAMIENTO.
    """
    log_messages.append("--- INICIANDO ANÁLISIS Y VALIDACIÓN DE PAQUETE CC (ULTRA RÁPIDO) ---")
    
    df = df_diario.copy()
    
    # --- PASO 0: NORMALIZACIÓN DE COLUMNAS DE CLIENTE/NIT ---
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

    # Limpieza vectorizada
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
    
    log_messages.append(f"⚙️ Analizando {len(df_meta)} asientos únicos...")
    
    # --- FASE 2: CLASIFICACIÓN ITERATIVA ---
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
    
    # --- FASE 3: INTELIGENCIA DE REVERSOS ---
    log_messages.append("🧠 Ejecutando cruce inteligente de reversos...")
    
    mask_reverso = df['Grupo'].astype(str).str.contains("Reverso", case=False, na=False)
    ids_reversos = df[mask_reverso]['Asiento'].unique()
    
    df_candidatos = df[~df['Asiento'].isin(ids_reversos)]
    candidatos_agrupados = df_candidatos.groupby(['Asiento'])['Monto_USD'].sum().round(2).reset_index()
    
    mapa_montos = {}
    for _, row in candidatos_agrupados.iterrows():
        m = row['Monto_USD']
        if m not in mapa_montos: mapa_montos[m] = []
        mapa_montos[m].append(row['Asiento'])
    
    mapa_cambio_grupo = {}
    procesados = set()
    
    for id_rev in ids_reversos:
        if id_rev in procesados: continue
        monto_rev = df_meta.loc[id_rev, 'Suma']
        monto_target = round(-monto_rev, 2)
        posibles = [p for p in mapa_montos.get(monto_target, []) if p not in procesados]
        if not posibles: continue
        
        ref_rev = df_meta.loc[id_rev, 'Ref'] + " " + df_meta.loc[id_rev, 'Fuente']
        numeros_clave = re.findall(r'\d+', ref_rev)
        match_final = None
        
        for cand_id in posibles:
            ref_cand = df_meta.loc[cand_id, 'Ref'] + " " + df_meta.loc[cand_id, 'Fuente']
            for num in numeros_clave:
                if len(num) > 3 and num in ref_cand:
                    match_final = cand_id; break
            if match_final: break
            
        if not match_final and len(posibles) == 1: match_final = posibles[0]
            
        if match_final:
            mapa_cambio_grupo[id_rev] = "Grupo 13: Operaciones Reversadas / Anuladas"
            mapa_cambio_grupo[match_final] = "Grupo 13: Operaciones Reversadas / Anuladas"
            procesados.update([id_rev, match_final])

    if mapa_cambio_grupo:
        df['Grupo'] = df['Asiento'].map(mapa_cambio_grupo).fillna(df['Grupo'])

    # --- FASE 4: VALIDACIÓN Y ORDENAMIENTO ---
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

    # 1. Calcular prioridad de orden (0=Rojo, 1=Blanco)
    df['Orden_Prioridad'] = df['Estado'].apply(lambda x: 1 if str(x).startswith('Conciliado') else 0)
    
    # 2. ORDENAR PRIMERO (Aquí estaba el error antes, ahora ordenamos mientras la columna existe)
    df = df.sort_values(by=['Grupo', 'Orden_Prioridad', 'Asiento'], ascending=[True, True, True])
    
    # 3. BORRAR COLUMNAS AUXILIARES DESPUÉS
    cols_drop = ['Ref_Str', 'Fuente_Str', 'Cuenta Contable Norm', 'Monto_USD', 'Orden_Prioridad']
    df_final = df.drop(columns=cols_drop, errors='ignore')
    
    log_messages.append("--- ANÁLISIS FINALIZADO CON ÉXITO ---")
    return df_final

# ==============================================================================
# LÓGICA PARA CUADRE CB - CG (TESORERÍA VS CONTABILIDAD) - VERSIÓN CORREGIDA
# ==============================================================================
import pdfplumber

# 1. DICCIONARIO MAESTRO DE NOMBRES (Unificado BEVAL + FEBECA)
NOMBRES_CUENTAS_OFICIALES = {
    # ... (Cuentas previas de Beval) ...
    '1.1.1.02.1.000': 'Bancos del País',
    '1.1.1.02.1.002': 'Banco Venezolano de Credito, S.A.',
    '1.1.1.02.1.003': 'Banco de Venezuela, S.A. Banco Universal',
    '1.1.1.02.1.004': 'Banco Provincial, S.A. Banco Universal',
    '1.1.1.02.1.005': 'Banesco, C.A. Banco Universal',
    '1.1.1.02.1.006': 'Bancaribe',
    '1.1.1.02.1.007': 'Banesco, C.A. Banco Universal', # Febeca
    '1.1.1.02.1.008': 'Banco Bicentenario, Banco Universal',
    '1.1.1.02.1.009': 'Banco Mercantil C.A. Banco Universal',
    '1.1.1.02.1.010': 'Banco del Caribe C.A. Banco Universal',
    '1.1.1.02.1.011': 'Banco del Caribe C.A. Banca Universal', # Febeca
    '1.1.1.02.1.015': 'Banco Exterior S.A. Banco Universal',
    '1.1.1.02.1.016': 'Banco de Venezuela, S.A. Banco Universal', # Febeca
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
    
    # Monedas Extranjeras
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
    
    # Bancos Exterior / Otros
    '1.1.1.03.6.000': 'Bancos del Exterior',
    '1.1.1.03.6.002': 'Amerant Bank, N.A.',
    '1.1.1.03.6.012': 'Banesco, S.A. Panamá',
    '1.1.1.03.6.015': 'Santander Private Banking',
    '1.1.1.03.6.024': 'Venezolano de Crédito Cayman Branch',
    '1.1.1.03.6.026': 'Banco Mercantil Panamá',
    '1.1.1.03.6.028': 'FACEBANK International',
    '1.1.1.03.6.031': 'Banesco USA', # Febeca
    '1.1.1.06.6.001': 'PayPal',
    '1.1.1.06.6.002': 'Creska',
    '1.1.1.06.6.003': 'Zinli',
    '1.1.4.01.7.020': 'Servicios de Administración de Fondos -Z',
    '1.1.4.01.7.021': 'Servicios de Administración de Fondos - USDT',
    '1.1.1.01.6.001': 'Cuenta Dolares',
    '1.1.1.01.6.002': 'Cuenta Euros'
}

# 2. MAPEO BEVAL
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

# 3. MAPEO FEBECA (NUEVO)
MAPEO_CB_CG_FEBECA = {
    "0102EU": {"cta": "1.1.1.02.6.213", "moneda": "EUR"},
    "0102L":  {"cta": "1.1.1.02.1.016", "moneda": "VES"},
    "0104L":  {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0104LP": {"cta": "1.1.1.02.1.116", "moneda": "VES"}, # Vencredito Pref
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
    "0134L2": {"cta": "1.1.1.02.1.007", "moneda": "VES"}, # Duplicado, apunta a la misma
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
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "2407E":  {"cta": "1.1.1.06.6.003", "moneda": "USD"},
}

def limpiar_monto_pdf(texto):
    """Convierte texto de moneda a float."""
    if not texto: return 0.0
    limpio = texto.replace('.', '').replace(',', '.')
    try: return float(limpio)
    except ValueError: return 0.0

def extraer_saldos_cb(archivo, log_messages):
    datos = {} 
    nombre_archivo = getattr(archivo, 'name', '').lower()
    
    if nombre_archivo.endswith('.pdf'):
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
                                    if any(c.isdigit() for c in p) and (',' in p or '.' in p or p=='0.00'):
                                        numeros_encontrados.append(p)
                                        indices_numeros.append(i)
                                if len(numeros_encontrados) >= 4:
                                    s_ini = limpiar_monto_pdf(numeros_encontrados[-4])
                                    s_deb = limpiar_monto_pdf(numeros_encontrados[-3])
                                    s_cre = limpiar_monto_pdf(numeros_encontrados[-2])
                                    s_fin = limpiar_monto_pdf(numeros_encontrados[-1])
                                    idx_inicio_nums = indices_numeros[-4]
                                    nombre_parts = parts[1:idx_inicio_nums]
                                    nombre_limpio_parts = []
                                    for p in nombre_parts:
                                        if not re.search(r'\d{2}/\d{2}/\d{4}', p) and not (p.isdigit() and len(p)==4):
                                            nombre_limpio_parts.append(p)
                                    nombre_banco = " ".join(nombre_limpio_parts)
                                    datos[codigo] = {'inicial': s_ini, 'debitos': s_deb, 'creditos': s_cre, 'final': s_fin, 'nombre': nombre_banco}
                            except: continue
        except Exception as e:
            log_messages.append(f"❌ Error leyendo PDF CB: {str(e)}")
            
    else:
        log_messages.append("📗 Procesando Reporte CB como Excel...")
        try:
            df = pd.read_excel(archivo)
            df.columns = [str(c).strip().upper() for c in df.columns]
            col_cta = next((c for c in df.columns if 'CUENTA' in c), None)
            col_nom = next((c for c in df.columns if 'NOMBRE' in c), None)
            col_fin = next((c for c in df.columns if 'FINAL' in c), None)
            col_ini = next((c for c in df.columns if 'INICIAL' in c), None)
            col_deb = next((c for c in df.columns if 'DEBITO' in c or 'DÉBITO' in c), None)
            col_cre = next((c for c in df.columns if 'CREDITO' in c or 'CRÉDITO' in c), None)
            if col_cta and col_fin:
                for _, row in df.iterrows():
                    codigo = str(row[col_cta]).strip()
                    nombre = str(row[col_nom]).strip() if col_nom else "SIN NOMBRE"
                    try:
                        s_fin = float(row[col_fin])
                        s_ini = float(row[col_ini]) if col_ini else 0.0
                        s_deb = float(row[col_deb]) if col_deb else 0.0
                        s_cre = float(row[col_cre]) if col_cre else 0.0
                        datos[codigo] = {'inicial': s_ini, 'debitos': s_deb, 'creditos': s_cre, 'final': s_fin, 'nombre': nombre}
                    except: pass
        except Exception as e:
            log_messages.append(f"❌ Error leyendo Excel CB: {str(e)}")
    return datos

def extraer_saldos_cg(archivo, log_messages):
    datos_cg = {}
    nombre_archivo = getattr(archivo, 'name', '').lower()
    
    if nombre_archivo.endswith('.pdf'):
        log_messages.append("📄 Procesando Balance CG como PDF...")
        try:
            with pdfplumber.open(archivo) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text: continue
                    for line in text.split('\n'):
                        parts = line.split()
                        if len(parts) < 3: continue
                        cuenta = parts[0].strip()
                        if not (cuenta.startswith('1.') and len(cuenta) > 10): continue
                        
                        if cuenta in NOMBRES_CUENTAS_OFICIALES:
                            descripcion = NOMBRES_CUENTAS_OFICIALES[cuenta]
                        else:
                            desc_parts = []
                            for p in parts[1:]:
                                p_clean = p.replace('.', '').replace(',', '').replace('-', '')
                                if p.upper() in ['DEUDOR', 'ACREEDOR', 'SALDO'] or (p_clean.isdigit() and len(p_clean)>0):
                                    break
                                desc_parts.append(p)
                            descripcion = " ".join(desc_parts) + " (Leído PDF)"

                        numeros = []
                        for p in parts[1:]:
                            p_clean = p.replace('.', '').replace(',', '').replace('-', '')
                            if p_clean.isdigit() and len(p_clean) > 0:
                                numeros.append(p)
                        
                        vals_ves = {'inicial':0.0, 'debitos':0.0, 'creditos':0.0, 'final':0.0}
                        vals_usd = {'inicial':0.0, 'debitos':0.0, 'creditos':0.0, 'final':0.0}
                        
                        if len(numeros) >= 4:
                            vals_ves = {
                                'inicial': limpiar_monto_pdf(numeros[0]),
                                'debitos': limpiar_monto_pdf(numeros[1]),
                                'creditos': limpiar_monto_pdf(numeros[2]),
                                'final': limpiar_monto_pdf(numeros[3])
                            }
                        if len(numeros) >= 8:
                            vals_usd = {
                                'inicial': limpiar_monto_pdf(numeros[4]),
                                'debitos': limpiar_monto_pdf(numeros[5]),
                                'creditos': limpiar_monto_pdf(numeros[6]),
                                'final': limpiar_monto_pdf(numeros[7])
                            }
                        datos_cg[cuenta] = {'VES': vals_ves, 'USD': vals_usd, 'descripcion': descripcion}
        except Exception as e:
            log_messages.append(f"❌ Error leyendo PDF CG: {str(e)}")
    else:
        # Placeholder Excel
        pass
    return datos_cg

def run_cuadre_cb_cg(file_cb, file_cg, nombre_empresa, log_messages):
    """
    Función Principal MULTI-EMPRESA.
    Selecciona el diccionario correcto según la empresa elegida en la App.
    """
    data_cb = extraer_saldos_cb(file_cb, log_messages)
    data_cg = extraer_saldos_cg(file_cg, log_messages)
    
    # --- SELECCIÓN DE DICCIONARIO ---
    if "FEBECA" in nombre_empresa.upper():
        mapeo_actual = MAPEO_CB_CG_FEBECA
        log_messages.append("🏢 Usando configuración de cuentas: FEBECA")
    else:
        mapeo_actual = MAPEO_CB_CG_BEVAL
        log_messages.append("🏢 Usando configuración de cuentas: BEVAL")
    # --------------------------------
    
    resultados = []
    
    for codigo_cb, config in mapeo_actual.items():
        cuenta_cg = config['cta']
        moneda = config['moneda']
        
        info_cb = data_cb.get(codigo_cb, {'inicial':0, 'debitos':0, 'creditos':0, 'final':0, 'nombre':'NO ENCONTRADO'})
        
        clave_cg = 'VES' if moneda == 'VES' else 'USD'
        info_cg_full = data_cg.get(cuenta_cg, {})
        info_cg = info_cg_full.get(clave_cg, {'inicial':0, 'debitos':0, 'creditos':0, 'final':0})
        desc_cg = info_cg_full.get('descripcion', NOMBRES_CUENTAS_OFICIALES.get(cuenta_cg, 'NO DEFINIDO'))
        
        saldo_cb = info_cb.get('final', 0)
        saldo_cg = info_cg.get('final', 0)
        
        dif_final = round(saldo_cb - saldo_cg, 2)
        estado = "OK" if dif_final == 0 else "DESCUADRE"
        
        resultados.append({
            'Moneda': moneda,
            'Banco (Tesorería)': codigo_cb, 
            'Cuenta Contable': cuenta_cg,
            'Descripción': desc_cg,
            'Saldo Final CB': saldo_cb,
            'Saldo Final CG': saldo_cg,
            'Diferencia': dif_final,
            'Estado': estado,
            'CB Inicial': info_cb.get('inicial', 0),
            'CB Débitos': info_cb.get('debitos', 0),
            'CB Créditos': info_cb.get('creditos', 0),
            'CG Inicial': info_cg.get('inicial', 0),
            'CG Débitos': info_cg.get('debitos', 0),
            'CG Créditos': info_cg.get('creditos', 0)
        })
        
    return pd.DataFrame(resultados)
