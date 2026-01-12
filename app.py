# ==============================================================================
# APP.PY - INTERFAZ DE USUARIO (FRONTEND)
# ==============================================================================
import streamlit as st
import pandas as pd
import traceback
from functools import partial

# --- 1. IMPORTACIONES DE MÓDULOS ---
from guides import (
    GUIA_GENERAL_ESPECIFICACIONES, 
    LOGICA_POR_CUENTA, 
    GUIA_COMPLETA_RETENCIONES,
    GUIA_PAQUETE_CC,
    GUIA_IMPRENTA,
    GUIA_GENERADOR,
    GUIA_PENSIONES
)

from logic import (
    # Conciliaciones
    run_conciliation_fondos_en_transito,
    run_conciliation_fondos_por_depositar,
    run_conciliation_devoluciones_proveedores,
    run_conciliation_viajes,
    run_conciliation_retenciones,
    run_conciliation_cobros_viajeros,
    run_conciliation_otras_cxp,
    run_conciliation_haberes_clientes,
    run_conciliation_cdc_factoring,
    run_conciliation_asientos_por_clasificar,
    run_conciliation_deudores_empleados_me,
    run_conciliation_deudores_empleados_bs,
    # Paquete CC
    run_analysis_paquete_cc,
    # Cuadre CB-CG
    run_cuadre_cb_cg,
    validar_coincidencia_empresa,
    # Imprenta
    run_cross_check_imprenta,
    generar_txt_retenciones_galac,
    # Pensiones
    procesar_calculo_pensiones
)

from utils import (
    cargar_y_limpiar_datos,
    generar_reporte_excel,
    generar_excel_saldos_abiertos,
    generar_reporte_paquete_cc,
    generar_reporte_cuadre,
    generar_reporte_imprenta,
    generar_reporte_auditoria_txt,
    generar_archivo_txt,
    generar_reporte_pensiones
)

from mappings import CODIGOS_EMPRESA

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Conciliador Automático", page_icon="🤖", layout="wide")

# --- ESTADO DE SESIÓN ---
if 'page' not in st.session_state: st.session_state.page = 'inicio'
if 'password_correct' not in st.session_state: st.session_state.password_correct = False
if 'processing_complete' not in st.session_state: st.session_state.processing_complete = False

# ==============================================================================
# FUNCIONES AUXILIARES DE UI
# ==============================================================================

def mostrar_error_amigable(e, contexto=""):
    """Traduce errores técnicos a mensajes amigables."""
    error_tecnico = str(e)
    mensaje = f"❌ Ocurrió un error en {contexto}."
    recomendacion = ""

    if "KeyError" in type(e).__name__:
        col = error_tecnico.replace("'", "").replace("KeyError", "").strip()
        mensaje = f"❌ Falta una columna obligatoria: '{col}'"
        recomendacion = "💡 Verifique los encabezados del archivo Excel."
    elif "BadZipFile" in error_tecnico:
        mensaje = "❌ El archivo Excel parece estar dañado."
    elif "The truth value of a Series is ambiguous" in error_tecnico:
        mensaje = "❌ Columnas duplicadas en el archivo."
        recomendacion = "💡 Verifique que no tenga dos columnas con el mismo nombre."
    
    st.error(mensaje)
    if recomendacion: st.info(recomendacion)
    with st.expander("Detalles Técnicos (Soporte)"):
        st.code(traceback.format_exc())

def run_conciliation_wrapper(func, df, log_messages, progress_bar=None):
    """Wrapper para funciones parciales."""
    return func(df, log_messages)

# ==============================================================================
# ESTRATEGIAS DE CONCILIACIÓN
# ==============================================================================
ESTRATEGIAS = {
    "111.04.1001 - Fondos en Tránsito": { 
        "id": "fondos_transito", 
        "funcion_principal": partial(run_conciliation_wrapper, run_conciliation_fondos_en_transito), 
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores", 
        "columnas_reporte": ['Asiento', 'Referencia', 'Fecha', 'Monto Dólar', 'Tasa', 'Bs.'], 
        "nombre_hoja_excel": "111.04.1001",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "111.04.6001 - Fondos por Depositar - ME": { 
        "id": "fondos_depositar", 
        "funcion_principal": run_conciliation_fondos_por_depositar, 
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores", 
        "columnas_reporte": ['Asiento', 'Referencia', 'Fecha', 'Monto Dólar', 'Tasa', 'Bs.'], 
        "nombre_hoja_excel": "111.04.6001",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Débito Dolar', 'Crédito Dolar']
    },
    "212.07.6009 - Devoluciones a Proveedores": { 
        "id": "devoluciones_proveedores", 
        "funcion_principal": partial(run_conciliation_wrapper, run_conciliation_devoluciones_proveedores),
        "label_actual": "Reporte de Devoluciones", "label_anterior": "Partidas pendientes", 
        "columnas_reporte": ['Fecha', 'Fuente', 'Referencia', 'Nombre del Proveedor', 'Monto USD', 'Monto Bs'], 
        "nombre_hoja_excel": "212.07.6009",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor']
    },
    "114.03.1002 - Cuentas de viajes - anticipos de gastos": {
        "id": "cuentas_viajes",
        "funcion_principal": run_conciliation_viajes,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['Asiento', 'NIT', 'Nombre del Proveedor', 'Referencia', 'Fecha', 'Monto_BS', 'Monto_USD', 'Tipo'],
        "nombre_hoja_excel": "114.03.1002",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'NIT']
    },
    "114.02.6006 - Deudores Empleados - Otros (ME)": {
        "id": "deudores_empleados_me",
        "funcion_principal": run_conciliation_deudores_empleados_me,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Referencia', 'Monto Dólar', 'Bs.', 'Tasa'],
        "nombre_hoja_excel": "114.02.6006",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Débito Dolar', 'Crédito Dolar']
    },
    "114.02.1006 - Deudores Empleados - Otros": {
        "id": "deudores_empleados_bs",
        "funcion_principal": run_conciliation_deudores_empleados_bs,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Referencia', 'Bs.', 'Monto Dólar', 'Tasa'],
        "nombre_hoja_excel": "114.02.1006",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "111.04.6003 - Fondos por Depositar - Cobros Viajeros - ME": {
        "id": "cobros_viajeros",
        "funcion_principal": run_conciliation_cobros_viajeros,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Asiento', 'Referencia', 'Fuente', 'Monto Dólar', 'Bs.', 'Tasa'],
        "nombre_hoja_excel": "111.04.6003",
        "columnas_requeridas": ['Asiento', 'Fuente', 'Fecha', 'Referencia', 'Nit', 'Débito Dolar', 'Crédito Dolar']
    },
    "212.05.1019 - Otras Cuentas por Pagar": {
        "id": "otras_cuentas_por_pagar",
        "funcion_principal": run_conciliation_otras_cxp,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Referencia', 'Numero_Envio', 'Monto Dólar', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1019",
        "columnas_requeridas": ['Asiento', 'Fuente', 'Fecha', 'Referencia', 'Nit', 'Debito Bolivar', 'Credito Bolivar']
    },
    "212.05.1108 - Haberes de Clientes": {
        "id": "haberes_clientes",
        "funcion_principal": run_conciliation_haberes_clientes,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha Origen Acreencia', 'Numero de Documento', 'Referencia', 'Monto Dólar', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1108",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Débito Bolivar', 'Crédito Bolivar', 'Fuente']
    },
    "212.07.9001 - CDC - Factoring": {
        "id": "cdc_factoring",
        "funcion_principal": run_conciliation_cdc_factoring,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['Contrato', 'Documento', 'Saldo USD', 'Tasa', 'Saldo Bs'], 
        "nombre_hoja_excel": "212.07.9001",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Fuente', 'Débito Dolar', 'Crédito Dolar']
    },
    "212.05.1005 - Asientos por clasificar": {
        "id": "asientos_por_clasificar",
        "funcion_principal": run_conciliation_asientos_por_clasificar,
        "label_actual": "Movimientos del mes", "label_anterior": "Saldos anteriores",
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Asiento', 'Referencia', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1005",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Débito Bolivar', 'Crédito Bolivar']
    }
}

# ==============================================================================
# AUTENTICACIÓN
# ==============================================================================
def password_entered():
    if st.session_state.get("password") == st.secrets.get("password"):
        st.session_state.password_correct = True
        del st.session_state["password"]
    else:
        st.session_state.password_correct = False

if not st.session_state.get("password_correct", False):
    _, col_main, _ = st.columns([1, 1.5, 1])
    with col_main:
        st.title("Portal de Herramientas Contables", anchor=False)
        with st.container(border=True):
            st.text_input("Contraseña", type="password", on_change=password_entered, key="password", label_visibility="collapsed")
            st.button("Ingresar", on_click=password_entered, type="primary", use_container_width=True)
            if st.session_state.get("password_correct") is False: st.error("Contraseña incorrecta.")
    st.stop()

# ==============================================================================
# NAVEGACIÓN Y RENDERIZADO
# ==============================================================================
def set_page(page_name): st.session_state.page = page_name

def render_inicio():
    st.title("🤖 Portal de Herramientas Contables")
    st.markdown("Seleccione una herramienta:")
    c1, c2 = st.columns(2)
    with c1:
        st.button("📄 Especificaciones", on_click=set_page, args=['especificaciones'], use_container_width=True)
        st.button("📦 Análisis Paquete CC", on_click=set_page, args=['paquete_cc'], use_container_width=True)
        st.button("🛡️ Cálculo Pensiones", on_click=set_page, args=['pensiones'], use_container_width=True)
        st.button("💵 Reservas y Apartados", disabled=True, use_container_width=True)
    with c2:
        st.button("⚖️ Cuadre CB - CG", on_click=set_page, args=['cuadre'], use_container_width=True)
        st.button("🧾 Relación Retenciones", on_click=set_page, args=['retenciones'], use_container_width=True)
        st.button("🖨️ Cruce Imprenta", on_click=set_page, args=['imprenta'], use_container_width=True)
        st.button("🔜 Próximamente", disabled=True, use_container_width=True)

def render_especificaciones():
    st.title('📄 Conciliación de Cuentas')
    if st.button("⬅️ Volver", key="b1"): set_page('inicio'); st.rerun()

    CASA_OPTIONS = ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"]
    c_sel = st.selectbox("Empresa:", CASA_OPTIONS, key="casa_spec")
    cta_sel = st.selectbox("Cuenta:", sorted(list(ESTRATEGIAS.keys())), key="cta_spec")
    
    with st.expander("📖 Guía de Uso"): st.markdown(GUIA_GENERAL_ESPECIFICACIONES + "\n\n" + LOGICA_POR_CUENTA.get(cta_sel, ""))

    est = ESTRATEGIAS[cta_sel]
    c1, c2 = st.columns(2)
    with c1: f_act = st.file_uploader(est["label_actual"], type="xlsx")
    with c2: f_ant = st.file_uploader(est["label_anterior"], type="xlsx")

    if f_act and f_ant:
        if st.button("▶️ Conciliar", type="primary"):
            log = []; prog = st.empty()
            try:
                df_full = cargar_y_limpiar_datos(f_act, f_ant, log)
                if df_full is not None:
                    prog.progress(0, "Conciliando...")
                    df_res = est["funcion_principal"](df_full.copy(), log, progress_bar=prog)
                    
                    st.session_state.df_open = df_res[~df_res['Conciliado']].copy()
                    st.session_state.df_closed = df_res[df_res['Conciliado']].copy()
                    
                    # Generar Nombres
                    cod = CODIGOS_EMPRESA.get(c_sel, "000")
                    num = cta_sel.split(" - ")[0].strip()
                    fecha = df_full['Fecha'].max()
                    f_txt = f"{['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'][fecha.month-1]}.{str(fecha.year)[-2:]}" if pd.notna(fecha) else "ND"
                    
                    name_rep = f"{cod}_{num} {f_txt}.xlsx"
                    name_saldo = f"Saldos_{cod}_{num} {f_txt}.xlsx"

                    xls_rep = generar_reporte_excel(df_full, st.session_state.df_open, st.session_state.df_closed, est, c_sel, cta_sel)
                    xls_sal = generar_excel_saldos_abiertos(st.session_state.df_open)
                    
                    st.success("✅ Listo!")
                    d1, d2 = st.columns(2)
                    d1.download_button("⬇️ Reporte", xls_rep, name_rep, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    d2.download_button("⬇️ Saldos Próximo Mes", xls_sal, name_saldo, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    
                    with st.expander("Log"): st.text('\n'.join(log))
            except Exception as e: mostrar_error_amigable(e, "Conciliación")

def render_paquete_cc():
    st.title('📦 Análisis Paquete CC')
    if st.button("⬅️ Volver", key="b2"): set_page('inicio'); st.rerun()
    
    with st.expander("📖 Guía"): st.markdown(GUIA_PAQUETE_CC)
    
    c_sel = st.selectbox("Empresa:", ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"])
    file_d = st.file_uploader("Diario (.xlsx)", type="xlsx")
    
    if file_d:
        if st.button("▶️ Analizar", type="primary"):
            log = []
            try:
                df = pd.read_excel(file_d)
                # Normalización Columnas
                std_cols = {'Débito Dolar':['Debito Dolar','Débito Dólar'], 'Crédito Dolar':['Credito Dolar'], 'Débito VES':['Debito VES','Debito Bolivar'], 'Crédito VES':['Credito VES','Credito Bolivar']}
                ren = {}
                for s, vars in std_cols.items():
                    for v in vars: 
                        if v in df.columns: ren[v] = s
                df.rename(columns=ren, inplace=True)

                res = run_analysis_paquete_cc(df, log)
                xls = generar_reporte_paquete_cc(res, c_sel)
                
                st.success("✅ Análisis Completado")
                st.download_button("⬇️ Reporte", xls, "Analisis_Paquete_CC.xlsx")
                with st.expander("Log"): st.text('\n'.join(log))
            except Exception as e: mostrar_error_amigable(e, "Paquete CC")

def render_cuadre():
    st.title("⚖️ Cuadre CB - CG")
    if st.button("⬅️ Volver", key="b3"): set_page('inicio'); st.rerun()
    
    c_sel = st.selectbox("Empresa:", ["MAYOR BEVAL, C.A", "FEBECA, C.A", "FEBECA, C.A (QUINCALLA)", "PRISMA, C.A", "SILLACA, C.A."])
    c1, c2 = st.columns(2)
    with c1: f_cb = st.file_uploader("1. Tesorería (PDF/XLS)", type=['pdf','xlsx'])
    with c2: f_cg = st.file_uploader("2. Contabilidad (PDF/XLS)", type=['pdf','xlsx'])
    
    if f_cb and f_cg:
        if st.button("Comparar", type="primary"):
            log = []
            try:
                # Validar Seguridad
                v1, m1 = validar_coincidencia_empresa(f_cb, c_sel)
                v2, m2 = validar_coincidencia_empresa(f_cg, c_sel)
                if not (v1 and v2): st.error(f"⛔ {m1 or m2}"); st.stop()
                
                with st.spinner("Procesando..."):
                    df_res, df_h = run_cuadre_cb_cg(f_cb, f_cg, c_sel, log)
                    
                st.dataframe(df_res[['Banco (Tesorería)','Descripción','Saldo Final CB','Saldo Final CG','Diferencia','Estado']], use_container_width=True)
                if not df_h.empty: st.error("⚠️ Cuentas Huérfanas detectadas"); st.dataframe(df_h)
                
                xls = generar_reporte_cuadre(df_res, df_h, c_sel)
                st.download_button("⬇️ Reporte Completo", xls, f"Cuadre_CB_CG_{c_sel}.xlsx")
                with st.expander("Log"): st.write(log)
            except Exception as e: mostrar_error_amigable(e, "Cuadre CB-CG")

def render_imprenta():
    st.title("🖨️ Gestión Imprenta")
    if st.button("⬅️ Volver", key="b4"): set_page('inicio'); st.rerun()
    
    t1, t2 = st.tabs(["Validar TXT", "Generar TXT"])
    
    with t1:
        c1, c2 = st.columns(2)
        f_s = st.file_uploader("Libro Ventas (.txt)", key="v_s")
        f_r = st.file_uploader("Retenciones (.txt)", key="v_r")
        if f_s and f_r and st.button("Validar", key="bv"):
            log = []
            try:
                df, txt = run_cross_check_imprenta(f_s, f_r, log)
                if not df.empty:
                    err = df[df['Estado']!='OK']
                    if not err.empty: st.error("❌ Errores encontrados"); st.dataframe(err)
                    else: st.success("✅ Todo OK")
                    st.download_button("⬇️ Excel", generar_reporte_imprenta(df), "Validacion.xlsx")
            except Exception as e: mostrar_error_amigable(e, "Validación")
            
    with t2:
        c1, c2 = st.columns(2)
        f_soft = st.file_uploader("Softland (.xlsx)", key="g_soft")
        f_gal = st.file_uploader("Libro Galac (.xlsx)", key="g_gal")
        if f_soft and f_gal and st.button("Generar", key="bg"):
            log = []
            try:
                txt_l, df_a = generar_txt_retenciones_galac(f_soft, f_gal, log)
                if df_a is not None:
                    st.success("✅ Generado")
                    st.download_button("⬇️ TXT", generar_archivo_txt(txt_l), "Retenciones.txt")
                    st.download_button("⬇️ Auditoría", generar_reporte_auditoria_txt(df_a), "Auditoria.xlsx")
                with st.expander("Log"): st.write(log)
            except Exception as e: mostrar_error_amigable(e, "Generación")

def render_retenciones():
    st.title("🧾 Auditoría Retenciones (Compras)")
    if st.button("⬅️ Volver", key="b5"): set_page('inicio'); st.rerun()
    
    with st.expander("Guía"): st.markdown(GUIA_COMPLETA_RETENCIONES)
    
    f_cp = st.file_uploader("Relacion CP", type="xlsx")
    f_cg = st.file_uploader("Diario CG", type="xlsx")
    f_iva = st.file_uploader("GALAC IVA", type="xlsx")
    f_islr = st.file_uploader("GALAC ISLR", type="xlsx")
    f_mun = st.file_uploader("GALAC MUNICIPAL", type="xlsx")
    
    if all([f_cp, f_cg, f_iva, f_islr, f_mun]) and st.button("Auditar", type="primary"):
        log = []
        try:
            xls = run_conciliation_retenciones(f_cp, f_cg, f_iva, f_islr, f_mun, log)
            if xls:
                st.success("✅ Auditoría lista")
                st.download_button("⬇️ Reporte", xls, "Auditoria_Retenciones.xlsx")
            with st.expander("Log"): st.text('\n'.join(log))
        except Exception as e: mostrar_error_amigable(e, "Retenciones")

def render_pensiones():
    st.title("🛡️ Cálculo Pensiones (9%)")
    if st.button("⬅️ Volver", key="b6"): set_page('inicio'); st.rerun()
    
    with st.expander("Guía"): st.markdown(GUIA_PENSIONES)
    
    c_sel = st.selectbox("Empresa:", ["FEBECA", "BEVAL", "PRISMA", "QUINCALLA"], key="pen_emp")
    c1, c2, c3 = st.columns([1.5, 1.5, 1])
    f_may = st.file_uploader("Mayor Contable", type="xlsx", key="p_m")
    f_nom = st.file_uploader("Nómina", type="xlsx", key="p_n")
    tasa = c3.number_input("Tasa", min_value=0.01, value=1.0, format="%.4f")
    
    if f_may and tasa > 0:
        if st.button("Calcular", type="primary"):
            log = []
            try:
                df_c, df_b, df_a, val = procesar_calculo_pensiones(f_may, f_nom, tasa, c_sel, log)
                if df_a is not None:
                    st.success(f"✅ Total a Pagar: {df_a['Crédito VES'].sum():,.2f}")
                    if val['estado'] != 'OK': st.warning("⚠️ Descuadre con Nómina detectado.")
                    
                    # Vista Previa
                    cols_v = ['Centro Costo','Cuenta Contable','Descripción','Débito VES','Crédito VES','Débito USD','Crédito USD','Tasa']
                    st.dataframe(df_a[cols_v], use_container_width=True)
                    
                    # Excel
                    f_cierre = pd.Timestamp.today() # Simplificado, logica real en utils
                    xls = generar_reporte_pensiones(df_c, df_b, df_a, val, c_sel, tasa, f_cierre)
                    st.download_button("⬇️ Reporte", xls, f"Pensiones_{c_sel}.xlsx")
                with st.expander("Log"): st.write(log)
            except Exception as e: mostrar_error_amigable(e, "Pensiones")

# --- MAIN ---
def main():
    router = {
        'inicio': render_inicio,
        'especificaciones': render_especificaciones,
        'retenciones': render_retenciones,
        'paquete_cc': render_paquete_cc,
        'cuadre': render_cuadre,
        'imprenta': render_imprenta,
        'pensiones': render_pensiones
    }
    router.get(st.session_state.page, render_inicio)()

if __name__ == "__main__":
    main()
