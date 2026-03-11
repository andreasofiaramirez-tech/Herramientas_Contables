# ==============================================================================
# I. INFRAESTRUCTURA, SEGURIDAD Y CONFIGURACIÓN
# ============================================================================== 
import streamlit as st
import pandas as pd
import traceback
from functools import partial
from datetime import datetime

# --- Configuración de la página de Streamlit ---
st.set_page_config(page_title="Conciliador Automático", page_icon="🤖", layout="wide")

# --- Inicialización del Estado de la Sesión ---
if 'page' not in st.session_state:
    st.session_state.page = 'inicio'
if 'password_correct' not in st.session_state:
    st.session_state.password_correct = False
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
    st.session_state.log_messages = []
    st.session_state.csv_output = None
    st.session_state.excel_output = None
    st.session_state.df_saldos_abiertos = pd.DataFrame()
    st.session_state.df_conciliados = pd.DataFrame()

# --- Bloque 1: Importación de Guías ---
from guides import (
    GUIA_GENERAL_ESPECIFICACIONES, 
    LOGICA_POR_CUENTA, 
    GUIA_COMPLETA_RETENCIONES,
    GUIA_PAQUETE_CC,
    GUIA_GENERADOR,
    GUIA_PENSIONES,
    GUIA_AJUSTES_USD,
    GUIA_DEBITO_FISCAL
)

# --- Bloque 2: Importación de Lógica Contable ---
from logic import (
    # Conciliaciones Mayoreo
    run_conciliation_fondos_en_transito,
    run_conciliation_fondos_por_depositar,
    run_conciliation_cobros_viajeros,
    run_conciliation_deudores_empleados_me,
    run_conciliation_viajes,
    run_conciliation_otras_cxp,
    run_conciliation_haberes_clientes,
    run_conciliation_asientos_por_clasificar,
    run_conciliation_devoluciones_proveedores,
    run_conciliation_proveedores_costos,
    run_conciliation_cdc_factoring,
    procesar_ajustes_balance_usd,

    # Conciliaciones COFERSA
    run_conciliation_envios_cofersa,
    run_conciliation_fondos_fondos_cofersa,
    run_conciliation_dev_proveedores_cofersa,

    # Procesos Fiscales y Auditoría
    run_conciliation_retenciones,
    procesar_calculo_pensiones,
    run_conciliation_debito_fiscal,
    run_analysis_paquete_cc,
    run_conciliation_comisiones_bancarias,
    procesar_calculo_locti,

    parsear_balance_softland, 
    conciliar_ciclo_apartados, 
    preparar_asiento_softland,

    # Helpers
    run_cuadre_cb_cg,
    validar_coincidencia_empresa,
    preparar_datos_softland_debito,
)

# --- Bloque 3: Importación de Utilidades (Reportes y Carga) ---
from utils import (
    cargar_y_limpiar_datos,
    generar_reporte_excel,
    generar_excel_saldos_abiertos,
    generar_reporte_paquete_cc,
    generar_reporte_cuadre,
    generar_reporte_pensiones,
    generar_cargador_asiento_pensiones,
    generar_reporte_cofersa,
    cargar_datos_cofersa,
    generar_reporte_ajustes_usd,
    generar_reporte_debito_fiscal,
    generar_hoja_pendientes_dev_cofersa,
    cargar_datos_fondos_cofersa,
    generar_reporte_auditoria_comisiones,
    generar_reporte_excel_locti,
    generar_cargador_softland_v2,
    generar_reporte_visual_liberaciones, 
    generar_reporte_maestro_apartados, 
    generar_excel_cargador_softland
)

# --- Bloque 4: Helpers de Interfaz ---
def mostrar_error_amigable(e, contexto=""):
    error_tecnico = str(e)
    mensaje_usuario = ""
    recomendacion = ""

    # 1. ERRORES DE COLUMNAS FALTANTES (KeyError)
    if "KeyError" in type(e).__name__ or "not in index" in error_tecnico:
        columna_faltante = error_tecnico.replace("'", "").replace("KeyError", "").strip()
        mensaje_usuario = f"❌ Falta una columna obligatoria en el archivo: '{columna_faltante}'"
        
        if "RIF" in columna_faltante or "Proveedor" in columna_faltante:
            recomendacion = "💡 **Posible Causa:** El archivo de Retenciones CP debe tener los encabezados en la **Fila 5**. Verifique que no estén en la fila 1."
        elif "Asiento" in columna_faltante:
            recomendacion = "💡 **Solución:** Verifique que la columna se llame 'Asiento' o 'ASIENTO'."
        else:
            recomendacion = "💡 **Solución:** Revise que el nombre de la columna esté escrito correctamente en el Excel."

    # 2. ERRORES DE LECTURA DE EXCEL (BadZipFile, ValueError)
    elif "BadZipFile" in error_tecnico:
        mensaje_usuario = "❌ El archivo cargado parece estar dañado o no es un Excel válido (.xlsx)."
        recomendacion = "💡 **Solución:** Intente abrir y volver a guardar el archivo en Excel antes de subirlo."
    
    elif "Excel file format cannot be determined" in error_tecnico:
        mensaje_usuario = "❌ Formato de archivo no reconocido."
        recomendacion = "💡 **Solución:** Asegúrese de subir archivos con extensión .xlsx (Excel moderno)."

    # 3. ERRORES DE LÓGICA / VACÍOS
    elif "The truth value of a Series is ambiguous" in error_tecnico:
        mensaje_usuario = "❌ Error de duplicidad en columnas."
        recomendacion = "💡 **Solución:** Su archivo Excel tiene dos columnas con el mismo nombre (ej: dos columnas 'RIF'). Elimine una."
    
    elif "No columns to parse" in error_tecnico:
        mensaje_usuario = "❌ El archivo parece estar vacío o no tiene datos legibles."

    # 4. ERROR GENÉRICO (Fallback)
    else:
        mensaje_usuario = f"❌ Ocurrió un error inesperado durante {contexto}."
        recomendacion = f"Detalle técnico: {error_tecnico}"

    # --- MOSTRAR EN PANTALLA ---
    st.error(mensaje_usuario)
    if recomendacion:
        st.info(recomendacion)
        
    # Mostrar el traceback solo si el usuario quiere verlo (para ti como soporte)
    with st.expander("Ver detalles técnicos del error (Solo para Soporte)"):
        st.code(traceback.format_exc())


def set_page(page_name):
    st.session_state.page = page_name
    
# --- Bloque 5: Autenticación ---
def password_entered():
    """Verifica la contraseña ingresada y actualiza el estado."""
    st.session_state.authentication_attempted = True
    if st.session_state.get("password") == st.secrets.get("password"):
        st.session_state.password_correct = True
        del st.session_state["password"]
    else:
        st.session_state.password_correct = False

if not st.session_state.get("password_correct", False):
    _ , col_main, _ = st.columns([1, 1.5, 1])
    with col_main:
        _ , col_logo, _ = st.columns([1, 2, 1])
        with col_logo:
            try:
                st.image("assets/logo_principal.png", use_container_width=True)  
            except:
                st.warning("No se encontró el logo principal en la carpeta 'assets'.")

        st.title("Bienvenido al Portal de Herramientas Contables", anchor=False)
        st.markdown("Una solución centralizada para el equipo de contabilidad.")
        
        with st.container(border=True):
            st.subheader("Acceso Exclusivo", anchor=False)
            
            # Campo de texto (Se activa con Enter)
            st.text_input(
                "Contraseña", 
                type="password", 
                on_change=password_entered, 
                key="password", 
                label_visibility="collapsed", 
                placeholder="Ingresa la contraseña"
            )

            st.button("Ingresar", on_click=password_entered, type="primary", use_container_width=True)
            
            if st.session_state.get("authentication_attempted", False):
                if not st.session_state.get("password_correct", False):
                    st.error("😕 Contraseña incorrecta.")
            else:
                st.markdown("") 
                st.info("Por favor, ingresa la contraseña para continuar.")

        st.divider()
        st.markdown("<p style='text-align: center; margin-bottom: 5px; font-size: 0.9rem;'>Una herramienta para las empresas del grupo:</p>", unsafe_allow_html=True)
        
        _, col1, col2, col3, _ = st.columns([1, 2, 2, 2, 1]) 
        
        logo_cols = [col1, col2, col3]
        logos_info = [
            {"path": "assets/logo_febeca.png", "fallback": "FEBECA, C.A."},
            {"path": "assets/logo_beval.png", "fallback": "MAYOR BEVAL, C.A."},
            {"path": "assets/logo_sillaca.png", "fallback": "SILLACA, C.A."}
        ]
        
        for i, col in enumerate(logo_cols):
            with col:
                try:
                    # Quitamos use_container_width o controlamos el tamaño con CSS
                    st.image(logos_info[i]["path"], width=150) # Ajusta el width a tu gusto (120-150 suele ser ideal)
                except:
                    st.markdown(f"<p style='text-align: center; font-size: small;'>{logos_info[i]['fallback']}</p>", unsafe_allow_html=True)
    st.stop()

# ==============================================================================
# II. ESTRATEGIAS DE CONCILIACIÓN
# ==============================================================================    
def run_conciliation_wrapper(func, df, log_messages, progress_bar=None):
    return func(df, log_messages)

ESTRATEGIAS = {
    "111.04.1001 - Fondos en Tránsito": { 
        "id": "fondos_transito", 
        "funcion_principal": partial(run_conciliation_wrapper, run_conciliation_fondos_en_transito), 
        "label_actual": "Movimientos del mes (Fondos en Tránsito)", 
        "label_anterior": "Saldos anteriores (Fondos en Tránsito)", 
        "columnas_reporte": ['Asiento', 'Referencia', 'Fecha', 'Monto Dólar', 'Tasa', 'Bs.'], 
        "nombre_hoja_excel": "111.04.1001",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
    },
    "111.04.6001 - Fondos por Depositar - ME": { 
        "id": "fondos_depositar", 
        "funcion_principal": run_conciliation_fondos_por_depositar, 
        "label_actual": "Movimientos del mes (Fondos por Depositar)", 
        "label_anterior": "Saldos anteriores (Fondos por Depositar)", 
        "columnas_reporte": ['Asiento', 'Referencia', 'Fecha', 'Monto Dólar', 'Tasa', 'Bs.'], 
        "nombre_hoja_excel": "111.04.6001",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
    },
    "212.07.6009 - Devoluciones a Proveedores": { 
        "id": "devoluciones_proveedores", 
        "funcion_principal": partial(run_conciliation_wrapper, run_conciliation_devoluciones_proveedores),
        "label_actual": "Reporte de Devoluciones (Proveedores)", 
        "label_anterior": "Partidas pendientes (Proveedores)", 
        "columnas_reporte": ['Fecha', 'Fuente', 'Referencia', 'Nombre del Proveedor', 'Monto USD', 'Monto Bs'], 
        "nombre_hoja_excel": "212.07.6009",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'NIT', 'Nombre del Proveedor', 'Fuente', 'Débito Dolar', 'Crédito Dolar']
    },
    "114.03.1002 - Cuentas de viajes - anticipos de gastos": {
        "id": "cuentas_viajes",
        "funcion_principal": run_conciliation_viajes,
        "label_actual": "Movimientos del mes (Viajes)",
        "label_anterior": "Saldos anteriores (Viajes)",
        "columnas_reporte": ['Asiento', 'NIT', 'Nombre del Proveedor', 'Referencia', 'Fecha', 'Monto_BS', 'Monto_USD', 'Tipo'],
        "nombre_hoja_excel": "114.03.1002",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor', 'NIT', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "114.02.6006 - Deudores Empleados - Otros (ME)": {
        "id": "deudores_empleados_me",
        "funcion_principal": run_conciliation_deudores_empleados_me,
        "label_actual": "Movimientos del mes (Deudores Empleados ME)",
        "label_anterior": "Saldos anteriores (Deudores Empleados ME)",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Referencia', 'Monto Dólar', 'Bs.', 'Tasa'],
        "nombre_hoja_excel": "114.02.6006",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Descripción Nit', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
    },
    "111.04.6003 - Fondos por Depositar - Cobros Viajeros - ME": {
        "id": "cobros_viajeros",
        "funcion_principal": run_conciliation_cobros_viajeros,
        "label_actual": "Movimientos del mes (Cobros Viajeros)",
        "label_anterior": "Saldos anteriores (Cobros Viajeros)",
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Asiento', 'Referencia', 'Fuente', 'Monto Dólar', 'Bs.', 'Tasa'],
        "nombre_hoja_excel": "111.04.6003",
        "columnas_requeridas": ['Asiento', 'Fuente', 'Fecha', 'Referencia', 'Nit', 'Descripcion NIT', 'Débito Dolar', 'Crédito Dolar']
    },
    "212.05.1019 - Otras Cuentas por Pagar": {
        "id": "otras_cuentas_por_pagar",
        "funcion_principal": run_conciliation_otras_cxp,
        "label_actual": "Movimientos del mes (Otras CxP)",
        "label_anterior": "Saldos anteriores (Otras CxP)",
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Referencia', 'Numero_Envio', 'Monto Dólar', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1019",
        "columnas_requeridas": ['Asiento', 'Fuente', 'Fecha', 'Referencia', 'Nit', 'Descripcion NIT', 'Debito Bolivar', 'Credito Bolivar']
    },
    "212.05.1108 - Haberes de Clientes": {
        "id": "haberes_clientes",
        "funcion_principal": run_conciliation_haberes_clientes,
        "label_actual": "Movimientos del mes (Haberes Clientes)",
        "label_anterior": "Saldos anteriores (Haberes Clientes)",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha Origen Acreencia', 'Numero de Documento', 'Referencia', 'Monto Bolivar', 'Monto Dólar'],
        "nombre_hoja_excel": "212.05.1108",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Descripción Nit', 'Débito Bolivar', 'Crédito Bolivar', 'Fuente']
    },
    "212.07.9001 - CDC - Factoring": {
        "id": "cdc_factoring",
        "funcion_principal": run_conciliation_cdc_factoring,
        "label_actual": "Movimientos del mes (Factoring)",
        "label_anterior": "Saldos anteriores (Factoring)",
        # Estas columnas son referenciales para el excel genérico, pero usaremos la función específica
        "columnas_reporte": ['Contrato', 'Documento', 'Saldo USD', 'Tasa', 'Saldo Bs'], 
        "nombre_hoja_excel": "212.07.9001",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Fuente', 'Débito Dolar', 'Crédito Dolar', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "212.05.1005 - Asientos por clasificar": {
        "id": "asientos_por_clasificar",
        "funcion_principal": run_conciliation_asientos_por_clasificar,
        "label_actual": "Movimientos del mes (Por Clasificar)",
        "label_anterior": "Saldos anteriores (Por Clasificar)",
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Asiento', 'Referencia', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1005",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Descripción Nit', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "212.07.1012 - Proveedores d/Mcia - Costos Causados": {
        "id": "proveedores_costos",
        "funcion_principal": run_conciliation_proveedores_costos,
        "label_actual": "Movimientos del mes (Costos Causados)",
        "label_anterior": "Saldos anteriores (Costos Causados)",
        "columnas_reporte": ['NIT', 'Proveedor y Descripcion', 'Fecha.', 'EMB', 'Saldo USD', 'Tasa', 'Bs.','OBSERVACION'],
        "nombre_hoja_excel": "212.07.1012",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Fuente', 'Débito Dolar', 'Crédito Dolar', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "101.01.03.00 - Fondos en Transito COFERSA": {
        "id": "fondos_transito_cofersa",
        "funcion_principal": run_conciliation_fondos_fondos_cofersa, # Nueva función
        "label_actual": "Movimientos del Mes (Fondos)",
        "label_anterior": "Saldos Anteriores (Fondos)",
        "columnas_reporte": ['Fecha', 'Asiento', 'Referencia', 'Monto Colones'],
        "nombre_hoja_excel": "101.01.03.00",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Fuente', 'Débito Local', 'Crédito Local']
    },
    "201081300 - Dev. a Prov. Pais (Colones)": {
        "id": "dev_prov_crc",
        "funcion_principal": lambda df, log, pb: run_conciliation_dev_proveedores_cofersa(df, log, 'CRC'),
        "label_actual": "Movimientos del Mes",
        "label_anterior": "Saldos Anteriores",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Asiento', 'Tipo', 'Referencia', 'Neto Colones', 'Neto Dólar'],
        "nombre_hoja_excel": "201081300",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'NIT', 'Débito Bolivar', 'Crédito Bolivar']
    },
    "201081400 - Dev. a Prov. Exterior (Dólares)": {
        "id": "dev_prov_usd_ext",
        "funcion_principal": lambda df, log, pb: run_conciliation_dev_proveedores_cofersa(df, log, 'USD'),
        "label_actual": "Movimientos del Mes",
        "label_anterior": "Saldos Anteriores",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Asiento', 'Tipo', 'Referencia', 'Neto Dólar', 'Neto Colones'],
        "nombre_hoja_excel": "201081400",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'NIT', 'Débito Dolar', 'Crédito Dolar']
    },
    "Dev. a Prov. Pais ME (Dólares)": {
        "id": "dev_prov_usd_me",
        "funcion_principal": lambda df, log, pb: run_conciliation_dev_proveedores_cofersa(df, log, 'USD'),
        "label_actual": "Movimientos del Mes",
        "label_anterior": "Saldos Anteriores",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Asiento', 'Tipo', 'Referencia', 'Neto Dólar', 'Neto Colones'],
        "nombre_hoja_excel": "DEV_PROV_ME",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'NIT', 'Débito Dolar', 'Crédito Dolar']
    },
    
}

# ==============================================================================
# III. PANEL DE CONTROL (HOME)
# ==============================================================================
def render_inicio():
    # --- CABECERA CON LOGOS ---
    st.markdown("<br>", unsafe_allow_html=True)
    _, col_logos, _ = st.columns([1, 10, 1])
    with col_logos:
        l1, l2, l3 = st.columns(3)
        with l1:
            try: st.image("assets/logo_febeca.png", use_container_width=True)
            except: st.write("**FEBECA**")
        with l2:
            try: st.image("assets/logo_beval.png", use_container_width=True)
            except: st.write("**BEVAL**")
        with l3:
            try: st.image("assets/logo_sillaca.png", use_container_width=True)
            except: st.write("**SILLACA**")
    st.divider()

    st.title("🤖 Portal de Herramientas Contables")
    st.subheader("Grupo Mayoreo", anchor=False)
    st.markdown("Seleccione una herramienta para comenzar:")

    c1, c2, c3 = st.columns(3, gap="medium")
    
    with c1:
        st.subheader("📊 Análisis y Conciliación")
        st.button("📄 Especificaciones", on_click=set_page, args=['especificaciones'], use_container_width=True)
        st.button("📦 Análisis Paquete CC", on_click=set_page, args=['paquete_cc'], use_container_width=True)
        st.button("🏦 Auditoría Comisiones", on_click=set_page, args=['comisiones'], use_container_width=True)

    with c2:
        st.subheader("📑 Cierres Mensuales") # <--- NUEVA SECCIÓN
        st.button("🏦 Cuadre CB - CG", on_click=set_page, args=['cuadre'], use_container_width=True)
        st.button("📈 Ajustes al Balance USD", on_click=set_page, args=['ajustes_usd'], use_container_width=True)
        st.button("📆 Liberaciones y Apartados", on_click=set_page, args=['apartados'], use_container_width=True)

    with c3:
        st.subheader("⚙️ Procesos Fiscales")
        st.button("🛡️ Cálculo Pensiones (9%)", on_click=set_page, args=['pensiones'], use_container_width=True)
        st.button("⚖️ Cálculo LOCTI (0.5%)", on_click=set_page, args=['locti'], use_container_width=True)
        st.button("📑 Verificación Débito Fiscal", on_click=set_page, args=['debito_fiscal'], use_container_width=True)
        st.button("🧾 Relación Retenciones", on_click=set_page, args=['retenciones'], use_container_width=True)


    st.divider()
    st.subheader("COFERSA", anchor=False)
    st.markdown("Seleccione una herramienta para comenzar:")

    # Usamos la misma estructura de 3 columnas que Mayoreo
    col_cof1, col_cof2, col_cof3 = st.columns(3, gap="medium")
    
    with col_cof1:
        st.subheader("📊 Análisis y Conciliación")
        st.button(
            "📄 Especificaciones", 
            on_click=set_page, 
            args=['especificaciones_cofersa'], # <--- Verifica que el nombre sea este exactamente
            use_container_width=True
        )

    with col_cof2:
        st.subheader("⚖️ Cierres Mensuales")
        # Espacio para futuras herramientas de cierre de Cofersa
        st.info("Próximamente nuevas utilidades de cierre.")

    with col_cof3:
        st.subheader("⚙️ Procesos Fiscales")
        # Espacio para futuras herramientas fiscales de Cofersa
        st.info("Próximamente utilidades fiscales.")

    st.markdown("---")
    st.caption("v2.6 - Sistema Integral de Automatización Contable.")
    

# ==============================================================================
# IV. CICLO DE CONCILIACIONES (ESTÁNDAR Y ESPECIALES)
# ==============================================================================
def render_especificaciones():
    st.title('📄 Herramienta de Conciliación de Cuentas', anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_spec"):
        set_page('inicio')
        st.session_state.processing_complete = False 
        st.rerun()
    st.markdown("Esta aplicación automatiza el proceso de conciliación de cuentas contables.")
    
    CASA_OPTIONS = ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"]
    CUENTA_OPTIONS = sorted([k for k in ESTRATEGIAS.keys() if "COFERSA" not in k])
    
    st.subheader("1. Seleccione la Empresa (Casa):", anchor=False)
    casa_seleccionada = st.selectbox("1. Seleccione la Empresa (Casa):", CASA_OPTIONS, label_visibility="collapsed")
    
    st.subheader("2. Seleccione la Cuenta Contable:", anchor=False)
    cuenta_seleccionada = st.selectbox("2. Seleccione la Cuenta Contable:", CUENTA_OPTIONS, label_visibility="collapsed")
    estrategia_actual = ESTRATEGIAS[cuenta_seleccionada]

    with st.expander("📖 Guía Completa: Cómo Usar y Entender la Conciliación", expanded=False):
        st.markdown(GUIA_GENERAL_ESPECIFICACIONES)
        st.divider()
        # Muestra la lógica específica de la cuenta seleccionada
        logica_especifica = LOGICA_POR_CUENTA.get(cuenta_seleccionada, "No hay una guía detallada para esta cuenta.")
        st.markdown(logica_especifica)

    st.subheader("3. Cargue los Archivos de Excel (.xlsx):", anchor=False)
    st.markdown("*Asegúrese de que los datos estén en la **primera hoja** y los **encabezados en la primera fila**.*")

    columnas = estrategia_actual.get("columnas_requeridas", [])
    if columnas:
        texto_columnas = "**Columnas Esenciales Requeridas:**\n" + "\n".join([f"- `{col}`" for col in columnas])
        texto_columnas += "\n\n*Nota: El archivo puede contener más columnas, pero las mencionadas son cruciales para el proceso.*"
        st.info(texto_columnas, icon="ℹ️")

    col1, col2 = st.columns(2)
    with col1:
        uploaded_actual = st.file_uploader(estrategia_actual["label_actual"], type="xlsx", key=f"actual_{estrategia_actual['id']}")
    with col2:
        uploaded_anterior = st.file_uploader(estrategia_actual["label_anterior"], type="xlsx", key=f"anterior_{estrategia_actual['id']}")
        
    if uploaded_actual and uploaded_anterior:
        if st.button("▶️ Iniciar Conciliación", type="primary", use_container_width=True):
            progress_container = st.empty()
            log_messages = []
            try:
                with st.spinner('Cargando y limpiando datos...'):
                    df_full = cargar_y_limpiar_datos(uploaded_actual, uploaded_anterior, log_messages)
                if df_full is not None:
                    progress_container.progress(0, text="Iniciando fases de conciliación...")
                    df_resultado = estrategia_actual["funcion_principal"](df_full.copy(), log_messages, progress_bar=progress_container)
                    progress_container.progress(1.0, text="¡Proceso completado!")
                    
                    st.session_state.df_saldos_abiertos = df_resultado[~df_resultado['Conciliado']].copy()
                    st.session_state.df_conciliados = df_resultado[df_resultado['Conciliado']].copy()
                    
                    # --- GENERACIÓN DE NOMBRE DE ARCHIVO ---
                    codigos_casa = {
                        "FEBECA, C.A": "004",
                        "MAYOR BEVAL, C.A": "207",
                        "PRISMA, C.A": "298",
                        "FEBECA, C.A (QUINCALLA)": "071"
                    }
                    
                    # 1. Código Casa
                    cod = codigos_casa.get(casa_seleccionada, "000")
                    
                    # 2. Número de Cuenta (quitamos la descripción)
                    num_cta = cuenta_seleccionada.split(" - ")[0].strip()
                    
                    # 3. Fecha (Mes y Año)
                    fecha_max = df_full['Fecha'].max()
                    if pd.notna(fecha_max):
                        meses_abr = {1:"ENE", 2:"FEB", 3:"MAR", 4:"ABR", 5:"MAY", 6:"JUN", 7:"JUL", 8:"AGO", 9:"SEP", 10:"OCT", 11:"NOV", 12:"DIC"}
                        fecha_txt = f"{meses_abr[fecha_max.month]}.{str(fecha_max.year)[-2:]}"
                    else:
                        fecha_txt = "SIN_FECHA"
                    
                    # Construir nombre: 071_212.05.1019 NOV.25.xlsx
                    nombre_final = f"{cod}_{num_cta} {fecha_txt}.xlsx"
                    st.session_state.nombre_archivo_salida = nombre_final
                    # ----------------------------------------------

                    st.session_state.excel_saldos_output = generar_excel_saldos_abiertos(st.session_state.df_saldos_abiertos)
                    
                    st.session_state.excel_output = generar_reporte_excel(
                        df_full, st.session_state.df_saldos_abiertos, st.session_state.df_conciliados,
                        estrategia_actual, casa_seleccionada, cuenta_seleccionada
                    )
                    st.session_state.log_messages = log_messages
                    st.session_state.processing_complete = True
                    st.rerun()
                    
            except Exception as e:
                mostrar_error_amigable(e, "la Conciliación")
                st.session_state.processing_complete = False
            finally:
                progress_container.empty()

    if st.session_state.get('processing_complete', False):
        st.success("✅ ¡Conciliación completada con éxito!")
        res_col1, res_col2 = st.columns(2, gap="small")
        with res_col1:
            st.metric("Movimientos Conciliados", len(st.session_state.df_conciliados))
            nombre_descarga = st.session_state.get('nombre_archivo_salida', 'reporte_conciliacion.xlsx')
            
            st.download_button(
                "⬇️ Descargar Reporte Completo (Excel)", 
                st.session_state.excel_output, 
                file_name=nombre_descarga,  # <--- CAMBIO
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True, 
                key="download_excel"
            )
            
        with res_col2:
            st.metric("Saldos Abiertos (Pendientes)", len(st.session_state.df_saldos_abiertos))
            
            # Opcional: También puedes nombrar este archivo parecido si quieres
            # Ej: Saldos_071_212.05.1019 NOV.25.xlsx
            nombre_saldos = "Saldos_" + st.session_state.get('nombre_archivo_salida', 'proximo_mes.xlsx')
            
            st.download_button(
                "⬇️ Descargar Saldos para Próximo Mes (Excel)", 
                st.session_state.excel_saldos_output, 
                file_name=nombre_saldos, # <--- CAMBIO SUGERIDO
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True, 
                key="download_saldos_xlsx"
            )
        
        st.info("**Instrucción de Ciclo Mensual:** Para el próximo mes, debe usar el archivo CSV descargado como el archivo de 'saldos anteriores'.")
        
        with st.expander("Ver registro detallado del proceso"):
            st.text_area("Log de Conciliación", '\n'.join(st.session_state.log_messages), height=300, key="log_area")
            
        st.subheader("Previsualización de Saldos Pendientes", anchor=False)
        df_vista_previa = st.session_state.df_saldos_abiertos.copy()
        
        if estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar', 'devoluciones_proveedores', 'cuentas_viajes']:
            columnas_numericas = ['Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar', 'Monto_BS', 'Monto_USD']
            for col in columnas_numericas:
                if col in df_vista_previa.columns:
                    df_vista_previa[col] = df_vista_previa[col].apply(lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else '')
            if 'Fecha' in df_vista_previa.columns:
                df_vista_previa['Fecha'] = pd.to_datetime(df_vista_previa['Fecha']).dt.strftime('%d/%m/%Y')
            st.dataframe(df_vista_previa, use_container_width=True)

        st.subheader("Previsualización de Movimientos Conciliados", anchor=False)
        df_conciliados_vista = st.session_state.df_conciliados.copy()
        
        if estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar', 'devoluciones_proveedores', 'cuentas_viajes']:
            columnas_numericas_conc = ['Monto_BS', 'Monto_USD']
            for col in columnas_numericas_conc:
                 if col in df_conciliados_vista.columns:
                    df_conciliados_vista[col] = df_conciliados_vista[col].apply(lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else '')
            if 'Fecha' in df_conciliados_vista.columns:
                df_conciliados_vista['Fecha'] = pd.to_datetime(df_conciliados_vista['Fecha']).dt.strftime('%d/%m/%Y')
            st.dataframe(df_conciliados_vista, use_container_width=True)

def render_especificaciones_cofersa():
    st.title('📄 Especificaciones de Cuentas: COFERSA', anchor=False)
    
    if st.button("⬅️ Volver al Inicio", key="back_from_spec_cof"):
        set_page('inicio')
        st.rerun()

    # FILTRO CLAVE: Aquí solo tomamos las cuentas que tienen la palabra "COFERSA"
    CUENTA_OPTIONS = sorted([k for k in ESTRATEGIAS.keys() if "COFERSA" in k])
    
    st.subheader("Seleccione la Cuenta Contable de COFERSA:", anchor=False)
    cuenta_seleccionada = st.selectbox("Cuenta:", CUENTA_OPTIONS, label_visibility="collapsed")
    estrategia_actual = ESTRATEGIAS[cuenta_seleccionada]

    # Guía de la cuenta
    with st.expander("📖 Guía de Conciliación", expanded=False):
        from guides import LOGICA_POR_CUENTA
        st.markdown(LOGICA_POR_CUENTA.get(cuenta_seleccionada, "Guía no disponible."))

    st.subheader("Cargue los Archivos (.xlsx o .xls):", anchor=False)
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_actual = st.file_uploader("Movimientos del Mes Actual", type=['xlsx', 'xls'], key="cof_act")
    with col2:
        uploaded_anterior = st.file_uploader("Saldos del Mes Anterior", type=['xlsx', 'xls'], key="cof_ant")
        
    if uploaded_actual and uploaded_anterior:
        if st.button("▶️ Iniciar Conciliación COFERSA", type="primary", use_container_width=True):
            log = []
            try:
                # Usamos el cargador especial de Cofersa que ya definimos
                from utils import cargar_datos_fondos_cofersa
                df_full = cargar_datos_fondos_cofersa(uploaded_actual, uploaded_anterior, log)
                
                if df_full is not None:
                    # Ejecutamos la lógica (la función_principal definida en ESTRATEGIAS)
                    df_res = estrategia_actual["funcion_principal"](df_full.copy(), log)
                    
                    # Generar Reporte
                    df_saldos = df_res[~df_res['Conciliado']]
                    df_conciliados = df_res[df_res['Conciliado']]
                    
                    from utils import generar_reporte_excel, generar_excel_saldos_abiertos
                    excel_reporte = generar_reporte_excel(
                        df_res, df_saldos, df_conciliados, estrategia_actual, "COFERSA", cuenta_seleccionada
                    )
                    
                    st.success("✅ Conciliación completada.")
                    
                    # Métricas y Descargas
                    c1, c2 = st.columns(2)
                    c1.metric("Movimientos Conciliados", len(df_conciliados))
                    c2.metric("Saldos Abiertos", len(df_saldos))
                    
                    st.download_button("⬇️ Descargar Reporte Final", excel_reporte, f"Conciliacion_{cuenta_seleccionada[:10]}.xlsx", use_container_width=True)
                    st.download_button("⬇️ Descargar Saldos Próximo Mes", generar_excel_saldos_abiertos(df_saldos), "Saldos_Anteriores.xlsx", use_container_width=True)
                    
                    with st.expander("Ver Log"): st.write(log)
            except Exception as e:
                st.error(f"Error: {e}")


def render_cuadre():
    st.title("⚖️ Cuadre de Disponibilidad (CB vs CG)", anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_cuadre"):
        set_page('inicio')
        st.rerun()
    
    # --- SELECTOR DE EMPRESA ---
    CASA_OPTIONS = ["MAYOR BEVAL, C.A", "FEBECA, C.A", "FEBECA, C.A (QUINCALLA)", "PRISMA, C.A"]
    col_emp, _ = st.columns([1, 1])
    with col_emp:
        empresa_sel = st.selectbox("Seleccione la Empresa:", CASA_OPTIONS, key="empresa_cuadre")
    
    st.info("Sube el Reporte de Tesorería (CB) y el Balance de Comprobación (CG). Pueden ser PDF o Excel.")
    
    # --- CARGA DE ARCHIVOS ---
    col1, col2 = st.columns(2)
    with col1:
        file_cb = st.file_uploader("1. Reporte Tesorería (CB)", type=['pdf', 'xlsx'])
    with col2:
        file_cg = st.file_uploader("2. Balance Contable (CG)", type=['pdf', 'xlsx'])
        
    # --- BOTÓN DE ACCIÓN ---
    if file_cb and file_cg:
        # Inicializamos la memoria manual si no existe
        if 'mapeo_manual' not in st.session_state:
            st.session_state.mapeo_manual = {}

        if st.button("Comparar Saldos", type="primary", use_container_width=True):
            log = []
            try:
                # --- SE MANTIENE TU VALIDACIÓN DE SEGURIDAD ---
                es_valido_cb, msg_cb = validar_coincidencia_empresa(file_cb, empresa_sel)
                if not es_valido_cb:
                    st.error(f"⛔ ALERTA DE SEGURIDAD (Tesorería): {msg_cb}")
                    st.stop()
                
                es_valido_cg, msg_cg = validar_coincidencia_empresa(file_cg, empresa_sel)
                if not es_valido_cg:
                    st.error(f"⛔ ALERTA DE SEGURIDAD (Contabilidad): {msg_cg}")
                    st.stop()
                # ----------------------------------------------

                with st.spinner("Analizando y cruzando saldos..."):
                    # Llamamos a la lógica enviando también el mapeo_manual
                    df_res, df_huerfanos = run_cuadre_cb_cg(
                        file_cb, file_cg, empresa_sel, log, st.session_state.mapeo_manual
                    )
                    
                    # Guardamos en sesión para que la tabla no desaparezca
                    st.session_state.res_cuadre = df_res
                    st.session_state.huerfanos_cuadre = df_huerfanos
                    st.session_state.log_cuadre = log

            except Exception as e:
                mostrar_error_amigable(e, "el Cuadre CB-CG")

    # --- MOSTRAR RESULTADOS (Fuera del botón para que no se borren) ---
    if 'res_cuadre' in st.session_state:
        df_res = st.session_state.res_cuadre
        df_huerfanos = st.session_state.huerfanos_cuadre
        
        # SI HAY HUÉRFANOS, MOSTRAMOS EL EDITOR PARA QUE EL USUARIO LOS ARREGLE
        if not df_huerfanos.empty:
            cb_orphans = df_huerfanos[df_huerfanos['Origen'] == 'TESORERÍA (CB)']
            if not cb_orphans.empty:
                st.warning("⚠️ Hay bancos en Tesorería que no están en el diccionario.")
                with st.expander("🛠️ ASIGNAR CUENTAS MANUALMENTE", expanded=True):
                    # Creamos tabla para editar
                    df_edit = pd.DataFrame({
                        'Código CB': cb_orphans['Código/Cuenta'],
                        'Nombre': cb_orphans['Descripción/Nombre'],
                        'Cuenta Contable (Escribir)': '',
                        'Moneda': 'VES'
                    })
                    
                    edited = st.data_editor(df_edit, key="editor_cuadre", hide_index=True, use_container_width=True,
                                          column_config={"Moneda": st.column_config.SelectboxColumn(options=["VES", "USD", "EUR"])})
                    
                    if st.button("🔄 Actualizar y volver a Calcular", type="primary", use_container_width=True):
                        for _, row in edited.iterrows():
                            # Solo agregamos si el usuario escribió algo en el campo de cuenta
                            if row['Cuenta Contable (Escribir)'].strip():
                                st.session_state.mapeo_manual[row['Código CB']] = {
                                    "cta": row['Cuenta Contable (Escribir)'].strip(),
                                    "moneda": row['Moneda']
                                }
                    
                    # RE-CALCULAR INMEDIATAMENTE
                    log_new = []
                    df_res_new, df_huerfanos_new = run_cuadre_cb_cg(
                        file_cb, file_cg, empresa_sel, log_new, st.session_state.mapeo_manual
                    )
                    
                    # Actualizamos la sesión con los datos limpios
                    st.session_state.res_cuadre = df_res_new
                    st.session_state.huerfanos_cuadre = df_huerfanos_new
                    st.session_state.log_cuadre = log_new
                    
                    st.success("✅ Reporte actualizado. Los huérfanos mapeados se han movido al resumen general.")
                    st.rerun()

        st.subheader("Resumen de Saldos", anchor=False)
        cols_pantalla = ['Moneda', 'Banco (Tesorería)', 'Cuenta Contable', 'Descripción', 'Saldo Final CB', 'Saldo Final CG', 'Diferencia', 'Estado']
        st.dataframe(df_res[cols_pantalla], use_container_width=True)
        
        # Mantenemos tu alerta de cuentas huérfanas original si aún quedan
        if not df_huerfanos.empty:
            st.error(f"⚠️ ATENCIÓN: Quedan {len(df_huerfanos)} cuentas sin configurar.")

        # --- GENERACIÓN DE EXCEL (Tu lógica original) ---
        excel_data = generar_reporte_cuadre(df_res, df_huerfanos, empresa_sel)
        st.download_button(label="⬇️ Descargar Reporte Final (Excel)", data=excel_data,
                         file_name=f"Cuadre_CB_CG_{empresa_sel}.xlsx", use_container_width=True)

def render_ajustes_usd():
    st.title("📈 Ajustes al Balance en USD", anchor=False)

    # 1. INICIALIZACIÓN DEL ESTADO (Para que no se borren los ajustes manuales al interactuar)
    if 'manual_adjustments' not in st.session_state:
        st.session_state.manual_adjustments = pd.DataFrame(columns=[
            'Cuenta a ajustar', 'Cuenta contrapartida', 'Monto en USD', 'Tasa de conversión'
        ])

    if st.button("⬅️ Volver al Inicio", key="btn_back_adj"):
        set_page('inicio')
        st.rerun()

    # --- SECCIÓN 1: CARGA DE ARCHIVOS ---
    st.subheader("1. Archivos de Entrada", anchor=False)
    col1, col2 = st.columns(2)
    with col1:
        f_cg = st.file_uploader("1. Balance de Comprobación (xlsx, xls, pdf)", type=['xlsx', 'xls', 'pdf'], key="adj_cg")
        f_cb = st.file_uploader("2. Reporte de Tesorería (xlsx, xls)", type=['xlsx', 'xls'], key="adj_cb")
    with col2:
        f_hab_usd = st.file_uploader("3. Reporte de Haberes USD (pdf)", type=['pdf'], key="adj_hab_usd")
        f_hab_ves = st.file_uploader("4. Reporte de Haberes VES (pdf)", type=['pdf'], key="adj_hab_ves")

    # --- SECCIÓN 2: AJUSTES MANUALES (Tabla Interactiva) ---
    st.subheader("2. Ajustes Manuales", anchor=False)
    st.markdown("Agregue aquí ajustes extraordinarios que la herramienta deba incluir en el reporte y cargador.")
    
    # 1. Inicializar la lista en el estado de la sesión si no existe
    if 'manual_adj_list' not in st.session_state:
        st.session_state.manual_adj_list = []

    # --- FUNCIÓN CALLBACK PARA AGREGAR ---
    def agregar_ajuste_callback():
        # Extraemos valores de las llaves (keys)
        cta = st.session_state.man_cta
        contra = st.session_state.man_contra
        moneda = st.session_state.man_moneda
        monto = st.session_state.man_monto
        tasa = st.session_state.get('man_tasa', 'N/A')

        if cta.strip() and contra.strip() and monto != 0:
            st.session_state.manual_adj_list.append({
                "cuenta": cta,
                "contrapartida": contra,
                "moneda": moneda,
                "monto": monto,
                "tasa_tipo": tasa
            })
            # LIMPIEZA: Ahora es seguro porque ocurre en el callback
            st.session_state.man_cta = ""
            st.session_state.man_contra = ""
            st.session_state.man_monto = 0.0
        else:
            # Usamos un flag temporal para mostrar error en el siguiente render
            st.session_state.error_manual = True

    # --- 2. FUNCIÓN CALLBACK PARA MODIFICAR (Solución al Error) ---
    def modificar_ajuste_callback(idx):
        item = st.session_state.manual_adj_list.pop(idx)
        # Seteamos los valores de los widgets ANTES del rerun
        st.session_state.man_cta = item['cuenta']
        st.session_state.man_contra = item['contrapartida']
        st.session_state.man_moneda = item['moneda']
        st.session_state.man_monto = item['monto']
        if item['moneda'] == "USD":
            st.session_state.man_tasa = item['tasa_tipo']

    # --- 3. FUNCIÓN CALLBACK PARA ELIMINAR ---
    def eliminar_ajuste_callback(idx):
        st.session_state.manual_adj_list.pop(idx)

    # --- 4. FORMULARIO DE ENTRADA ---
    with st.container(border=True):
        col_m1, col_m2, col_m3 = st.columns([2, 2, 1])
        with col_m1:
            st.text_input("Cuenta a ajustar", placeholder="Ej: 1.1.02...", key="man_cta")
            st.text_input("Cuenta contrapartida", placeholder="Ej: 2.1.02...", key="man_contra")
        with col_m2:
            st.selectbox("Moneda del ajuste", ["USD", "BS"], key="man_moneda")
            # Accedemos al valor actual de la moneda para el label
            curr_mon = st.session_state.get('man_moneda', 'USD')
            st.number_input(f"Monto en {curr_mon}", format="%.2f", key="man_monto")
        with col_m3:
            if st.session_state.get('man_moneda') == "USD":
                st.selectbox("Tasa de conversión", ["BCV", "CORP"], key="man_tasa")
            st.markdown("<br>", unsafe_allow_html=True)
            
            # EL BOTÓN LLAMA AL CALLBACK
            st.button("➕ Agregar", 
                      use_container_width=True, 
                      type="secondary", 
                      on_click=agregar_ajuste_callback)

    if st.session_state.get('error_manual'):
        st.error("Complete los campos obligatorios (Cuentas y Monto distinto a 0)")
        st.session_state.error_manual = False   

    # --- 5. LISTADO DE AJUSTES AGREGADOS ---
    if st.session_state.manual_adj_list:
        st.markdown("---")
        st.write("**Listado de Ajustes Extraordinarios:**")
        
        for idx, item in enumerate(st.session_state.manual_adj_list):
            with st.expander(f"📌 {item['cuenta']} vs {item['contrapartida']} | {item['monto']} {item['moneda']}", expanded=True):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.write(f"**Monto:** {item['monto']:,} {item['moneda']} ({item['tasa_tipo']})")
                with c2:
                    # MODIFICAR: Carga valores y elimina de la lista
                    st.button("📝 Modificar", key=f"edit_{idx}", on_click=modificar_ajuste_callback, args=(idx,))
                with c3:
                    st.button("🗑️ Eliminar", key=f"del_{idx}", on_click=eliminar_ajuste_callback, args=(idx,))

    # --- SECCIÓN 3: PARÁMETROS ---
    st.subheader("3. Parámetros de Cálculo", anchor=False)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tasa_bcv = st.number_input("Tasa BCV (Cierre)", min_value=0.01, value=1.0, format="%.4f", key="val_tasa_bcv")
    with c2:
        tasa_corp = st.number_input("Tasa CORP (Interna)", min_value=0.01, value=1.0, format="%.4f", key="val_tasa_corp")
    with c3:
        empresa_sel = st.selectbox("Empresa", ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "SILLACA, C.A"], key="val_empresa")
    with c4:
        n_asiento = st.text_input("Número de Asiento", value="CG0000", key="val_asiento")

    # --- SECCIÓN 4: EJECUCIÓN ---
    st.divider()
    if st.button("🚀 Calcular Ajustes y Generar Reporte", type="primary", use_container_width=True, key="btn_run_adj"):
        if not f_cg or not f_cb:
            st.error("⚠️ El Balance de Comprobación y el Reporte de Tesorería son obligatorios para este proceso.")
        else:
            log_messages = []
            try:
                # 1. Procesar Lógica
                with st.spinner("Ejecutando motor de ajustes bimonetarios..."):
                    # Importamos aquí para asegurar que los cambios en logic.py se apliquen
                    from logic import procesar_ajustes_balance_usd
                    from utils import generar_reporte_ajustes_usd

                    # 1. Convertimos la lista de ajustes manuales acumulada en un DataFrame
                    # Si la lista está vacía, creará un DataFrame vacío que la lógica sabe manejar.
                    df_manual_para_procesar = pd.DataFrame(st.session_state.manual_adj_list)

                    # 2. Llamamos a la lógica con el nuevo DataFrame
                    df_res, df_banc, df_asiento, df_raw, val_data = procesar_ajustes_balance_usd(
                        f_cb, f_cg, f_hab_usd, f_hab_ves, 
                        tasa_bcv, tasa_corp, empresa_sel, n_asiento, 
                        df_manual_para_procesar, # <--- Variable corregida
                        log_messages
                    )

                # 2. Mostrar Resultados y Descarga
                if not df_asiento.empty:
                    st.success("✅ Ajustes procesados exitosamente.")
                    
                    # Generar binario del Excel
                    excel_bin = generar_reporte_ajustes_usd(
                        df_res, df_banc, df_asiento, df_raw, empresa_sel, val_data
                    )

                    st.download_button(
                        label="📥 Descargar Reporte de Ajustes y Asiento",
                        data=excel_bin,
                        file_name=f"Ajustes_Balance_USD_{empresa_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                    with st.expander("🔍 Ver Vista Previa del Asiento"):
                        st.dataframe(df_asiento, use_container_width=True)
                else:
                    st.warning("El proceso finalizó pero no se generaron movimientos de ajuste relevantes.")

                # Mostrar Log
                with st.expander("📄 Ver Log de Extracción y Proceso"):
                    for msg in log_messages:
                        st.text(msg)

            except Exception as e:
                # Función de error amigable que ya tienes en app.py
                mostrar_error_amigable(e, "el cálculo de Ajustes al Balance")





def render_paquete_cc():
    st.title('📦 Herramienta de Análisis de Paquete CC', anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_paquete"):
        set_page('inicio')
        if 'processing_paquete_complete' in st.session_state:
            del st.session_state['processing_paquete_complete']
        st.rerun()
    
    st.markdown("Esta herramienta analiza el diario contable para clasificar y agrupar los asientos.")

    with st.expander("📖 Manual de Usuario: Criterios de Análisis y Errores Comunes", expanded=False):
        st.markdown(GUIA_PAQUETE_CC)
    
    CASA_OPTIONS = ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"]
    st.subheader("1. Seleccione la Empresa (Casa):", anchor=False)
    casa_seleccionada = st.selectbox("Empresa", CASA_OPTIONS, label_visibility="collapsed", key="casa_paquete_cc")
    
    st.subheader("2. Cargue el Archivo de Movimientos del Diario (.xlsx):", anchor=False)
    
    columnas_requeridas = ['Asiento', 'Fecha', 'Fuente', 'Cuenta Contable', 'Descripción de Cuenta', 'Referencia', 'Débito Dolar', 'Crédito Dolar', 'Débito VES', 'Crédito VES']
    texto_columnas = "**Columnas Esenciales Requeridas:**\n" + "\n".join([f"- `{col}`" for col in columnas_requeridas])
    st.info(texto_columnas, icon="ℹ️")
    
    uploaded_diario = st.file_uploader("Movimientos del Diario Contable", type="xlsx", label_visibility="collapsed")
    
    if uploaded_diario:
        if st.button("▶️ Iniciar Análisis", type="primary", use_container_width=True):
            with st.spinner('Ejecutando análisis de asientos... Este proceso puede tardar unos momentos.'):
                log_messages = []
                try:
                    df_diario = pd.read_excel(uploaded_diario)
                    
                    # Mapeo robusto para estandarizar nombres de columnas
                    # Nombres estándar que la lógica espera
                    standard_names = {
                        'Débito Dolar': ['Debito Dolar', 'Débitos Dolar', 'Débito Dólar', 'Debito Dólar'],
                        'Crédito Dolar': ['Credito Dolar', 'Créditos Dolar', 'Crédito Dólar', 'Credito Dólar'],
                        'Débito VES': ['Debito VES', 'Débitos VES', 'Débito Bolivar', 'Debito Bolivar', 'Débito Bs', 'Debito Bs'],
                        'Crédito VES': ['Credito VES', 'Créditos VES', 'Crédito Bolivar', 'Credito Bolivar', 'Crédito Bs', 'Credito Bs'],
                        'Descripción de Cuenta': ['Descripcion de Cuenta', 'Descripción de la Cuenta', 'Descripcion de la Cuenta', 'Descripción de la Cuenta Contable', 'Descripcion de la Cuenta Contable']
                    }

                    # Crear un diccionario para el renombrado final
                    rename_map = {}
                    # Normalizar las columnas del DataFrame para una comparación más fácil
                    df_columns_normalized = {col.strip(): col for col in df_diario.columns}

                    for standard, variations in standard_names.items():
                        # Añadir el nombre estándar a su propia lista de variaciones
                        all_variations = [standard] + variations
                        for var_name in all_variations:
                            if var_name in df_columns_normalized:
                                # Si encontramos una variación, la mapeamos al nombre estándar
                                rename_map[df_columns_normalized[var_name]] = standard
                                break # Pasamos al siguiente nombre estándar

                    # Aplicar el renombrado
                    df_diario.rename(columns=rename_map, inplace=True)
                    log_messages.append(f"✔️ Columnas estandarizadas. Mapeo aplicado: {rename_map}")

                    df_resultado = run_analysis_paquete_cc(df_diario, log_messages)
                    
                    st.session_state.reporte_paquete_output = generar_reporte_paquete_cc(df_resultado, casa_seleccionada)
                    st.session_state.log_messages_paquete = log_messages
                    st.session_state.processing_paquete_complete = True
                    st.rerun()

                except Exception as e:
                    mostrar_error_amigable(e, "el Análisis de Paquete CC")
                    st.session_state.processing_paquete_complete = False

    if st.session_state.get('processing_paquete_complete', False):
        st.success("✅ ¡Análisis de Paquete CC completado con éxito!")
        st.download_button(
            "⬇️ Descargar Reporte de Análisis (Excel)",
            st.session_state.reporte_paquete_output,
            "Reporte_Analisis_Paquete_CC.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        with st.expander("Ver registro detallado del proceso de análisis"):
            st.text_area("Log de Análisis", '\n'.join(st.session_state.log_messages_paquete), height=400)



# ==============================================================================
# V. CICLO FISCAL Y DE AUDITORÍA
# ==============================================================================
def render_retenciones():
    st.title("🧾 Herramienta de Auditoría de Retenciones", anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_ret"):
        set_page('inicio')
        if 'processing_ret_complete' in st.session_state:
            del st.session_state['processing_ret_complete']
        st.rerun()

    st.markdown("""
    Esta herramienta audita el proceso de retenciones cruzando la **Preparación Contable (CP)**, 
    la **Fuente Oficial (GALAC)** y el **Diario Contable (CG)** para identificar discrepancias.
    """)

    # --- El expander ahora lee el texto desde el archivo guides.py ---
    with st.expander("📖 Guía Completa: Cómo Usar y Entender la Herramienta de Auditoría", expanded=True):
        st.markdown(GUIA_COMPLETA_RETENCIONES)

    st.subheader("1. Cargue los Archivos de Excel (.xlsx):", anchor=False)
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("Archivos de Preparación y Registro")
        file_cp = st.file_uploader("1. Relacion_Retenciones_CP.xlsx", type="xlsx")
        file_cg = st.file_uploader("2. Transacciones_Diario_CG.xlsx", type="xlsx")

    with col2:
        st.info("Archivos Oficiales (Fuente GALAC)")
        file_iva = st.file_uploader("3. Retenciones_IVA.xlsx", type="xlsx")
        file_islr = st.file_uploader("4. Retenciones_ISLR.xlsx", type="xlsx")
        file_mun = st.file_uploader("5. Retenciones_Municipales.xlsx", type="xlsx")

    if all([file_cp, file_cg, file_iva, file_islr, file_mun]):
        if st.button("▶️ Iniciar Auditoría de Retenciones", type="primary", use_container_width=True):
            with st.spinner('Ejecutando auditoría... Este proceso puede tardar unos momentos.'):
                log_messages = []
                
                try:
                    reporte_resultado = run_conciliation_retenciones(
                        file_cp, file_cg, file_iva, file_islr, file_mun, log_messages
                    )
                    
                    if reporte_resultado is None:
                        raise Exception("Error interno: La lógica devolvió un resultado vacío.")

                    st.session_state.reporte_ret_output = reporte_resultado
                    st.session_state.log_messages_ret = log_messages
                    st.session_state.processing_ret_complete = True
                    st.rerun()

                except Exception as e:
                    mostrar_error_amigable(e, "la Auditoría de Retenciones")
                    st.session_state.log_messages_ret = log_messages
                    # No activamos processing_ret_complete en error para no mostrar el botón de descarga vacío,
                    # pero sí guardamos los logs por si quieres verlos.
                    st.session_state.processing_ret_complete = True 
                    # Importante: Si hubo error, reporte_ret_output debe ser None
                    st.session_state.reporte_ret_output = None 

    # --- BLOQUE 2: MOSTRAR RESULTADOS (Fuera del if del botón) ---
    if st.session_state.get('processing_ret_complete', False):
        
        # Solo mostramos Éxito y Descarga si hay un reporte generado
        if st.session_state.get('reporte_ret_output') is not None:
            st.success("✅ ¡Auditoría de retenciones completada con éxito!")
            st.download_button(
                "⬇️ Descargar Reporte de Auditoría (Excel)",
                st.session_state.reporte_ret_output,
                "Reporte_Auditoria_Retenciones.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        # El log lo mostramos siempre (sea éxito o error controlado)
        if 'log_messages_ret' in st.session_state and st.session_state.log_messages_ret:
            with st.expander("Ver registro detallado del proceso de auditoría"):
                st.text_area("Log de Auditoría de Retenciones", '\n'.join(st.session_state.log_messages_ret), height=400)


def render_pensiones():
    st.title("⚖️ Cálculo Ley Protección Pensiones (9%)", anchor=False)

    # --- 1. ESTILOS CSS PARA REPLICAR LOCTI (DARK PREMIUM) ---
    st.markdown("""
        <style>
        .stContainer {
            background-color: #1a1c24;
            border-radius: 15px;
            padding: 20px;
            border: 1px solid #333;
        }
        .card-header {
            padding: 10px;
            border-radius: 10px 10px 0 0;
            font-weight: bold;
            text-align: center;
            color: white;
            margin-bottom: -10px;
        }
        .purple-header { background-color: #4b0082; }
        .green-header { background-color: #1e5631; }
        
        /* Ajuste para que los file uploaders se vean como en la imagen */
        div[data-testid="stFileUploader"] {
            background-color: #1a1c24 !important;
            border: 1px solid #333 !important;
            border-radius: 0 0 15px 15px !important;
        }
        </style>
    """, unsafe_allow_html=True)

    if st.button("⬅️ Volver al Inicio", key="back_pen"):
        set_page('inicio')
        st.rerun()

    # --- 2. SECCIÓN DE PARÁMETROS (3 Columnas - Fecha Eliminada) ---
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            EMPRESAS_NOMINA = ["FEBECA", "BEVAL", "PRISMA", "QUINCALLA"]
            empresa_sel = st.selectbox("🏢 Seleccione la Filial:", EMPRESAS_NOMINA)
        with c2:
            tasa = st.number_input("💵 Tasa de Cambio:", min_value=0.01, value=1.0, format="%.4f")
        with c3:
            num_asiento = st.text_input("🔢 N° Asiento:", value="CG0000")
        
        analista_nombre = st.text_input("👤 Hecho por:", placeholder="Nombre del analista")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"📥 Carga de Balances: {empresa_sel}", anchor=False)

    # --- 3. SECCIÓN DE CARGA (Tarjetas de Colores) ---
    col_file1, col_file2 = st.columns(2)

    with col_file1:
        st.markdown('<div class="card-header purple-header">📊 Balance VENTAS / MAYOR</div>', unsafe_allow_html=True)
        with st.container(border=True):
            file_mayor = st.file_uploader("Subir Mayor Analítico", type=['xlsx'], key="pen_mayor", label_visibility="collapsed")
        st.caption("📍 Debe contener las cuentas 7.1.1.01 y 7.1.1.09")

    with col_file2:
        st.markdown('<div class="card-header green-header">💰 Balance INGRESOS / NÓMINA</div>', unsafe_allow_html=True)
        with st.container(border=True):
            file_nomina = st.file_uploader("Subir Resumen RRHH", type=['xlsx'], key="pen_nom", label_visibility="collapsed")
        st.caption("📍 La pestaña debe coincidir con el Mes/Año del cálculo")

    st.divider()

    # --- 4. BOTÓN DE PROCESO ---
    if file_mayor and tasa > 0:
        if st.button("🚀 Calcular Impuesto y Generar Asiento", type="primary", use_container_width=True):
            log = []
            try:
                from logic import procesar_calculo_pensiones
                from utils import generar_reporte_pensiones, generar_cargador_asiento_pensiones
                import datetime
                
                with st.spinner("Analizando bases imponibles..."):
                    df_calc, df_base, df_asiento, dict_val = procesar_calculo_pensiones(file_mayor, file_nomina, tasa, empresa_sel, log, num_asiento)
                
                if df_asiento is not None and not df_asiento.empty:
                    # Determinamos la fecha automáticamente para no pedirla en la interfaz
                    # Si el proceso detectó una fecha en el mayor, la usamos, sino usamos hoy
                    fecha_proceso = pd.Timestamp.now()
                    
                    st.success("✅ Cálculo finalizado satisfactoriamente.")
                    
                    # Alertas de Auditoría (Estilo LOCTI)
                    if dict_val.get('estado') != 'OK':
                        st.error(f"⚠️ Descuadre detectado contra nómina: {dict_val.get('dif_base_total'):,.2f} Bs.")
                    
                    if dict_val.get('tiene_cc_genericos'):
                        st.warning(f"🚨 Centros de costo .00 detectados: {dict_val.get('lista_cc_genericos')}")

                    # Bloque de Descargas
                    c_down1, c_down2 = st.columns(2)
                    with c_down1:
                        cargador_bin = generar_cargador_asiento_pensiones(df_asiento, fecha_proceso)
                        st.download_button("📥 Descargar Cargador Softland", cargador_bin, f"CARGADOR_{num_asiento}.xlsx", use_container_width=True)
                    
                    with c_down2:
                        excel_data = generar_reporte_pensiones(df_calc, df_base, df_asiento, dict_val, empresa_sel, tasa, fecha_proceso, analista_nombre)
                        st.download_button("📊 Descargar Memoria de Cálculo", excel_data, f"PENSIONES_{empresa_sel}.xlsx", use_container_width=True)

                    # Vista previa para control del usuario
                    with st.expander("🔍 Ver Vista Previa del Asiento"):
                        st.dataframe(df_asiento, use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Error en el proceso: {str(e)}")
    else:
        st.info("💡 Por favor, cargue los archivos necesarios para habilitar el cálculo.")

def render_debito_fiscal():
    st.title("📑 Verificación de Débito Fiscal (Bs.)", anchor=False)
    
    # Botón para regresar al inicio
    if st.button("⬅️ Volver al Inicio"): 
        set_page('inicio')
        st.rerun()

    # Guía de uso (Cargada desde guides.py)
    with st.expander("📖 Guía de Uso: Preparación y Reglas de Negocio", expanded=False):
        st.markdown(GUIA_DEBITO_FISCAL)
    
    st.info("Cruce de auditoría: Softland (Diario + Mayor) vs Libro de Ventas (Imprenta)")
    
    # --- SECCIÓN DE PARÁMETROS ---
    col_a, col_b = st.columns(2)
    with col_a:
        casa_sel = st.selectbox("Seleccione la Empresa:", ["FEBECA (FB + SC)", "BEVAL", "PRISMA"])
    with col_b:
        tolerancia = st.number_input("Margen de Tolerancia en Bs.:", min_value=0.0, value=50.0, help="Diferencias menores a este monto se marcarán como OK.")

    st.divider()

    # --- SECCIÓN DE CARGA DE ARCHIVOS ---
    # Caso 1: Consolidado Febeca + Sillaca (Requiere 4 archivos de Softland)
    if "FEBECA" in casa_sel:
        st.subheader("📁 Archivos Softland: Febeca + Sillaca")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Casa Febeca (FB)**")
            f_fb_d = st.file_uploader("Transacciones Diario (FB)", type=['xlsx'], key="fb_d")
            f_fb_m = st.file_uploader("Transacciones Mayor (FB)", type=['xlsx'], key="fb_m")
        with c2:
            st.markdown("**Casa Sillaca (SC)**")
            f_sc_d = st.file_uploader("Transacciones Diario (SC)", type=['xlsx'], key="sc_d")
            f_sc_m = st.file_uploader("Transacciones Mayor (SC)", type=['xlsx'], key="sc_m")
        
        st.subheader("📄 Libro de Ventas")
        f_imp = st.file_uploader("Archivo de Imprenta (Excel)", type=['xlsx'], key="imp_f")
        ready = all([f_fb_d, f_fb_m, f_sc_d, f_sc_m, f_imp])

    # Caso 2: Otras filiales (Solo requiere 2 archivos de Softland)
    else:
        st.subheader(f"📁 Archivos Softland: {casa_sel}")
        c1, c2 = st.columns(2)
        with c1:
            f_d = st.file_uploader("Transacciones del Diario", type=['xlsx'], key="std_d")
            f_m = st.file_uploader("Transacciones del Mayor", type=['xlsx'], key="std_m")
        with c2:
            st.subheader("📄 Libro de Ventas")
            f_imp = st.file_uploader("Libro de Ventas (Imprenta)", type=['xlsx'], key="std_i")
        ready = all([f_d, f_m, f_imp])

    # --- PROCESAMIENTO ---
    if ready:
        if st.button("▶️ Ejecutar Verificación Cruzada", type="primary", use_container_width=True):
            log = []
            try:
                with st.spinner("Procesando auditoría..."):
                    # 1. Cargamos y preparamos los datos de Softland
                    if "FEBECA" in casa_sel:
                        # Preparamos Febeca y Sillaca por separado y luego unimos
                        soft_fb = preparar_datos_softland_debito(pd.read_excel(f_fb_d), pd.read_excel(f_fb_m), "FB")
                        soft_sc = preparar_datos_softland_debito(pd.read_excel(f_sc_d), pd.read_excel(f_sc_m), "SC")
                        soft_total = pd.concat([soft_fb, soft_sc], ignore_index=True)
                    else:
                        # Preparación normal para Beval o Prisma
                        soft_total = preparar_datos_softland_debito(pd.read_excel(f_d), pd.read_excel(f_m), casa_sel[:2].upper())

                    # 2. Cargamos el archivo de Imprenta
                    # Raw para la hoja de copia fiel y Logic con el header en fila 8 (index 7)
                    df_imp_raw = pd.read_excel(f_imp, header=None)
                    df_imp_logic = pd.read_excel(f_imp, header=7)
                    df_imp_logic.dropna(how='all', inplace=True)

                    # 3. DETERMINACIÓN DINÁMICA DE EXCLUSIÓN (MEJORA SOLICITADA)
                    # Si es BEVAL, obviamos Beval. Si es FEBECA, obviamos Febeca.
                    if "FEBECA" in casa_sel:
                        tag_obviar = "FEBECA"
                    elif "BEVAL" in casa_sel:
                        tag_obviar = "BEVAL"
                    elif "PRISMA" in casa_sel:
                        tag_obviar = "PRISMA"
                    else:
                        tag_obviar = ""

                    # 4. Ejecutamos la lógica de conciliación con el tag dinámico
                    df_res = run_conciliation_debito_fiscal(soft_total, df_imp_logic, tolerancia, log, tag_obviar)
                    
                    # 5. Generamos el reporte Excel
                    excel_bin = generar_reporte_debito_fiscal(df_res, soft_total, df_imp_raw)
                    
                    # --- RESULTADOS ---
                    st.success(f"✅ Auditoría finalizada. Se han identificado las discrepancias excluyendo registros de '{tag_obviar}'.")
                    
                    st.download_button(
                        label="⬇️ Descargar Reporte de Auditoría (Excel)",
                        data=excel_bin,
                        file_name=f"Auditoria_Fiscal_{casa_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # Mostramos el log por si hay advertencias de columnas
                    with st.expander("Ver Log de Proceso"):
                        for m in log:
                            st.text(m)

            except Exception as e:
                mostrar_error_amigable(e, "la Verificación de Débito Fiscal")

# Función para que el bot analice los resultados
def asistente_contable_inteligente(pregunta, df=None):
    p = pregunta.lower()
    
    # 1. VERIFICAR SI HAY DATOS
    if df is None or df.empty:
        if any(w in p for w in ["error", "total", "monto", "asiento"]):
            return "Aún no he analizado los datos. Por favor, carga los archivos y pulsa 'Iniciar Análisis' para poder responderte sobre este reporte."
        return "Hola! Soy tu asistente contable. ¿En qué puedo ayudarte hoy?"

    # 2. LÓGICA DE ERRORES (Detecta "error", "errores", "fallas", "mal")
    if any(w in p for w in ["error", "falla", "malo", "no coincide"]):
        # Buscamos en la columna de Monto OK o Banco OK
        errores_monto = len(df[df['Monto Coincide (CB vs CG)'].str.contains("❌")])
        errores_banco = len(df[df['Cuenta de Banco Correcta'].str.contains("❌")])
        
        if errores_monto == 0 and errores_banco == 0:
            return "¡Buenas noticias! No detecté errores de monto ni de cuentas bancarias en este reporte."
        
        msg = f"He encontrado {errores_monto + errores_banco} incidencias: \n"
        if errores_monto > 0: msg += f"- {errores_monto} diferencias de dinero entre Tesorería y Contabilidad.\n"
        if errores_banco > 0: msg += f"- {errores_banco} errores en la cuenta contable del banco utilizada.\n"
        msg += "Los verás resaltados en el Excel de descarga."
        return msg

    # 3. LÓGICA DE TOTALES (Detecta "total", "suma", "monto", "cuanto")
    if any(w in p for w in ["total", "suma", "monto", "cuánto", "cuanto"]):
        total_cb = df['Monto en Tesorería (CB)'].sum()
        total_cg = df['Monto en Contabilidad (CG)'].sum()
        
        if "dolar" in p or "usd" in p or "$" in p:
            # Solo sumamos los que el reporte marcó como USD
            usd_sum = df[df['Moneda'] == 'USD']['Monto en Tesorería (CB)'].sum()
            return f"El total de comisiones en Dólares (USD) es de ${usd_sum:,.2f}."
            
        return f"El total reportado por Tesorería es Bs. {total_cb:,.2f} y lo registrado en Contabilidad es Bs. {total_cg:,.2f}."

    # 4. CUENTAS CONTABLES
    if "cuenta" in p or "donde" in p:
        if "usd" in p or "dolar" in p:
            return "Para comisiones en USD (Exterior), la cuenta correcta es 7.1.3.50.1.002."
        return "Para comisiones en VES (País), la cuenta correcta es 7.1.3.50.1.001."

    # 5. CASO POR DEFECTO
    return "Entiendo tu pregunta, pero necesito que seas más específico. Puedes preguntarme por 'errores en el reporte', 'total de comisiones' o 'cuentas contables'."
    
def render_comisiones_bancarias():
    # --- IDENTIDAD VISUAL CORPORATIVA (Basada en la propuesta original) ---
    CONFIG_EMPRESAS = {
        "MAYOR BEVAL, C.A": {"borde": "#28A745", "fondo": "#EAFAF1", "tag": "BEVAL"},
        "FEBECA, C.A":      {"borde": "#2196F3", "fondo": "#E8F4FD", "tag": "FEBECA"},
        "PRISMA, C.A":      {"borde": "#566573", "fondo": "#F8F9F9", "tag": "PRISMA"},
        "FEBECA, C.A (QUINCALLA)": {"borde": "#FF00FF", "fondo": "#FDE9F9", "tag": "QUINCALLA"}
    }

    # Barra lateral: Asistente Virtual del Departamento
    with st.sidebar:
        st.title("🤖 Asistente de Comisiones")
        if "messages_com" not in st.session_state:
            st.session_state.messages_com = [{"role": "assistant", "content": "Bienvenido al módulo de Comisiones. Por favor, seleccione la empresa y cargue los archivos correspondientes."}]
        
        for m in st.session_state.messages_com:
            with st.chat_message(m["role"]): st.markdown(m["content"])
            
        if prompt := st.chat_input("Consulte sobre este proceso..."):
            st.session_state.messages_com.append({"role": "user", "content": prompt})
    
            # Le pasamos el dataframe de resultados si ya existe
            res_df = st.session_state.get('df_res_comisiones') 
            respuesta = asistente_contable_inteligente(prompt, res_df)
    
            st.session_state.messages_com.append({"role": "assistant", "content": respuesta})
            st.rerun()

    # Cabecera Institucional
    st.title("💰 Conciliación de Comisiones Bancarias")
    if st.button("⬅️ Volver al Panel Principal"): 
        set_page('inicio')
        st.rerun()

    # Selección de empresa y aplicación de estilos dinámicos
    st.subheader("Configuración de Proceso", anchor=False)
    casa_sel = st.selectbox("Empresa a procesar:", list(CONFIG_EMPRESAS.keys()), key="empresa_com")
    tema = CONFIG_EMPRESAS[casa_sel]

    # Inyección de Estilos (Mantiene los cuadros de colores del diseño original)
    st.markdown(f"""
        <style>
        .header-box {{ background-color: {tema['borde']}; color: white; padding: 12px; border-radius: 10px 10px 0 0; font-weight: bold; text-align: center; font-size: 1.1rem; }}
        [data-testid="stFileUploader"] {{ background-color: {tema['fondo']} !important; border: 2px solid {tema['borde']} !important; border-radius: 0 0 15px 15px !important; }}
        div.stButton > button {{ background-color: {tema['borde']} !important; color: white !important; border-radius: 12px; height: 3.5em; font-weight: bold; width: 100%; border: none; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        </style>
        """, unsafe_allow_html=True)

    st.divider()

    # Área de Carga de Archivos
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f'<div class="header-box">Comisiones Tesorería (CB)</div>', unsafe_allow_html=True)
        f_cb = st.file_uploader("Subir Reporte de Bancos", type=['xlsx'], key="com_cb", label_visibility="collapsed")
    with col2:
        st.markdown(f'<div class="header-box">Transacciones Diario (CG)</div>', unsafe_allow_html=True)
        f_cg = st.file_uploader("Subir Diario Contable", type=['xlsx'], key="com_cg", label_visibility="collapsed")

    # Acción de Procesamiento
    if f_cb and f_cg:
        if st.button(f"⚡ Iniciar Análisis de Comisiones - {tema['tag']}"):
            log = []
            try:
                with st.spinner("Cruzando Tesorería vs Contabilidad..."):
                    # Cargamos el CB crudo (header=None) para la réplica perfecta de la hoja de consulta
                    df_cb_replica = pd.read_excel(f_cb, header=None)
                    df_cg_replica = pd.read_excel(f_cg) # El mayor suele ser estándar

                    # Llamamos a la lógica (que usa sus propios procesos de limpieza internos)
                    df_res = run_conciliation_comisiones_bancarias(
                        pd.read_excel(f_cb), 
                        pd.read_excel(f_cg), 
                        casa_sel, # <--- Pasamos la empresa seleccionada
                        log
                    )
                    st.session_state['df_res_comisiones'] = df_res 
                    st.success(f"✅ Proceso completado exitosamente para {casa_sel}")
                    st.dataframe(df_res, use_container_width=True)
                    
                    # Reporte Excel personalizado con el color de la empresa
                    excel_bin = generar_reporte_auditoria_comisiones(
                        df_res,      # El resultado de la auditoría
                        df_cg_replica,  # El mayor original (subido por el usuario)
                        df_cb_replica,   # El reporte de tesorería original (subido por el usuario)
                        casa_sel, 
                        tema['borde']
                    )
                    st.download_button(
                        label=f"📥 Descargar Reporte Final ({tema['tag']})", 
                        data=excel_bin, 
                        file_name=f"Conciliacion_Comisiones_{tema['tag']}.xlsx",
                        use_container_width=True
                    )
            except Exception as e:
                mostrar_error_amigable(e, "el módulo de Comisiones")

def render_locti():
    # --- 1. CABECERA E IDENTIDAD VISUAL ---
    st.title("⚖️ Cálculo LOCTI")
    
    if st.button("⬅️ Volver al Inicio"): 
        set_page('inicio')
        st.rerun()

    # --- IDENTIDAD VISUAL DE LUSI (Réplica Exacta) ---
    st.markdown("""
        <style>
        /* Estilo para los encabezados de color */
        .titulo-reporte {
            padding: 12px; 
            border-radius: 10px 10px 0px 0px; /* Redondeado solo arriba */
            color: white; 
            font-weight: bold;
            text-align: center; 
            font-size: 1.1em;
            box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
        }
        .ventas { background-color: #1a237e; } /* Azul Marino */
        .ingresos { background-color: #1b5e20; } /* Verde Oscuro */
        .reserva { background-color: #b71c1c; } /* Rojo Intenso */
        
        /* Ajuste del cargador de archivos para que se pegue al título */
        [data-testid="stFileUploader"] {
            border: 1px solid #ddd !important;
            border-top: none !important; /* Quita el borde superior para unirlo al título */
            border-radius: 0px 0px 10px 10px !important; /* Redondeado solo abajo */
            padding: 10px;
            background-color: #f8f9fa;
        }
        
        /* Quitar el padding extra que Streamlit pone entre elementos */
        .stMarkdown { margin-bottom: -15px; }
        
        /* Estilo para los mensajes con el pin */
        .pin-guia {
            font-size: 0.85rem;
            color: #666;
            margin-top: 10px;
            margin-left: 5px;
        }
        </style>
    """, unsafe_allow_html=True)

    # --- 2. BARRA DE PARÁMETROS (Configuración de Cierre) ---
    dict_filiales = {
        "BEVAL, C.A.": "271", "FEBECA, C.A.": "004",
        "SILLACA, C.A.": "071", "PRISMA SISTEMAS": "298"
    }
    
    with st.container(border=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        filial = c1.selectbox("🏢 Seleccione la Filial:", list(dict_filiales.keys()))
        fecha_rep = c2.date_input("📅 Mes de Cierre:", value=datetime(2026, 1, 31))
        tasa = c3.number_input("💵 Tasa de Cambio:", min_value=0.01, value=1.0, format="%.4f")
        num_asto = c4.text_input("🔢 N° Asiento:", value="CG0000")
        analista = c3.text_input("👤 Hecho por:", value=" ").upper()

    st.markdown("---")
    # --- 3. ÁREA DE CARGA DE ARCHIVOS (Diseño Lusi) ---
    st.markdown(f"### 📥 Carga de Balances de Comprobación: **{filial}**")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="titulo-reporte ventas">🏢 Balance VENTAS</div>', unsafe_allow_html=True)
        f_v = st.file_uploader("Subir Ventas", type=["xlsx", "xls"], key="lv", label_visibility="collapsed")
        st.markdown('<p class="pin-guia">📍 Debe contener la cuenta 4.0.0.00.0.000</p>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="titulo-reporte ingresos">💰 Balance INGRESOS</div>', unsafe_allow_html=True)
        f_i = st.file_uploader("Subir Ingresos", type=["xlsx", "xls"], key="li", label_visibility="collapsed")
        st.markdown('<p class="pin-guia">📍 Debe contener los grupos 6.1.1.X</p>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="titulo-reporte reserva">📥 RESERVA ANTERIOR</div>', unsafe_allow_html=True)
        f_r = st.file_uploader("Subir Reserva", type=["xlsx", "xls"], key="lr", label_visibility="collapsed")
        st.markdown('<p class="pin-guia">📍 Debe contener la cuenta 7.1.3.57.1.001</p>', unsafe_allow_html=True)

    # --- 4. EJECUCIÓN Y RESULTADOS ---
    if f_v and f_i and f_r:
        if st.button("🚀 Iniciar Cálculo y Generar Cargador", type="primary", use_container_width=True):
            log = []
            try:
                with st.spinner("Procesando balances..."):
                    # Llamada a logic.py (Motor de Lusi)
                    res, df_asiento = procesar_calculo_locti(f_v, f_i, f_r, tasa, num_asto, log)
                
                if res:
                    st.success(f"✅ Cálculo finalizado para {filial}")
                    
                    # Métricas de Resumen
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Base Gravable", f"Bs. {res['base_mes']:,.2f}")
                    m2.metric("Aporte Mes (0.5%)", f"Bs. {res['aporte_mes']:,.2f}")
                    m3.metric("Saldo Acumulado", f"Bs. {res['acum_directo']:,.2f}")
                    m4.metric("Diferencia", f"Bs. {res['diferencia']:,.2f}", delta_color="inverse")

                    # Cuadro de Conciliación Visual (Diseño Lusi)
                    cuadra = res['diferencia'] < 1.0
                    color_bg = '#d4edda' if cuadra else '#f8d7da'
                    st.markdown(f"""
                        <div style="background-color: {color_bg}; padding: 20px; border-radius: 10px; border: 1px solid #ccc; text-align: center; color: black;">
                            <h4>{'✅ CONCILIACIÓN EXITOSA' if cuadra else '❌ DESCUADRE DETECTADO'}</h4>
                            <p>Anterior ({res['res_ant']:,.2f}) + Mes ({res['aporte_mes']:,.2f}) = <b>{res['proyectado']:,.2f} Bs.</b></p>
                        </div>
                    """, unsafe_allow_html=True)

                    # --- 5. GENERACIÓN DE DESCARGAS ---
                    st.divider()
                    st.subheader("📥 Descarga de Archivos")
                    d_col1, d_col2 = st.columns(2)

                    # A. Reporte Excel de Lusi
                    meta_data = {
                        "filial": filial, "usuario": analista,
                        "fecha_str": fecha_rep.strftime("%d/%m/%Y"),
                        "mes_nombre": fecha_rep.strftime("%B %Y").upper(),
                        "mes_corto": fecha_rep.strftime("%b.%y").upper(), "num_casa": dict_filiales[filial]
                    }
                    excel_rep = generar_reporte_excel_locti(res, df_asiento, meta_data)
                    d_col1.download_button(
                        "📊 Descargar Informe LOCTI", 
                        excel_rep, 
                        f"Reporte_LOCTI_{filial.split(',')[0]}_{meta_data['mes_corto']}.xlsx",
                        use_container_width=True
                    )

                    # B. Cargador Softland (Usando la FUNCIÓN UNIVERSAL de utils.py)
                    cargador_bin = generar_cargador_softland_v2(df_asiento, fecha_rep)
                    d_col2.download_button(
                        "📥 Descargar Cargador para Sistema", 
                        cargador_bin, 
                        f"CARGADOR_LOCTI_{num_asto}.xlsx",
                        use_container_width=True
                    )

                    with st.expander("Ver Log de Extracción"):
                        for m in log: st.text(m)

            except Exception as e:
                mostrar_error_amigable(e, "el proceso LOCTI")
    else:
        st.info("💡 Por favor, cargue los tres balances de comprobación para habilitar el cálculo.")





def render_apartados_liberaciones():
    # --- ENCABEZADO Y BOTÓN VOLVER ---
    col_t1, col_t2 = st.columns([8, 2])
    with col_t1:
        st.title("⚖️ Gestión de Apartados y Liberaciones", anchor=False)
    with col_t2:
        if st.button("⬅️ Volver al Inicio", key="back_apt_lib"):
            set_page('inicio')
            st.rerun()
        

    # --- 1. ESTILOS CSS (DARK PREMIUM - ESTILO LOCTI) ---
    st.markdown("""
        <style>
        .stContainer {
            background-color: #1a1c24;
            border-radius: 15px;
            padding: 20px;
            border: 1px solid #333;
        }
        .card-header {
            padding: 10px;
            border-radius: 10px 10px 0 0;
            font-weight: bold;
            text-align: center;
            color: white;
            margin-bottom: -10px;
        }
        .purple-header { background-color: #4b0082; }
        .green-header { background-color: #1e5631; }
        
        div[data-testid="stFileUploader"] {
            background-color: #1a1c24 !important;
            border: 1px solid #333 !important;
            border-radius: 0 0 15px 15px !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # --- 2. PANEL DE PARÁMETROS SUPERIOR ---
    with st.container(border=True):
        # Fila 1: Filial y Analista
        r1_c1, r1_c2 = st.columns(2)
        with r1_c1:
            empresa_sel = st.selectbox("🏢 Seleccione la Filial:", ["FEBECA", "BEVAL", "PRISMA", "QUINCALLA"])
        with r1_c2:
            analista_nombre = st.text_input("👤 Hecho por:", placeholder="Nombre del analista que procesa")

        # Fila 2: Tasas de Cambio
        r2_c1, r2_c2 = st.columns(2)
        with r2_c1:
            tasa_bcv = st.number_input("💵 Tasa BCV (Gastos en USD):", min_value=0.01, value=1.0, format="%.4f")
        with r2_c2:
            tasa_corp = st.number_input("📈 Tasa Corp. (Gastos en Bs):", min_value=0.01, value=1.0, format="%.4f")

        # Fila 3: Códigos de Asiento (Prefijo CG solicitado)
        r3_c1, r3_c2 = st.columns(2)
        with r3_c1:
            num_asiento_lib = st.text_input("🔢 N° Asiento Liberaciones:", value="CG000")
        with r3_c2:
            num_asiento_apt = st.text_input("🔢 N° Asiento Apartados:", value="CG000")

    st.markdown("<br>", unsafe_allow_html=True)

    # --- 3. CARGA DE ARCHIVOS (MAESTRO Y BALANCE) ---
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.markdown('<div class="card-header purple-header">📂 Maestro Anterior</div>', unsafe_allow_html=True)
        with st.container(border=True):
            f_maestro = st.file_uploader("Subir Maestro", type=['xlsx', 'xls'], key="up_maestro", label_visibility="collapsed")
        st.caption("📍 El archivo con las pestañas MES.XX y el Histórico.")

    with col_f2:
        st.markdown('<div class="card-header green-header">📄 Balance Analítico</div>', unsafe_allow_html=True)
        with st.container(border=True):
            f_balance = st.file_uploader("Subir Balance", type=['xlsx', 'xls'], key="up_balance", label_visibility="collapsed")
        st.caption("📍 Reporte de Softland con las cuentas 7 detalladas.")

    # --- 4. LÓGICA DE PROCESAMIENTO ---
    if f_maestro and f_balance:
        try:
            # A. Identificación de pestañas del Maestro
            xls_maestro = pd.ExcelFile(f_maestro)
            hojas_disponibles = xls_maestro.sheet_names
            
            # Radar flexible para buscar pestañas con un punto (ENE.26, etc)
            hojas_mes = [h.strip() for h in hojas_disponibles if "." in h]
            if not hojas_mes: hojas_mes = hojas_disponibles

            with st.container(border=True):
                st.write("📂 **Configuración de Pestañas Detectadas**")
                c_m1, c_m2 = st.columns(2)
                with c_m1:
                    mes_anterior_sel = st.selectbox("Seleccione la Portada del mes pasado:", hojas_mes)
                with c_m2:
                    hoja_hist_name = next((h for h in hojas_disponibles if "APARTADO" in h.upper()), "APARTADO VALENCIA")
                    st.info(f"Hoja histórica detectada: `{hoja_hist_name}`")

            # B. Carga de datos con detección de encabezado automática
            nombre_real_hoja = hojas_disponibles[hojas_mes.index(mes_anterior_sel)]
            df_temp_m = pd.read_excel(f_maestro, sheet_name=nombre_real_hoja, header=None)
            h_row_m = 0
            for i in range(min(20, len(df_temp_m))):
                fila_titles = [str(x).upper() for x in df_temp_m.iloc[i].values]
                if "CTA" in fila_titles or "CUENTA" in fila_titles:
                    h_row_m = i
                    break
            
            df_m = pd.read_excel(f_maestro, sheet_name=nombre_real_hoja, header=h_row_m)
            df_b_raw = pd.read_excel(f_balance, header=None)

            # C. Ejecución de la IA (Logic)
            with st.spinner("La IA está analizando el balance para sugerir liberaciones..."):
                
                df_movs_real = parsear_balance_softland(df_b_raw)
                df_propuesta = conciliar_ciclo_apartados(df_m, df_movs_real)

            # --- VISTA: CARRITO DE LIBERACIONES ---
            st.divider()
            st.subheader("🛒 Fase 1: Selección de Liberaciones", anchor=False)
            st.write("Marque las partidas que ya llegaron en el balance para generar el reverso:")
            
            df_editado = st.data_editor(
                df_propuesta,
                column_config={
                    "Liberar": st.column_config.CheckboxColumn("¿LIBERAR?", default=False),
                    "Estado": st.column_config.TextColumn("Sugerencia IA"),
                    "Monto_Original_BS": st.column_config.NumberColumn("Apartado (Bs)", format="%.2f"),
                    "Monto_Real_Encontrado": st.column_config.NumberColumn("Encontrado (Bs)", format="%.2f"),
                },
                disabled=["Cuenta", "CC", "Descripcion", "Monto_Original_BS", "Monto_Real_Encontrado", "Estado"],
                hide_index=True, use_container_width=True, key="editor_liberaciones"
            )

            # --- VISTA: NUEVOS APARTADOS ---
            st.divider()
            st.subheader("➕ Fase 2: Nuevos Apartados del Mes", anchor=False)
            if 'lista_nuevos' not in st.session_state: st.session_state.lista_nuevos = []

            with st.expander("Abrir Formulario de Registro"):
                with st.form("nuevo_gasto_form"):
                    c_n1, c_n2 = st.columns(2)
                    n_cta = c_n1.text_input("Cuenta Contable (7.x.x)")
                    n_cc = c_n2.text_input("Centro de Costo (xx.xx.xxx)")
                    n_desc = st.text_input("Descripción y Periodo (Ej: MOVISTAR FEB.26)")
                    
                    c_n3, c_n4 = st.columns(2)
                    n_monto = c_n3.number_input("Monto:", min_value=0.0)
                    n_moneda = c_n4.radio("Moneda:", ["BS", "USD"], horizontal=True)
                    
                    if st.form_submit_button("Añadir a la lista"):
                        st.session_state.lista_nuevos.append({
                            'Cuenta': n_cta, 'CC': n_cc, 'Descripcion': n_desc,
                            'Monto_USD': n_monto if n_moneda == 'USD' else 0,
                            'Monto_BS': n_monto if n_moneda == 'BS' else 0,
                            'Moneda': n_moneda, 'Tasa_Original': tasa_bcv if n_moneda == 'USD' else tasa_corp
                        })
                        st.rerun()

            if st.session_state.lista_nuevos:
                st.write("**Partidas nuevas para apartar:**")
                st.dataframe(pd.DataFrame(st.session_state.lista_nuevos), use_container_width=True)
                if st.button("🗑️ Limpiar lista de nuevos"):
                    st.session_state.lista_nuevos = []
                    st.rerun()

            # --- BOTÓN DE CIERRE Y GENERACIÓN ---
            st.divider()
            if st.button("🚀 PROCESAR CIERRE: GENERAR TODO", type="primary", use_container_width=True):
                fecha_cierre_hoy = pd.Timestamp.now()
                st.success("✅ Documentos generados satisfactoriamente.")
                
                col_d1, col_d2 = st.columns(2)
                
                with col_d1:
                    st.write("**📊 Soporte de Gestión**")
                    # 1. Reporte Verde (Liberaciones)
                    bin_lib = generar_reporte_visual_liberaciones(df_editado, empresa_sel, fecha_cierre_hoy, analista_nombre)
                    st.download_button("📊 Descargar Soporte Liberaciones (Verde)", bin_lib, f"SOPORTE_LIB_{empresa_sel}.xlsx", use_container_width=True)
                    
                    # 2. Maestro Actualizado (Pestaña nueva + Histórico)
                    df_nuevos_final = pd.DataFrame(st.session_state.lista_nuevos)
                    df_maestro_proximo = pd.concat([df_editado[df_editado['Liberar']==False], df_nuevos_final], ignore_index=True)
                    # El nombre de la hoja será el mes actual (ej: MAR.26)
                    mes_actual_nombre = f"{fecha_cierre_hoy.strftime('%b').upper()}.{fecha_cierre_hoy.strftime('%y')}"
                    bin_maestro = generar_reporte_maestro_apartados(xls_maestro, df_maestro_proximo, mes_actual_nombre, hoja_hist_name, empresa_sel, fecha_cierre_hoy)
                    st.download_button("📂 Descargar Nuevo Maestro (Amarillo)", bin_maestro, f"MAESTRO_NUEVO_{empresa_sel}.xlsx", use_container_width=True)

                with col_d2:
                    st.write("**📥 Cargadores para Softland**")
                    # 3. Cargador REVERSOS
                    df_reversar = df_editado[df_editado['Liberar'] == True]
                    if not df_reversar.empty:
                        as_rev = preparar_asiento_softland(df_reversar, "REVERSO", tasa_bcv, num_asiento_lib)
                        bin_as_rev = generar_excel_cargador_softland(as_rev, fecha_cierre_hoy)
                        st.download_button("📥 Descargar Cargador REVERSOS", bin_as_rev, f"CARGADOR_REV_{num_asiento_lib}.xlsx", use_container_width=True)
                    
                    # 4. Cargador APARTADOS
                    if not df_nuevos_final.empty:
                        as_new = preparar_asiento_softland(df_nuevos_final, "NUEVO", tasa_bcv, num_asiento_apt)
                        bin_as_new = generar_excel_cargador_softland(as_new, fecha_cierre_hoy)
                        st.download_button("📥 Descargar Cargador APARTADOS", bin_as_new, f"CARGADOR_APT_{num_asiento_apt}.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"❌ Error en la lectura del archivo: {str(e)}")
            st.info("💡 Asegúrese de subir el Maestro correcto con su estructura de portadas.")
    else:
        st.info("💡 Por favor, cargue el Maestro del mes pasado y el Balance del mes actual para habilitar el análisis.")


# ==============================================================================
# VI. ENRUTAMIENTO FINAL (ROUTER)
# ==============================================================================
def main():
    page_map = {
        'inicio': render_inicio,
        'especificaciones': render_especificaciones,
        'especificaciones_cofersa': render_especificaciones_cofersa,
        'retenciones': render_retenciones,
        'paquete_cc': render_paquete_cc, 
        'cuadre': render_cuadre,
        'pensiones': render_pensiones,
        'ajustes_usd' : render_ajustes_usd,
        'comisiones': render_comisiones_bancarias,
        'debito_fiscal': render_debito_fiscal,
        'locti': render_locti,
        'apartados': render_apartados_liberaciones
    }
    
    current_page = st.session_state.get('page', 'inicio')
    render_function = page_map.get(current_page, render_inicio)
    render_function()

if __name__ == "__main__":
    main()
