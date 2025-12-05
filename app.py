
# app.py

# ==============================================================================
# 1. IMPORTACIÓN DE LIBRERÍAS Y CONFIGURACIÓN INICIAL
# ==============================================================================
import streamlit as st
import pandas as pd
import traceback
from guides import (
    GUIA_GENERAL_ESPECIFICACIONES, 
    LOGICA_POR_CUENTA, 
    GUIA_COMPLETA_RETENCIONES,
    GUIA_PAQUETE_CC
)

from functools import partial

from logic import (
    run_conciliation_fondos_en_transito,
    run_conciliation_fondos_por_depositar,
    run_conciliation_devoluciones_proveedores,
    run_conciliation_viajes,
    run_conciliation_retenciones,
    run_analysis_paquete_cc,
    run_conciliation_cobros_viajeros,
    run_conciliation_otras_cxp,
    run_conciliation_haberes_clientes,
    run_conciliation_cdc_factoring,
    run_conciliation_asientos_por_clasificar,
    run_conciliation_deudores_empleados_me,
    run_cuadre_cb_cg_beval
)
from utils import (
    cargar_y_limpiar_datos,
    generar_reporte_excel,
    generar_excel_saldos_abiertos,
    generar_reporte_paquete_cc,
    generar_reporte_cuadre
)


def mostrar_error_amigable(e, contexto=""):
    """
    Traduce errores técnicos de Python a mensajes amigables para el usuario contable.
    """
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

# ==============================================================================
# BLOQUE DE AUTENTICACIÓN
# ==============================================================================
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
            st.text_input(
                "Contraseña", type="password", on_change=password_entered, key="password", label_visibility="collapsed", placeholder="Ingresa la contraseña"
            )
            
            if st.session_state.get("authentication_attempted", False):
                if not st.session_state.get("password_correct", False):
                    st.error("😕 Contraseña incorrecta.")
            else:
                st.info("Por favor, ingresa la contraseña para continuar.")

        st.divider()

        st.markdown("<p style='text-align: center;'>Una herramienta para las empresas del grupo:</p>", unsafe_allow_html=True)
        
        logo_cols = st.columns(3)
        logos_info = [
            {"path": "assets/logo_febeca.png", "fallback": "FEBECA, C.A."},
            {"path": "assets/logo_beval.png", "fallback": "MAYOR BEVAL, C.A."},
            {"path": "assets/logo_sillaca.png", "fallback": "SILLACA, C.A."}
        ]
        
        for i, col in enumerate(logo_cols):
            with col:
                try:
                    st.image(logos_info[i]["path"], use_container_width=True)
                except:
                    st.markdown(f"<p style='text-align: center; font-size: small;'>{logos_info[i]['fallback']}</p>", unsafe_allow_html=True)

    st.stop()

# ==============================================================================
# DICCIONARIO CENTRAL DE ESTRATEGIAS (EL "CEREBRO")
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
        "columnas_reporte": ['NIT', 'Descripcion NIT', 'Fecha', 'Referencia', 'Numero_Envio', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1019",
        "columnas_requeridas": ['Asiento', 'Fuente', 'Fecha', 'Referencia', 'Nit', 'Descripcion NIT', 'Debito Bolivar', 'Credito Bolivar']
    },
    "212.05.1108 - Haberes de Clientes": {
        "id": "haberes_clientes",
        "funcion_principal": run_conciliation_haberes_clientes,
        "label_actual": "Movimientos del mes (Haberes Clientes)",
        "label_anterior": "Saldos anteriores (Haberes Clientes)",
        "columnas_reporte": ['NIT', 'Descripción Nit', 'Fecha', 'Fuente', 'Referencia', 'Monto Bolivar'],
        "nombre_hoja_excel": "212.05.1108",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nit', 'Descripción Nit', 'Débito Bolivar', 'Crédito Bolivar']
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
    }
    
}

# ==============================================================================
# RENDERIZADO DE VISTAS (PÁGINAS)
# ==============================================================================
def set_page(page_name):
    st.session_state.page = page_name

def render_inicio():
    st.title("🤖 Portal de Herramientas Contables")
    st.markdown("Seleccione la herramienta que desea utilizar:")
    col1, col2 = st.columns(2)
    with col1:
        st.button("📄 Especificaciones", on_click=set_page, args=['especificaciones'], use_container_width=True)
        st.button("📦 Análisis de Paquete CC", on_click=set_page, args=['paquete_cc'], use_container_width=True)
        st.button("💵 Reservas y Apartados", on_click=set_page, args=['reservas'], use_container_width=True, disabled=True)
        
    with col2:
        st.button("⚖️ Cuadre CB - CG", on_click=set_page, args=['cuadre'], use_container_width=True)
        st.button("🧾 Relación de Retenciones", on_click=set_page, args=['retenciones'], use_container_width=True)
        st.button("🔜 Próximamente", on_click=set_page, args=['proximamente'], use_container_width=True, disabled=True)    

def render_proximamente(titulo):
    st.title(f"🛠️ {titulo}")
    st.info("Esta funcionalidad estará disponible en futuras versiones.")
    st.button("⬅️ Volver al Inicio", on_click=set_page, args=['inicio'])

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
                
                # --- TRY / EXCEPT QUE ACABAMOS DE HACER ---
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

def render_especificaciones():
    st.title('📄 Herramienta de Conciliación de Cuentas', anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_spec"):
        set_page('inicio')
        st.session_state.processing_complete = False 
        st.rerun()
    st.markdown("Esta aplicación automatiza el proceso de conciliación de cuentas contables.")
    
    CASA_OPTIONS = ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"]
    CUENTA_OPTIONS = sorted(list(ESTRATEGIAS.keys()))
    
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
                    progress_container.progress(0, text="Iniciando fases de conciliación. Esto puede tardar unos momentos...")
                    df_resultado = estrategia_actual["funcion_principal"](df_full.copy(), log_messages, progress_bar=progress_container)
                    progress_container.progress(1.0, text="¡Proceso completado!")
                    st.session_state.df_saldos_abiertos = df_resultado[~df_resultado['Conciliado']].copy()
                    st.session_state.df_conciliados = df_resultado[df_resultado['Conciliado']].copy()
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
            st.download_button("⬇️ Descargar Reporte Completo (Excel)", st.session_state.excel_output, f"reporte_conciliacion_{estrategia_actual['id']}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_excel")
        with res_col2:
            st.metric("Saldos Abiertos (Pendientes)", len(st.session_state.df_saldos_abiertos))
            st.download_button(
                label="⬇️ Descargar Saldos para Próximo Mes (Excel)",
                data=st.session_state.excel_saldos_output,
                file_name="saldos_para_proximo_mes.xlsx",
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
    # ------------------------------------------
    
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

def render_cuadre():
    st.title("⚖️ Cuadre de Disponibilidad (CB vs CG)", anchor=False)
    
    # --- SELECTOR DE EMPRESA ---
    CASA_OPTIONS = ["MAYOR BEVAL, C.A", "FEBECA, C.A", "FEBECA, C.A (QUINCALLA)", "PRISMA, C.A"]
    col_emp, _ = st.columns([1, 1])
    with col_emp:
        empresa_sel = st.selectbox("Seleccione la Empresa:", CASA_OPTIONS, key="empresa_cuadre")
    # ---------------------------

    st.info("Sube el Reporte de Tesorería (CB) y el Balance de Comprobación (CG). Pueden ser PDF o Excel.")
    
    col1, col2 = st.columns(2)
    with col1:
        file_cb = st.file_uploader("1. Reporte Tesorería (CB)", type=['pdf', 'xlsx'])
    with col2:
        file_cg = st.file_uploader("2. Balance Contable (CG)", type=['pdf', 'xlsx'])
        
    if file_cb and file_cg:
        if st.button("Comparar Saldos"):
            log = []
            try:
                # --- CAMBIO: Se pasa 'empresa_sel' a la función ---
                # Nota: Asegúrate de importar 'run_cuadre_cb_cg' (el nuevo nombre) en los imports de arriba
                # o renómbralo en logic.py si prefieres mantener el nombre viejo.
                # Aquí asumo que en logic.py ahora se llama 'run_cuadre_cb_cg' como puse arriba.
                from logic import run_cuadre_cb_cg 
                
                df_res = run_cuadre_cb_cg(file_cb, file_cg, empresa_sel, log)
                # --------------------------------------------------
                
                cols_pantalla = ['Moneda', 'Banco (Tesorería)', 'Cuenta Contable', 'Descripción', 'Saldo Final CB', 'Saldo Final CG', 'Diferencia', 'Estado']
                st.dataframe(df_res[cols_pantalla], use_container_width=True)
                
                excel_data = generar_reporte_cuadre(df_res)
                st.download_button("⬇️ Descargar Reporte Completo (Excel)", excel_data, "Cuadre_CB_CG.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                with st.expander("Ver Log de Extracción"):
                    st.write(log)
                    
            except Exception as e:
                mostrar_error_amigable(e, "el Cuadre CB-CG")
                

# ==============================================================================
# FLUJO PRINCIPAL DE LA APLICACIÓN (ROUTER)
# ==============================================================================
def main():
    page_map = {
        'inicio': render_inicio,
        'especificaciones': render_especificaciones,
        'retenciones': render_retenciones,
        'paquete_cc': render_paquete_cc, 
        'cuadre': render_cuadre,
        'reservas': lambda: render_proximamente("Reservas y Apartados"),
        'proximamente': lambda: render_proximamente("Próximamente")
    }
    
    current_page = st.session_state.get('page', 'inicio')
    render_function = page_map.get(current_page, render_inicio)
    render_function()

if __name__ == "__main__":
    main()
