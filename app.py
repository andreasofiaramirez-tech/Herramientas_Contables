
# app.py

# ==============================================================================
# 1. IMPORTACIÓN DE LIBRERÍAS Y CONFIGURACIÓN INICIAL
# ==============================================================================
import streamlit as st
import pandas as pd
import traceback
from functools import partial

# --- BLOQUE 1: IMPORTAR GUÍAS (Verifica las comas) ---
from guides import (
    GUIA_GENERAL_ESPECIFICACIONES, 
    LOGICA_POR_CUENTA, 
    GUIA_COMPLETA_RETENCIONES,
    GUIA_PAQUETE_CC,
    GUIA_IMPRENTA,
    GUIA_GENERADOR,
    GUIA_PENSIONES
)

# --- BLOQUE 2: IMPORTAR LÓGICA (Verifica las comas) ---
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
    run_analysis_paquete_cc,
    run_cuadre_cb_cg,
    validar_coincidencia_empresa,
    run_cross_check_imprenta,
    generar_txt_retenciones_galac,
    procesar_calculo_pensiones
)

# --- BLOQUE 3: IMPORTAR UTILS ---
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
            
            # Campo de texto (Se activa con Enter)
            st.text_input(
                "Contraseña", 
                type="password", 
                on_change=password_entered, 
                key="password", 
                label_visibility="collapsed", 
                placeholder="Ingresa la contraseña"
            )
            
            # --- NUEVO BOTÓN DE INGRESAR ---
            # Se activa con Clic. Llama a la misma función de validación.
            st.button("Ingresar", on_click=password_entered, type="primary", use_container_width=True)
            # -------------------------------
            
            if st.session_state.get("authentication_attempted", False):
                if not st.session_state.get("password_correct", False):
                    st.error("😕 Contraseña incorrecta.")
            else:
                # Pequeño espacio visual
                st.markdown("") 
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
        st.button("🪪 Calculo Pensiones", on_click=set_page, args=['pensiones'], use_container_width=True)
        st.button("💵 Reservas y Apartados", on_click=set_page, args=['reservas'], use_container_width=True, disabled=True)
        
    with col2:
        st.button("⚖️ Cuadre CB - CG", on_click=set_page, args=['cuadre'], use_container_width=True)
        st.button("🧾 Relación de Retenciones", on_click=set_page, args=['retenciones'], use_container_width=True)
        st.button("🖨️ Cruce Imprenta", on_click=set_page, args=['imprenta'], use_container_width=True)
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
                    # ... (Lógica de conciliación existente) ...
                    progress_container.progress(0, text="Iniciando fases de conciliación...")
                    df_resultado = estrategia_actual["funcion_principal"](df_full.copy(), log_messages, progress_bar=progress_container)
                    progress_container.progress(1.0, text="¡Proceso completado!")
                    
                    st.session_state.df_saldos_abiertos = df_resultado[~df_resultado['Conciliado']].copy()
                    st.session_state.df_conciliados = df_resultado[df_resultado['Conciliado']].copy()
                    
                    # --- NUEVO: GENERACIÓN DE NOMBRE DE ARCHIVO ---
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
            
            # --- USAMOS EL NOMBRE DINÁMICO AQUÍ ---
            nombre_descarga = st.session_state.get('nombre_archivo_salida', 'reporte_conciliacion.xlsx')
            
            st.download_button(
                "⬇️ Descargar Reporte Completo (Excel)", 
                st.session_state.excel_output, 
                file_name=nombre_descarga,  # <--- CAMBIO
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True, 
                key="download_excel"
            )
            # --------------------------------------

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
    
    # --- BOTÓN VOLVER AL INICIO ---
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
        if st.button("Comparar Saldos", type="primary", use_container_width=True):
            log = []
            try:
                # Importamos funciones necesarias (incluyendo la nueva validación)
                from logic import run_cuadre_cb_cg, validar_coincidencia_empresa
                from utils import generar_reporte_cuadre
                
                # --- FASE 0: VALIDACIÓN DE SEGURIDAD ---
                # 1. Verificar archivo Tesorería
                es_valido_cb, msg_cb = validar_coincidencia_empresa(file_cb, empresa_sel)
                if not es_valido_cb:
                    st.error(f"⛔ ALERTA DE SEGURIDAD (Tesorería): {msg_cb}")
                    st.warning("Por favor verifique que seleccionó la empresa correcta en el menú.")
                    st.stop() # Detiene la ejecución aquí para proteger los datos
                
                # 2. Verificar archivo Contabilidad
                es_valido_cg, msg_cg = validar_coincidencia_empresa(file_cg, empresa_sel)
                if not es_valido_cg:
                    st.error(f"⛔ ALERTA DE SEGURIDAD (Contabilidad): {msg_cg}")
                    st.warning("Por favor verifique que seleccionó la empresa correcta en el menú.")
                    st.stop() # Detiene la ejecución aquí
                # ---------------------------------------

                # --- FASE 1: PROCESAMIENTO ---
                with st.spinner("Analizando y cruzando saldos..."):
                    df_res, df_huerfanos = run_cuadre_cb_cg(file_cb, file_cg, empresa_sel, log)
                
                # --- FASE 2: MOSTRAR RESULTADOS EN PANTALLA ---
                st.subheader("Resumen de Saldos", anchor=False)
                
                # Mostramos solo columnas clave para no saturar la vista
                cols_pantalla = ['Moneda', 'Banco (Tesorería)', 'Cuenta Contable', 'Descripción', 'Saldo Final CB', 'Saldo Final CG', 'Diferencia', 'Estado']
                st.dataframe(df_res[cols_pantalla], use_container_width=True)
                
                # Si hay cuentas huérfanas (no configuradas), mostramos alerta
                if not df_huerfanos.empty:
                    st.error(f"⚠️ ATENCIÓN: Se detectaron {len(df_huerfanos)} cuentas con saldo que NO están configuradas. Revisa la 3ra pestaña del Excel.")
                    st.dataframe(df_huerfanos, use_container_width=True)
                
                # --- FASE 3: GENERAR EXCEL ---
                excel_data = generar_reporte_cuadre(df_res, df_huerfanos, empresa_sel)
                
                st.download_button(
                    label="⬇️ Descargar Reporte Completo (Excel)",
                    data=excel_data,
                    file_name=f"Cuadre_CB_CG_{empresa_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Log técnico al final
                with st.expander("Ver Log de Extracción"):
                    st.write(log)
                    
            except Exception as e:
                mostrar_error_amigable(e, "el Cuadre CB-CG")
                
def render_imprenta():
    st.title("🖨️ Gestión de Imprenta (Retenciones de IVA)", anchor=False)
    
    if st.button("⬅️ Volver al Inicio", key="back_from_imprenta"):
        set_page('inicio')
        st.rerun()

    tab1, tab2 = st.tabs(["✅ 1. Validar TXTs Existentes", "⚙️ 2. Generar TXT desde Softland"])
    
    # --- PESTAÑA 1: VALIDACIÓN ---
    with tab1:
        st.info("Valida integridad entre el Libro de Ventas y el archivo TXT de Retenciones ya generado.")
        c1, c2 = st.columns(2)
        with c1: f_sales = st.file_uploader("1. Libro de Ventas (.txt)", type=['txt'], key="v_sales")
        with c2: f_ret = st.file_uploader("2. Archivo Retenciones (.txt)", type=['txt'], key="v_ret")
            
        if f_sales and f_ret:
            if st.button("Validar Archivos", type="primary", key="btn_val"):
                log = []
                try:
                    from logic import run_cross_check_imprenta
                    from utils import generar_reporte_imprenta
                    
                    df, txt = run_cross_check_imprenta(f_sales, f_ret, log)
                    if not df.empty:
                        err = df[df['Estado'].str.contains('ERROR')]
                        if not err.empty: st.error(f"❌ {len(err)} errores."); st.dataframe(err)
                        else: st.success("✅ Validación Exitosa."); st.dataframe(df.head())
                        
                        st.download_button("⬇️ Excel Resultados", generar_reporte_imprenta(df), "Validacion.xlsx")
                    with st.expander("Log"): st.write(log)
                except Exception as e: mostrar_error_amigable(e, "Validación")

    # --- PESTAÑA 2: GENERACIÓN ---
    with tab2:
        st.info("Calcula y genera el TXT cruzando Softland con el Libro de Ventas (Excel).")
        c1, c2 = st.columns(2)
        with c1: f_soft = st.file_uploader("1. Mayor Softland (Excel)", type=['xlsx'], key="g_soft")
        with c2: f_book = st.file_uploader("2. Libro Ventas GALAC (Excel)", type=['xlsx'], key="g_book")
            
        if f_soft and f_book:
            if st.button("Generar TXT", type="primary", key="btn_gen"):
                log = []
                try:
                    # IMPORTANTE: Estos nombres deben coincidir con logic.py
                    from logic import generar_txt_retenciones_galac
                    from utils import generar_archivo_txt, generar_reporte_auditoria_txt
                    
                    txt_lines, df_audit = generar_txt_retenciones_galac(f_soft, f_book, log)
                    
                    if df_audit is not None and not df_audit.empty:
                        st.success(f"✅ Procesado. {len(df_audit)} registros.")
                        st.dataframe(df_audit.head())
                        
                        col_a, col_b = st.columns(2)
                        col_a.download_button("⬇️ TXT para GALAC", generar_archivo_txt(txt_lines), "Retenciones_GALAC.txt")
                        col_b.download_button("⬇️ Auditoría Excel", generar_reporte_auditoria_txt(df_audit), "Auditoria_Imprenta.xlsx")
                    else:
                        st.warning("⚠️ No se generaron datos.")
                    
                    with st.expander("Log"): st.write(log)
                except Exception as e: mostrar_error_amigable(e, "Generación")

def render_pensiones():
    st.title("🛡️ Cálculo Ley Protección Pensiones (9%)", anchor=False)
    
    with st.expander("📖 Guía de Uso"):
        st.markdown(GUIA_PENSIONES)

    if st.button("⬅️ Volver al Inicio", key="back_pen"):
        set_page('inicio')
        st.rerun()
        
    # 1. Configuración de Empresa
    EMPRESAS_NOMINA = ["FEBECA", "BEVAL", "PRISMA", "QUINCALLA"]
    col_emp, _ = st.columns([1, 1])
    with col_emp:
        empresa_sel = st.selectbox("Seleccione la Empresa:", EMPRESAS_NOMINA, key="empresa_pensiones")

    # 2. Carga de Archivos
    c1, c2, c3 = st.columns([1.5, 1.5, 1])
    with c1:
        file_mayor = st.file_uploader("1. Mayor Contable (Excel)", type=['xlsx'], key="pen_mayor")
    with c2:
        file_nomina = st.file_uploader("2. Resumen Nómina (Validación)", type=['xlsx'], key="pen_nom")
    with c3:
        tasa = st.number_input("Tasa de Cambio", min_value=0.01, value=1.0, format="%.4f", key="pen_tasa")

    # 3. Botón de Acción
    if file_mayor and tasa > 0:
        if st.button("Calcular Impuesto", type="primary", use_container_width=True, key="btn_calc_pen"):
            log = []
            try:
                from logic import procesar_calculo_pensiones
                from utils import generar_reporte_pensiones
                
                with st.spinner("Procesando mayor contable y cruzando con nómina..."):
                    # Ejecutar lógica principal (Recibe 4 variables ahora)
                    df_calc, df_base, df_asiento, dict_val = procesar_calculo_pensiones(file_mayor, file_nomina, tasa, empresa_sel, log)

                # --- NUEVA ESTRUCTURA DE ALERTA ---
                if dict_val.get('estado') == 'OK':
                    st.success(f"✅ Cálculo exitoso. Nómina cuadra perfectamente.")
                else:
                    st.warning(
                        f"⚠️ Atención: Descuadres detectados (Ver Hoja 1).\n"
                        f"• Dif. Salarios: {dict_val.get('dif_salario', 0):,.2f}\n"
                        f"• Dif. Tickets: {dict_val.get('dif_ticket', 0):,.2f}\n"
                        f"• Dif. Impuesto: {dict_val.get('dif_imp', 0):,.2f}"
                    )
                
                if df_asiento is not None and not df_asiento.empty:
                    # Mostrar resultados en pantalla
                    total_pagar = df_asiento['Crédito VES'].sum()
                    st.success(f"✅ Cálculo exitoso para {empresa_sel}. Total a Pagar: Bs. {total_pagar:,.2f}")

                    st.subheader("Vista Previa del Asiento")

                # --- MEJORA DE VISUALIZACIÓN ---
                # 1. Definir orden lógico de columnas (Débito al lado de Crédito)
                cols_orden = [
                    'Centro Costo', 'Cuenta Contable', 'Descripción', 
                    'Débito VES', 'Crédito VES', 
                    'Débito USD', 'Crédito USD', 'Tasa'
                ]
                
                # Crear copia para visualización
                df_view = df_asiento[cols_orden].copy()
                
                # 2. Calcular Totales
                totales = {
                    'Centro Costo': 'TOTALES',
                    'Cuenta Contable': '', 'Descripción': '',
                    'Débito VES': df_view['Débito VES'].sum(),
                    'Crédito VES': df_view['Crédito VES'].sum(),
                    'Débito USD': df_view['Débito USD'].sum(),
                    'Crédito USD': df_view['Crédito USD'].sum(),
                    'Tasa': ''
                }
                
                # Agregar fila de totales al final
                df_view = pd.concat([df_view, pd.DataFrame([totales])], ignore_index=True)
                
                # 3. Aplicar Formato Venezolano (1.000,00)
                # Función auxiliar local para formatear
                def fmt_ve(x):
                    if isinstance(x, (float, int)):
                        # Formato: Miles con punto, Decimales con coma
                        return "{:,.2f}".format(x).replace(",", "X").replace(".", ",").replace("X", ".")
                    return x

                cols_num = ['Débito VES', 'Crédito VES', 'Débito USD', 'Crédito USD', 'Tasa']
                for col in cols_num:
                    df_view[col] = df_view[col].apply(fmt_ve)

                # Mostrar tabla mejorada
                st.dataframe(df_view, use_container_width=True, hide_index=True)
                # -------------------------------
                    
                    # --- PREPARACIÓN PARA EXCEL ---
                    # Intentamos detectar la fecha de cierre basada en el archivo cargado
                    fecha_cierre = pd.Timestamp.today()
                    try:
                        # Buscamos columna fecha sin importar mayúsculas
                        col_fecha = next((c for c in df_base.columns if 'FECHA' in c.upper()), None)
                        if col_fecha:
                            # Tomamos la primera fecha válida y calculamos el último día de ese mes
                            primera_fecha = pd.to_datetime(df_base[col_fecha].iloc[0])
                            fecha_cierre = primera_fecha + pd.offsets.MonthEnd(0)
                    except:
                        pass # Si falla, usa fecha de hoy
                    
                    # Generar Reporte Excel (Pasando dict_val también)
                    excel_data = generar_reporte_pensiones(df_calc, df_base, df_asiento, dict_val, empresa_sel, tasa, fecha_cierre)
                    
                    st.download_button(
                        "⬇️ Descargar Reporte Completo (Excel)",
                        excel_data,
                        f"Calculo_Pensiones_{empresa_sel}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("No se pudo generar el cálculo. Por favor revisa el log para ver el detalle del error.")

                # Mostrar Log
                with st.expander("Ver Log de Proceso"):
                    st.write(log)
                    
            except Exception as e:
                mostrar_error_amigable(e, "el Cálculo de Pensiones")

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
        'imprenta': render_imprenta,
        'pensiones': render_pensiones,
        'reservas': lambda: render_proximamente("Reservas y Apartados"),
        'proximamente': lambda: render_proximamente("Próximamente")
    }
    
    current_page = st.session_state.get('page', 'inicio')
    render_function = page_map.get(current_page, render_inicio)
    render_function()

if __name__ == "__main__":
    main()
