# app.py

# ==============================================================================
# 1. IMPORTACIÓN DE LIBRERÍAS Y CONFIGURACIÓN INICIAL
# ==============================================================================
import streamlit as st
import pandas as pd
from functools import partial

# --- Importaciones desde nuestros módulos ---
from logic import (
    run_conciliation_fondos_en_transito,
    run_conciliation_fondos_por_depositar,
    run_conciliation_devoluciones_proveedores,
    run_conciliation_viajes,
    run_conciliation_retenciones
)
from utils import (
    cargar_y_limpiar_datos,
    generar_reporte_excel,
    generar_csv_saldos_abiertos
)

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
    
    # --- INICIO DE LA MODIFICACIÓN DEL DISEÑO ---

    # 1. Ajustamos la proporción de las columnas para hacer el cuadro central un poco más estrecho,
    #    lo que compacta el contenido verticalmente.
    _ , col_main, _ = st.columns([1, 1.5, 1])

    with col_main:
        # 2. Centramos el logo principal de forma más responsiva.
        _ , col_logo, _ = st.columns([1, 2, 1])
        with col_logo:
            try:
                # Usar 'use_column_width' es mejor que un ancho fijo para la adaptabilidad.
                st.image("assets/logo_principal.png", use_container_width=True)  
            except:
                st.warning("No se encontró el logo principal en la carpeta 'assets'.")

        st.title("Bienvenido al Portal de Herramientas Contables", anchor=False)
        st.markdown("Una solución centralizada para el equipo de contabilidad.")
        
        # Contenedor para el campo de contraseña
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
                    # Texto alternativo si el logo no se encuentra
                    st.markdown(f"<p style='text-align: center; font-size: small;'>{logos_info[i]['fallback']}</p>", unsafe_allow_html=True)

    # --- FIN DE LA MODIFICACIÓN DEL DISEÑO ---

    st.stop() # CRÍTICO: Detiene la ejecución del resto de la app.

# ==============================================================================
# DICCIONARIO CENTRAL DE ESTRATEGIAS (EL "CEREBRO")
# ==============================================================================
from functools import partial

# --- Creamos funciones parciales para las estrategias que no usan la barra de progreso ---
# Esto evita tener que modificar todas las funciones en logic.py
def run_conciliation_wrapper(func, df, log_messages, progress_bar=None):
    # Esta función simple llama a la función de lógica original, ignorando progress_bar
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
        "funcion_principal": run_conciliation_viajes, # La nueva función maestra
        "label_actual": "Movimientos del mes (Viajes)",
        "label_anterior": "Saldos anteriores (Viajes)",
        # El orden de columnas que solicitaste para el reporte
        "columnas_reporte": ['Asiento', 'NIT', 'Nombre del Proveedor', 'Referencia', 'Fecha', 'Monto_BS', 'Monto_USD', 'Tipo'],
        "nombre_hoja_excel": "114.03.1002",
        "columnas_requeridas": ['Fecha', 'Asiento', 'Referencia', 'Nombre del Proveedor', 'NIT', 'Débito Bolivar', 'Crédito Bolivar']
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
        st.button("📄 Especificaciones", on_click=set_page, args=['especificaciones'], use_container_width=True, type="primary")
        st.button("📦 Reservas y Apartados", on_click=set_page, args=['reservas'], use_container_width=True)
    with col2:
        st.button("🧾 Relaciones de Retenciones", on_click=set_page, args=['retenciones'], use_container_width=True)
        st.button("🔜 Próximamente", on_click=set_page, args=['proximamente'], use_container_width=True)

def render_proximamente(titulo):
    st.title(f"🛠️ {titulo}")
    st.info("Esta funcionalidad estará disponible en futuras versiones.")
    st.button("⬅️ Volver al Inicio", on_click=set_page, args=['inicio'])

def render_retenciones():
    st.title("🧾 Herramienta de Conciliación de Retenciones", anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_ret"):
        set_page('inicio')
        st.session_state.processing_ret_complete = False 
        st.rerun()

    st.markdown("""
    Esta herramienta audita el proceso de retenciones cruzando la **Preparación Contable (CP)**, 
    la **Fuente Oficial (GALAC)** y el **Diario Contable (CG)** para identificar discrepancias.
    """)

    # --- Carga de Archivos ---
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

    # --- Ejecución del Proceso ---
    if all([file_cp, file_cg, file_iva, file_islr, file_mun]):
        if st.button("▶️ Iniciar Auditoría de Retenciones", type="primary", use_container_width=True):
            with st.spinner('Ejecutando auditoría... Este proceso puede tardar unos momentos.'):
                log_messages = []
                reporte_resultado = run_conciliation_retenciones(
                    file_cp, file_cg, file_iva, file_islr, file_mun, log_messages
                )
                
                st.session_state.reporte_ret_output = reporte_resultado
                st.session_state.log_messages_ret = log_messages
                
                # --- CORRECCIÓN CLAVE ---
                # Esta variable AHORA se establece en True sin importar el resultado.
                # Su única función es indicar que el proceso ya se ejecutó.
                st.session_state.processing_ret_complete = True

    # --- Visualización de Resultados ---
    # Esta condición ahora se cumplirá siempre después de hacer clic en el botón.
    if st.session_state.get('processing_ret_complete', False):
        
        # La decisión de mostrar éxito o error se basa directamente en si hay un reporte.
        if st.session_state.reporte_ret_output:
            st.success("✅ ¡Auditoría de retenciones completada con éxito!")
            st.download_button(
                "⬇️ Descargar Reporte de Auditoría (Excel)",
                st.session_state.reporte_ret_output,
                "Conciliacion_Retenciones_Resultado.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            # Si no hay reporte, significa que hubo un error.
            st.error("❌ La auditoría finalizó con un error. Revisa el registro para más detalles.")

        # El registro detallado ahora SIEMPRE se mostrará, permitiendo la depuración.
        with st.expander("Ver registro detallado del proceso de auditoría"):
            st.text_area("Log de Conciliación de Retenciones", '\n'.join(st.session_state.log_messages_ret), height=400)

def render_especificaciones():
    st.title('🤖 Herramienta de Conciliación Automática', anchor=False)
    if st.button("⬅️ Volver al Inicio", key="back_from_spec"):
        set_page('inicio')
        st.session_state.processing_complete = False 
        st.rerun()
    st.markdown("Esta aplicación automatiza el proceso de conciliación de cuentas contables.")
    
    # --- Selección de Parámetros ---
    CASA_OPTIONS = ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"]
    CUENTA_OPTIONS = list(ESTRATEGIAS.keys())
    
    st.subheader("1. Seleccione la Empresa (Casa):", anchor=False)
    casa_seleccionada = st.selectbox("1. Seleccione la Empresa (Casa):", CASA_OPTIONS, label_visibility="collapsed")
    
    st.subheader("2. Seleccione la Cuenta Contable:", anchor=False)
    cuenta_seleccionada = st.selectbox("2. Seleccione la Cuenta Contable:", CUENTA_OPTIONS, label_visibility="collapsed")
    estrategia_actual = ESTRATEGIAS[cuenta_seleccionada]

    # --- Carga de Archivos ---
    st.subheader("3. Cargue los Archivos de Excel (.xlsx):", anchor=False)
    st.markdown("*Asegúrese de que los datos estén en la **primera hoja** y los **encabezados en la primera fila**.*")

    columnas = estrategia_actual.get("columnas_requeridas", [])
    if columnas:
        texto_columnas = "**Columnas Esenciales Requeridas:**\n"
        texto_columnas += "\n".join([f"- `{col}`" for col in columnas])
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
                    st.session_state.csv_output = generar_csv_saldos_abiertos(st.session_state.df_saldos_abiertos)
                    st.session_state.excel_output = generar_reporte_excel(
                        df_full, st.session_state.df_saldos_abiertos, st.session_state.df_conciliados,
                        estrategia_actual, casa_seleccionada, cuenta_seleccionada
                    )
                    st.session_state.log_messages = log_messages
                    st.session_state.processing_complete = True
            except Exception as e:
                st.error(f"❌ Ocurrió un error crítico durante el procesamiento: {e}")
                import traceback
                st.code(traceback.format_exc())
                st.session_state.processing_complete = False
            finally:
                progress_container.empty()

    if st.session_state.processing_complete:
        st.success("✅ ¡Conciliación completada con éxito!")
        res_col1, res_col2 = st.columns(2, gap="small")
        with res_col1:
            st.metric("Movimientos Conciliados", len(st.session_state.df_conciliados))
            st.download_button("⬇️ Descargar Reporte Completo (Excel)", st.session_state.excel_output, f"reporte_conciliacion_{estrategia_actual['id']}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_excel")
        with res_col2:
            st.metric("Saldos Abiertos (Pendientes)", len(st.session_state.df_saldos_abiertos))
            st.download_button("⬇️ Descargar Saldos para Próximo Mes (CSV)", st.session_state.csv_output, "saldos_para_proximo_mes.csv", "text/csv", use_container_width=True, key="download_csv")
        
        st.info("**Instrucción de Ciclo Mensual:** Para el próximo mes, debe usar el archivo CSV descargado como el archivo de 'saldos anteriores'.")
        
        with st.expander("Ver registro detallado del proceso"):
            st.text_area("Log de Conciliación", '\n'.join(st.session_state.log_messages), height=300, key="log_area")
            
        # --- PREVISUALIZACIÓN DE SALDOS PENDIENTES ---
        st.subheader("Previsualización de Saldos Pendientes", anchor=False)
        df_vista_previa = st.session_state.df_saldos_abiertos.copy()
        
        if estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar']:
            columnas_a_mostrar = ['Asiento', 'Referencia', 'Fecha', 'Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
            columnas_existentes = [col for col in columnas_a_mostrar if col in df_vista_previa.columns]
            df_vista_previa = df_vista_previa[columnas_existentes]
            if 'Fecha' in df_vista_previa.columns:
                df_vista_previa['Fecha'] = pd.to_datetime(df_vista_previa['Fecha']).dt.strftime('%d/%m/%Y')
            columnas_numericas = ['Débito Bolivar', 'Crédito Bolivar', 'Débito Dolar', 'Crédito Dolar']
            for col in columnas_numericas:
                if col in df_vista_previa.columns:
                    df_vista_previa[col] = df_vista_previa[col].apply(
                        lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else ''
                    )
        
        # --- NUEVO BLOQUE ELIF PARA PROVEEDORES ---
        elif estrategia_actual['id'] == 'devoluciones_proveedores':
            df_vista_previa.rename(columns={'Monto_BS': 'Monto Bolivar', 'Monto_USD': 'Monto Dolar'}, inplace=True)
            columnas_a_mostrar = ['Asiento', 'Referencia', 'Fecha', 'Nombre del Proveedor', 'NIT', 'Monto Bolivar', 'Monto Dolar']
            columnas_existentes = [col for col in columnas_a_mostrar if col in df_vista_previa.columns]
            df_vista_previa = df_vista_previa[columnas_existentes]
            if 'Fecha' in df_vista_previa.columns:
                df_vista_previa['Fecha'] = pd.to_datetime(df_vista_previa['Fecha']).dt.strftime('%d/%m/%Y')
            columnas_numericas = ['Monto Bolivar', 'Monto Dolar']
            for col in columnas_numericas:
                if col in df_vista_previa.columns:
                    df_vista_previa[col] = df_vista_previa[col].apply(
                        lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else ''
                    )

        st.dataframe(df_vista_previa, use_container_width=True)
        
        # --- PREVISUALIZACIÓN DE MOVIMIENTOS CONCILIADOS ---
        st.subheader("Previsualización de Movimientos Conciliados", anchor=False)
        df_conciliados_vista = st.session_state.df_conciliados.copy()

        if estrategia_actual['id'] in ['fondos_transito', 'fondos_depositar']:
            df_conciliados_vista.rename(columns={'Monto_BS': 'Monto Bolivar', 'Monto_USD': 'Monto Dolar'}, inplace=True)
            columnas_conciliados_mostrar = ['Asiento', 'Referencia', 'Fecha', 'Monto Bolivar', 'Monto Dolar', 'Grupo_Conciliado']
            columnas_existentes = [col for col in columnas_conciliados_mostrar if col in df_conciliados_vista.columns]
            df_conciliados_vista = df_conciliados_vista[columnas_existentes]
            if 'Fecha' in df_conciliados_vista.columns:
                df_conciliados_vista['Fecha'] = pd.to_datetime(df_conciliados_vista['Fecha']).dt.strftime('%d/%m/%Y')
            columnas_numericas = ['Monto Bolivar', 'Monto Dolar']
            for col in columnas_numericas:
                if col in df_conciliados_vista.columns:
                    df_conciliados_vista[col] = df_conciliados_vista[col].apply(
                        lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else ''
                    )
        
        # --- NUEVO BLOQUE ELIF PARA PROVEEDORES ---
        elif estrategia_actual['id'] == 'devoluciones_proveedores':
            df_conciliados_vista.rename(columns={'Monto_BS': 'Monto Bolivar', 'Monto_USD': 'Monto Dolar'}, inplace=True)
            columnas_conciliados_mostrar = ['Asiento', 'Referencia', 'Fecha', 'Nombre del Proveedor', 'NIT', 'Monto Bolivar', 'Monto Dolar', 'Grupo_Conciliado']
            columnas_existentes = [col for col in columnas_conciliados_mostrar if col in df_conciliados_vista.columns]
            df_conciliados_vista = df_conciliados_vista[columnas_existentes]
            if 'Fecha' in df_conciliados_vista.columns:
                df_conciliados_vista['Fecha'] = pd.to_datetime(df_conciliados_vista['Fecha']).dt.strftime('%d/%m/%Y')
            columnas_numericas = ['Monto Bolivar', 'Monto Dolar']
            for col in columnas_numericas:
                if col in df_conciliados_vista.columns:
                    df_conciliados_vista[col] = df_conciliados_vista[col].apply(
                        lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else ''
                    )
                    
        # --- NUEVO BLOQUE ELIF PARA VIAJES ---
        elif estrategia_actual['id'] == 'cuentas_viajes':
            df_conciliados_vista.rename(columns={'Monto_BS': 'Saldo Bs', 'Monto_USD': 'Saldo USD', 'Nombre del Proveedor': 'Nombre'}, inplace=True)
            columnas_conciliados_mostrar = ['Asiento', 'NIT', 'Nombre', 'Referencia', 'Fecha', 'Saldo Bs', 'Saldo USD', 'Tipo', 'Grupo_Conciliado']
            columnas_existentes = [col for col in columnas_conciliados_mostrar if col in df_conciliados_vista.columns]
            df_conciliados_vista = df_conciliados_vista[columnas_existentes]
            if 'Fecha' in df_conciliados_vista.columns:
                df_conciliados_vista['Fecha'] = pd.to_datetime(df_conciliados_vista['Fecha']).dt.strftime('%d/%m/%Y')
            columnas_numericas = ['Saldo Bs', 'Saldo USD']
            for col in columnas_numericas:
                if col in df_conciliados_vista.columns:
                    df_conciliados_vista[col] = df_conciliados_vista[col].apply(
                        lambda x: f"{x:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.') if pd.notna(x) else ''
                    )

        st.dataframe(df_conciliados_vista, use_container_width=True)
        
# ==============================================================================
# FLUJO PRINCIPAL DE LA APLICACIÓN (ROUTER)
# ==============================================================================
def main():
    if st.session_state.page == 'inicio':
        render_inicio()
    elif st.session_state.page == 'especificaciones':
        render_especificaciones()
    elif st.session_state.page == 'retenciones':
        render_retenciones()
    elif st.session_state.page == 'reservas':
        render_proximamente("Reservas y Apartados")
    elif st.session_state.page == 'proximamente':
        render_proximamente("Próximamente")
    else:
        st.session_state.page = 'inicio'
        st.experimental_rerun()

if __name__ == "__main__":
    main()
