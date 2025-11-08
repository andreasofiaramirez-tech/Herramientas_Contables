# app.py

# ==============================================================================
# 1. IMPORTACIÓN DE LIBRERÍAS Y CONFIGURACIÓN INICIAL
# ==============================================================================
import streamlit as st
import pandas as pd

# --- Importaciones desde nuestros módulos ---
from logic import (
    run_conciliation_fondos_en_transito,
    run_conciliation_fondos_por_depositar,
    run_conciliation_devoluciones_proveedores
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
    _ , col2, _ = st.columns([1, 2, 1])
    with col2:
        try:
            st.image("assets/logo_principal.png", width=300)
        except:
            st.warning("No se encontró el logo principal en la carpeta 'assets'.")
        st.title("Bienvenido al Portal de Herramientas Contables")
        st.markdown("Una solución centralizada para el equipo de contabilidad.")
        with st.container(border=True):
            st.subheader("Acceso Exclusivo")
            st.text_input(
                "Contraseña", type="password", on_change=password_entered, key="password", label_visibility="collapsed"
            )
            if st.session_state.get("authentication_attempted", False):
                if not st.session_state.get("password_correct", False):
                    st.error("😕 Contraseña incorrecta.")
            else:
                st.info("Por favor, ingresa la contraseña para continuar.")
        st.divider()
        st.markdown("<p style='text-align: center;'>Una herramienta para las empresas del grupo:</p>", unsafe_allow_html=True)
        logo_cols = st.columns(3)
        with logo_cols[0]:
            try: st.image("assets/logo_febeca.png")
            except: st.markdown("<p style='text-align: center;'>FEBECA, C.A.</p>", unsafe_allow_html=True)
        with logo_cols[1]:
            try: st.image("assets/logo_beval.png")
            except: st.markdown("<p style='text-align: center;'>MAYOR BEVAL, C.A.</p>", unsafe_allow_html=True)
        with logo_cols[2]:
            try: st.image("assets/logo_sillaca.png")
            except: st.markdown("<p style='text-align: center;'>SILLACA, C.A.</p>", unsafe_allow_html=True)
    st.stop()

# ==============================================================================
# DICCIONARIO CENTRAL DE ESTRATEGIAS (EL "CEREBRO")
# ==============================================================================
from functools import partial

# ... (resto de importaciones)

# --- Creamos funciones parciales para las estrategias que no usan la barra de progreso ---
# Esto evita tener que modificar todas las funciones en logic.py
def run_conciliation_wrapper(func, df, log_messages, progress_bar=None):
    # Esta función simple llama a la función de lógica original, ignorando progress_bar
    return func(df, log_messages)

ESTRATEGIAS = {
    "111.04.1001 - Fondos en Tránsito": { 
        "id": "fondos_transito", 
        # Usamos partial para "pre-configurar" el wrapper con la función correcta
        "funcion_principal": partial(run_conciliation_wrapper, run_conciliation_fondos_en_transito), 
        "label_actual": "Movimientos del mes (Fondos en Tránsito)", 
        "label_anterior": "Saldos anteriores (Fondos en Tránsito)", 
        "columnas_reporte": ['Asiento', 'Referencia', 'Fecha', 'Monto Dólar', 'Tasa', 'Bs.'], 
        "nombre_hoja_excel": "111.04.1001" 
    },
    "111.04.6001 - Fondos por Depositar - ME": { 
        "id": "fondos_depositar", 
        # Esta es la función que sí usa la barra de progreso, la dejamos como está
        "funcion_principal": run_conciliation_fondos_por_depositar, 
        "label_actual": "Movimientos del mes (Fondos por Depositar)", 
        "label_anterior": "Saldos anteriores (Fondos por Depositar)", 
        "columnas_reporte": ['Asiento', 'Referencia', 'Fecha', 'Monto Dólar', 'Tasa', 'Bs.'], 
        "nombre_hoja_excel": "111.04.6001" 
    },
    "212.07.6009 - Devoluciones a Proveedores": { 
        "id": "devoluciones_proveedores", 
        "funcion_principal": partial(run_conciliation_wrapper, run_conciliation_devoluciones_proveedores),
        "label_actual": "Reporte de Devoluciones (Proveedores)", 
        "label_anterior": "Partidas pendientes (Proveedores)", 
        "columnas_reporte": ['Fecha', 'Fuente', 'Referencia', 'Nombre del Proveedor', 'Monto USD', 'Monto Bs'], 
        "nombre_hoja_excel": "212.07.6009" 
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

def render_especificaciones():
    st.title('🤖 Herramienta de Conciliación Automática')
    if st.button("⬅️ Volver al Inicio", key="back_from_spec"):
        set_page('inicio')
        st.session_state.processing_complete = False 
        st.rerun()
    st.markdown("Esta aplicación automatiza el proceso de conciliación de cuentas contables.")
    CASA_OPTIONS = ["FEBECA, C.A", "MAYOR BEVAL, C.A", "PRISMA, C.A", "FEBECA, C.A (QUINCALLA)"]
    CUENTA_OPTIONS = list(ESTRATEGIAS.keys())
    casa_seleccionada = st.selectbox("**1. Seleccione la Empresa (Casa):**", CASA_OPTIONS)
    cuenta_seleccionada = st.selectbox("**2. Seleccione la Cuenta Contable:**", CUENTA_OPTIONS)
    estrategia_actual = ESTRATEGIAS[cuenta_seleccionada]
    st.markdown("""**3. Cargue los Archivos de Excel (.xlsx):**
    *Asegúrese de que los datos estén en la **primera hoja** y los **encabezados en la primera fila**.*""")
    col1, col2 = st.columns(2)
    with col1:
        uploaded_actual = st.file_uploader(estrategia_actual["label_actual"], type="xlsx", key=f"actual_{estrategia_actual['id']}")
    with col2:
        uploaded_anterior = st.file_uploader(estrategia_actual["label_anterior"], type="xlsx", key=f"anterior_{estrategia_actual['id']}")
    if uploaded_actual and uploaded_anterior:
        if st.button("▶️ Iniciar Conciliación", type="primary", use_container_width=True):
            with st.spinner('Procesando...'):
                log_messages = []
                try:
                    df_full = cargar_y_limpiar_datos(uploaded_actual, uploaded_anterior, log_messages)
                    if df_full is not None:
                        df_resultado = estrategia_actual["funcion_principal"](df_full.copy(), log_messages)
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
        st.subheader("Previsualización de Saldos Pendientes")
        st.dataframe(st.session_state.df_saldos_abiertos, use_container_width=True)
        st.subheader("Previsualización de Movimientos Conciliados")
        st.dataframe(st.session_state.df_conciliados, use_container_width=True)
        
# ==============================================================================
# FLUJO PRINCIPAL DE LA APLICACIÓN (ROUTER)
# ==============================================================================
def main():
    if st.session_state.page == 'inicio':
        render_inicio()
    elif st.session_state.page == 'especificaciones':
        render_especificaciones()
    elif st.session_state.page == 'retenciones':
        render_proximamente("Relaciones de Retenciones")
    elif st.session_state.page == 'reservas':
        render_proximamente("Reservas y Apartados")
    elif st.session_state.page == 'proximamente':
        render_proximamente("Próximamente")
    else:
        st.session_state.page = 'inicio'
        st.experimental_rerun()

if __name__ == "__main__":
    main()
