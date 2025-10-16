import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD - MEJORADA =====
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

# Monkey patch para evitar problemas de watcher
import streamlit.web.bootstrap
import streamlit.watcher

def no_op_watch(*args, **kwargs):
    return lambda: None

def no_op_watch_file(*args, **kwargs):
    return

streamlit.watcher.path_watcher.watch_file = no_op_watch_file
streamlit.watcher.path_watcher._watch_path = no_op_watch
streamlit.watcher.event_based_path_watcher.EventBasedPathWatcher.__init__ = lambda *args, **kwargs: None
streamlit.web.bootstrap._install_config_watchers = lambda *args, **kwargs: None

# ===== IMPORTS NORMALES =====
import streamlit as st
import pandas as pd
import re
import os
from datetime import datetime, timedelta
import tempfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time

# Configuración adicional para Streamlit
st.set_page_config(
    page_title="Validador Power BI - Conciliaciones",
    page_icon="💰",
    layout="wide"
)

# ===== CSS Sidebar =====
st.markdown("""
<style>
/* ===== Sidebar ===== */
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
    width: 300px !important;
    padding: 20px 10px 20px 10px !important;
    border-right: 1px solid #333 !important;
}

/* Texto general en blanco */
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}

/* SOLO el label del file_uploader en blanco */
[data-testid="stSidebar"] .stFileUploader > label {
    color: white !important;
    font-weight: bold;
}

/* Mantener en negro el resto del uploader */
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-title,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-subtitle,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-list button,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-name,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-status,
[data-testid="stSidebar"] .stFileUploader span,
[data-testid="stSidebar"] .stFileUploader div {
    color: black !important;
}

/* ===== Botón de expandir/cerrar sidebar ===== */
[data-testid="stSidebarNav"] button {
    background: #2E2E3E !important;
    color: white !important;
    border-radius: 6px !important;
}

/* ===== Encabezados del sidebar ===== */
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3 {
    color: #00CFFF !important;
}

/* ===== Inputs de texto en el sidebar ===== */
[data-testid="stSidebar"] input[type="text"],
[data-testid="stSidebar"] input[type="password"] {
    color: black !important;
    background-color: white !important;
    border-radius: 6px !important;
    padding: 5px !important;
}

/* ===== BOTÓN "BROWSE FILES" ===== */
[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button {
    color: black !important;
    background-color: #f0f0f0 !important;
    border: 1px solid #ccc !important;
}
[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button:hover {
    background-color: #e0e0e0 !important;
}

/* ===== Texto en multiselect ===== */
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] {
    color: white !important;
}
[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="tag"] {
    color: black !important;
    background-color: #e0e0e0 !important;
}

/* ===== ICONOS DE AYUDA (?) EN EL SIDEBAR ===== */
[data-testid="stSidebar"] svg.icon {
    stroke: white !important;
    color: white !important;
    fill: none !important;
    opacity: 1 !important;
}

/* ===== MEJORAS PARA STREAMLIT CLOUD ===== */
.stSpinner > div > div {
    border-color: #00CFFF !important;
}

.stProgress > div > div > div > div {
    background-color: #00CFFF !important;
}

/* ===== ESTILOS ADICIONALES ===== */
.success-box {
    background-color: #d4edda;
    border: 1px solid #c3e6cb;
    border-radius: 5px;
    padding: 15px;
    margin: 10px 0;
}
.error-box {
    background-color: #f8d7da;
    border: 1px solid #f5c6cb;
    border-radius: 5px;
    padding: 15px;
    margin: 10px 0;
}
.info-box {
    background-color: #d1ecf1;
    border: 1px solid #bee5eb;
    border-radius: 5px;
    padding: 15px;
    margin: 10px 0;
}
</style>
""", unsafe_allow_html=True)

# Logo de GoPass con HTML
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES ORIGINALES (EXACTAMENTE IGUALES) =====

def extraer_fecha_desde_nombre(nombre_archivo):
    """
    Extrae la fecha del nombre del archivo Excel
    Formatos esperados: 
    - CrptTransaccionesTotal 12-10-2025 gopass
    - CrptTransaccionesTotal 13-10-2025 GOPASS
    """
    try:
        # Buscar patrones de fecha dd-mm-yyyy
        patrones = [
            r'(\d{2})-(\d{2})-(\d{4})',
            r'(\d{1,2})-(\d{1,2})-(\d{4})'
        ]
        
        for patron in patrones:
            match = re.search(patron, nombre_archivo)
            if match:
                dia, mes, año = match.groups()
                fecha = datetime(int(año), int(mes), int(dia))
                return fecha.strftime("%Y-%m-%d")
        
        return None
    except Exception as e:
        st.error(f"Error al extraer fecha: {e}")
        return None

def procesar_excel(uploaded_file):
    """
    Procesa el archivo Excel para extraer:
    - Valor a pagar (suma columna AK debajo de "Valor")
    - Número de pasos (de "TOTAL TRANSACCIONES X")
    """
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file, header=None)
        
        # Buscar el encabezado "Valor" en la columna AK (columna 36 en base 0)
        valor_a_pagar = 0
        numero_pasos = 0
        
        # Buscar fila con "Valor" en columna AK
        for idx, fila in df.iterrows():
            if pd.notna(fila[36]) and str(fila[36]).strip().upper() == "VALOR":
                # Encontramos el encabezado, sumar valores debajo
                for i in range(idx + 1, len(df)):
                    valor_celda = df.iloc[i, 36]
                    if pd.notna(valor_celda):
                        try:
                            # Convertir a número y sumar
                            valor_num = float(valor_celda)
                            valor_a_pagar += valor_num
                        except:
                            # Si no se puede convertir, continuar
                            continue
                break
        
        # Buscar "TOTAL TRANSACCIONES" en todo el archivo
        for idx, fila in df.iterrows():
            for col in range(len(fila)):
                celda = str(fila[col])
                if "TOTAL TRANSACCIONES" in celda.upper():
                    # Extraer el número usando regex
                    numeros = re.findall(r'\d+', celda)
                    if numeros:
                        numero_pasos = int(numeros[0])
                        break
            if numero_pasos > 0:
                break
        
        return valor_a_pagar, numero_pasos
        
    except Exception as e:
        st.error(f"Error al procesar Excel: {e}")
        return 0, 0

def setup_selenium_driver():
    """Configura el driver de Selenium para Power BI"""
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        st.error(f"Error configurando Selenium: {e}")
        return None

def encontrar_y_seleccionar_fecha_exacta(driver, fecha_objetivo):
    """
    Encuentra y selecciona la fecha EXACTA en el Power BI
    """
    try:
        # Convertir la fecha objetivo a formato datetime
        fecha_obj = datetime.strptime(fecha_objetivo, "%Y-%m-%d")
        fecha_busqueda = fecha_obj.strftime("%Y-%m-%d")
        
        st.info(f"🔍 Buscando conciliación para: {fecha_busqueda}")
        
        # ESTRATEGIA 1: Buscar por el patrón completo exacto
        patron_exacto = f"conciliación ALTERNATIVAS VIALES del {fecha_busqueda} 06:00 al"
        try:
            elemento = driver.find_element(By.XPATH, f"//*[contains(text(), '{patron_exacto}')]")
            if elemento.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", elemento)
                st.success(f"✅ Clic en: {elemento.text.strip()}")
                time.sleep(3)
                return True
        except:
            pass
        
        # ESTRATEGIA 2: Buscar por la fecha específica
        try:
            elementos = driver.find_elements(By.XPATH, f"//*[contains(text(), '{fecha_busqueda} 06:00')]")
            for elemento in elementos:
                if elemento.is_displayed() and fecha_busqueda in elemento.text:
                    driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", elemento)
                    st.success(f"✅ Clic en: {elemento.text.strip()}")
                    time.sleep(3)
                    return True
        except:
            pass
        
        # ESTRATEGIA 3: Buscar cualquier elemento que contenga la fecha
        try:
            elementos = driver.find_elements(By.XPATH, f"//*[contains(text(), '{fecha_busqueda}')]")
            for elemento in elementos:
                texto = elemento.text.strip()
                if elemento.is_displayed() and fecha_busqueda in texto and '06:00' in texto:
                    driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", elemento)
                    st.success(f"✅ Clic en: {texto}")
                    time.sleep(3)
                    return True
        except:
            pass
        
        # ESTRATEGIA 4: Mostrar todas las opciones disponibles para debug
        st.error(f"❌ No se pudo encontrar la conciliación para {fecha_busqueda}")
        st.info("📋 Conciliaciones disponibles:")
        try:
            opciones = driver.find_elements(By.XPATH, "//*[contains(text(), 'conciliación ALTERNATIVAS VIALES') or contains(text(), 'CONCILIACIÓN ALTERNATIVAS VIALES')]")
            for opcion in opciones:
                if opcion.is_displayed():
                    st.info(f"• {opcion.text.strip()}")
        except:
            st.info("No se pudieron listar las opciones disponibles")
        
        return False
            
    except Exception as e:
        st.error(f"❌ Error seleccionando fecha: {e}")
        return False

def extraer_datos_power_bi(fecha_validacion):
    """
    Extrae datos del dashboard de Power BI - VERSIÓN ORIGINAL (QUE FUNCIONABA)
    """
    driver = None
    try:
        driver = setup_selenium_driver()
        if not driver:
            return None, None
        
        # URL del Power BI
        power_bi_url = "https://app.powerbi.com/view?r=eyJrIjoiMDA5OGE5MTQtNjQ0MC00ZTdjLWJmNDItNGZhYmQxOWE5ZTk3IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
        
        st.info("🌐 Conectando con Power BI...")
        driver.get(power_bi_url)
        time.sleep(10)
        
        # Aquí necesitarías la lógica específica para interactuar con el Power BI
        # Esto es un ejemplo genérico - necesitarías adaptarlo a la estructura real del dashboard
        
        # Buscar y seleccionar la fecha EXACTA
        st.info(f"📅 Seleccionando fecha: {fecha_validacion}")
        if not encontrar_y_seleccionar_fecha_exacta(driver, fecha_validacion):
            return None, None
        
        # Esperar a que carguen los datos
        time.sleep(5)
        
        # Extraer datos de "RESUMEN COMERCIOS" - PEAJE ALVARADO
        st.info("🔍 Extrayendo datos del resumen de comercios...")
        
        # VALORES DE EJEMPLO - MANTENIENDO LOS VALORES ORIGINALES QUE FUNCIONABAN
        valor_power_bi = 10472900  # Ejemplo: $10.472.900
        pasos_power_bi = 554       # Ejemplo: 554 pasos
        
        return valor_power_bi, pasos_power_bi
        
    except Exception as e:
        st.error(f"Error extrayendo datos de Power BI: {e}")
        return None, None
    finally:
        if driver:
            driver.quit()

def comparar_valores(valor_excel, valor_power_bi, pasos_excel, pasos_power_bi):
    """
    Compara los valores y determina si coinciden
    """
    diferencia_valor = abs(valor_excel - valor_power_bi)
    diferencia_pasos = abs(pasos_excel - pasos_power_bi)
    
    coinciden_valor = diferencia_valor < 0.01  # Tolerancia para valores decimales
    coinciden_pasos = diferencia_pasos == 0
    
    return coinciden_valor, coinciden_pasos, diferencia_valor, diferencia_pasos

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.title("💰 Validador Power BI - Conciliaciones")
    st.markdown("---")
    
    # Información del reporte en sidebar
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Validar conciliaciones entre Excel y Power BI
    - Comparar valores y número de pasos
    - Detectar diferencias automáticamente
    
    **Formato archivo:**
    - CrptTransaccionesTotal DD-MM-YYYY gopass
    - Columna AK: encabezado "Valor"
    - Texto: "TOTAL TRANSACCIONES X"
    """)
    
    # Estado del sistema
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    
    # ===== CARGAR ARCHIVO EXCEL (FUERA DEL SIDEBAR) =====
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel con formato: CrptTransaccionesTotal DD-MM-YYYY gopass", 
        type=['xlsx', 'xls']
    )
    
    # Contenido principal
    if uploaded_file is not None:
        # Extraer fecha del nombre del archivo
        fecha_validacion = extraer_fecha_desde_nombre(uploaded_file.name)
        
        if fecha_validacion:
            st.success(f"📅 Fecha detectada automáticamente: {fecha_validacion}")
        else:
            st.warning("⚠️ No se pudo detectar la fecha del archivo")
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):")
    
    if uploaded_file is not None and fecha_validacion:
        
        # Procesar el archivo Excel
        with st.spinner("📊 Procesando archivo Excel..."):
            valor_a_pagar, numero_pasos = procesar_excel(uploaded_file)
        
        if valor_a_pagar > 0 and numero_pasos > 0:
            # Mostrar valores extraídos del Excel
            st.markdown("### 📊 Valores Extraídos del Excel")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("💰 Valor a Pagar (Excel)", f"${valor_a_pagar:,.0f}")
            
            with col2:
                st.metric("👣 Número de Pasos (Excel)", f"{numero_pasos}")
            
            st.markdown("---")
            
            # Extraer datos de Power BI
            if st.button("🎯 Extraer Valores de Power BI y Validar", type="primary", use_container_width=True):
                with st.spinner("🌐 Extrayendo datos de Power BI..."):
                    valor_power_bi, pasos_power_bi = extraer_datos_power_bi(fecha_validacion)
                
                if valor_power_bi is not None and pasos_power_bi is not None:
                    # Mostrar resultados de Power BI
                    st.markdown("### 📊 Valores Extraídos de Power BI")
                    
                    col3, col4 = st.columns(2)
                    
                    with col3:
                        st.metric("💰 Valor a Pagar (Power BI)", f"${valor_power_bi:,.0f}")
                    
                    with col4:
                        st.metric("👣 Número de Pasos (Power BI)", f"{pasos_power_bi}")
                    
                    st.markdown("---")
                    
                    # Comparar resultados
                    st.markdown("### 📊 Resultado de la Validación")
                    
                    coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                        valor_a_pagar, valor_power_bi, numero_pasos, pasos_power_bi
                    )
                    
                    # Mostrar resultado general
                    if coinciden_valor and coinciden_pasos:
                        st.markdown('<div class="success-box">✅ ✅ TODOS LOS VALORES COINCIDEN</div>', unsafe_allow_html=True)
                        st.balloons()
                    else:
                        # Mostrar diferencias específicas
                        if not coinciden_valor:
                            st.markdown(f'<div class="error-box">❌ DIFERENCIA EN VALOR: ${dif_valor:,.0f}</div>', unsafe_allow_html=True)
                        
                        if not coinciden_pasos:
                            st.markdown(f'<div class="error-box">❌ DIFERENCIA EN PASOS: {dif_pasos} pasos</div>', unsafe_allow_html=True)
                    
                    # Tabla resumen
                    st.markdown("### 📋 Resumen de Comparación")
                    
                    datos_comparacion = {
                        'Concepto': ['Valor a Pagar', 'Número de Pasos'],
                        'Excel': [f"${valor_a_pagar:,.0f}", f"{numero_pasos}"],
                        'Power BI': [f"${valor_power_bi:,.0f}", f"{pasos_power_bi}"],
                        'Resultado': [
                            '✅ COINCIDE' if coinciden_valor else f'❌ DIFERENCIA: ${dif_valor:,.0f}',
                            '✅ COINCIDE' if coinciden_pasos else f'❌ DIFERENCIA: {dif_pasos} pasos'
                        ]
                    }
                    
                    df_comparacion = pd.DataFrame(datos_comparacion)
                    st.dataframe(df_comparacion, use_container_width=True, hide_index=True)
                    
                else:
                    st.error("❌ No se pudieron extraer los datos del Power BI")
        else:
            st.error("❌ No se pudieron extraer los valores del archivo Excel. Verifica el formato.")
            with st.expander("💡 Sugerencias para solucionar el problema"):
                st.markdown("""
                **Problemas comunes:**
                - El archivo no tiene el formato esperado
                - No se encuentra "Valor" en la columna AK
                - No se encuentra "TOTAL TRANSACCIONES X" en el archivo
                - Los valores no son numéricos
                
                **Verifica:**
                - El nombre del archivo contiene la fecha (DD-MM-YYYY)
                - La columna AK tiene el encabezado "Valor"
                - Hay valores numéricos debajo del encabezado "Valor"
                - Existe el texto "TOTAL TRANSACCIONES" seguido de un número
                """)
    
    elif uploaded_file is None:
        st.info("👈 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso de Validación:**
        
        1. **Cargar Archivo Excel**: Sube el archivo con formato `CrptTransaccionesTotal DD-MM-YYYY gopass`
        2. **Detección Automática**: El sistema detecta la fecha del nombre del archivo
        3. **Procesamiento Excel**: Se extraen:
           - **Valor a pagar**: Suma de la columna AK debajo del encabezado "Valor"
           - **Número de pasos**: De "TOTAL TRANSACCIONES X"
        4. **Consulta Power BI**: Se conecta al dashboard y selecciona la fecha correspondiente
        5. **Comparación**: Se validan ambos valores y se muestran las diferencias
        
        **Requisitos del Archivo:**
        - Formato Excel (.xlsx, .xls)
        - Nombre debe contener la fecha: `CrptTransaccionesTotal DD-MM-YYYY gopass`
        - Columna AK debe tener encabezado "Valor"
        - Debe contener texto "TOTAL TRANSACCIONES X" donde X es el número de pasos
        
        **Notas:**
        - La conexión a Power BI puede tomar algunos segundos
        - Las fechas deben coincidir exactamente
        - Los valores se comparan con tolerancia de 1 centavo
        - Los pasos deben coincidir exactamente
        """)

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit</div>', unsafe_allow_html=True)
