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
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import tempfile
from datetime import datetime

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

/* ===== ESTILOS ADICIONALES PARA LA NUEVA APP ===== */
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

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI =====

def setup_driver():
    """Configurar ChromeDriver para Selenium - VERSIÓN COMPATIBLE"""
    try:
        chrome_options = Options()
        
        # Opciones para mejor compatibilidad
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent real
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            return driver
        except Exception as e:
            st.error(f"❌ Error al configurar ChromeDriver: {e}")
            return None
            
    except Exception as e:
        st.error(f"❌ Error crítico al configurar ChromeDriver: {e}")
        return None

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
        st.error(f"❌ Error al extraer fecha: {e}")
        return None

def procesar_excel(uploaded_file):
    """
    Procesa el archivo Excel para extraer:
    - Valor a pagar (suma columna AK debajo de "Valor" en fila 38)
    - Número de pasos (de "TOTAL TRANSACCIONES X")
    """
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file, header=None)
        
        # Buscar el encabezado "Valor" en la columna AK (columna 36 en base 0)
        valor_a_pagar = 0
        numero_pasos = 0
        
        # Buscar fila 38 (índice 37) con "Valor" en columna AK
        try:
            fila_38 = df.iloc[37]  # Fila 38 (base 0 es 37)
            if pd.notna(fila_38[36]) and str(fila_38[36]).strip().upper() == "VALOR":
                # Encontramos el encabezado, sumar valores debajo
                for i in range(38, len(df)):  # Empezar desde fila 39
                    valor_celda = df.iloc[i, 36]
                    if pd.notna(valor_celda):
                        try:
                            # Convertir a número y sumar
                            valor_num = float(valor_celda)
                            valor_a_pagar += valor_num
                        except:
                            # Si no se puede convertir, continuar
                            continue
        except:
            # Si no encuentra en fila específica, buscar en todo el archivo
            for idx, fila in df.iterrows():
                if pd.notna(fila[36]) and str(fila[36]).strip().upper() == "VALOR":
                    for i in range(idx + 1, len(df)):
                        valor_celda = df.iloc[i, 36]
                        if pd.notna(valor_celda):
                            try:
                                valor_num = float(valor_celda)
                                valor_a_pagar += valor_num
                            except:
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
        st.error(f"❌ Error al procesar Excel: {e}")
        return 0, 0

def extraer_datos_power_bi(fecha_validacion):
    """
    Extrae datos del dashboard de Power BI para PEAJE ALVARADO
    """
    driver = None
    try:
        driver = setup_driver()
        if not driver:
            return None, None
        
        # URL del Power BI proporcionada
        power_bi_url = "https://app.powerbi.com/view?r=eyJrIjoiMDA5OGE5MTQtNjQ0MC00ZTdjLWJmNDItNGZhYmQxOWE5ZTk3IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
        
        st.info("🌐 Conectando con Power BI...")
        driver.get(power_bi_url)
        time.sleep(10)
        
        # Tomar screenshot inicial
        driver.save_screenshot("powerbi_inicial.png")
        
        # Buscar y seleccionar la fecha específica
        st.info(f"📅 Buscando conciliación para: {fecha_validacion}")
        
        # Buscar el elemento de la conciliación por fecha
        elemento_conciliacion = None
        selectors = [
            f"//*[contains(text(), 'Conciliación ALTERNATIVAS VIALES del {fecha_validacion}')]",
            f"//*[contains(text(), 'CONCILIACIÓN ALTERNATIVAS VIALES DEL {fecha_validacion}')]",
            f"//*[contains(text(), '{fecha_validacion}')]",
        ]
        
        for selector in selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    elemento_conciliacion = elemento
                    break
            except:
                continue
        
        if elemento_conciliacion:
            # Hacer clic en la conciliación
            driver.execute_script("arguments[0].scrollIntoView(true);", elemento_conciliacion)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(5)
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            return None, None
        
        # Tomar screenshot después de selección
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        # Buscar tabla "RESUMEN COMERCIOS" y extraer datos de PEAJE ALVARADO
        st.info("🔍 Extrayendo datos de RESUMEN COMERCIOS - PEAJE ALVARADO...")
        
        # Buscar la tabla RESUMEN COMERCIOS
        resumen_element = None
        resumen_selectors = [
            "//*[contains(text(), 'RESUMEN COMERCIOS')]",
            "//*[contains(text(), 'Resumen Comercios')]",
        ]
        
        for selector in resumen_selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    resumen_element = elemento
                    break
            except:
                continue
        
        if not resumen_element:
            st.error("❌ No se encontró la tabla RESUMEN COMERCIOS")
            return None, None
        
        # Buscar PEAJE ALVARADO en la tabla
        peaje_alvarado_element = None
        peaje_selectors = [
            "//*[contains(text(), 'PEAJE ALVARADO')]",
            "//*[contains(text(), 'Peaje Alvarado')]",
        ]
        
        for selector in peaje_selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    peaje_alvarado_element = elemento
                    break
            except:
                continue
        
        if not peaje_alvarado_element:
            st.error("❌ No se encontró PEAJE ALVARADO en el resumen")
            return None, None
        
        # Extraer valores de Cant Pasos y Valor a Pagar para PEAJE ALVARADO
        try:
            # Buscar el contenedor padre que contiene los datos
            container = peaje_alvarado_element.find_element(By.XPATH, "./ancestor::tr | ./ancestor::div")
            
            # Obtener todo el texto del contenedor
            container_text = container.text
            
            # Buscar número de pasos (formato: 554)
            pasos_match = re.search(r'(\d{1,6})\s+Cant\s+Pasos|\b(\d{1,6})\b.*Pasos', container_text)
            if pasos_match:
                pasos_power_bi = pasos_match.group(1) if pasos_match.group(1) else pasos_match.group(2)
            else:
                # Buscar cualquier número que parezca cantidad de pasos
                numeros = re.findall(r'\b\d{3,6}\b', container_text)
                if numeros:
                    pasos_power_bi = numeros[0]  # Primer número encontrado
                else:
                    pasos_power_bi = 0
            
            # Buscar valor monetario (formato: $10.472.900)
            valor_match = re.search(r'\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)', container_text)
            if valor_match:
                valor_power_bi_texto = valor_match.group(0)
                # Convertir a número
                valor_limpio = valor_power_bi_texto.replace('$', '').replace('.', '').replace(',', '.')
                valor_power_bi = float(valor_limpio)
            else:
                valor_power_bi = 0
            
            # Tomar screenshot final
            driver.save_screenshot("powerbi_final.png")
            
            return valor_power_bi, int(pasos_power_bi)
            
        except Exception as e:
            st.error(f"❌ Error extrayendo datos específicos: {e}")
            return None, None
        
    except Exception as e:
        st.error(f"❌ Error extrayendo datos de Power BI: {e}")
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
    
    # Información del reporte
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con formato específico
    - Extraer Valor a Pagar (columna AK) y Número de Pasos
    - Comparar con Power BI - PEAJE ALVARADO
    
    **Formato archivo:**
    - CrptTransaccionesTotal DD-MM-YYYY gopass
    - Columna AK, fila 38: encabezado "Valor"
    - Texto: "TOTAL TRANSACCIONES X"
    
    **Estado:** ✅ ChromeDriver Compatible
    **Versión:** v1.0 - Validación Conciliaciones
    """)
    
    # Estado del sistema
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    st.sidebar.info(f"✅ Streamlit {st.__version__}")
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel (Formato: CrptTransaccionesTotal DD-MM-YYYY gopass)", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Extraer fecha del nombre del archivo
        fecha_validacion = extraer_fecha_desde_nombre(uploaded_file.name)
        
        if fecha_validacion:
            st.success(f"📅 Fecha detectada automáticamente: {fecha_validacion}")
        else:
            st.warning("⚠️ No se pudo detectar la fecha del archivo")
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):", value="2025-10-16")
        
        if fecha_validacion:
            # Procesar el archivo Excel
            with st.spinner("📊 Procesando archivo Excel..."):
                valor_a_pagar, numero_pasos = procesar_excel(uploaded_file)
            
            if valor_a_pagar > 0 and numero_pasos > 0:
                # ========== MOSTRAR RESUMEN DE VALORES EXCEL ==========
                st.markdown("### 📊 Valores Extraídos del Excel")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("💰 Valor a Pagar (Excel)", f"${valor_a_pagar:,.0f}".replace(",", "."))
                
                with col2:
                    st.metric("👣 Número de Pasos (Excel)", f"{numero_pasos:,}")
                
                st.markdown("---")
                
                # ========== SECCIÓN CONSULTA POWER BI ==========
                st.subheader("🌐 Consulta Power BI")
                
                if st.button("🎯 Extraer Valores de Power BI y Validar", type="primary", use_container_width=True):
                    with st.spinner("🌐 Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
                        valor_power_bi, pasos_power_bi = extraer_datos_power_bi(fecha_validacion)
                    
                    if valor_power_bi is not None and pasos_power_bi is not None:
                        # ========== SECCIÓN RESULTADOS POWER BI ==========
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("💰 Valor a Pagar (Power BI)", f"${valor_power_bi:,.0f}".replace(",", "."))
                        
                        with col2:
                            st.metric("👣 Número de Pasos (Power BI)", f"{pasos_power_bi:,}")
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN COMPARACIÓN ==========
                        st.markdown("### 📊 Resultado de la Validación")
                        
                        # Comparar resultados
                        coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                            valor_a_pagar, valor_power_bi, numero_pasos, pasos_power_bi
                        )
                        
                        # Mostrar resultados de comparación
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if coinciden_valor:
                                st.success("✅ VALOR COINCIDE")
                            else:
                                st.error(f"❌ DIFERENCIA EN VALOR: ${dif_valor:,.0f}".replace(",", "."))
                        
                        with col2:
                            if coinciden_pasos:
                                st.success("✅ PASOS COINCIDEN")
                            else:
                                st.error(f"❌ DIFERENCIA EN PASOS: {dif_pasos:,}")
                        
                        # Resultado general
                        st.markdown("---")
                        st.markdown("### 📋 Resultado Final")
                        
                        if coinciden_valor and coinciden_pasos:
                            st.success("🎉 **VALIDACIÓN EXITOSA** - Todos los valores coinciden")
                            st.balloons()
                        else:
                            st.error("❌ **VALIDACIÓN FALLIDA** - Existen diferencias")
                        
                        # ========== TABLA COMPARATIVA ==========
                        st.markdown("### 📊 Resumen Comparativo")
                        
                        datos_comparacion = {
                            'Concepto': ['Valor a Pagar', 'Número de Pasos'],
                            'Excel': [
                                f"${valor_a_pagar:,.0f}".replace(",", "."), 
                                f"{numero_pasos:,}"
                            ],
                            'Power BI': [
                                f"${valor_power_bi:,.0f}".replace(",", "."), 
                                f"{pasos_power_bi:,}"
                            ],
                            'Resultado': [
                                '✅ COINCIDE' if coinciden_valor else f'❌ DIFERENCIA: ${dif_valor:,.0f}'.replace(",", "."),
                                '✅ COINCIDE' if coinciden_pasos else f'❌ DIFERENCIA: {dif_pasos:,}'
                            ]
                        }
                        
                        df_comparacion = pd.DataFrame(datos_comparacion)
                        st.dataframe(df_comparacion, use_container_width=True, hide_index=True)
                        
                        # ========== DETALLES ADICIONALES ==========
                        with st.expander("🔍 Ver Detalles Completos y Capturas"):
                            # Tabla detallada
                            st.markdown("#### 📊 Tabla Detallada")
                            resumen_data = []
                            
                            # Valor a Pagar
                            resumen_data.append({
                                'Concepto': 'Valor a Pagar',
                                'Excel': f"${valor_a_pagar:,.2f}".replace(",", "."),
                                'Power BI': f"${valor_power_bi:,.2f}".replace(",", "."),
                                'Diferencia': f"${dif_valor:,.2f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden_valor else '❌ No coincide'
                            })
                            
                            # Número de Pasos
                            resumen_data.append({
                                'Concepto': 'Número de Pasos',
                                'Excel': f"{numero_pasos:,}",
                                'Power BI': f"{pasos_power_bi:,}",
                                'Diferencia': f"{dif_pasos:,}",
                                'Estado': '✅ Coincide' if coinciden_pasos else '❌ No coincide'
                            })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                            
                            # Screenshots
                            st.markdown("#### 📸 Capturas del Proceso Power BI")
                            col1, col2, col3 = st.columns(3)
                            
                            if os.path.exists("powerbi_inicial.png"):
                                with col1:
                                    st.image("powerbi_inicial.png", caption="Vista Inicial", use_column_width=True)
                            
                            if os.path.exists("powerbi_despues_seleccion.png"):
                                with col2:
                                    st.image("powerbi_despues_seleccion.png", caption="Tras Selección", use_column_width=True)
                            
                            if os.path.exists("powerbi_final.png"):
                                with col3:
                                    st.image("powerbi_final.png", caption="Vista Final", use_column_width=True)
                        
                    else:
                        st.error("❌ No se pudieron extraer los datos del Power BI")
            else:
                st.error("❌ No se pudieron extraer los valores del archivo Excel")
                with st.expander("💡 Sugerencias para solucionar el problema"):
                    st.markdown("""
                    **Problemas comunes:**
                    - El archivo no tiene el formato esperado
                    - No se encuentra "Valor" en la columna AK, fila 38
                    - No se encuentra "TOTAL TRANSACCIONES X" en el archivo
                    - Los valores no son numéricos
                    
                    **Verifica:**
                    - El nombre del archivo contiene la fecha (DD-MM-YYYY)
                    - La columna AK tiene el encabezado "Valor" en la fila 38
                    - Hay valores numéricos debajo del encabezado "Valor"
                    - Existe el texto "TOTAL TRANSACCIONES" seguido de un número
                    """)
    
    else:
        st.info("👈 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso de Validación:**
        
        1. **Cargar Archivo Excel**: 
           - Formato: `CrptTransaccionesTotal DD-MM-YYYY gopass`
           - Ejemplo: `CrptTransaccionesTotal 12-10-2025 gopass.xlsx`
        
        2. **Extracción Automática**:
           - **Fecha**: Se detecta del nombre del archivo
           - **Valor a Pagar**: Suma de columna AK debajo de "Valor" (fila 38)
           - **Número de Pasos**: De "TOTAL TRANSACCIONES X"
        
        3. **Consulta Power BI**:
           - Se conecta al dashboard de Power BI
           - Selecciona la fecha correspondiente
           - Extrae datos de "RESUMEN COMERCIOS" - PEAJE ALVARADO
           - Obtiene "Cant Pasos" y "Valor a Pagar"
        
        4. **Comparación**:
           - Valida coincidencias entre Excel y Power BI
           - Muestra diferencias si existen
        
        **Requisitos del Archivo Excel:**
        - Formato: .xlsx o .xls
        - Nombre debe contener fecha: `DD-MM-YYYY`
        - Columna AK, fila 38: debe decir "Valor"
        - Debajo de "Valor" deben haber valores numéricos
        - Debe contener "TOTAL TRANSACCIONES X" (X = número de pasos)
        
        **Notas:**
        - La conexión a Power BI puede tomar 1-2 minutos
        - Las fechas deben coincidir exactamente
        - Los valores se comparan con tolerancia de 1 centavo
        - Los pasos deben coincidir exactamente
        """)

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v1.0</div>', unsafe_allow_html=True)
