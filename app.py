import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD =====
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
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Configuración Streamlit
st.set_page_config(
    page_title="Validador Power BI - ALVARADO",
    page_icon="💰",
    layout="wide"
)

# ===== CSS =====
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
    width: 300px !important;
    padding: 20px 10px 20px 10px !important;
    border-right: 1px solid #333 !important;
}

[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}

[data-testid="stSidebar"] .stFileUploader > label {
    color: white !important;
    font-weight: bold;
}

[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-title,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-subtitle,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-list button,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-name,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-status,
[data-testid="stSidebar"] .stFileUploader span,
[data-testid="stSidebar"] .stFileUploader div {
    color: black !important;
}

[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button {
    color: black !important;
    background-color: #f0f0f0 !important;
    border: 1px solid #ccc !important;
}

[data-testid="stSidebar"] svg.icon {
    stroke: white !important;
    color: white !important;
    fill: none !important;
    opacity: 1 !important;
}
</style>
""", unsafe_allow_html=True)

# Logo
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES MEJORADAS (ADAPTADAS DE APP GICA) =====

def extraer_fecha_desde_nombre(nombre_archivo):
    """Extrae la fecha del nombre del archivo Excel"""
    try:
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
    """Procesa el archivo Excel para extraer valor a pagar y número de pasos"""
    try:
        df = pd.read_excel(uploaded_file, header=None)
        
        valor_a_pagar = 0
        numero_pasos = 0
        
        # Buscar fila con "Valor" en columna AK (índice 36)
        for idx, fila in df.iterrows():
            if pd.notna(fila[36]) and str(fila[36]).strip().upper() == "VALOR":
                # Sumar valores debajo del encabezado
                for i in range(idx + 1, len(df)):
                    valor_celda = df.iloc[i, 36]
                    if pd.notna(valor_celda):
                        try:
                            valor_num = float(valor_celda)
                            valor_a_pagar += valor_num
                        except:
                            continue
                break
        
        # Buscar "TOTAL TRANSACCIONES"
        for idx, fila in df.iterrows():
            for col in range(len(fila)):
                celda = str(fila[col])
                if "TOTAL TRANSACCIONES" in celda.upper():
                    numeros = re.findall(r'\d+', celda)
                    if numeros:
                        numero_pasos = int(numeros[0])
                        break
            if numero_pasos > 0:
                break
        
        return valor_a_pagar, numero_pasos
        
    except Exception as e:
        st.error(f"❌ Error procesando Excel: {e}")
        return 0, 0

def setup_driver():
    """Configurar ChromeDriver - ADAPTADO DE APP GICA"""
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        st.error(f"❌ Error configurando ChromeDriver: {e}")
        return None

def click_conciliacion_date(driver, fecha_objetivo):
    """Hacer clic en la conciliación específica por fecha - ADAPTADO DE APP GICA"""
    try:
        # Buscar el elemento que contiene la fecha exacta
        selectors = [
            f"//*[contains(text(), 'conciliación ALTERNATIVAS VIALES del {fecha_objetivo}')]",
            f"//*[contains(text(), 'CONCILIACIÓN ALTERNATIVAS VIALES DEL {fecha_objetivo}')]",
            f"//*[contains(text(), '{fecha_objetivo} 06:00')]",
            f"//div[contains(text(), '{fecha_objetivo}')]",
            f"//span[contains(text(), '{fecha_objetivo}')]",
        ]
        
        elemento_conciliacion = None
        for selector in selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    elemento_conciliacion = elemento
                    st.success(f"✅ Encontrado: {elemento.text.strip()}")
                    break
            except:
                continue
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView(true);", elemento_conciliacion)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(3)
            return True
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            return False
            
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliación: {str(e)}")
        return False

def find_alvarado_card(driver):
    """
    NUEVA FUNCIÓN: Buscar la tarjeta/tabla de PEAJE ALVARADO
    Adaptada de las estrategias exitosas de APP GICA
    """
    try:
        # Buscar por diferentes patrones del título
        titulo_selectors = [
            "//*[contains(text(), 'PEAJE ALVARADO')]",
            "//*[contains(text(), 'Peaje Alvarado')]",
            "//*[contains(text(), 'PEAJE') and contains(text(), 'ALVARADO')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if 'ALVARADO' in texto.upper():
                            titulo_element = elemento
                            st.success(f"✅ Encontrado título: {texto}")
                            break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            st.error("❌ No se encontró 'PEAJE ALVARADO' en el reporte")
            return None, None
        
        # ESTRATEGIA 1: Buscar en el mismo contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            valores_encontrados = []
            
            for elem in all_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto):
                    # Extraer números
                    numeros = re.findall(r'\d+(?:\.\d{3})*|\d+', texto)
                    for num in numeros:
                        num_limpio = num.replace('.', '')
                        if num_limpio.isdigit():
                            valores_encontrados.append(int(num_limpio))
            
            # Filtrar valores: el más pequeño son pasos, el más grande es valor
            if len(valores_encontrados) >= 2:
                valores_encontrados.sort()
                
                # Pasos típicamente son < 10,000
                pasos = None
                for val in valores_encontrados:
                    if val < 10000 and val > 100:
                        pasos = val
                        break
                
                # Valor típicamente es > 100,000
                valor = None
                for val in reversed(valores_encontrados):
                    if val > 100000:
                        valor = val
                        break
                
                if pasos and valor:
                    st.success(f"✅ Estrategia 1: Pasos={pasos}, Valor=${valor:,.0f}")
                    return valor, pasos
        except Exception as e:
            st.warning(f"⚠️ Estrategia 1 falló: {e}")
        
        # ESTRATEGIA 2: Buscar en elementos hermanos
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            valores_encontrados = []
            
            for sibling in siblings:
                if sibling != titulo_element and sibling.is_displayed():
                    texto = sibling.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        numeros = re.findall(r'\d+(?:\.\d{3})*|\d+', texto)
                        for num in numeros:
                            num_limpio = num.replace('.', '')
                            if num_limpio.isdigit():
                                valores_encontrados.append(int(num_limpio))
            
            if len(valores_encontrados) >= 2:
                valores_encontrados.sort()
                
                pasos = None
                for val in valores_encontrados:
                    if val < 10000 and val > 100:
                        pasos = val
                        break
                
                valor = None
                for val in reversed(valores_encontrados):
                    if val > 100000:
                        valor = val
                        break
                
                if pasos and valor:
                    st.success(f"✅ Estrategia 2: Pasos={pasos}, Valor=${valor:,.0f}")
                    return valor, pasos
        except Exception as e:
            st.warning(f"⚠️ Estrategia 2 falló: {e}")
        
        # ESTRATEGIA 3: Buscar en la tabla RESUMEN COMERCIOS
        try:
            # Buscar la tabla completa
            resumen_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'RESUMEN COMERCIOS')]")
            
            for resumen_elem in resumen_elements:
                if resumen_elem.is_displayed():
                    container = resumen_elem.find_element(By.XPATH, "./ancestor::div[position()<=10]")
                    container_text = container.text
                    
                    # Buscar la sección de PEAJE ALVARADO
                    if 'PEAJE ALVARADO' in container_text:
                        start_idx = container_text.find('PEAJE ALVARADO')
                        remaining_text = container_text[start_idx:]
                        
                        # Encontrar el final de la sección
                        end_markers = ['PEAJE ARMERO', 'PEAJE HONDA', 'TOTAL', 'Select Row']
                        end_idx = len(remaining_text)
                        for marker in end_markers:
                            idx = remaining_text.find(marker)
                            if idx != -1 and idx < end_idx:
                                end_idx = idx
                        
                        alvarado_section = remaining_text[:end_idx].strip()
                        st.info(f"📊 Sección ALVARADO: {alvarado_section}")
                        
                        # Extraer números de la sección
                        numeros = re.findall(r'\d+(?:\.\d{3})*|\d+', alvarado_section)
                        valores_encontrados = []
                        
                        for num in numeros:
                            num_limpio = num.replace('.', '')
                            if num_limpio.isdigit():
                                valores_encontrados.append(int(num_limpio))
                        
                        if len(valores_encontrados) >= 2:
                            valores_encontrados.sort()
                            
                            pasos = None
                            for val in valores_encontrados:
                                if val < 10000 and val > 100:
                                    pasos = val
                                    break
                            
                            valor = None
                            for val in reversed(valores_encontrados):
                                if val > 100000:
                                    valor = val
                                    break
                            
                            if pasos and valor:
                                st.success(f"✅ Estrategia 3: Pasos={pasos}, Valor=${valor:,.0f}")
                                return valor, pasos
        except Exception as e:
            st.warning(f"⚠️ Estrategia 3 falló: {e}")
        
        st.error("❌ No se pudieron extraer los valores de PEAJE ALVARADO")
        return None, None
        
    except Exception as e:
        st.error(f"❌ Error buscando PEAJE ALVARADO: {str(e)}")
        return None, None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - MEJORADA"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiMDA5OGE5MTQtNjQ0MC00ZTdjLWJmNDItNGZhYmQxOWE5ZTk3IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None, None
    
    try:
        # 1. Navegar al reporte
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        # 2. Tomar screenshot inicial
        driver.save_screenshot("powerbi_inicial.png")
        
        # 3. Hacer clic en la conciliación específica
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None, None
        
        # 4. Esperar a que cargue la selección
        time.sleep(5)
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        # 5. Buscar datos de PEAJE ALVARADO
        with st.spinner("🔍 Extrayendo datos de PEAJE ALVARADO..."):
            valor_power_bi, pasos_power_bi = find_alvarado_card(driver)
        
        # 6. Tomar screenshot final
        driver.save_screenshot("powerbi_final.png")
        
        return valor_power_bi, pasos_power_bi
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None, None
    finally:
        driver.quit()

def comparar_valores(valor_excel, valor_power_bi, pasos_excel, pasos_power_bi):
    """Compara los valores y determina si coinciden"""
    try:
        diferencia_valor = abs(valor_excel - valor_power_bi) if valor_power_bi else valor_excel
        diferencia_pasos = abs(pasos_excel - pasos_power_bi) if pasos_power_bi else pasos_excel
        
        coinciden_valor = diferencia_valor < 1.0
        coinciden_pasos = diferencia_pasos == 0
        
        return coinciden_valor, coinciden_pasos, diferencia_valor, diferencia_pasos
        
    except Exception as e:
        st.error(f"❌ Error comparando valores: {e}")
        return False, False, 0, 0

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.title("💰 Validador Power BI - PEAJE ALVARADO")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Validar conciliaciones entre Excel y Power BI
    - Extraer datos de PEAJE ALVARADO
    - Comparar valores y número de pasos
    
    **Estado:** ✅ Mejorado con estrategias de APP GICA
    **Versión:** v3.0 - Extracción Mejorada
    """)
    
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel (CrptTransaccionesTotal DD-MM-YYYY gopass)", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Extraer fecha del nombre
        fecha_validacion = extraer_fecha_desde_nombre(uploaded_file.name)
        
        if fecha_validacion:
            st.success(f"📅 Fecha detectada: {fecha_validacion}")
        else:
            st.warning("⚠️ No se pudo detectar la fecha")
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):")
        
        if fecha_validacion:
            # Procesar Excel
            with st.spinner("📊 Procesando archivo Excel..."):
                valor_excel, pasos_excel = procesar_excel(uploaded_file)
            
            if valor_excel > 0 and pasos_excel > 0:
                # Mostrar valores del Excel
                st.markdown("### 📊 Valores Extraídos del Excel")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("💰 Valor a Pagar", f"${valor_excel:,.0f}")
                with col2:
                    st.metric("👣 Número de Pasos", f"{pasos_excel}")
                
                st.markdown("---")
                
                # Botón para extraer de Power BI
                if st.button("🎯 Extraer Valores de Power BI y Validar", type="primary", use_container_width=True):
                    with st.spinner("🌐 Extrayendo datos de Power BI..."):
                        valor_power_bi, pasos_power_bi = extract_powerbi_data(fecha_validacion)
                    
                    if valor_power_bi is not None and pasos_power_bi is not None:
                        # Mostrar resultados de Power BI
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col3, col4 = st.columns(2)
                        with col3:
                            st.metric("💰 Valor a Pagar (Power BI)", f"${valor_power_bi:,.0f}")
                        with col4:
                            st.metric("👣 Número de Pasos (Power BI)", f"{pasos_power_bi}")
                        
                        st.markdown("---")
                        
                        # Comparar
                        st.markdown("### 📊 Resultado de la Validación")
                        
                        coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                            valor_excel, valor_power_bi, pasos_excel, pasos_power_bi
                        )
                        
                        if coinciden_valor and coinciden_pasos:
                            st.success("🎉 ✅ TODOS LOS VALORES COINCIDEN")
                            st.balloons()
                        else:
                            if not coinciden_valor:
                                st.error(f"❌ DIFERENCIA EN VALOR: ${dif_valor:,.0f}")
                            if not coinciden_pasos:
                                st.error(f"❌ DIFERENCIA EN PASOS: {dif_pasos} pasos")
                        
                        # Tabla resumen
                        st.markdown("### 📋 Resumen de Comparación")
                        
                        datos = {
                            'Concepto': ['Valor a Pagar', 'Número de Pasos'],
                            'Excel': [f"${valor_excel:,.0f}", f"{pasos_excel}"],
                            'Power BI': [f"${valor_power_bi:,.0f}", f"{pasos_power_bi}"],
                            'Resultado': [
                                '✅ COINCIDE' if coinciden_valor else f'❌ DIFERENCIA: ${dif_valor:,.0f}',
                                '✅ COINCIDE' if coinciden_pasos else f'❌ DIFERENCIA: {dif_pasos} pasos'
                            ]
                        }
                        
                        df = pd.DataFrame(datos)
                        st.dataframe(df, use_container_width=True, hide_index=True)
                        
                        # Screenshots
                        with st.expander("📸 Ver Capturas del Proceso"):
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
                        st.error("❌ No se pudieron extraer los datos de Power BI")
            else:
                st.error("❌ No se pudieron extraer los valores del Excel")
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar")
    
    # Ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. Cargar archivo Excel con formato `CrptTransaccionesTotal DD-MM-YYYY gopass`
        2. Detección automática de fecha
        3. Extracción de valores del Excel (columna AK)
        4. Conexión con Power BI y selección de fecha
        5. Extracción de datos de PEAJE ALVARADO
        6. Comparación y validación
        
        **Mejoras v3.0:**
        - ✅ Estrategias de extracción adaptadas de APP GICA
        - ✅ Búsqueda múltiple en contenedores
        - ✅ Identificación inteligente de valores y pasos
        - ✅ Filtrado por rangos numéricos razonables
        - ✅ Capturas de pantalla del proceso
        """)

if __name__ == "__main__":
    main()
    
    st.markdown("---")
    st.markdown('<div style="text-align: center;">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v3.0</div>', unsafe_allow_html=True)
    
