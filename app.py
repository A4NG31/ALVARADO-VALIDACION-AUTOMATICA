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

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI (USANDO LAS FUNCIONES PROBADAS) =====

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

def click_conciliacion_date(driver, fecha_objetivo):
    """Hacer clic en la conciliación específica por fecha - FUNCIÓN PROBADA"""
    try:
        # Buscar el elemento que contiene la fecha exacta
        selectors = [
            f"//*[contains(text(), 'Conciliación APP GICA del {fecha_objetivo}')]",
            f"//*[contains(text(), 'CONCILIACIÓN APP GICA DEL {fecha_objetivo}')]",
            f"//*[contains(text(), '{fecha_objetivo} 00:00 al {fecha_objetivo} 11:59')]",
            f"//div[contains(text(), '{fecha_objetivo}')]",
            f"//span[contains(text(), '{fecha_objetivo}')]",
        ]
        
        elemento_conciliacion = None
        for selector in selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    elemento_conciliacion = elemento
                    break
            except:
                continue
        
        if elemento_conciliacion:
            # Hacer clic en el elemento
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

def find_valor_a_pagar_comercio_card(driver):
    """Buscar la tarjeta/table 'VALOR A PAGAR A COMERCIO' - FUNCIÓN PROBADA"""
    try:
        # Buscar por diferentes patrones del título
        titulo_selectors = [
            "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]",
            "//*[contains(text(), 'Valor a pagar a comercio')]",
            "//*[contains(text(), 'VALOR A PAGAR') and contains(text(), 'COMERCIO')]",
            "//*[contains(text(), 'Valor A Pagar') and contains(text(), 'Comercio')]",
            "//*[contains(text(), 'PAGAR A COMERCIO')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if "PAGAR" in texto.upper() and "COMERCIO" in texto.upper():
                            titulo_element = elemento
                            break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            st.error("❌ No se encontró 'VALOR A PAGAR A COMERCIO' en el reporte")
            return None
        
        # Buscar el valor numérico debajo del título
        # Estrategia 1: Buscar en el mismo contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and texto != titulo_element.text:
                    return texto
        except:
            pass
        
        # Estrategia 2: Buscar en elementos hermanos
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        return texto
        except:
            pass
        
        # Estrategia 3: Buscar debajo del título
        try:
            following_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]/following::*")
            
            for elem in following_elements[:10]:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and len(texto) < 50:
                    return texto
        except:
            pass
        
        st.error("❌ No se pudo encontrar el valor numérico")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando valor: {str(e)}")
        return None

def find_cantidad_pasos_card(driver):
    """Buscar la tarjeta/table 'CANTIDAD PASOS' - FUNCIÓN PROBADA"""
    try:
        # Buscar por diferentes patrones del título - MÁS ESPECÍFICO
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad de Pasos')]",
            "//*[contains(text(), 'CANTIDAD') and contains(text(), 'PASOS')]",
            "//*[text()='CANTIDAD PASOS']",
            "//*[text()='Cantidad Pasos']",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if any(palabra in texto.upper() for palabra in ['CANTIDAD', 'PASOS']):
                            titulo_element = elemento
                            break
                if titulo_element:
                    break
            except Exception as e:
                continue
        
        if not titulo_element:
            st.warning("❌ No se encontró el título 'CANTIDAD PASOS'")
            return None
        
        # ESTRATEGIA MEJORADA: Buscar en el mismo contenedor o contenedores cercanos
        try:
            # Buscar en el contenedor padre
            container = titulo_element.find_element(By.XPATH, "./..")
            
            # Buscar TODOS los elementos numéricos en el contenedor
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in all_elements:
                texto = elem.text.strip()
                # Verificar si es un número (contiene dígitos pero no texto largo)
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and 
                    texto != titulo_element.text and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    # Verificar formato numérico (puede tener comas, puntos, pero ser principalmente números)
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:  # Al menos un dígito
                        st.success(f"✅ Valor numérico encontrado: {texto}")
                        return texto
                        
        except Exception as e:
            st.warning(f"⚠️ Estrategia 1 falló: {e}")
        
        # ESTRATEGIA 2: Buscar elementos hermanos específicamente
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                        
                        digit_count = sum(char.isdigit() for char in texto)
                        if digit_count >= 1:
                            st.success(f"✅ Valor encontrado en hermano: {texto}")
                            return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 2 falló: {e}")
        
        # ESTRATEGIA 3: Buscar elementos que siguen al título
        try:
            # Buscar elementos que están después del título
            following_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), 'CANTIDAD PASOS')]/following::*")
            
            for i, elem in enumerate(following_elements[:20]):  # Buscar en los primeros 20 elementos siguientes
                texto = elem.text.strip()
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:
                        return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 3 falló: {e}")
        
        st.error("❌ No se pudo encontrar el valor numérico de CANTIDAD PASOS")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando cantidad de pasos: {str(e)}")
        return None

def buscar_cantidad_pasos_alternativo(driver):
    """Búsqueda alternativa y más agresiva para CANTIDAD PASOS - FUNCIÓN PROBADA"""
    try:
        # Buscar todos los elementos que contengan números
        all_elements = driver.find_elements(By.XPATH, "//*[text()]")
        
        for elem in all_elements:
            texto = elem.text.strip()
            # Buscar patrones numéricos que parezcan cantidades (4,452, 4452, etc.)
            if (texto and 
                any(char.isdigit() for char in texto) and
                3 <= len(texto) <= 10 and
                not any(word in texto.upper() for word in ['$', 'TOTAL', 'VALOR', 'PAGAR', 'COMERCIO'])):
                
                # Verificar si es un número con formato de cantidad (puede tener comas)
                clean_text = texto.replace(',', '').replace('.', '')
                if clean_text.isdigit():
                    num_value = int(clean_text)
                    # Verificar si está en un rango razonable para cantidad de pasos
                    if 100 <= num_value <= 999999:
                        st.success(f"✅ Valor alternativo encontrado: {texto}")
                        return texto
        
        return None
    except Exception as e:
        st.warning(f"⚠️ Búsqueda alternativa falló: {e}")
        return None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - USANDO FUNCIONES PROBADAS"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiMDA5OGE5MTQtNjQ0MC00ZTdjLWJmNDItNGZhYmQxOWE5ZTk3IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        # 1. Navegar al reporte
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        # 2. Tomar screenshot inicial
        driver.save_screenshot("powerbi_inicial.png")
        
        # 3. Hacer clic en la conciliación específica
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None
        
        # 4. Esperar a que cargue la selección
        time.sleep(3)
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        # 5. Buscar tarjeta "VALOR A PAGAR A COMERCIO" y extraer valor
        valor_texto = find_valor_a_pagar_comercio_card(driver)
        
        # 6. Buscar "CANTIDAD PASOS" 
        cantidad_pasos_texto = find_cantidad_pasos_card(driver)
        
        # Si no se encuentra, intentar una búsqueda más agresiva
        if not cantidad_pasos_texto or cantidad_pasos_texto == 'No encontrado':
            st.warning("🔄 Intentando búsqueda alternativa para CANTIDAD PASOS...")
            cantidad_pasos_texto = buscar_cantidad_pasos_alternativo(driver)
        
        # 9. Tomar screenshot final
        driver.save_screenshot("powerbi_final.png")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto or 'No encontrado',
            'screenshots': {
                'inicial': 'powerbi_inicial.png',
                'seleccion': 'powerbi_despues_seleccion.png',
                'final': 'powerbi_final.png'
            }
        }
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None
    finally:
        if driver:
            driver.quit()

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL =====

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

# ===== FUNCIONES DE COMPARACIÓN =====

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            # Limpiar el string
            cleaned = currency_string.strip()
            
            # Remover símbolos de moneda y espacios
            cleaned = cleaned.replace('$', '').replace(' ', '')
            
            # Manejar formato colombiano (puntos para miles, coma para decimales)
            if '.' in cleaned and ',' in cleaned:
                # Formato: 1.000.000,00 -> quitar puntos, cambiar coma por punto
                cleaned = cleaned.replace('.', '').replace(',', '.')
            elif '.' in cleaned and cleaned.count('.') > 1:
                # Formato: 1.000.000 -> quitar todos los puntos
                cleaned = cleaned.replace('.', '')
            elif ',' in cleaned:
                # Formato: 1,000,000 o 1,000,000.00
                if cleaned.count(',') == 2 and '.' in cleaned:
                    # Formato internacional: 1,000,000.00
                    cleaned = cleaned.replace(',', '')
                elif cleaned.count(',') == 1:
                    # Podría ser decimal: 1000,50
                    cleaned = cleaned.replace(',', '.')
                else:
                    # Múltiples comas como separadores de miles
                    cleaned = cleaned.replace(',', '')
            
            # Convertir a float
            return float(cleaned) if cleaned else 0.0
            
        return float(currency_string)
        
    except Exception as e:
        st.error(f"❌ Error convirtiendo moneda: '{currency_string}' - {e}")
        return 0.0

def comparar_valores(valor_excel, valor_power_bi, pasos_excel, pasos_power_bi):
    """
    Compara los valores y determina si coinciden
    """
    # Convertir valores de Power BI a números
    try:
        if isinstance(valor_power_bi, str):
            valor_power_bi_num = convert_currency_to_float(valor_power_bi)
        else:
            valor_power_bi_num = float(valor_power_bi)
            
        if isinstance(pasos_power_bi, str):
            # Limpiar string de pasos (quitar comas, puntos, etc.)
            pasos_limpio = re.sub(r'[^\d]', '', pasos_power_bi)
            pasos_power_bi_num = int(pasos_limpio) if pasos_limpio else 0
        else:
            pasos_power_bi_num = int(pasos_power_bi) if pasos_power_bi else 0
    except Exception as e:
        st.error(f"❌ Error convirtiendo valores de Power BI: {e}")
        return False, False, 0, 0
    
    diferencia_valor = abs(valor_excel - valor_power_bi_num)
    diferencia_pasos = abs(pasos_excel - pasos_power_bi_num)
    
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
    - Comparar con Power BI
    
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
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):", value="2025-10-12")
        
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
                        resultados = extract_powerbi_data(fecha_validacion)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_power_bi_texto = resultados['valor_texto']
                        cantidad_pasos_texto = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        
                        # ========== SECCIÓN RESULTADOS POWER BI ==========
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("💰 VALOR A PAGAR A COMERCIO", valor_power_bi_texto)
                        
                        with col2:
                            st.metric("👣 CANTIDAD DE PASOS", cantidad_pasos_texto)
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN COMPARACIÓN ==========
                        st.markdown("### 📊 Resultado de la Validación")
                        
                        # Comparar resultados
                        coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                            valor_a_pagar, valor_power_bi_texto, numero_pasos, cantidad_pasos_texto
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
                                valor_power_bi_texto, 
                                str(cantidad_pasos_texto)
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
                            
                            # Convertir valores para mostrar
                            valor_power_bi_num = convert_currency_to_float(valor_power_bi_texto)
                            if isinstance(cantidad_pasos_texto, str):
                                pasos_power_bi_num = int(re.sub(r'[^\d]', '', cantidad_pasos_texto)) if re.sub(r'[^\d]', '', cantidad_pasos_texto) else 0
                            else:
                                pasos_power_bi_num = int(cantidad_pasos_texto)
                            
                            resumen_data = []
                            
                            # Valor a Pagar
                            resumen_data.append({
                                'Concepto': 'Valor a Pagar',
                                'Excel': f"${valor_a_pagar:,.2f}".replace(",", "."),
                                'Power BI': f"${valor_power_bi_num:,.2f}".replace(",", "."),
                                'Diferencia': f"${dif_valor:,.2f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden_valor else '❌ No coincide'
                            })
                            
                            # Número de Pasos
                            resumen_data.append({
                                'Concepto': 'Número de Pasos',
                                'Excel': f"{numero_pasos:,}",
                                'Power BI': f"{pasos_power_bi_num:,}",
                                'Diferencia': f"{dif_pasos:,}",
                                'Estado': '✅ Coincide' if coinciden_pasos else '❌ No coincide'
                            })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                            
                            # Screenshots
                            st.markdown("#### 📸 Capturas del Proceso Power BI")
                            col1, col2, col3 = st.columns(3)
                            
                            screenshots = resultados.get('screenshots', {})
                            
                            if 'inicial' in screenshots and os.path.exists(screenshots['inicial']):
                                with col1:
                                    st.image(screenshots['inicial'], caption="Vista Inicial", use_column_width=True)
                            
                            if 'seleccion' in screenshots and os.path.exists(screenshots['seleccion']):
                                with col2:
                                    st.image(screenshots['seleccion'], caption="Tras Selección", use_column_width=True)
                            
                            if 'final' in screenshots and os.path.exists(screenshots['final']):
                                with col3:
                                    st.image(screenshots['final'], caption="Vista Final", use_column_width=True)
                        
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
           - Extrae "VALOR A PAGAR A COMERCIO" y "CANTIDAD PASOS"
        
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
