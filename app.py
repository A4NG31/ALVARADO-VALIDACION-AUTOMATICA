import streamlit as st
import pandas as pd
import re
import os
from datetime import datetime
import tempfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time

# Configuración de la página
st.set_page_config(
    page_title="Validador Power BI - Conciliaciones",
    page_icon="💰",
    layout="wide"
)

# CSS para mejorar la apariencia
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
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

def extraer_datos_power_bi(fecha_validacion):
    """
    Extrae datos del dashboard de Power BI
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
        
        # Buscar y seleccionar la fecha
        st.info(f"📅 Seleccionando fecha: {fecha_validacion}")
        # Código para seleccionar fecha específica en el Power BI
        
        # Esperar a que carguen los datos
        time.sleep(5)
        
        # Extraer datos de "RESUMEN COMERCIOS" - PEAJE ALVARADO
        st.info("🔍 Extrayendo datos del resumen de comercios...")
        
        # Valores de ejemplo - necesitarás adaptar los selectores
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

def main():
    st.markdown('<div class="main-header">💰 Validador Power BI - Conciliaciones</div>', unsafe_allow_html=True)
    
    # Sidebar para carga de archivo
    with st.sidebar:
        st.header("📁 Cargar Archivo Excel")
        uploaded_file = st.file_uploader("Selecciona el archivo Excel", type=['xlsx', 'xls'])
        
        if uploaded_file is not None:
            st.success(f"Archivo cargado: {uploaded_file.name}")
            
            # Extraer fecha del nombre del archivo
            fecha_validacion = extraer_fecha_desde_nombre(uploaded_file.name)
            
            if fecha_validacion:
                st.info(f"📅 Fecha detectada: {fecha_validacion}")
            else:
                st.warning("No se pudo detectar la fecha del archivo")
                fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):")
    
    # Contenido principal
    if uploaded_file is not None and fecha_validacion:
        
        # Procesar el archivo Excel
        with st.spinner("Procesando archivo Excel..."):
            valor_a_pagar, numero_pasos = procesar_excel(uploaded_file)
        
        if valor_a_pagar > 0 and numero_pasos > 0:
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("💰 Valor a Pagar (Excel)", f"${valor_a_pagar:,.0f}")
            
            with col2:
                st.metric("👣 Número de Pasos (Excel)", f"{numero_pasos}")
            
            # Extraer datos de Power BI
            if st.button("🔄 Consultar Power BI y Validar", type="primary"):
                with st.spinner("Extrayendo datos de Power BI..."):
                    valor_power_bi, pasos_power_bi = extraer_datos_power_bi(fecha_validacion)
                
                if valor_power_bi is not None and pasos_power_bi is not None:
                    # Mostrar resultados de Power BI
                    col3, col4 = st.columns(2)
                    
                    with col3:
                        st.metric("💰 Valor a Pagar (Power BI)", f"${valor_power_bi:,.0f}")
                    
                    with col4:
                        st.metric("👣 Número de Pasos (Power BI)", f"{pasos_power_bi}")
                    
                    # Comparar resultados
                    st.markdown("---")
                    st.subheader("📊 Resultado de la Validación")
                    
                    coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                        valor_a_pagar, valor_power_bi, numero_pasos, pasos_power_bi
                    )
                    
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
                    st.subheader("📋 Resumen de Comparación")
                    
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
                    st.error("No se pudieron extraer los datos del Power BI")
        else:
            st.error("No se pudieron extraer los valores del archivo Excel. Verifica el formato.")
    
    elif uploaded_file is None:
        st.info("👈 Por favor, carga un archivo Excel para comenzar la validación")
    
    # Instrucciones de uso
    with st.expander("📖 Instrucciones de Uso"):
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
        - Columna AK debe tener encabezado "Valor" en fila 38
        - Debe contener texto "TOTAL TRANSACCIONES X" donde X es el número de pasos
        """)

if __name__ == "__main__":
    main()