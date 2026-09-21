import streamlit as st
import pandas as pd
import io
import re

# Configuración de la página
st.set_page_config(page_title="Automatización J3 - Estación de Servicio", layout="wide")

st.title("⛽ Automatización de Integración Contable para J3")
st.markdown("""
Esta aplicación reemplaza el proceso manual de la *Hoja formulada para archivo final*. 
Carga los reportes de ventas y el modelo de integración para generar el archivo plano de J3.
""")

# 1. Zona de carga de archivos
st.header("1. Carga de Archivos Input")
col1, col2 = st.columns(2)

with col1:
    ventas_files = st.file_uploader(
        "Sube los Reportes de Ventas (Ej: ajustado_Repporte_Ventas_...)", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )

with col2:
    integracion_file = st.file_uploader(
        "Sube el Modelo de Integración (Ej: ajustado_modelo_integracion_...)", 
        type=["xlsx", "xls"]
    )

st.header("2. Ingreso Manual de Pagos (QR y Datáfono)")
st.info("Dado que los valores exactos de QR y Datáfono vienen de talonarios físicos, puedes hacer los ajustes aquí.")
# En una versión más avanzada, esto podría ser una tabla editable para ingresar números de ticket y su medio de pago real.

if ventas_files and integracion_file:
    if st.button("🚀 Procesar y Generar Archivo Final", type="primary"):
        try:
            with st.spinner("Procesando archivos..."):
                
                # --- PASO A: CONSOLIDAR VENTAS ---
                df_ventas_list = [pd.read_excel(f) for f in ventas_files]
                df_ventas = pd.concat(df_ventas_list, ignore_index=True)
                
                # --- PASO B: LIMPIAR INTEGRACIÓN CONTABLE ---
                df_integ = pd.read_excel(integracion_file)
                
                # Filtrar solo facturas electrónicas (FEI). 
                # TODO: Cambia 'Prefijo' por el nombre real de tu columna en el Excel
                if 'Prefijo' in df_integ.columns:
                    df_integ = df_integ[df_integ['Prefijo'].str.contains('FEI', na=False)]
                
                # Extraer número de ticket (simula el Texto en Columnas de la columna AF/O/R)
                # TODO: Cambia 'Columna_Texto' por el nombre real donde viene el ticket encriptado
                def extraer_ticket(texto):
                    # Busca la primera secuencia de números que represente el ticket
                    match = re.search(r'\d+', str(texto))
                    return match.group(0) if match else texto
                
                # Ejemplo de aplicación (Descomenta y ajusta el nombre de la columna):
                # df_integ['Numero_Ticket'] = df_integ['Columna_Texto'].apply(extraer_ticket)
                
                # --- PASO C: CRUZAR DATOS (SIMULACIÓN HOJA 4) ---
                # TODO: Reemplaza 'Numero_Ticket' por la llave común en ambos archivos
                # df_final = pd.merge(df_ventas, df_integ, on='Numero_Ticket', how='left')
                
                # Para efectos del código base, asumiremos que df_final es la mezcla de ambos.
                # Aquí se filtrarían los #N/D (compras sin cliente en J3)
                df_final = df_ventas.copy() # Placeholder hasta definir nombres de columnas
                
                # --- PASO D: MAPEO DE FORMAS DE PAGO ---
                # Transformar los textos a códigos para J3 (QR, Datáfono, Efectivo)
                # TODO: Cambia 'Forma_Pago' por el nombre de la columna (ej. Columna M)
                mapeo_pagos = {
                    'QR': 1,           # Código ejemplo de J3 para QR
                    'EFECTIVO': 2,     # Código ejemplo para Efectivo
                    'DATAFONO': 3      # Código ejemplo para Datáfono
                }
                if 'Forma_Pago' in df_final.columns:
                    df_final['Codigo_J3'] = df_final['Forma_Pago'].str.upper().map(mapeo_pagos)
                
                # --- PASO E: FORMATEO FINAL COMO TEXTO PLANO ---
                # Eliminar columnas sobrantes (ejemplo)
                columnas_a_eliminar = ['ColumnaInnecesaria1', 'ColumnaInnecesaria2']
                df_final = df_final.drop(columns=[c for c in columnas_a_eliminar if c in df_final.columns])
                
                # Convertir todo a texto plano para evitar fórmulas en J3
                df_final = df_final.astype(str)
                df_final = df_final.replace('nan', '')

            st.success("✅ Archivos procesados con éxito. Listo para importar en J3.")
            st.dataframe(df_final.head(10)) # Vista previa

            # --- PASO F: DESCARGA DEL ARCHIVO EXCEL ---
            buffer = io.BytesIO()
            # Usamos xlsxwriter. En J3 tal vez debas guardar como libro 97-2003 (.xls) manualmente 
            # tras descargarlo si el sistema es sumamente antiguo, pero xlsxwriter asegura que los datos van limpios.
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Hoja1')
            
            st.download_button(
                label="📥 Descargar Archivo Final para J3",
                data=buffer.getvalue(),
                file_name="Calculado_final_J3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Ocurrió un error al procesar los archivos. Verifica que el formato coincida. Detalle: {e}")