import os
import smtplib
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Importación del nuevo SDK unificado de Google
from google import genai

# ==========================================
# 1. FUNCIÓN DE CORREO CLOUD-SAFE (Envia .docx usando MIME)
# ==========================================
def despachar_correo_docx(remitente, password, destinatario, proyecto, ruta_docx):
    """Despacha el archivo Word nativo, compatible con servidores Linux/Cloud."""
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 Reporte de Control Semanal - {proyecto}"
    msg['From'] = remitente
    msg['To'] = destinatario
    
    cuerpo = f"Estimados,\n\nAdjunto sírvase encontrar el reporte automatizado (.docx) con el avance físico y financiero del proyecto {proyecto}.\n\nAtentamente,\nPM Copilot en la Nube"
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    with open(ruta_docx, "rb") as f:
        adjunto = MIMEApplication(f.read(), _subtype="docx")
        adjunto.add_header('Content-Disposition', 'attachment', filename=os.path.basename(ruta_docx))
        msg.attach(adjunto)
        
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)

# ==========================================
# 2. CONFIGURACIÓN DE LA INTERFAZ WEB
# ==========================================
st.set_page_config(page_title="PM Copilot Cloud", page_icon="☁️", layout="wide")

st.title("☁️ Portafolio Inteligente & Centro de Automatización")
st.caption("Despliegue Público: Dashboards Interactivos, Consultor IA y Despacho de Correos")

# ==========================================
# 3. SEGURIDAD Y CREDENCIALES (BARRA LATERAL)
# ==========================================
st.sidebar.header("🔑 1. Credenciales IA")
api_key = st.sidebar.text_input("Ingresa tu Gemini API Key:", type="password")

client = None
if api_key:
    # Instanciamos el cliente con el nuevo SDK
    client = genai.Client(api_key=api_key)
    st.sidebar.success("✅ Conexión con Gemini lista.")
else:
    st.sidebar.warning("⚠️ Introduce la API Key para habilitar la IA.")

st.sidebar.header("✉️ 2. Servidor SMTP")
correo_emisor = st.sidebar.text_input("Tu Correo (Gmail):", value="tu_correo_aqui@gmail.com")
clave_app = st.sidebar.text_input("Contraseña de Aplicación:", type="password")

st.sidebar.header("📂 3. Base de Datos")
archivo_cargado = st.sidebar.file_uploader("Sube la matriz (Excel .xlsx):", type=["xlsx"])

# ==========================================
# 4. PROCESAMIENTO MÁSTER
# ==========================================
if archivo_cargado is not None:
    lector_excel = pd.ExcelFile(archivo_cargado)
    hoja_seleccionada = st.sidebar.selectbox("📄 Proyecto para Análisis de Dashboard e IA:", lector_excel.sheet_names)
    
    df_proyecto = pd.read_excel(archivo_cargado, sheet_name=hoja_seleccionada)
    
    # Inyección de métricas EVM
    if 'Costo_Planificado' in df_proyecto.columns and 'Avance_Trabajador_%' in df_proyecto.columns:
        df_proyecto['EV'] = df_proyecto['Costo_Planificado'] * (df_proyecto['Avance_Trabajador_%'] / 100)
        df_proyecto['CPI'] = df_proyecto.apply(
            lambda r: r['EV'] / r['Costo_Real_Actual'] if r.get('Costo_Real_Actual', 0) > 0 else 1.0, axis=1
        )
    
    # Cálculos globales
    total_pv = df_proyecto['Costo_Planificado'].sum() if 'Costo_Planificado' in df_proyecto.columns else 0
    total_ac = df_proyecto.get('Costo_Real_Actual', pd.Series([0])).sum()
    total_ev = df_proyecto.get('EV', pd.Series([0])).sum()
    cpi_global = total_ev / total_ac if total_ac > 0 else 1.0

    # ==========================================
    # 5. SISTEMA DE PESTAÑAS (MÓDULOS INTEGRADOS)
    # ==========================================
    tab1, tab2, tab3 = st.tabs(["📈 Módulo 5: Dashboard Interactivo", "🧠 Módulo 3: Consultor Predictivo IA", "📬 Módulo 4: Despacho Masivo (Cloud)"])
    
    # --- TAB 1: DASHBOARD INTERACTIVO ---
    with tab1:
        st.subheader(f"Dashboard de Salud Financiera: {hoja_seleccionada}")
        
        # Tarjetas de métricas rápidas
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Valor Planificado (PV)", f"${total_pv:,.2f}")
        col2.metric("Costo Real (AC)", f"${total_ac:,.2f}")
        col3.metric("Valor Ganado (EV)", f"${total_ev:,.2f}")
        col4.metric("CPI Global", f"{cpi_global:.2f}", delta="Sobrecosto" if cpi_global < 1 else "Saludable", delta_color="inverse")
        
        # Gráfico Plotly interactivo
        df_plot = pd.DataFrame({
            "Métrica": ["PV (Planificado)", "AC (Real)", "EV (Ganado)"],
            "Monto": [total_pv, total_ac, total_ev]
        })
        fig = px.bar(df_plot, x="Métrica", y="Monto", color="Métrica", title="Comparativa EVM", text_auto='.2s')
        
        st.plotly_chart(fig, use_container_width=True, theme=None)
        st.dataframe(df_proyecto, use_container_width=True)

    # --- TAB 2: CONSULTOR IA ---
    with tab2:
        st.subheader("Consultor Predictivo de Riesgos")
        opcion_analisis = st.selectbox(
            "Enfoque del reporte:",
            ["Análisis de Desvíos Críticos", "Proyección Financiera (EAC/CPI)", "Plan de Mitigación"]
        )
        
        if st.button("🤖 Generar Informe con IA"):
            if not client:
                st.error("Configura tu API Key en la barra lateral.")
            else:
                with st.spinner("Redactando informe ejecutivo..."):
                    prompt = f"Analiza esta matriz de tareas de proyecto: \n{df_proyecto.to_string(index=False)}\nGenera un diagnóstico enfocado en: {opcion_analisis}. Sé conciso y usa viñetas."
                    try:
                        # Lógica del nuevo SDK (google-genai)
                        response = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=prompt
                        )
                        st.session_state['reporte_ia'] = response.text
                    except Exception as e:
                        st.error(f"Error de API: {e}")

        if 'reporte_ia' in st.session_state:
            st.info("📋 Reporte Emitido:")
            st.markdown(st.session_state['reporte_ia'])

    # --- TAB 3: DESPACHO MASIVO CLOUD-SAFE ---
    with tab3:
        st.subheader("Motor de Compilación y Distribución en la Nube")
        st.write("Este módulo lee todas las hojas del Excel, genera documentos en formato Word (.docx) y los envía por correo.")
        
        PLANTILLA = "Plantilla_Reporte.docx"
        if not os.path.exists(PLANTILLA):
            st.error(f"🚨 Recuerda subir el archivo '{PLANTILLA}' a tu repositorio de GitHub para habilitar esta función.")
        else:
            # Diccionario de Distribución
            lista_distribucion = {
                "INDUMENTECH_Software": correo_emisor,
                "SULFTECH_Mina": correo_emisor,
                "INTERCONNECTIUM_Redes": correo_emisor,
                "MINASAFE_Soporte": correo_emisor
            }
            
            if st.button("🚀 Ejecutar Despacho Masivo", use_container_width=True):
                if not clave_app:
                    st.error("⚠️ Falta la Contraseña de Aplicación de tu correo en la barra lateral.")
                else:
                    status_bar = st.status("Iniciando automatización...", expanded=True)
                    dict_proyectos = pd.read_excel(archivo_cargado, sheet_name=None)
                    
                    for hoja, df_hoja in dict_proyectos.items():
                        status_bar.write(f"⚙️ Procesando: **{hoja}**")
                        
                        pv_h = df_hoja['Costo_Planificado'].sum() if 'Costo_Planificado' in df_hoja.columns else 0
                        ac_h = df_hoja.get('Costo_Real_Actual', pd.Series([0])).sum()
                        ev_h = (df_hoja['Costo_Planificado'] * (df_hoja['Avance_Trabajador_%'] / 100)).sum() if 'Costo_Planificado' in df_hoja.columns and 'Avance_Trabajador_%' in df_hoja.columns else 0
                        
                        criticas_h = df_hoja[df_hoja['Avance_Trabajador_%'] < 50] if 'Avance_Trabajador_%' in df_hoja.columns else pd.DataFrame()
                        alertas_txt = "\n".join([f"- {r.get('Actividad', 'Tarea')} ({r.get('Avance_Trabajador_%', 0)}%)" for _, r in criticas_h.iterrows()]) if not criticas_h.empty else "✅ Cronograma al día."
                        
                        datos_plantilla = {
                            "{{PROYECTO}}": hoja,
                            "{{FECHA}}": datetime.now().strftime('%Y-%m-%d'),
                            "{{AVANCE_GLOBAL}}": f"{df_hoja['Avance_Trabajador_%'].mean():.1f}" if 'Avance_Trabajador_%' in df_hoja.columns else "0.0",
                            "{{COSTO_PLAN}}": f"{pv_h:,.2f}",
                            "{{COSTO_REAL}}": f"{ac_h:,.2f}",
                            "{{CPI_GLOBAL}}": f"{(ev_h/ac_h if ac_h > 0 else 1):.2f}",
                            "{{TAREAS_ALERTAS}}": alertas_txt
                        }
                        
                        # Generación de Word nativo
                        doc_template = Document(PLANTILLA)
                        for p in doc_template.paragraphs:
                            for marcador, valor in datos_plantilla.items():
                                p.text = p.text.replace(marcador, str(valor))
                                
                        marca_t = datetime.now().strftime("%H%M%S")
                        nombre_docx = os.path.abspath(f"Reporte_{hoja}_{marca_t}.docx")
                        doc_template.save(nombre_docx)
                        
                        correo_destino = lista_distribucion.get(hoja, correo_emisor)
                        status_bar.write(f"📧 Despachando email a: `{correo_destino}`")
                        try:
                            despachar_correo_docx(correo_emisor, clave_app, correo_destino, hoja, nombre_docx)
                        except Exception as e:
                            status_bar.write(f"❌ Error enviando correo: {e}")
                            
                        # Limpieza del documento temporal en el servidor
                        if os.path.exists(nombre_docx):
                            os.remove(nombre_docx)
                    
                    status_bar.update(label="🏆 ¡Despacho Masivo Completado!", state="complete", expanded=False)
                    st.balloons()

else:
    st.info("ℹ️ Sube un archivo Excel maestro para iniciar el despliegue del portafolio en la nube.")