import os
import smtplib
import streamlit as st
from docx import Document
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def enviar_correo_nube(remitente, password, destinatario, proyecto, ruta_docx):
    """Despacha el archivo Word nativo usando la lógica clásica de MIME."""
    # 1. Crear el sobre principal
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 Reporte Ejecutivo desde la Nube - {proyecto}"
    msg['From'] = remitente
    msg['To'] = destinatario
    
    # 2. Agregar el cuerpo del texto
    cuerpo = f"Estimados,\n\nSe adjunta el reporte oficial de estado (EVM) para el proyecto {proyecto}, generado desde nuestro servidor en la nube.\n\nAtentamente,\nPM Copilot System"
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    # 3. Leer y adjuntar el archivo Word (.docx)
    with open(ruta_docx, 'rb') as f:
        # Usamos subtype="docx" o la extensión estándar para archivos de Word
        adjunto = MIMEApplication(f.read(), _subtype="docx")
        adjunto.add_header('Content-Disposition', 'attachment', filename=os.path.basename(ruta_docx))
        msg.attach(adjunto)
        
    # 4. Enviar el correo
    print(f"🔄 Conectando a SMTP para enviar a {destinatario}...")
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
    print("✅ ¡Correo despachado con éxito!")

# --- PRUEBA EN STREAMLIT ---
st.title("☁️ Prueba de Correo en la Nube")
if st.button("Enviar Correo de Prueba"):
    # Genera un Word temporal
    doc = Document()
    doc.add_heading('Prueba de Despliegue', 0)
    archivo_prueba = "Reporte_Prueba.docx"
    doc.save(archivo_prueba)
    
    try:
        # Reemplaza con tus datos reales para probar en internet
        enviar_correo_nube("norbilllumpo346@gmail.com", "rmnm zpqj jqpc xbzn", "norbilllumpo346@gmail.com", "Alfa", archivo_prueba)
        st.success("✅ ¡Correo despachado con éxito desde la nube!")
    except Exception as e:
        st.error(f"Error al enviar: {e}")
        
    os.remove(archivo_prueba)
