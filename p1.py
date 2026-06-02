# pip install plotly
import streamlit as st
import pandas as pd
import plotly.express as px

st.title("📊 Mi Primer Dashboard EVM")

# Simulamos la carga de un proyecto
datos = {
    "Métrica": ["Valor Planificado (PV)", "Costo Real (AC)", "Valor Ganado (EV)"],
    "Monto ($)": [15000, 16500, 14000]
}
df = pd.DataFrame(datos)

# Crear gráfico interactivo
figura = px.bar(
    df, 
    x="Métrica", 
    y="Monto ($)", 
    color="Métrica",
    title="Estado de Salud Financiera del Proyecto",
    text_auto=True
)

# Mostrar en Streamlit
st.plotly_chart(figura, use_container_width=True)
st.info("💡 En la web, los usuarios podrán pasar el ratón sobre las barras para ver los valores exactos.")