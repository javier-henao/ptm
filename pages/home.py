import streamlit as st

col1, col2 = st.columns([3, 1], vertical_alignment='center')
with col1:
    st.title("PTM: Lista de Chequeo")
with col2:
    st.subheader("Planta Yumbo")

st.divider()

st.markdown("""
### Bienvenid@s

Esta aplicación permite registrar y monitorear las variables operativas
de los equipos de la planta en tiempo real.

**Módulos disponibles:**
- 🌿 **Alimentación** — Mesa de alimentación y variables
- 🕒 **Tiempos Perdidos** — Reporte de tiempos perdidos de turno
- 🪵 **Descortezado** — Equipos de descortezado
""")

st.info("Selecciona un módulo en el menú superior para comenzar.", icon=":material/info:")