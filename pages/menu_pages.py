import streamlit as st


def build_navigation():
    home  = st.Page(title="Inicio",            page="pages/home.py",
                    icon=":material/home:",    default=True)
    page1 = st.Page(title="Mesa Alimentacion", page="pages/page1.py",
                    icon=":material/check_circle:")
    page2 = st.Page(title="Reporte de Tiempos Perdidos", page="pages/page2.py",
                    icon=":material/check_circle:")
    page3 = st.Page(title="Page3", page="pages/page3.py",
                    icon=":material/check_circle:")

    return st.navigation(
        {
            "":[home],
            "Lista de Chequeo": [page1],
            "Tiempos Perdidos": [page2], },
        position="top"
    )
    
    # return st.navigation(
    #     {"Alimentacion": [page1, page2],
    #      "Descortezado": [page3], },
    #     position="top"
    # )
