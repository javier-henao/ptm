import streamlit as st
from pages.menu_pages import build_navigation

st.set_page_config(
    page_title="PTM: Lista de Chequeo",
    page_icon=":material/forest:"
)

page = build_navigation()
page.run()