import streamlit as st

st.set_page_config(page_title="Rent a Hobby", layout="wide", initial_sidebar_state="expanded")

# Sidebar-Navigation (immer sichtbar)
page = st.sidebar.radio("", ["Startseite", "Hobbys", "Kontakt"])

# Hauptinhalt
st.title("Rent a Hobby")

if page == "Startseite":
    st.write("Willkommen bei Rent a Hobby! Entdecke neue Hobbys und miete sie für kurze Zeit.")
elif page == "Hobbys":
    st.write("Hier findest du alle verfügbaren Hobbys zum Mieten.")
elif page == "Kontakt":
    st.write("Kontaktiere uns für weitere Informationen.")

