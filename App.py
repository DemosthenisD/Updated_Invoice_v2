import streamlit as st
import toml
from streamlit_extras.switch_page_button import switch_page


# Load secrets
user = st.secrets["credentials"]["USER"]
password = st.secrets["credentials"]["PASSWORD"]


st.sidebar.page_link("App.py" , label="Log-IN", icon="🔐")

tab1, tab2 = st.tabs(['Login', 'Register'])
with tab1:
    username = st.text_input("Username")
    password_input = st.text_input("Password", type="password")
    submit = st.button("Submit")
    if submit:
       if username == user and password_input == password:
            st.session_state.username       = username
            st.session_state.password_input = password_input
            switch_page('home')
       else:
            st.error("Please Provide valid UserID")    
with tab2:
    st.warning("Registration is currently not allowed")


