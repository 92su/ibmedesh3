import streamlit as st
#from multiapp import MultiApp
from st_vizzu import *
from multiapp import MultiApp
from ipyvizzu import Chart,Data, Config, Style,DisplayTarget
from streamlit.components.v1 import html
import pandas as pd
from apps import home,data
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

#st.set_page_config(layout='wide',initial_sidebar_state='expanded')

st.markdown("""
<style>
    [data-testid=stSidebar] {
        background-color: #3F034A;
        
    }
</style>
""", unsafe_allow_html=True)

with open ('style.css') as f:
    st.markdown(f'<style>{f.read()}</style>',unsafe_allow_html=True)

app = MultiApp()

st.markdown("""
# Research Paper Publication Analysis

""")

# Add all your application here
app.add_app("Home", home.app)
app.add_app("Data", data.app)


app.run()
