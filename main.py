from convert import *
import streamlit as st

st.title("Convert your xls roster to ics!")

input_file = st.file_uploader("Upload your file here!", type=['xls'])
if input_file is not None:
    data = convert_file(input_file)
    st.download_button(
        label="Completed! Download ics here!",
        data=data,
        file_name="roster.ics",
        mime="text"
    )