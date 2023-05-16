import streamlit as st

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")
st.title('📧Contact', anchor='contact_page')

report_issues = '[Report issues here](https://github.com/ansys-satyajeet/pcb_thermal_analyzer/issues)'
st.markdown(report_issues, unsafe_allow_html=True)

