import streamlit as st

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")

st.title('🌡️PCB Thermal Analyzer ')


st.markdown('''PCB Thermal Analyzer is a fully automated web application to predict temperature distribution on a 
Printed Circuit Board (PCB) and identify potential hot spots. To run the thermal simulations, the app uses the power 
of [***Icepak in Ansys Electronics Desktop***](https://www.ansys.com/products/electronics/ansys-icepak) in the 
background.

The app is built using the PyAEDT, an open source Python framework, for automating workflows in 
Ansys Electronics Desktop. PyAEDT is a part of the greater PyAnsys framework which can be used 
to automate simulation workflows using Ansys tools.
''')

link_icepak_alh = '[Training Content for Icepak in Ansys Electronics Desktop](' \
                  'https://jam8.sapjam.com/groups/tEgwx0OWCayodWqIHHv9Rq/overview_page/oZ7yhpQk0C07VL9LsaKTe3)'
link_pyaedt = '[Read more about PyAEDT](https://aedt.docs.pyansys.com/)'
st.markdown(link_icepak_alh, unsafe_allow_html=True)
st.markdown(link_pyaedt, unsafe_allow_html=True)

st.markdown('''Please find the system requirements below to run PCB Thermal Analyzer App.

**Requirements**:
* A licensed copy of Ansys Electronics Desktop 2022R2 or above must be installed on the machine running the simulation.
* Python 3.10 must be installed on the machine running the simulation.
''')

st.warning('⚠️ Do not refresh the app (browser tab) during a run session. The session state will be reset and all '
           'progress will be lost!')
