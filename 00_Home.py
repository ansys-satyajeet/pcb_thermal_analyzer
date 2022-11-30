import streamlit as st

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")

st.title('🌡️PCB Thermal Analyzer ')


st.markdown('''PCB Thermal Analyzer is a fully automated web application to predict temperature distribution 
on a Printed Circuit Board (PCB) and identify potential hot spots. To run the thermal simulations, the app uses the power of
[***Icepak in Ansys Electronics Desktop***](https://www.ansys.com/products/electronics/ansys-icepak) in the background.

The app is built using the PyAEDT, an open source Python framework, for automating workflows in 
Ansys Electronics Desktop. PyAEDT is a part of the greater PyAnsys framework which can be used 
to automate simulation workflows using Ansys tools.
''')

link_icepak_alh = '[Training Content for Icepak in Ansys Electronics Desktop](https://jam8.sapjam.com/groups/tEgwx0OWCayodWqIHHv9Rq/overview_page/oZ7yhpQk0C07VL9LsaKTe3)'
link_pyaedt = '[Read more about PyAEDT](https://aedt.docs.pyansys.com/)'
st.markdown(link_icepak_alh, unsafe_allow_html=True)
st.markdown(link_pyaedt, unsafe_allow_html=True)

st.markdown('''Please find the system requirements below to run PCB Thermal Analyzer App.

**Requirements**:
* A licensed copy of Ansys Electronics Desktop 2022 R2 and above must be installed on the machine running the simulation.
* Python 3.10 must be installed on the machine running the simulation.
''')

st.warning('⚠️ Do not refresh the app (browser tab) during a run session. The session state will be reset and all progress will be lost!')

# st.markdown('''
# A brief description of the sections are written below. **👇Please read before proceeding.👇** 
# ''')

# with st.expander("📝 Modify Boundary Conditions"):
#     st.write("""
#         ℹ️ This section imports an IDF board file and parses it to extract the component names in the file.
#         The names are then written out to a CSV file which can be edited.
#         The columns of the CSV file include:
#         * __Include__: Include or filter out component (_0 = filter; 1 = include_)
#         * __Package_name__: Name of package
#         * __Part_name__: Part name
#         * __Instance_name__: Instance name
#         * __Placement__: Location of component (TOP or BOTTOM) w.r.t. board
#         * __BC_type__: Boundary Condition (_block; network; hollow_)
#         * __Power [W]__: Power on component in Watts
#         * __Theta_jb [C/W]__: Junction to board thermal resistance in 2-R network in C/W
#         * __Theta_jc [C/W]__: Junction to case thermal resistance in 2-R network in C/W
#         * __Monitor Point__: Create temperature monitor point at object center (_0 = no; 1 = yes_)
#         * __Material__: Name of material of component
        
#         ⚠️ Modify the CSV file and save it as CSV file only. Other file types such *.txt, *.xlxs (Excel) are not allowed.

#         ⚠️ Do not change the name of the CSV file.
#     """)
    
# with st.expander("🖥️ Simulate in AEDT Icepak"):
#     st.write("""
#         ℹ️ This section lets the user use the boundary conditions CSV file to assign boundary conditions to the 
#         imported IDF file. The users can setup the problem and simulate it in a much simpler manner.
#     """)

# with st.expander("📊 Postprocessing"):
#     st.write("""
#         ℹ️ If the simulation was executed and completed in the Simulation stage, this stage allows the users 
#         to generate reports, postprocessed images and much more.
#     """)