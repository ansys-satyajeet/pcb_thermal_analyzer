import os
import re
import pandas as pd
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from ctypes import windll
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")
st.title('📝Create Boundary Conditions File')

# Fix blur issue in tkinter window panels
windll.shcore.SetProcessDpiAwareness(1)

if 'idf' not in st.session_state:
    st.session_state['idf'] = False

if 'idf_file' not in st.session_state:
    st.session_state['idf_file'] = False

if 'dataframe' not in st.session_state:
    st.session_state['dataframe'] = False

if 'generate_bc_csv' not in st.session_state:
    st.session_state['generate_bc_csv'] = False

if 'idf_csv_file' not in st.session_state:
    st.session_state['idf_csv_file'] = False

if 'matcsv' not in st.session_state:
    st.session_state['matcsv'] = False

if 'mat_csvfile' not in st.session_state:
    st.session_state['mat_csvfile'] = False

if 'mat_dataframe' not in st.session_state:
    st.session_state['mat_dataframe'] = False

if 'workdir' not in st.session_state:
    st.session_state['workdir'] = False

c1, c2 = st.columns([3, 1])
c1.markdown(f'''**Select working directory:**''')
workdir_button = c2.button('Select Folder')
if workdir_button:
    st.session_state['workdir'] = True
    root0 = tk.Tk()
    root0.attributes("-topmost", True)
    root0.withdraw()
    try:
        workdir = filedialog.askdirectory(initialdir=os.getcwd())
        st.session_state['workdir'] = workdir
    except RuntimeWarning:
        pass

if st.session_state['workdir']:
    os.chdir(st.session_state['workdir'])
ph0 = st.empty()
if st.session_state['workdir']:
    ph0.markdown(f'''📁 ```{os.path.abspath(st.session_state['workdir'])}```''')
    # ph0.markdown(f'''**Current Working Directory:** ```{st.session_state['workdir']}```''')
else:
    ph0.markdown(f'''📁 ```{os.getcwd()}```''')

# Read board file from Windows Explorer dialog box
#
col01, col02, col03 = st.columns([2, 1, 1])
col01.markdown(f'''**Select IDF Board file:**''')
idf_type = col02.selectbox('Select IDF Board file type:', ('*.emn', '*.bdf'), label_visibility='collapsed')
idf_button = col03.button('Select Board File')
if idf_button:
    st.session_state['idf'] = True
    root1 = tk.Tk()
    root1.attributes("-topmost", True)
    root1.withdraw()
    try:
        if idf_type == '*.emn':
            files = filedialog.askopenfilenames(parent=root1,
                                                filetypes=[('EMN File', '*.emn')])
            idf_file = os.path.basename(files[0])
            st.session_state['idf_file'] = idf_file
        if idf_type == '*.bdf':
            files = filedialog.askopenfilenames(parent=root1,
                                                filetypes=[('BDF File', '*.bdf')])
            idf_file = os.path.basename(files[0])
            st.session_state['idf_file'] = idf_file
    except RuntimeWarning:
        pass

# Placeholder for IDF file input information
ph1 = st.empty()
if st.session_state['idf_file']:
    ph1.markdown(f'''📝 ```{os.path.abspath(st.session_state['idf_file'])}```''')
else:
    ph1.markdown(f'''⚠️*No IDF file selected.*''')

# Read board file from Windows Explorer dialog box
#
include_matfile = st.checkbox(f'''**Read Materials as CSV File?**''')
if include_matfile:
    col04, col05 = st.columns([3, 1])
    col04.markdown(f'''**Please select materials CSV file:**''')
    matfile_button = col05.button('Select CSV File')
    if matfile_button:
        st.session_state['matcsv'] = True
        root2 = tk.Tk()
        root2.attributes("-topmost", True)
        root2.withdraw()
        try:
            files = filedialog.askopenfilenames(parent=root2,
                                                filetypes=[('Microsoft Excel Comma Separated Values File', '*.csv')])
            mat_csvfile = os.path.basename(files[0])
            st.session_state['mat_csvfile'] = mat_csvfile
        except RuntimeWarning:
            pass
    ph2 = st.empty()
    if st.session_state['mat_csvfile']:
        ph2.markdown(f'''📝 ```{os.path.abspath(st.session_state['mat_csvfile'])}```''')
    else:
        ph2.markdown(f'''⚠️*No Materials CSV file selected.*''')

if st.session_state['idf_file']:
    filename_no_ext = os.path.splitext(st.session_state['idf_file'])[0]

    # Board file and Library file
    if idf_type == '*.emn':
        board_file = os.path.abspath(filename_no_ext + '.emn')
        lib_file = os.path.abspath(filename_no_ext + '.emp')
    else:
        board_file = os.path.abspath(filename_no_ext + '.bdf')
        lib_file = os.path.abspath(filename_no_ext + '.ldf')

    # Generate CSV of boundary conditions
    st.markdown('**Generate boundary conditions table as CSV file**')
    st.session_state['generate_bc_csv'] = st.button('Generate')

    if st.session_state['generate_bc_csv']:
        # Board components
        components = []
        component_names = []
        component_placement = []
        with open(board_file) as emn:
            for line in emn:
                if line.strip() == ".PLACEMENT":
                    break
            for line in emn:
                if line.strip() == ".END_PLACEMENT":
                    break
                components.append(line.strip())
        component_names = components[::2]
        component_placement = components[1::2]

        designator_list = []
        for i in range(len(component_names)):
            if component_names[i].find('\"\"') > 0:
                n = re.findall(r'[^"\s]\S*|".+?', component_names[i])
            else:
                n = re.findall(r'[^"\s]\S*|".+?"', component_names[i])
            if n[1] == '""':
                n[1] = 'NOPARTNAME'
            p = list(filter(None, component_placement[i].split(' ')))
            p = p[4]
            n.append(p)
            designator_list.append(n)

        for i in range(len(designator_list)):
            for j in range(len(designator_list[i])):
                designator_list[i][j] = designator_list[i][j].replace(",", "_")

        # Export designator list as csv
        st.session_state['idf_csv_file'] = filename_no_ext + '_bcs.csv'
        with open(st.session_state['idf_csv_file'], 'w') as f:
            header = 'Include,Package_Name,Part_Name,Instance_Name,Placement,BC_Type,Power [W],R_jb [C/W],R_jc [C/W],' \
                     'Monitor_Point,Material\n'
            f.write(header)
            for i in designator_list:
                line = 'YES,' + ','.join(map(str, i)) + ',block,0,0,0,NO,Ceramic_material\n'
                f.write(line)

        # Add reference designator type and height information
        df = pd.read_csv(st.session_state['idf_csv_file'])
        refdes = df.iloc[:, 3]
        destype = []
        for i in refdes:
            if re.match(r'^U\d', i):
                ic = re.match(r'^U\d', i)
                destype.append('INTEGRATED CIRCUIT')
            elif re.match(r'^R\d', i):
                res = re.match(r'^R\d', i)
                destype.append('RESISTOR')
            elif re.match(r'^C\d', i):
                cap = re.match(r'^C\d', i)
                destype.append('CAPACITOR')
            elif re.match(r'^L\d', i):
                ind = re.match(r'^L\d', i)
                destype.append('INDUCTOR')
            else:
                destype.append('MISC')

        df.insert(loc=4, column='Designator_Type', value=destype)

        comp_hts = []
        with open(lib_file) as emp:
            for line in emp:
                if line.startswith('.ELECTRICAL'):
                    x = next(emp)
                    if x.find('\"\"') > 0:
                        n = re.findall(r'[^"\s]\S*|".+?', x)
                    else:
                        n = re.findall(r'[^"\s]\S*|".+?"', x)
                    if n[1] == '""':
                        n[1] = 'NOPARTNAME'
                    comp_hts.append(n)

        for i in comp_hts:
            i.remove('THOU')

        for i in range(len(comp_hts)):
            for j in range(len(comp_hts[i])):
                comp_hts[i][j] = comp_hts[i][j].replace(",", "_")

        part_names = df.iloc[:, 2]
        component_height = [0] * df.shape[0]
        for i in comp_hts:
            for j in range(len(part_names)):
                if i[1] == part_names[j]:
                    component_height[j] = float(i[2]) * 0.0254

        df.insert(loc=5, column='Height [mm]', value=component_height)

        try:
            df.to_csv(st.session_state['idf_csv_file'], index=False)
        except RuntimeWarning:
            st.write('Something went wrong!')
        st.write('👍 Boundary Conditions CSV File Generated.')

if st.session_state['idf_csv_file']:
    st.info('ℹ️ To export the table as CSV, right click on any cell in table, then Export -> CSV Export.')
    st.session_state['dataframe'] = pd.read_csv(st.session_state['idf_csv_file'])
    include_dropdownlist = ('YES', 'NO')
    bc_dropdownlist = ('block', 'network', 'hollow')

    gb = GridOptionsBuilder.from_dataframe(st.session_state['dataframe'])
    gb.configure_default_column(editable=True)
    # gb.configure_grid_options(domLayout='normal')
    gb.configure_column('Include', editable=True, cellEditor='agSelectCellEditor',
                        cellEditorParams={'values': include_dropdownlist}, singleClickEdit=True)
    gb.configure_column('Package_Name', editable=False)
    gb.configure_column('Part_Name', editable=False)
    gb.configure_column('Instance_Name', editable=False)
    gb.configure_column('Placement', editable=False)
    gb.configure_column('BC_Type', editable=True, cellEditor='agSelectCellEditor',
                        cellEditorParams={'values': bc_dropdownlist}, singleClickEdit=True)
    gb.configure_column('Power [W]', editable=True, type=["numericColumn", "numberColumnFilter", "customNumericFormat"],
                        precision=2)
    gb.configure_column('R_jb [C/W]', editable=True,
                        type=["numericColumn", "numberColumnFilter", "customNumericFormat"], precision=2)
    gb.configure_column('R_jc [C/W]', editable=True,
                        type=["numericColumn", "numberColumnFilter", "customNumericFormat"], precision=2)
    gb.configure_column('Monitor_Point', editable=True, cellEditor='agSelectCellEditor',
                        cellEditorParams={'values': include_dropdownlist}, singleClickEdit=True)
    if st.session_state['mat_csvfile']:
        st.session_state['mat_dataframe'] = pd.read_csv(st.session_state['mat_csvfile'])
        mat_dropdownlist = tuple(st.session_state['mat_dataframe'].iloc[:, 0])
        gb.configure_column('Material', editable=True, cellEditor='agSelectCellEditor',
                            cellEditorParams={'values': mat_dropdownlist}, singleClickEdit=True)
    grid_options = gb.build()
    grid_height = 400
    grid_response = AgGrid(
        st.session_state['dataframe'],
        gridOptions=grid_options,
        height=grid_height,
        width='100%',
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.GRID_CHANGED
    )

    df = grid_response['data']
    selected = grid_response['selected_rows']
    selected_df = pd.DataFrame(selected).apply(pd.to_numeric, errors='coerce')
