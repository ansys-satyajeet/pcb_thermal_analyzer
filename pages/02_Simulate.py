import os
import re
import csv
import shutil
import pyaedt
import signal
import numpy as np
import pandas as pd
from ctypes import windll
import streamlit as st
import tkinter as tk
from tkinter import filedialog as fd

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")
st.title('🖥️Simulate')


# Function definitions

# Function to create a forced convection problem setup with default entries
def forced_convection_setup(setup_name, flow_regime, turb_model='ZeroEquation'):
    """ Default settings for forced convection problem
        Parameters
        ----------
        setup_name: str
            Name of setup. Use only alphanumeric characters (letter, numbers and underscores)
        flow_regime: str
            Laminar or Turbulent
        turb_model: str, optional
            default = 'ZeroEquation'
            'TwoEquation' uses 'Enhanced Realizable k-epsilon' turbulence model
    """
    setup = ipk.create_setup(setup_name)
    setup.props['Enabled'] = True
    flow_regime = flow_regime.casefold()
    turb_model = turb_model.casefold()
    if flow_regime == 'laminar':
        setup.props['Flow Regime'] = 'Laminar'
    else:
        setup.props['Flow Regime'] = 'Turbulent'
        if turb_model == 'zeroequation':
            setup.props['Turbulent Model Eqn'] = 'ZeroEquation'
        else:
            setup.props['Turbulent Model Eqn'] = 'EnhancedRealizableTwoEquation'
    setup.props['Include Temperature'] = True
    setup.props['Include Flow'] = True
    setup.props['Include Gravity'] = False
    setup.props['Include Solar'] = False
    setup.props['Solution Initialization - X Velocity'] = "0m_per_sec"
    setup.props['Solution Initialization - Y Velocity'] = "0m_per_sec"
    setup.props['Solution Initialization - Z Velocity'] = "0m_per_sec"
    setup.props['Solution Initialization - Use Model Based Flow Initialization'] = False
    setup.props['Convergence Criteria - Flow'] = '1e-4'
    setup.props['Convergence Criteria - Energy'] = '1e-10'
    setup.props['IsEnabled'] = False
    setup.props['Radiation Model'] = 'Off'
    setup.props['Under-relaxation - Pressure'] = '0.3'
    setup.props['Under-relaxation - Momentum'] = '0.7'
    setup.props['Under-relaxation - Temperature'] = '1'
    setup.props['Discretization Scheme - Pressure'] = 'Standard'
    setup.props['Discretization Scheme - Momentum'] = 'First'
    setup.props['Discretization Scheme - Temperature'] = 'First'
    setup.props['Secondary Gradient'] = False
    setup.props['Sequential Solve of Flow and Energy Equations'] = True
    setup.props['Convergence Criteria - Max Iterations'] = 300
    setup.update()


# Function to create a natural convection problem setup with default entries
def natural_convection_setup(setup_name, gravity_dir, flow_regime, turb_model='ZeroEquation', ambient_temp=20):
    """ Default settings for natural convection problem
        Parameters
        ----------
        setup_name: str
            Name of setup. Use only alphanumeric characters (letter, numbers and underscores)
        gravity_dir: str
            Direction of gravity: -x, +x, -y, +y, -z, +z
        flow_regime: str
            Laminar or Turbulent
        turb_model: str, optional
            default = 'ZeroEquation'
            'TwoEquation' uses 'Enhanced Realizable k-epsilon' turbulence model
        ambient_temp: float, optional
            default = 20
    """
    setup = ipk.create_setup(setup_name)
    setup.props['Enabled'] = True
    flow_regime = flow_regime.casefold()
    turb_model = turb_model.casefold()
    if flow_regime == 'laminar':
        setup.props['Flow Regime'] = 'Laminar'
    else:
        setup.props['Flow Regime'] = 'Turbulent'
        if turb_model == 'zeroequation':
            setup.props['Turbulent Model Eqn'] = 'ZeroEquation'
        else:
            setup.props['Turbulent Model Eqn'] = 'EnhancedRealizableTwoEquation'
    setup.props['Include Temperature'] = True
    setup.props['Include Flow'] = True
    setup.props['Include Gravity'] = True
    setup.props['Include Solar'] = False
    setup.props['Solution Initialization - X Velocity'] = "0m_per_sec"
    setup.props['Solution Initialization - Y Velocity'] = "0m_per_sec"
    setup.props['Solution Initialization - Z Velocity'] = "0m_per_sec"
    gravity_dir = gravity_dir.casefold()
    ambient_temp = str(ambient_temp) + 'cel'
    if gravity_dir == "-x":
        ipk.apply_icepak_settings(ambienttemp=ambient_temp, gravityDir=0)
        ipk.modeler.edit_region_dimensions([250, 50, 50, 50, 200, 200])
        setup.props['Solution Initialization - X Velocity'] = "0.00098m_per_sec"
    elif gravity_dir == "-y":
        ipk.apply_icepak_settings(ambienttemp=ambient_temp, gravityDir=1)
        ipk.modeler.edit_region_dimensions([50, 50, 250, 50, 200, 200])
        setup.props['Solution Initialization - Y Velocity'] = "0.00098m_per_sec"
    elif gravity_dir == "-z":
        ipk.apply_icepak_settings(ambienttemp=ambient_temp, gravityDir=2)
        ipk.modeler.edit_region_dimensions([50, 50, 50, 50, 250, 50])
        setup.props['Solution Initialization - Z Velocity'] = "0.00098m_per_sec"
    elif gravity_dir == "+x":
        ipk.apply_icepak_settings(ambienttemp=ambient_temp, gravityDir=3)
        ipk.modeler.edit_region_dimensions([50, 250, 50, 50, 200, 200])
        setup.props['Solution Initialization - X Velocity'] = "-0.00098m_per_sec"
    elif gravity_dir == "+y":
        ipk.apply_icepak_settings(ambienttemp=ambient_temp, gravityDir=4)
        ipk.modeler.edit_region_dimensions([50, 50, 50, 250, 200, 200])
        setup.props['Solution Initialization - Y Velocity'] = "-0.00098m_per_sec"
    else:
        ipk.apply_icepak_settings(ambienttemp=ambient_temp, gravityDir=5)
        ipk.modeler.edit_region_dimensions([50, 50, 50, 50, 50, 250])
        setup.props['Solution Initialization - Z Velocity'] = "-0.00098m_per_sec"
    setup.props['Solution Initialization - Use Model Based Flow Initialization'] = False
    setup.props['Convergence Criteria - Flow'] = '1e-4'
    setup.props['Convergence Criteria - Energy'] = '1e-10'
    setup.props['IsEnabled'] = True
    setup.props['Radiation Model'] = 'Discrete Ordinates Model'
    setup.props['Flow Iteration Per Radiation Iteration'] = '10'
    setup.props['ThetaDivision'] = 2
    setup.props['PhiDivision'] = 2
    setup.props['ThetaPixels'] = 2
    setup.props['PhiPixels'] = 2
    setup.props['Convergence Criteria - Discrete Ordinates'] = '1e-6'
    setup.props['Under-relaxation - Pressure'] = '0.7'
    setup.props['Under-relaxation - Momentum'] = '0.3'
    setup.props['Under-relaxation - Temperature'] = '1'
    setup.props['Discretization Scheme - Pressure'] = 'Standard'
    setup.props['Discretization Scheme - Momentum'] = 'First'
    setup.props['Discretization Scheme - Temperature'] = 'First'
    setup.props['Discretization Scheme - Discrete Ordinates'] = 'First'
    setup.props['Secondary Gradient'] = False
    setup.props['Linear Solver Type - Pressure'] = 'V'
    setup.props['Linear Solver Type - Momentum'] = 'flex'
    setup.props['Linear Solver Type - Temperature'] = 'F'
    setup.props['Sequential Solve of Flow and Energy Equations'] = False
    setup.props['Convergence Criteria - Max Iterations'] = 500
    setup.update()


# Function to create opening boundary condition
def assign_opening_boundary(name, face_id, flow_type,
                            xvel="0m_per_sec", yvel="0m_per_sec", zvel="0m_per_sec",
                            pressure="AmbientPressure", temperature="AmbientTemp"):
    """ Function to create opening boundary condition
        Parameters
        ----------
        name: str
            name of opening boundary condition, e.g., inlet/outlet
        face_id: int
            face ID of opening
        flow_type: str
            velocity or pressure
        xvel: str, optional
            velocity in x-direction
        yvel: str, optional
            velocity in y-direction
        zvel: str, optional
            velocity in z-direction
        pressure: str, optional
            pressure at opening boundary
        temperature: str, optional
            temperature at opening boundary
    """
    props = {"Faces": [face_id]}
    if flow_type == 'velocity':
        props['Inlet Type'] = "Velocity"
        props['Static Pressure'] = pressure
        props['X Velocity'] = xvel
        props['Y Velocity'] = yvel
        props['Z Velocity'] = zvel
        props['Temperature'] = temperature
    else:
        props['Inlet Type'] = "Pressure"
        props['Total Pressure'] = pressure
        props['Temperature'] = temperature
    bound = pyaedt.modules.Boundary.BoundaryObject(ipk, name, props, 'Opening')
    if bound.create():
        ipk.boundaries.append(bound)
        return bound


# Function to add slack
def add_slack(box_name, minx, maxx, miny, maxy, minz, maxz):
    obj_ref = ipk.modeler.get_object_from_name(box_name)
    obj_ref.bottom_face_x.move_with_offset(minx)
    obj_ref.bottom_face_y.move_with_offset(miny)
    obj_ref.bottom_face_z.move_with_offset(minz)
    obj_ref.top_face_x.move_with_offset(maxx)
    obj_ref.top_face_y.move_with_offset(maxy)
    obj_ref.top_face_z.move_with_offset(maxz)
    return None


# Function to remove aedt files and project folders
def cleanup_files(proj_name):
    proj_path = os.path.join(os.getcwd(), proj_name)
    proj_name_no_ext = os.path.splitext(proj_name)[0]
    if os.path.exists(proj_path):
        os.remove(proj_path)
    if os.path.exists(os.path.join(os.getcwd(), proj_name + ".lock")):
        os.remove(os.path.join(os.getcwd(), proj_name + ".lock"))
    # Delete aedt results folder
    if os.path.exists(os.path.join(os.getcwd(), proj_name_no_ext + ".aedtresults")):
        try:
            shutil.rmtree(os.path.join(os.getcwd(), proj_name_no_ext + ".aedtresults"))
        except RuntimeError:
            print('Error deleting aedtresults directory')
    # Delete pyaedt folder
    if os.path.exists(os.path.join(os.getcwd(), proj_name_no_ext + ".pyaedt")):
        try:
            shutil.rmtree(os.path.join(os.getcwd(), proj_name_no_ext + ".pyaedt"))
        except RuntimeError:
            print('Error deleting pyaedt directory')


def quit_aedt():
    ipk.save_project()
    pid = desktop.aedt_process_id
    os.kill(pid, signal.SIGTERM)
    file_list = os.listdir(os.getcwd())
    for item in file_list:
        if item.endswith('.lock'):
            os.remove(item)


# Fix blur issue in tkinter window panels
windll.shcore.SetProcessDpiAwareness(1)

if 'idf' not in st.session_state:
    st.session_state['idf'] = False
if 'ecad' not in st.session_state:
    st.session_state['ecad'] = False
if 'bcs' not in st.session_state:
    st.session_state['bcs'] = False
if 'mats' not in st.session_state:
    st.session_state['mats'] = False
if 'idf_type' not in st.session_state:
    st.session_state['idf_type'] = False
if 'ecad_type' not in st.session_state:
    st.session_state['ecad_type'] = False

if 'idf_file' not in st.session_state:
    st.session_state['idf_file'] = False
if 'ecad_file' not in st.session_state:
    st.session_state['ecad_file'] = False
if 'bc_filename' not in st.session_state:
    st.session_state['bc_filename'] = False
if 'materials_filename' not in st.session_state:
    st.session_state['materials_filename'] = False
if 'workdir' not in st.session_state:
    st.session_state['workdir'] = False
if 'pid' not in st.session_state:
    st.session_state['pid'] = False

# Get working directory from Windows Explorer dialog box
c1, c2 = st.columns([3, 1])
c1.write('Select working directory:')
workdir_button = c2.button('Select Folder')
if workdir_button:
    st.session_state['workdir'] = True
    root0 = tk.Tk()
    root0.attributes("-topmost", True)
    root0.withdraw()
    try:
        workdir = fd.askdirectory(parent=root0, initialdir=os.getcwd(), title='Select Folder')
        st.session_state['workdir'] = workdir
    except RuntimeWarning:
        pass

if st.session_state['workdir']:
    os.chdir(st.session_state['workdir'])

# Read IDF board file from Windows Explorer dialog box
col01, col02, col03 = st.columns([2, 1, 1])
col01.write('Select IDF Board file type:')
st.session_state['idf_type'] = col02.selectbox('Select IDF Board file type:', ('*.emn', '*.bdf'),
                                               label_visibility='collapsed')
idf_button = col03.button('Select Board File')
if idf_button:
    st.session_state['idf'] = True
    root1 = tk.Tk()
    root1.attributes("-topmost", True)
    root1.withdraw()
    try:
        if st.session_state['idf_type'] == '*.emn':
            files = fd.askopenfilenames(parent=root1, initialdir=os.getcwd(), filetypes=[('EMN File', '*.emn')])
            idf_file = os.path.basename(files[0])
            st.session_state['idf_file'] = idf_file
        if st.session_state['idf_type'] == '*.bdf':
            files = fd.askopenfilenames(parent=root1, initialdir=os.getcwd(), filetypes=[('BDF File', '*.bdf')])
            idf_file = os.path.basename(files[0])
            st.session_state['idf_file'] = idf_file
    except RuntimeWarning:
        pass

# Read ECAD file from Windows Explorer dialog box
col04, col05, col06 = st.columns([2, 1, 1])
col04.write('Please select ECAD file type:')
st.session_state['ecad_type'] = col05.selectbox('Select ECAD type:', ('EDB Folder', 'ODB++ File', 'BRD File'),
                                                label_visibility='collapsed')
ecad_button = col06.button('Select ECAD')
if ecad_button:
    st.session_state['ecad'] = True
    root2 = tk.Tk()
    root2.attributes("-topmost", True)
    root2.withdraw()
    try:
        if st.session_state['ecad_type'] == 'EDB Folder':
            ecad_file = fd.askdirectory(parent=root2, initialdir=os.getcwd())
            st.session_state['ecad_file'] = ecad_file
        elif st.session_state['ecad_type'] == 'ODB++ File':
            ecad_file = fd.askopenfilename(parent=root2, initialdir=os.getcwd(), filetypes=[('TGZ File', '*.tgz')])
            st.session_state['ecad_file'] = ecad_file
        elif st.session_state['ecad_type'] == 'BRD File':
            ecad_file = fd.askopenfilename(parent=root2, initialdir=os.getcwd(), filetypes=[('BRD File', '*.brd')])
            st.session_state['ecad_file'] = ecad_file
        else:
            st.error('Something went wrong!')
    except RuntimeWarning:
        pass

# Read Boundary Conditions CSV file from Windows Explorer dialog box
col07, col08 = st.columns([3, 1])
col07.write('Please select boundary conditions CSV file:')
bc_file_button = col08.button('Select BC CSV file')
if bc_file_button:
    st.session_state['bcs'] = True
    bcs = True
    root3 = tk.Tk()
    root3.attributes("-topmost", True)
    root3.withdraw()
    try:
        bc_file = fd.askopenfilenames(parent=root3, initialdir=os.getcwd(),
                                      filetypes=[('Microsoft Excel Comma Separated Values File', '*.csv')])
        bc_filename = os.path.basename(bc_file[0])
        st.session_state['bc_filename'] = bc_filename
    except RuntimeWarning:
        pass

# Read material file from Windows Explorer dialog box
include_matfile = st.checkbox('Read Materials as CSV File?')
if include_matfile:
    col09, col10 = st.columns([3, 1])
    col09.write('Please select materials CSV file:')
    mats_button = col10.button('Select Materials CSV')
    if mats_button:
        st.session_state['mats'] = True
        root4 = tk.Tk()
        root4.attributes("-topmost", True)
        root4.withdraw()
        try:
            materials_file = fd.askopenfilenames(parent=root4, initialdir=os.getcwd(), filetypes=[
                ('Microsoft Excel Comma Separated Values File', '*.csv')])
            materials_filename = os.path.basename(materials_file[0])
            st.session_state['materials_filename'] = materials_filename
        except RuntimeWarning:
            pass

with st.expander('List of Inputs'):
    if st.session_state['workdir']:
        st.markdown(f'''**Selected working directory:** ```{os.path.abspath(st.session_state['workdir'])}```''')
    else:
        st.markdown(f'''**Selected working directory** ```{os.getcwd()}```''')
    if st.session_state['idf']:
        st.markdown(f'''**Selected IDF file:** ```{os.path.abspath(st.session_state['idf_file'])}```''')
    else:
        st.markdown(f'''**Selected IDF file:** None''')

    if st.session_state['ecad_file']:
        st.markdown(f'''**Selected ECAD file:** ```{st.session_state['ecad_file']}```''')
    else:
        st.markdown(f'''**Selected ECAD file:** None''')

    if st.session_state['bc_filename']:
        st.markdown(
            f'''**Selected Boundary Conditions CSV file:** ```{os.path.abspath(st.session_state['bc_filename'])}```''')
    else:
        st.markdown(f'''**Selected Boundary Conditions CSV file:** None''')

    if include_matfile:
        if st.session_state['materials_filename']:
            st.markdown(
                f'''**Selected Materials CSV file:** ```{os.path.abspath(st.session_state['materials_filename'])}```''')
        else:
            st.markdown(f'''**Selected Materials CSV file:** None''')

# Setup Options
st.markdown('---')
st.markdown('**Setup Options**')

all_points = st.checkbox('Create points at board side face centers for all components',
                         help='Can increase model setup time!')

if all_points:
    st.write(':information_source: Temperatures at the created points will be written out to a table/file.')

delete_filtered = st.checkbox('Delete filtered objects?', help='Deleted objects cannot be recovered.')

if delete_filtered:
    st.write(':information_source: Filtered objects will be deleted.')

# Solution Settings
st.markdown('---')
st.markdown('**Solution Settings**')

conv_type = st.selectbox('Select convection mode:', ('Forced', 'Natural'))

conv_cond = False
air_temp = str(20.0)
vel_dir = '+X'
vel = 0.0
gravity_direction = '+Z'

if conv_type == 'Forced':
    col11, col12 = st.columns(2)
    vel = col11.text_input('Velocity Magnitude [m/s]:')
    vel_dir = col12.selectbox('Direction:', ('+X', '-X', '+Y', '-Y', '+Z', '-Z'))
    Tin = st.text_input('Inlet Temperature [C]:')
    air_temp = Tin
    conv_cond = True

if conv_type == 'Natural':
    gravity_direction = st.selectbox('Direction of gravity:', ('+X', '-X', '+Y', '-Y', '+Z', '-Z'))
    Tamb = st.text_input('Ambient Temperature [C]:')
    air_temp = Tamb
    conv_cond = True

# Mesh Settings
st.markdown('---')
st.markdown('**Mesh Settings**')
mesh_fidelity = st.select_slider('Select mesh resolution:', options=['Coarse', 'Medium', 'Fine'])

# Solve Settings
st.markdown('---')
st.markdown('**Solve Settings**')
col13, col14 = st.columns(2)
num_cores = col13.number_input('Number of Processors', min_value=1, max_value=128, value=1, step=1, format='%d')
mode = col14.radio('Mode:', ('Graphical', 'Non-Graphical'))

project_name = st.text_input('Enter Project Name:',
                             help='Only letters (A-Z,a-z), numbers (0-9) and underscores are allowed.')

aedt_version = st.selectbox('Select AEDT release:', ('2023 R1', '2023 R2'))

analyze_setup = st.checkbox('Setup problem and proceed to solve')
if analyze_setup:
    sim_button_text = '**Simulate**'
else:
    sim_button_text = '**Setup Only**'
setup_analyze_button = st.button(sim_button_text)
placeholder = st.empty()
analysis_complete = False
analysis_setup = 'Icepak_Analysis'

# Main Code Execution
#
if st.session_state['idf_file'] and st.session_state['ecad_file'] and st.session_state['bc_filename'] \
        and conv_cond and air_temp and project_name:
    placeholder.success('The setup is ready for analysis', icon="✅")
    filename_no_ext = os.path.splitext(st.session_state['idf_file'])[0]

    # Board File and Library File
    if st.session_state['idf_type'] == '*.emn':
        board_filename = os.path.abspath(filename_no_ext + '.emn')
        lib_filename = os.path.abspath(filename_no_ext + '.emp')
    else:
        board_filename = os.path.abspath(filename_no_ext + '.bdf')
        lib_filename = os.path.abspath(filename_no_ext + '.ldf')

    ecad_foldername = st.session_state['ecad_file']
    bc_filename = st.session_state['bc_filename']
    materials_filename = st.session_state['materials_filename']

    # Launch new AEDT Icepak session
    project_name = project_name + '.aedt'
    project_path = os.path.join(os.getcwd(), project_name)
    if setup_analyze_button:
        aedt_release = re.sub(' R', '.', aedt_version)
        cleanup_files(project_name)
        placeholder.info('AEDT Icepak session in progress...', icon="🏃🏽")

        if mode == 'Graphical':
            non_graphical_mode = False
        else:
            non_graphical_mode = True

        # Start AEDT Desktop session
        desktop = pyaedt.Desktop(aedt_release, non_graphical=non_graphical_mode)

        # Insert ECAD Import code
        ecad_file_path = st.session_state['ecad_file']
        ecad_type = st.session_state['ecad_type']
        ecad_file_name = os.path.basename(ecad_file_path)
        ecad_file_name_no_ext = os.path.splitext(ecad_file_name)[0]

        ecad_project_name = ecad_file_name_no_ext + '.aedt'
        ecad_project_path = os.path.join(os.getcwd(), ecad_project_name)
        ecad_project_name_no_ext = os.path.splitext(ecad_project_name)[0]

        cleanup_files(ecad_project_name)

        h3d = pyaedt.Hfss3dLayout()
        if ecad_type == 'EDB Folder':
            h3d.import_edb(ecad_file_path)
        if ecad_type == 'ODB++ File':
            h3d.import_odb(ecad_file_path)
        if ecad_type == 'BRD File':
            h3d.import_brd(ecad_file_path)

        h3d.save_project()

        # Delete empty project
        project_list = desktop.project_list()
        for i in project_list:
            if i != ecad_file_name_no_ext:
                desktop.odesktop.DeleteProject(i)

        # ECAD design name from HFSS 3D Layout
        ecad_design = h3d.design_list[0]

        # Get name of outline polygon
        outline_poly = []
        for key in h3d.modeler.polygons.keys():
            if h3d.modeler.polygons[key].placement_layer == 'Outline':
                outline_poly.append(key)

        # Insert Icepak design and rename project to user-specified project name
        ipk = pyaedt.Icepak()
        ipk.save_project()
        ipk.oproject.Rename(os.path.join(ipk.project_path, project_name), True)

        # Create PCB object in Icepak from HFSS 3D Layout
        ipk.create_pcb_from_3dlayout(component_name=ecad_file_name_no_ext,
                                     project_name=None,
                                     design_name=ecad_design,
                                     resolution=3,
                                     extent_type='Polygon',
                                     outline_polygon=outline_poly[0],
                                     close_linked_project_after_import=False)
        # Import IDF file into Icepak
        ipk.import_idf(board_filename)

        # Fit all and save
        ipk.modeler.fit_all()
        ipk.save_project()
        ipk.autosave_disable()

        # List of boundary conditions
        list_bcs = ipk.odesign.GetChildObject('Thermal').GetChildNames()

        # Delete all boundary conditions
        omodule = ipk.odesign.GetModule("BoundarySetup")
        if list_bcs:
            for i in list_bcs:
                omodule.DeleteBoundaries([i])

        # Delete all points
        for i in ipk.modeler.points:
            ipk.modeler.points[i].delete()

        # Import Modified CSV file
        rows = []
        with open(bc_filename, 'r') as csvFile:
            csvReader = csv.reader(csvFile)
            fields = next(csvReader)
            for row in csvReader:
                rows.append(row)

        # Read material properties file (if provided)
        if st.session_state['materials_filename']:
            rows_mat = []
            with open(materials_filename, 'r', encoding='utf-8-sig') as matFile:
                csvReader = csv.reader(matFile)
                fields_mat = next(csvReader)
                for row in csvReader:
                    rows_mat.append(row)

            # Create materials in AEDT
            for i in rows_mat:
                mat = ipk.materials.add_material(i[0])
                mat.thermal_conductivity = i[1]
                mat.update()

        # Delete empty part numbers/NOREFDES instances
        for i in ipk.modeler.solid_bodies:
            if i.startswith('idf_mech'):
                ipk.modeler.delete(i)

        # Remove any gap between board and components
        top_components = []
        bottom_components = []
        for i in range(len(rows)):
            if rows[i][3] != 'NOREFDES':
                block_name = re.sub(r"\W", "_", rows[i][3])
                if rows[i][6] == 'TOP':
                    top_components.append(block_name)
                if rows[i][6] == 'BOTTOM':
                    bottom_components.append(block_name)

        tc_z = ipk.modeler.get_object_from_name(top_components[0]).bottom_face_z.center[2]
        bc_z = ipk.modeler.get_object_from_name(bottom_components[0]).top_face_z.center[2]

        pcb = ipk.modeler.primitives.user_defined_component_names
        pcb_layers = sorted(ipk.modeler.get_3d_component_object_list(pcb[0]))
        top_layer_z_bound = ipk.modeler.get_object_from_name(pcb_layers[0]).top_face_z.center[2]
        bottom_layer_z_bound = ipk.modeler.get_object_from_name(pcb_layers[-1]).bottom_face_z.center[2]

        move_top = tc_z - top_layer_z_bound
        if move_top > 0:
            ipk.modeler.move(objid=top_components, vector=[0, 0, -move_top])
        else:
            ipk.modeler.move(objid=top_components, vector=[0, 0, move_top])

        move_bottom = bc_z - bottom_layer_z_bound
        if move_bottom < 0:
            ipk.modeler.move(objid=bottom_components, vector=[0, 0, -move_bottom])
        else:
            ipk.modeler.move(objid=bottom_components, vector=[0, 0, move_bottom])

        # Create dictionary of points at board side of all components
        points_dict = {}
        for i in range(len(rows)):
            if rows[i][3] != 'NOREFDES':
                block_name = re.sub(r"\W", "_", rows[i][3])
                block_handle = ipk.modeler.get_object_from_name(block_name)
                pcb_top_layer = ipk.modeler.get_object_from_name(pcb_layers[0])
                pcb_bottom_layer = ipk.modeler.get_object_from_name(pcb_layers[-1])
                if block_handle.get_touching_faces(pcb_top_layer):
                    block_board_side = block_handle.get_touching_faces(pcb_top_layer)
                else:
                    block_board_side = block_handle.get_touching_faces(pcb_bottom_layer)
                point_name = 'point_' + block_name
                if all_points:
                    mon_point = ipk.modeler.primitives.get_face_center(block_board_side[0].id)
                    points_dict[point_name] = mon_point

        # Delete filtered objects or make them non-model
        for i in range(len(rows)):
            if rows[i][0] == 'NO':
                if rows[i][3] != 'NOREFDES':
                    block_name = re.sub(r"\W", "_", rows[i][3])
                    block_handle = ipk.modeler.get_object_from_name(block_name)
                    if delete_filtered:
                        ipk.modeler.delete(block_handle.name)
                    else:
                        block_handle.model = False

        # Priority assignments based on volume of objects
        obj_dict = {}
        for i in ipk.modeler.solid_bodies:
            if i != 'Region':
                obj_dict[i] = ipk.modeler.get_object_from_name(i).volume
        vol_sorted_objs = sorted(obj_dict.items(), key=lambda x: x[1], reverse=True)
        vol_sorted_obj_list = []
        for i in vol_sorted_objs:
            vol_sorted_obj_list.append(i[0])

        priority_num = 2
        args = ["NAME:UpdatePriorityListData"]
        for i in vol_sorted_obj_list:
            if i != 'Region':
                prio = [
                    "NAME:PriorityListParameters",
                    "EntityType:=", "Object",
                    "EntityList:=", i,
                    "PriorityNumber:=", priority_num,
                    "PriorityListType:=", "3D"
                ]
                args.append(prio)
                priority_num = priority_num + 1
        ipk.modeler.oeditor.UpdatePriorityList(args)

        # Clear Desktop messages
        desktop.clear_messages()

        # Save project
        ipk.save_project()

        # Make board that comes with the IDF file as non-model object
        board_handle = ipk.modeler.get_object_from_name('IDF_BoardOutline')
        board_handle.model = False

        # Clear Desktop messages
        desktop.clear_messages()

        # List all model objects in design
        model_objects = ipk.modeler.model_objects
        model_objects.remove('Region')

        # List of primitive objects
        primitive_objects = [x for x in model_objects if x not in pcb_layers]

        # Set mesh dimensions
        dim_x = []
        dim_y = []
        dim_z = []
        for i in primitive_objects:
            obj_handle = ipk.modeler.get_object_from_name(i)
            dim_x.append(obj_handle.bounding_dimension[0])
            dim_y.append(obj_handle.bounding_dimension[1])
        for i in pcb_layers:
            obj_handle = ipk.modeler.get_object_from_name(i)
            dim_z.append(obj_handle.bounding_dimension[2])

        pcb_dim_x = ipk.modeler.get_object_from_name(pcb_layers[0]).bounding_dimension[0]
        pcb_dim_y = ipk.modeler.get_object_from_name(pcb_layers[0]).bounding_dimension[1]
        pcb_min_x = ipk.modeler.get_object_from_name(pcb_layers[0]).bounding_box[0]
        pcb_min_y = ipk.modeler.get_object_from_name(pcb_layers[0]).bounding_box[1]
        pcb_max_x = ipk.modeler.get_object_from_name(pcb_layers[0]).bounding_box[3]
        pcb_max_y = ipk.modeler.get_object_from_name(pcb_layers[0]).bounding_box[4]

        tx = np.histogram(dim_x, bins=10)
        ty = np.histogram(dim_y, bins=10)
        max_val_index_x = np.argmax(tx[0])
        max_val_index_y = np.argmax(ty[0])

        if mesh_fidelity == 'Coarse':
            mesh_mult_xy = 0.5
            mesh_mult_z = 8
        elif mesh_fidelity == 'Medium':
            mesh_mult_xy = 0.25
            mesh_mult_z = 4
        else:
            mesh_mult_xy = 0.1
            mesh_mult_z = 2

        # Max element size in x, y, z direction based on mesh fidelity
        mesh_x = mesh_mult_xy * (tx[1][max_val_index_x] + tx[1][max_val_index_x + 1])
        mesh_y = mesh_mult_xy * (ty[1][max_val_index_y] + ty[1][max_val_index_y + 1])
        mesh_z = mesh_mult_z * min(dim_z)

        # Find extent of all objects in z-direction
        minzs = []
        maxzs = []
        for i in primitive_objects:
            obj_handle = ipk.modeler.get_object_from_name(i)
            minzs.append(obj_handle.bounding_box[2])
            maxzs.append(obj_handle.bounding_box[5])
        z_extent_min = min(minzs)
        z_extent_max = max(maxzs)
        z_extent = z_extent_max - z_extent_min

        # slack values
        slack_x = 0.1 * pcb_dim_x
        slack_y = 0.1 * pcb_dim_y
        slack_z = 0.25 * z_extent

        # Add mesh region
        meshregion_box = ipk.modeler.create_box([pcb_min_x, pcb_min_y, z_extent_min], [pcb_dim_x, pcb_dim_y, z_extent],
                                                'meshregion_all_objs')
        add_slack('meshregion_all_objs', slack_x, slack_x, slack_y, slack_y, slack_z, slack_z)
        meshregion_box.model = False
        mesh_box = 'meshregion_all_objs'
        mesh_region = ipk.mesh.assign_mesh_region([mesh_box], 5, False, 'meshregion_all_objs')

        # Set user defined settings in mesh region
        mesh_region.UserSpecifiedSettings = True
        mesh_region.MaxElementSizeX = str(mesh_x) + ipk.modeler.model_units
        mesh_region.MaxElementSizeY = str(mesh_y) + ipk.modeler.model_units
        mesh_region.MaxElementSizeZ = str(mesh_z) + ipk.modeler.model_units
        mesh_region.MinElementsInGap = 2
        mesh_region.MinElementsOnEdge = 2
        mesh_region.MaxSizeRatio = 2
        mesh_region.NoOGrids = True
        mesh_region.StairStepMeshing = False
        mesh_region.MinGapX = '0.0001mm'
        mesh_region.MinGapY = '0.0001mm'
        mesh_region.MinGapZ = '0.0001mm'
        mesh_region.EnableMLM = True
        mesh_region.MaxLevels = 2
        mesh_region.BufferLayers = 1
        mesh_region.EnforeMLMType = "2D"
        mesh_region.Enable2DCutCell = True
        mesh_region.UniformMeshParametersType = "Average"
        mesh_region.DMLMType = "2DMLM_XY"
        mesh_region.Objects = [mesh_box]
        mesh_region.update()

        # Add mesh operation to primitives, mesh level = 2
        mesh_levels_primitives = {}
        for i in primitive_objects:
            mesh_levels_primitives[i] = 2
        ipk.mesh.assign_mesh_level(mesh_levels_primitives, "mesh_levels_primitives")

        # Add mesh operation to primitives, mesh level = 1
        mesh_levels_3dcomps = {}
        for i in pcb_layers:
            mesh_levels_3dcomps[i] = 1
        ipk.mesh.assign_mesh_level(mesh_levels_3dcomps, "mesh_levels_pcb_layers")

        # Global mesh dimensions
        domain = ipk.modeler.get_bounding_dimension()
        global_max_x = 4 * mesh_x
        global_max_y = 4 * mesh_y
        global_max_z = 4 * mesh_z

        # Apply global mesh settings.
        ipk.mesh.global_mesh_region.UserSpecifiedSettings = True
        ipk.mesh.global_mesh_region.MaxElementSizeX = str(global_max_x) + ipk.modeler.model_units
        ipk.mesh.global_mesh_region.MaxElementSizeY = str(global_max_y) + ipk.modeler.model_units
        ipk.mesh.global_mesh_region.MaxElementSizeZ = str(global_max_z) + ipk.modeler.model_units
        ipk.mesh.global_mesh_region.MinElementsInGap = 3
        ipk.mesh.global_mesh_region.MinElementsOnEdge = 2
        ipk.mesh.global_mesh_region.MaxSizeRatio = 2
        ipk.mesh.global_mesh_region.NoOGrids = True
        ipk.mesh.global_mesh_region.StairStepMeshing = False
        ipk.mesh.global_mesh_region.MinGapX = '0.0001mm'
        ipk.mesh.global_mesh_region.MinGapY = '0.0001mm'
        ipk.mesh.global_mesh_region.MinGapZ = '0.0001mm'
        ipk.mesh.global_mesh_region.EnableMLM = False
        ipk.mesh.global_mesh_region.UniformMeshParametersType = "None"
        ipk.mesh.global_mesh_region.OptimizePCBMesh = True
        ipk.mesh.global_mesh_region.update()

        # Assign Boundary Conditions
        def name_cleanup(name):
            return re.sub(r"\W", "_", name)


        colnames = ['Include', 'Package_Name', 'Part_Name', 'Instance_Name', 'Designator_Type', 'Height [mm]',
                    'Placement', 'BC_Type', 'Power [W]', 'R_jb [C/W]', 'R_jc [C/W]', 'Monitor_Point', 'Material']
        df = pd.DataFrame(rows, columns=colnames)
        df2 = df.copy()
        df2 = df2[(df2['Include'] == 'YES')]
        df2 = df2[(df2['Instance_Name'] != 'NOREFDES')]
        df3 = pd.DataFrame()
        df3['block_name'] = df2.apply(lambda x: name_cleanup(x.Instance_Name), axis=1)
        df3['bc_type'] = df2['BC_Type']
        df3['power'] = df2['Power [W]']
        df3['rjb'] = df2['R_jb [C/W]']
        df3['rjc'] = df2['R_jc [C/W]']
        df3['monpt'] = df2['Monitor_Point']
        df3['mat_type'] = df2['Material']

        for ind in df3.index:
            block_handle = ipk.modeler.get_object_from_name(df3['block_name'][ind])
            if df3['bc_type'][ind] == "block":
                if df3['power'][ind] != 0:
                    ipk.create_source_block(df3['block_name'][ind], df3['power'][ind] + "W", assign_material=False,
                                            use_object_for_name=True)
                # Assign material property
                if df3['mat_type'][ind] != "":
                    block_handle.material_name = df3['mat_type'][ind]
                    block_handle.surface_material_name = 'Ceramic-surface'
            elif df3['bc_type'][ind] == "network":
                ipk.create_two_resistor_network_block(object_name=df3['block_name'][ind], pcb=pcb[0],
                                                      power=df3['power'][ind] + "W",
                                                      rjb=df3['rjb'][ind], rjc=df3['rjc'][ind])
            elif df3['bc_type'][ind] == "hollow":
                ipk.create_source_block(df3['block_name'][ind], df3['power'][ind] + "W", assign_material=False,
                                        use_object_for_name=True)
                ipk.modeler.primitives[df3['block_name'][ind]].solve_inside = False
            else:
                e = RuntimeError('Error! Incorrect block boundary condition.')
                st.exception(e)
            # Monitor points
            if df3['monpt'][ind] == "YES":
                pcb_top_layer = ipk.modeler.get_object_from_name(pcb_layers[0])
                pcb_bottom_layer = ipk.modeler.get_object_from_name(pcb_layers[-1])
                if block_handle.get_touching_faces(pcb_top_layer):
                    block_board_side = block_handle.get_touching_faces(pcb_top_layer)
                else:
                    block_board_side = block_handle.get_touching_faces(pcb_bottom_layer)
                point_name = 'point_' + df3['block_name'][ind]
                mon_point = ipk.modeler.primitives.get_face_center(block_board_side[0].id)
                ipk.assign_point_monitor(mon_point, monitor_type='Temperature', monitor_name=point_name)

        # Insert forced convection setup
        if conv_type == 'Forced':
            analysis_setup = 'forced_conv_setup'
            forced_convection_setup(analysis_setup, 'Turbulent')

            # Assign velocity inlet and pressure outlet boundary conditions
            region = ipk.modeler.primitives["Region"]
            if vel_dir == '+X':
                ipk.modeler.edit_region_dimensions([100, 100, 50, 50, 50, 50])
                inlet_opening_face_id = region.bottom_face_x.id
                outlet_opening_face_id = region.top_face_x.id
                speed = str(vel) + 'm_per_sec'
                air_temp = str(air_temp) + 'cel'
                assign_opening_boundary('inlet', inlet_opening_face_id, flow_type='velocity', xvel=speed,
                                        temperature=air_temp)
                assign_opening_boundary('outlet', outlet_opening_face_id, flow_type='pressure')
            elif vel_dir == '-X':
                ipk.modeler.edit_region_dimensions([100, 100, 50, 50, 50, 50])
                inlet_opening_face_id = region.top_face_x.id
                outlet_opening_face_id = region.bottom_face_x.id
                speed = str(vel) + 'm_per_sec'
                air_temp = str(air_temp) + 'cel'
                assign_opening_boundary('inlet', inlet_opening_face_id, flow_type='velocity', xvel=speed,
                                        temperature=air_temp)
                assign_opening_boundary('outlet', outlet_opening_face_id, flow_type='pressure')
            elif vel_dir == '+Y':
                ipk.modeler.edit_region_dimensions([50, 50, 100, 100, 50, 50])
                inlet_opening_face_id = region.bottom_face_y.id
                outlet_opening_face_id = region.top_face_y.id
                speed = str(vel) + 'm_per_sec'
                air_temp = str(air_temp) + 'cel'
                assign_opening_boundary('inlet', inlet_opening_face_id, flow_type='velocity', yvel=speed,
                                        temperature=air_temp)
                assign_opening_boundary('outlet', outlet_opening_face_id, flow_type='pressure')
            elif vel_dir == '-Y':
                ipk.modeler.edit_region_dimensions([50, 50, 100, 100, 50, 50])
                inlet_opening_face_id = region.top_face_y.id
                outlet_opening_face_id = region.bottom_face_y.id
                speed = str(vel) + 'm_per_sec'
                air_temp = str(air_temp) + 'cel'
                assign_opening_boundary('inlet', inlet_opening_face_id, flow_type='velocity', yvel=speed,
                                        temperature=air_temp)
                assign_opening_boundary('outlet', outlet_opening_face_id, flow_type='pressure')
            elif vel_dir == '+Z':
                ipk.modeler.edit_region_dimensions([50, 50, 50, 50, 100, 100])
                inlet_opening_face_id = region.bottom_face_z.id
                outlet_opening_face_id = region.top_face_z.id
                speed = str(vel) + 'm_per_sec'
                air_temp = str(air_temp) + 'cel'
                assign_opening_boundary('inlet', inlet_opening_face_id, flow_type='velocity', zvel=speed,
                                        temperature=air_temp)
                assign_opening_boundary('outlet', outlet_opening_face_id, flow_type='pressure')
            else:
                ipk.modeler.edit_region_dimensions([50, 50, 50, 50, 100, 100])
                inlet_opening_face_id = region.top_face_z.id
                outlet_opening_face_id = region.bottom_face_z.id
                speed = str(vel) + 'm_per_sec'
                air_temp = str(air_temp) + 'cel'
                assign_opening_boundary('inlet', inlet_opening_face_id, flow_type='velocity', zvel=speed,
                                        temperature=air_temp)
                assign_opening_boundary('outlet', outlet_opening_face_id, flow_type='pressure')

        if conv_type == 'Natural':
            analysis_setup = 'natural_conv_setup'
            natural_convection_setup(analysis_setup, gravity_dir=gravity_direction, flow_regime='Turbulent',
                                     ambient_temp=air_temp)
            for i in ipk.modeler.get_object_faces('Region'):
                outlet_name = 'outlet_' + str(i)
                assign_opening_boundary(outlet_name, i, flow_type='pressure')

        # Create monitor points at all object bases
        list_mon_pts = ipk.odesign.GetChildObject("Monitor").GetChildNames()
        for pt in points_dict:
            if pt not in list_mon_pts:
                ipk.assign_point_monitor(points_dict[pt], monitor_type='Temperature', monitor_name=pt)

        ipk.modeler.refresh_all_ids()
        ipk.modeler.refresh()

        if analyze_setup:
            ipk.mesh.generate_mesh(analysis_setup)
            # Solve the model.
            num_tasks = num_cores
            ipk.analyze_setup(analysis_setup, num_cores, num_tasks)
            quit_aedt()
        analysis_complete = True
        if analysis_complete:
            placeholder.success('AEDT Icepak run completed.', icon="✅")
            st.markdown(f'''**Project saved to:** ```{project_path}```''')
            st.session_state['pid'] = desktop.aedt_process_id
else:
    e = RuntimeError('One or more input files are missing')
    st.exception(e)

if st.session_state['pid']:
    close_aedt = st.button('Close AEDT')
    if close_aedt:
        try:
            os.kill(st.session_state['pid'], signal.SIGTERM)
            files = os.listdir(os.getcwd())
            for file in files:
                if file.endswith('.lock'):
                    os.remove(file)
        except RuntimeWarning:
            st.warning('⚠️ No active AEDT sessions!')
