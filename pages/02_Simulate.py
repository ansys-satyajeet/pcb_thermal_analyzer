import os
import re
import csv
import shutil
import pyaedt
import signal
import numpy as np
import pandas as pd
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from collections import OrderedDict

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")
st.title('🖥️Simulate')

# Function definitions

# Function to create 2R network - to be replaced with built-in method
def create_2R_network_BC(object_handle, power, rjb, rjc, board_side):
    """ Function to create 2-Resistor network object
        Parameters
        ----------
        object_handle: str
            handle of the object (3D block primitive) on which 2-R network is created
        power: float
            junction power in [W]
        rjb: float
            Junction to board thermal resistance in [K/W]
        rjc: float
            Junction to case thermal resistance in [K/W]
        board_side: str
            location of board w.r.t. block. 
            Acceptable entries are "minx","miny","minz","maxx","maxy","maxz"
    """
    board_side = board_side.casefold()
    if board_side == "minx":
        board_faceID = object_handle.bottom_face_x.id
        case_faceID = object_handle.top_face_x.id
        case_side = "maxx"
    elif board_side == "maxx":
        board_faceID = object_handle.top_face_x.id
        case_faceID = object_handle.bottom_face_x.id
        case_side = "minx"
    elif board_side == "miny":
        board_faceID = object_handle.bottom_face_y.id
        case_faceID = object_handle.top_face_y.id
        case_side = "maxy"
    elif board_side == "maxy":
        board_faceID = object_handle.top_face_y.id
        case_faceID = object_handle.bottom_face_y.id
        case_side = "miny"
    elif board_side == "minz":
        board_faceID = object_handle.bottom_face_z.id
        case_faceID = object_handle.top_face_z.id
        case_side = "maxz"
    else:
        board_faceID = object_handle.top_face_z.id
        case_faceID = object_handle.bottom_face_z.id
        case_side = "minz"
        
    # Define network properties in props directory
    props = {}
    props["Faces"] = [board_faceID, case_faceID]
    
    props["Nodes"] = OrderedDict(
        {
            "Case_side(" + case_side + ")": [case_faceID, "NoResistance"],
            "Board_side(" + board_side + ")": [board_faceID, "NoResistance"],
            "Internal": [power+'W'],
        }
    )
    
    props["Links"] = OrderedDict(
        {
            "Rjc": ["Case_side(" + case_side + ")", "Internal", "R", str(rjc) + "cel_per_w"],
            "Rjb": ["Board_side(" + board_side + ")", "Internal", "R", str(rjb) + "cel_per_w"],
        }
    )
    
    props["SchematicData"] = ({})
    
    # Default material is Ceramic Material
    ipk.modeler.primitives[object_handle.name].material_name = "Ceramic_material"
    
    # Create boundary condition and set Solve Inside to No
    bound = pyaedt.modules.Boundary.BoundaryObject(ipk, object_handle.name, props, "Network")
    if bound.create():
        ipk.boundaries.append(bound)
        ipk.modeler.primitives[object_handle.name].solve_inside = False

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
    setup.props['Convergence Criteria - Max Iterations'] = 250
    setup.update()

# Function to create a natural convection problem setup with default entries
def natural_convection_setup(setup_name, gravity_dir, flow_regime, turb_model='ZeroEquation', ambient_temp = 20):
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
        ipk.apply_icepak_settings(ambienttemp = ambient_temp, gravityDir = 0)
        ipk.modeler.edit_region_dimensions([250,50,50,50,200,200])
        setup.props['Solution Initialization - X Velocity'] = "0.00098m_per_sec"
    elif gravity_dir == "-y":
        ipk.apply_icepak_settings(ambienttemp = ambient_temp, gravityDir = 1)
        ipk.modeler.edit_region_dimensions([50,50,250,50,200,200])
        setup.props['Solution Initialization - Y Velocity'] = "0.00098m_per_sec"
    elif gravity_dir == "-z":
        ipk.apply_icepak_settings(ambienttemp = ambient_temp, gravityDir = 2)
        ipk.modeler.edit_region_dimensions([50,50,50,50,250,50])
        setup.props['Solution Initialization - Z Velocity'] = "0.00098m_per_sec"
    elif gravity_dir == "+x":
        ipk.apply_icepak_settings(ambienttemp = ambient_temp, gravityDir = 3)
        ipk.modeler.edit_region_dimensions([50,250,50,50,200,200])
        setup.props['Solution Initialization - X Velocity'] = "-0.00098m_per_sec"
    elif gravity_dir == "+y":
        ipk.apply_icepak_settings(ambienttemp = ambient_temp, gravityDir = 4)
        ipk.modeler.edit_region_dimensions([50,50,50,250,200,200])
        setup.props['Solution Initialization - Y Velocity'] = "-0.00098m_per_sec"
    else:
        ipk.apply_icepak_settings(ambienttemp = ambient_temp, gravityDir = 5)
        ipk.modeler.edit_region_dimensions([50,50,50,50,50,250])
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
    setup.props['Convergence Criteria - Max Iterations'] = 250
    setup.update()

# Function to create opening boundary condition
def assign_opening_BC(name, face_id, flow_type,
                      xvel = "0m_per_sec", yvel = "0m_per_sec", zvel = "0m_per_sec",
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
        xvel: float, optional
            velocity in x-direction
        yvel: float, optional
            velocity in y-direction
        zvel: float, optional
            velocity in z-direction
        pressure: float, optional
            pressure at opening boundary
        temperature: float, optional
            temperature at opening boundary
    """
    props = {}
    props["Faces"] = [face_id]
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
def add_slack(box_name,minx,maxx,miny,maxy,minz,maxz):
    obj_handle = ipk.modeler.get_object_from_name(box_name)
    obj_handle.bottom_face_x.move_with_offset(minx)
    obj_handle.bottom_face_y.move_with_offset(miny)
    obj_handle.bottom_face_z.move_with_offset(minz)
    obj_handle.top_face_x.move_with_offset(maxx)
    obj_handle.top_face_y.move_with_offset(maxy)
    obj_handle.top_face_z.move_with_offset(maxz)
    return None

# Function to remove aedt files and project folders
def cleanup_files(project_name):
    project_path = os.path.join(os.getcwd(), project_name)
    project_name_no_ext = os.path.splitext(project_name)[0]
    if os.path.exists(project_path):
        os.remove(project_path)
    if os.path.exists(os.path.join(os.getcwd(),project_name + ".lock")):
        os.remove(os.path.join(os.getcwd(),project_name + ".lock"))
    # Delete aedt results folder
    if os.path.exists(os.path.join(os.getcwd(),project_name_no_ext + ".aedtresults")):
        try:
            shutil.rmtree(os.path.join(os.getcwd(),project_name_no_ext + ".aedtresults"))
        except:
            print('Error deleting aedtresults directory')
    # Delete pyaedt folder
    if os.path.exists(os.path.join(os.getcwd(),project_name_no_ext + ".pyaedt")):
        try:
            shutil.rmtree(os.path.join(os.getcwd(),project_name_no_ext + ".pyaedt"))
        except:
            print('Error deleting pyaedt directory')
    # # Delete aedb folder
    # if os.path.exists(os.path.join(os.getcwd(),project_name_no_ext + ".aedb")):
    #     try:
    #         shutil.rmtree(os.path.join(os.getcwd(),project_name_no_ext + ".aedb"))
    #     except:
    #         print('Error deleting aedb directory')
          
# Function to import ECAD
def import_ecad(ecad_file_path, ecad_type):
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
    ecad_design = h3d.design_list[0]

    # Get name of outline polygon
    outline_poly = []
    for key in h3d.modeler.polygons.keys():
        if h3d.modeler.polygons[key].placement_layer == 'Outline':
            outline_poly.append(key)

    ipk.create_pcb_from_3dlayout(component_name=ecad_file_name_no_ext,
                                    project_name=ecad_project_path,
                                    design_name=ecad_design,
                                    close_linked_project_after_import=True,
                                    extenttype='Polygon',
                                    outlinepolygon=outline_poly[0],
                                    resolution=3)
                                    
    for i in ipk.design_list:
        design = desktop.design_type(project_name=ipk.project_name,design_name=i)
        if design != 'Icepak':
            ipk.delete_design(i)

def quit_aedt():
    ipk.save_project()
    pid = desktop.aedt_process_id
    os.kill(pid,signal.SIGTERM)
    files = os.listdir(os.getcwd())
    for file in files:
        if file.endswith('.lock'):
            os.remove(file)
    

# Fix blur issue in tkinter window panels
from ctypes import windll
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

c1, c2 = st.columns([3,1])
c1.write('Select working directory:')
workdir_button = c2.button('Select Folder')
if workdir_button:
    st.session_state['workdir'] = True
    root0 = tk.Tk()
    root0.attributes("-topmost", True)
    root0.withdraw()
    try:
        workdir = filedialog.askdirectory(master=root0)
        st.session_state['workdir'] = workdir
    except:
        pass

if st.session_state['workdir']:
    os.chdir(st.session_state['workdir'])

# Read board file from windows explorer dialog box
# 
col01, col02, col03 = st.columns([2,1,1])
col01.write('Select IDF Board file type:')
st.session_state['idf_type'] = col02.selectbox('Select IDF Board file type:', ('*.emn','*.bdf'), label_visibility='collapsed')
idf_button = col03.button('Select Board File')
if idf_button:
    st.session_state['idf'] = True
    root1 = tk.Tk()
    root1.attributes("-topmost", True)
    root1.withdraw()
    try:
        if st.session_state['idf_type'] == '*.emn':
            files = filedialog.askopenfilenames(master=root1,filetypes=[('EMN File','*.emn')])
        if st.session_state['idf_type'] == '*.bdf':
            files = filedialog.askopenfilenames(master=root1,filetypes=[('BDF File', '*.bdf')])
        idf_file = os.path.basename(files[0])
        st.session_state['idf_file'] = idf_file
    except:
        pass

# Read ECAD file from windows explorer dialog box
#
col04, col05, col06 = st.columns([2,1,1])
col04.write('Please select ECAD file type:')
st.session_state['ecad_type'] = col05.selectbox('Select ECAD type:', ('EDB Folder','ODB++ File','BRD File'), label_visibility='collapsed')
ecad_button = col06.button('Select ECAD')
if ecad_button:
    st.session_state['ecad'] = True
    root2 = tk.Tk()
    root2.attributes("-topmost", True)
    root2.withdraw()
    try:
        if st.session_state['ecad_type'] == 'EDB Folder':
            ecad_file = filedialog.askdirectory(master=root2,initialdir = os.getcwd())
            st.session_state['ecad_file'] = ecad_file
        elif st.session_state['ecad_type'] == 'ODB++ File':
            ecad_file = filedialog.askopenfilename(master=root2,filetypes=[('TGZ File','*.tgz')],initialdir = os.getcwd())
            st.session_state['ecad_file'] = ecad_file
        elif st.session_state['ecad_type'] == 'BRD File':
            ecad_file = filedialog.askopenfilename(master=root2,filetypes=[('BRD File','*.brd')],initialdir = os.getcwd())
            st.session_state['ecad_file'] =ecad_file
        else:
            st.error('Something went wrong!')
    except:
        pass

# Read Boundary Conditions CSV file from windows explorer dialog box
#
col07, col08 = st.columns([3,1])
col07.write('Please select boundary conditions CSV file:')
bc_file_button = col08.button('Select BC CSV file')
if bc_file_button:
    st.session_state['bcs'] = True
    bcs = True
    root3 = tk.Tk()
    root3.attributes("-topmost", True)
    root3.withdraw()
    try:
        bc_file = filedialog.askopenfilenames(master=root3,filetypes=[('Microsoft Excel Comma Separated Values File','*.csv')],initialdir = os.getcwd())
        bc_filename = os.path.basename(bc_file[0])
        st.session_state['bc_filename'] = bc_filename
    except:
        pass

# Read material file from windows explorer dialog box
#
include_matfile = st.checkbox('Read Materials as CSV File?')
if include_matfile:
    col09, col10 = st.columns([3,1])
    col09.write('Please select materials CSV file:')
    mats_button = col10.button('Select Materials CSV')
    if mats_button:
        st.session_state['mats'] = True
        root4 = tk.Tk()
        root4.attributes("-topmost", True)
        root4.withdraw()
        try:
            materials_file = filedialog.askopenfilenames(master=root4,filetypes=[('Microsoft Excel Comma Separated Values File','*.csv')],initialdir = os.getcwd())
            materials_filename = os.path.basename(materials_file[0])
            st.session_state['materials_filename'] = materials_filename
        except:
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
        st.markdown(f'''**Selected Boundary Conditions CSV file:** ```{os.path.abspath(st.session_state['bc_filename'])}```''')
    else:
        st.markdown(f'''**Selected Boundary Conditions CSV file:** None''')
    
    if include_matfile:
        if st.session_state['materials_filename']:
            st.markdown(f'''**Selected Materials CSV file:** ```{os.path.abspath(st.session_state['materials_filename'])}```''')
        else:
            st.markdown(f'''**Selected Materials CSV file:** None''')
    


# Setup Options
st.markdown('---')
st.markdown('**Setup Options**')

all_points = st.checkbox('Create points at board side face centers for all components', help='Can increase model setup time!')

if all_points:
    st.write(':information_source: Temperatures at the created points will be written out to a table/file.')

delete_filtered = st.checkbox('Delete filtered objects?', help='Deleted objects cannot be recovered.')

if delete_filtered:
    st.write(':information_source: Filtered objects will be deleted.')

# Solution Settings
st.markdown('---')
st.markdown('**Solution Settings**')

conv_type = st.selectbox('Select convection mode:', ('Forced','Natural'))

conv_cond = False

if conv_type == 'Forced':
    col11, col12 = st.columns(2)
    vel = col11.text_input('Velocity Magnitude [m/s]:')
    vel_dir = col12.selectbox('Direction:', ('+X','-X','+Y','-Y','+Z','-Z'))
    Tin = st.text_input('Inlet Temperature [C]:')
    conv_cond = True

if conv_type == 'Natural':
    gravity_direction = st.selectbox('Direction of gravity:', ('+X','-X','+Y','-Y','+Z','-Z'))
    Tamb = st.text_input('Ambient Temperature [C]:')
    conv_cond = True

# Mesh Settings
st.markdown('---')
st.markdown('**Mesh Settings**')
mesh_fidelity = st.select_slider('Select mesh resolution:', options=['Coarse', 'Medium', 'Fine'])

# Solve Settings
st.markdown('---')
st.markdown('**Solve Settings**')
col13, col14 = st.columns(2)
num_cores = col13.number_input('Number of Processors', min_value=1, max_value=128,value=1,step=1,format='%d')
mode = col14.radio('Mode:',('Graphical', 'Non-Graphical'))

project_name = st.text_input('Enter Project Name:', help='Only letters (A-Z,a-z), numbers (0-9) and underscores are allowed.')

aedt_version = st.selectbox('Select AEDT release:', ('2022 R2','2023 R1'))

analyze_setup = st.checkbox('Setup problem and proceed to solve')
if analyze_setup:
    sim_button_text = '**Simulate**'
else:
    sim_button_text = '**Setup Only**'
setup_analyze = st.button(sim_button_text)


if conv_type == 'Forced':
    airtemp = Tin
else:
    airtemp = Tamb

placeholder = st.empty()
analysis_complete = False

if st.session_state['idf_file'] and st.session_state['ecad_file'] and st.session_state['bc_filename'] and conv_cond and airtemp and project_name: 
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
    if setup_analyze:
        # if aedt_version == '2022 R2':
        #     aedt_release = '2022.2'
        # elif aedt_version == '2023 R1':
        #     aedt_release = '2023.1'
        # else:
        #     aedt_release = '2022.2'

        aedt_release = re.sub(' R', '.', aedt_version)

        cleanup_files(project_name)
        placeholder.info('AEDT Icepak session in progress...', icon="🏃🏽")

        if mode == 'Graphical':
            non_graphical_mode = False
        else:
            non_graphical_mode = True

        # Start AEDT Desktop session
        desktop = pyaedt.Desktop(aedt_release, non_graphical=non_graphical_mode)

        # Import ECAD file
        # import_ecad(ecad_file_path=st.session_state['ecad_file'], ecad_type=st.session_state['ecad_type'])

        # Insert ECAD Import code
        ecad_file_path=st.session_state['ecad_file']
        ecad_type=st.session_state['ecad_type']
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

        ecad_design = h3d.design_list[0]

        # Get name of outline polygon
        outline_poly = []
        for key in h3d.modeler.polygons.keys():
            if h3d.modeler.polygons[key].placement_layer == 'Outline':
                outline_poly.append(key)

        ipk = pyaedt.Icepak()
        ipk.save_project()

        # project_path = os.path.join(os.getcwd(), project_name)
        ipk.oproject.Rename(os.path.join(ipk.project_path, project_name), True)

        # Create PCB object in Icepak
        ipk.create_pcb_from_3dlayout(component_name=ecad_file_name_no_ext,
                                        project_name=None,
                                        design_name=ecad_design,
                                        close_linked_project_after_import=False,
                                        extenttype='Polygon',
                                        outlinepolygon=outline_poly[0],
                                        resolution=3)

        # Import IDF file
        ipk.import_idf(board_filename)

        # Fit all and save
        ipk.modeler.fit_all()
        ipk.save_project()
        ipk.autosave_disable()

        # List of boundary conditions
        bc_list = ipk.odesign.GetChildObject('Thermal').GetChildNames()

        # Delete all boundary conditions
        omodule = ipk.odesign.GetModule("BoundarySetup")
        if bc_list:
            for i in bc_list:
                omodule.DeleteBoundaries([i])

        # Delete all points
        for i in ipk.modeler.points:
            ipk.modeler.points[i].delete()

        # Import Modified CSV file 
        fields = []
        rows = []
        with open(bc_filename, 'r') as csvFile:
            csvReader = csv.reader(csvFile)
            fields = next(csvReader)
            for row in csvReader:
                rows.append(row)
        
        if st.session_state['materials_filename']:
            # Read material properties file
            fields_mat = []
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
                block_name = re.sub("\W", "_", rows[i][3])
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

        
        ## Create monitor points
        points_list = []
        mon_point_list = []
        for i in range(len(rows)):
            if rows[i][3] != 'NOREFDES':
                block_name = re.sub("\W", "_", rows[i][3])
                block_handle = ipk.modeler.get_object_from_name(block_name)
                board = "IDF_BoardOutline"
                block_board_side = block_handle.get_touching_faces(board)
                point_name = 'point_' + block_name
                
                if rows[i][11] == 'YES':
                    mon_point = ipk.modeler.primitives.get_face_center(block_board_side[0].id)
                    ipk.modeler.primitives.create_point(mon_point, point_name)
                    mon_point_list.append(point_name)
                    ipk.assign_point_monitor(mon_point, monitor_type='Temperature',monitor_name=point_name)
                    
                if all_points:
                    if point_name not in ipk.modeler.point_names:
                        mon_point = ipk.modeler.primitives.get_face_center(block_board_side[0].id)
                        ipk.modeler.primitives.create_point(mon_point, point_name)
                        points_list.append(point_name)
                

        # # Create points at all object locations
        # points_list = []
        # if all_points:
        #     for i in range(len(rows)):
        #         block_name = rows[i][3]
        #         block_name = re.sub("\W","_",block_name)
        #         block_handle = ipk.modeler.get_object_from_name(rows[i][3])
        #         plusZ_side = ipk.modeler.primitives.get_face_center(block_handle.top_face_z.id)[2]
        #         minusZ_side = ipk.modeler.primitives.get_face_center(block_handle.bottom_face_z.id)[2]
        #         board = "IDF_BoardOutline"
        #         board_handle = ipk.modeler.get_object_from_name(board)
        #         board_top = ipk.modeler.primitives.get_face_center(board_handle.top_face_z.id)[2]
        #         board_bottom = ipk.modeler.primitives.get_face_center(board_handle.bottom_face_z.id)[2]
        #         if plusZ_side == board_top:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.top_face_z.id)
        #         elif plusZ_side == board_bottom:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.top_face_z.id)
        #         elif minusZ_side == board_top:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.bottom_face_z.id)
        #         elif minusZ_side == board_bottom:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.bottom_face_z.id)
        #         else:
        #             mon_point = -1
        #         point_name = 'point_' + block_name
        #         ipk.modeler.primitives.create_point(mon_point, point_name)
        #         points_list.append(point_name)

        # mon_point_list = []
        # for i in range(len(rows)):
        #     if rows[i][11] == 'YES':
        #         block_name = rows[i][3]
        #         block_name = re.sub("\W","_",block_name)
        #         block_handle = ipk.modeler.get_object_from_name(rows[i][3])
        #         plusZ_side = ipk.modeler.primitives.get_face_center(block_handle.top_face_z.id)[2]
        #         minusZ_side = ipk.modeler.primitives.get_face_center(block_handle.bottom_face_z.id)[2]
        #         board = "IDF_BoardOutline"
        #         board_handle = ipk.modeler.get_object_from_name(board)
        #         board_top = ipk.modeler.primitives.get_face_center(board_handle.top_face_z.id)[2]
        #         board_bottom = ipk.modeler.primitives.get_face_center(board_handle.bottom_face_z.id)[2]
        #         if plusZ_side == board_top:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.top_face_z.id)
        #         elif plusZ_side == board_bottom:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.top_face_z.id)
        #         elif minusZ_side == board_top:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.bottom_face_z.id)
        #         elif minusZ_side == board_bottom:
        #             mon_point = ipk.modeler.primitives.get_face_center(block_handle.bottom_face_z.id)
        #         else:
        #             mon_point = -1
        #         point_name = 'point_' + block_name
        #         ipk.modeler.primitives.create_point(mon_point, point_name)
        #         # Create monitor points as specified by user
        #         ipk.assign_point_monitor(mon_point, monitor_type='Temperature',monitor_name=point_name)
        #         mon_point_list.append(point_name)

        # Delete filtered objects or make them non-model
        for i in range(len(rows)):
            if rows[i][0] == 'NO':
                if rows[i][3] != 'NOREFDES':
                    block_name = re.sub("\W", "_", rows[i][3])
                    block_handle = ipk.modeler.get_object_from_name(block_name)
                    if delete_filtered:
                        ipk.modeler.delete(block_handle.name)
                    else:
                        block_handle.model = False
        
        # Assign BCs
        for i in range(len(rows)):
            if rows[i][0] == 'YES':
                if rows[i][3] != 'NOREFDES':
                    block_name = re.sub("\W", "_", rows[i][3])
                    block_handle = ipk.modeler.get_object_from_name(block_name)
                    if rows[i][7] == "block":
                        if rows[i][8] != 0:
                            ipk.create_source_block(block_name, rows[i][8] + "W", assign_material=False, use_object_for_name=True)
                        # Assign material property
                        if rows[i][12] != "":
                            block_handle.material_name = rows[i][12]
                            block_handle.surface_material_name = 'Ceramic-surface'
                    elif rows[i][7] == "network":
                        if rows[i][6] == 'TOP':
                            board_side = 'minz'
                        else:
                            board_side = 'maxz'
                        # if rows[i][4] == 'BOTTOM':
                        #     board_side == 'maxz'
                        create_2R_network_BC(block_handle, rows[i][8], rows[i][9], rows[i][10], board_side)
                    elif rows[i][7] == "hollow":
                        ipk.create_source_block(block_name, rows[i][8] + "W", assign_material=False, use_object_for_name=True)
                        ipk.modeler.primitives[block_name].solve_inside = False
                    else:
                        e = RuntimeError('Error! Incorrect block boundary condition.')
                        st.exception(e)

        # Save project
        ipk.save_project()

        # Import ECAD file
        # import_ecad(ecad_file_path=st.session_state['ecad_file'], ecad_type=st.session_state['ecad_type'])

        # Make board that comes with the IDF file as non-model object
        board_handle = ipk.modeler.get_object_from_name('IDF_BoardOutline')
        board_handle.model = False

        # Insert forced convection setup
        if conv_type == 'Forced':
            analysis_setup = 'forced_conv_setup'
            forced_convection_setup(analysis_setup, 'Turbulent')

            # Assign velocity inlet and pressure outlet boundary conditions
            region = ipk.modeler.primitives["Region"]
            if vel_dir == '+X':
                inlet_opening_face_id = region.bottom_face_x.id
                outlet_opening_face_id = region.top_face_x.id
                speed = str(vel) + 'm_per_sec'
                airtemp = str(airtemp) + 'cel' 
                assign_opening_BC('inlet', inlet_opening_face_id, flow_type = 'velocity', xvel = speed, temperature = airtemp)
                assign_opening_BC('outlet', outlet_opening_face_id, flow_type = 'pressure')
                ipk.modeler.edit_region_dimensions([100,100,50,50,50,50])
            elif vel_dir == '-X':
                inlet_opening_face_id = region.top_face_x.id
                outlet_opening_face_id = region.bottom_face_x.id
                speed = str(vel) + 'm_per_sec'
                airtemp = str(airtemp) + 'cel'
                assign_opening_BC('inlet', inlet_opening_face_id, flow_type = 'velocity', xvel = speed, temperature = airtemp)
                assign_opening_BC('outlet', outlet_opening_face_id, flow_type = 'pressure')
                ipk.modeler.edit_region_dimensions([100,100,50,50,50,50])
            elif vel_dir == '+Y':
                inlet_opening_face_id = region.bottom_face_y.id
                outlet_opening_face_id = region.top_face_y.id
                speed = str(vel) + 'm_per_sec'
                airtemp = str(airtemp) + 'cel'
                assign_opening_BC('inlet', inlet_opening_face_id, flow_type = 'velocity', yvel = speed, temperature = airtemp)
                assign_opening_BC('outlet', outlet_opening_face_id, flow_type = 'pressure')
                ipk.modeler.edit_region_dimensions([50,50,100,100,50,50])
            elif vel_dir == '-Y':
                inlet_opening_face_id = region.top_face_y.id
                outlet_opening_face_id = region.bottom_face_y.id
                speed = str(vel) + 'm_per_sec'
                airtemp = str(airtemp) + 'cel'
                assign_opening_BC('inlet', inlet_opening_face_id, flow_type = 'velocity', yvel = speed, temperature = airtemp)
                assign_opening_BC('outlet', outlet_opening_face_id, flow_type = 'pressure')
                ipk.modeler.edit_region_dimensions([50,50,100,100,50,50])
            elif vel_dir == '+Z':
                inlet_opening_face_id = region.bottom_face_z.id
                outlet_opening_face_id = region.top_face_z.id
                speed = str(vel) + 'm_per_sec'
                airtemp = str(airtemp) + 'cel'
                assign_opening_BC('inlet', inlet_opening_face_id, flow_type = 'velocity', zvel = speed, temperature = airtemp)
                assign_opening_BC('outlet', outlet_opening_face_id, flow_type = 'pressure')
                ipk.modeler.edit_region_dimensions([50,50,50,50,100,100])
            else:
                inlet_opening_face_id = region.top_face_z.id
                outlet_opening_face_id = region.bottom_face_z.id
                speed = str(vel) + 'm_per_sec'
                airtemp = str(airtemp) + 'cel'
                assign_opening_BC('inlet', inlet_opening_face_id, flow_type = 'velocity', zvel = speed, temperature = airtemp)
                assign_opening_BC('outlet', outlet_opening_face_id, flow_type = 'pressure')
                ipk.modeler.edit_region_dimensions([50,50,50,50,100,100])

        if conv_type == 'Natural':
            analysis_setup = 'natural_conv_setup'
            natural_convection_setup(analysis_setup, gravity_dir=gravity_direction, flow_regime='Turbulent', ambient_temp=airtemp)
            for i in ipk.modeler.get_object_faces('Region'):
                outlet_name = 'outlet_' + str(i)
                assign_opening_BC(outlet_name, i, flow_type='pressure')

        # Clear Desktop messages
        desktop.clear_messages()

        # Priority assignments based on volume of objects
        obj_dict = {}
        for i in ipk.modeler.solid_bodies:
            if i != 'Region':
                obj_dict[i] = ipk.modeler.get_object_from_name(i).volume
        vol_sorted_objs = sorted(obj_dict.items(), key=lambda x:x[1], reverse=True)
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


        # Validation Check for Icepak design
        # ipk.odesign.ValidateDesign()
        # x = ipk.odesktop.GetMessages(ipk.project_name, ipk.design_name, 2)
        # desktop.clear_messages()
        # intersecting_objs = re.findall('Parts\s"(.*?)"\sand\s"(.*?)"\sintersect', ''.join(x))
        # p = 2
        # for i in intersecting_objs:
        #     obj_list = [i[0]]
        #     p = p + 1
        #     ipk.mesh.add_priority(1,obj_list,"None",p)
        # ipk.save_project()

        # List of objects
        # list_of_objects = ipk.modeler.solid_bodies
        # model_objects = []
        # for i in list_of_objects:
        #     obj_handle = ipk.modeler.get_object_from_name(i)
        #     if obj_handle.model == True:
        #         model_objects.append(i)
        model_objects = ipk.modeler.model_objects
        model_objects.remove('Region')

        # 3D Components 
        # comp3d_name = ipk.odesign.GetChildObject('3D Modeler').Get3DComponentDefinitionNames()[0]
        # comp3d_instance_name = ipk.odesign.GetChildObject('3D Modeler').Get3DComponentInstanceNames(comp3d_name)[0]
        # comp3d_part_names = list(ipk.odesign.GetChildObject('3D Modeler').Get3DComponentPartNames(comp3d_instance_name))

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
        # pcb_dim_z = sum(dim_z)
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
        mesh_x = mesh_mult_xy*(tx[1][max_val_index_x] + tx[1][max_val_index_x + 1])  
        mesh_y = mesh_mult_xy*(ty[1][max_val_index_y] + ty[1][max_val_index_y + 1])  
        mesh_z = mesh_mult_z*min(dim_z)

        # Find extent of all objects
        # in z-direction
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
        slack_x = 0.1*pcb_dim_x
        slack_y = 0.1*pcb_dim_y
        slack_z = 0.25*z_extent

        # Add mesh region
        meshregion_box = ipk.modeler.create_box([pcb_min_x,pcb_min_y,z_extent_min],[pcb_dim_x,pcb_dim_y,z_extent],'meshregion_all_objs')
        add_slack('meshregion_all_objs', slack_x,slack_x,slack_y,slack_y,slack_z,slack_z)
        meshregion_box.model = False
        mesh_box = 'meshregion_all_objs'
        mesh_region = ipk.mesh.assign_mesh_region([mesh_box],5,False,'meshregion_all_objs')

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
        mesh_region.UniformMeshParametersType="Average"
        mesh_region.DMLMType = "2DMLM_XY"
        mesh_region.Objects = [mesh_box]
        mesh_region.update()

        # Add mesh operation to primitives, mesh level = 2
        mesh_levels_primitives = {}
        for i in primitive_objects:
            mesh_levels_primitives[i] = 2
        ipk.mesh.assign_mesh_level(mesh_levels_primitives,"mesh_levels_primitives")

        # Add mesh operation to primitives, mesh level = 1
        mesh_levels_3dcomps = {}
        for i in pcb_layers:
            mesh_levels_3dcomps[i] = 1
        ipk.mesh.assign_mesh_level(mesh_levels_3dcomps,"mesh_levels_pcb_layers")

        # Global mesh dimensions
        domain = ipk.modeler.get_bounding_dimension()
        # global_max_x = domain[0]/25
        # global_max_y = domain[1]/25
        # global_max_z = domain[2]/25
        global_max_x = 4*mesh_x
        global_max_y = 4*mesh_y
        global_max_z = 4*mesh_z

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

        ipk.modeler.refresh_all_ids()
        ipk.modeler.refresh()

        if analyze_setup:
            ipk.mesh.generate_mesh(analysis_setup)
            # Solve the model.
            num_tasks = num_cores
            ipk.analyze_setup(analysis_setup,num_cores,num_tasks)
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
            os.kill(st.session_state['pid'],signal.SIGTERM)
            files = os.listdir(os.getcwd())
            for file in files:
                if file.endswith('.lock'):
                    os.remove(file)
        except:
            st.warning('⚠️ No active AEDT sessions!')