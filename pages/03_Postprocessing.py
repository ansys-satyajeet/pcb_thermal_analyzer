import os
import re
import signal
import pyaedt
import pandas as pd
import streamlit as st
import tkinter as tk
from PIL import Image
from tkinter import filedialog as fd
from ctypes import windll

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")
st.title('📊Postprocessing')

# Fix blur issue in tkinter window panels
windll.shcore.SetProcessDpiAwareness(1)


# Function to get solution name
def get_solution_name():
    # sol_name = st.session_state.ipk.get_setups()[0] + ':' + st.session_state.ipk.post.post_solution_type
    sol_name = st.session_state.ipk.existing_analysis_sweeps[0]
    return sol_name


# Function to get boundary conditions types
def get_boundary_condition_type():
    list_bcs = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildNames()
    thermal_bc_types = {}
    for i in list_bcs:
        type_bc = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildObject(i).GetPropValue('Type')
        if type_bc in thermal_bc_types:
            if not isinstance(thermal_bc_types[type_bc], list):
                thermal_bc_types[type_bc] = [thermal_bc_types[type_bc]]
            thermal_bc_types[type_bc].append(i)
        else:
            thermal_bc_types[type_bc] = [i]
    return thermal_bc_types


# Function to get boundary conditions associated with objects
def get_boundary_condition_association():
    thermal_bcs = {}
    omodule = st.session_state.ipk.odesign.GetModule("BoundarySetup")
    oeditor = st.session_state.ipk.odesign.SetActiveEditor("3D Modeler")
    obj_bcs = ('Solid Block', 'Hollow Block', 'Source')
    face_bcs = ('Network', 'Opening', 'Conducting Plate', 'Grille')
    list_bcs = st.session_state.ipk.odesign.GetChildObject("Thermal").GetChildNames()
    for bc in list_bcs:
        type_bc = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildObject(bc).GetPropValue('Type')
        obj_bc_dict = {}
        if type_bc in obj_bcs:
            block = omodule.GetBoundaryAssignment(bc)
            objname = [oeditor.GetObjectNameByID(x) for x in block]
            obj_bc_dict[type_bc] = objname
        if type_bc in face_bcs:
            sheet = omodule.GetBoundaryAssignment(bc)
            objname = [oeditor.GetObjectNameByFaceID(sheet[0])]
            obj_bc_dict[type_bc] = objname
        thermal_bcs[bc] = obj_bc_dict
    return thermal_bcs


# Function to get monitor point temperatures
def get_monitor_point_temperatures(sol_name):
    mon_point_list = list(st.session_state.ipk.odesign.GetChildObject('Monitor').GetChildNames())
    mon_point_quant = []
    for i in mon_point_list:
        x = i + '.Temperature'
        mon_point_quant.append(x)
    a = ["X:=", ["All"]]
    b = ["X Component:=", "X", "Y Component:=", mon_point_quant]
    mon_point_table = 'Monitor_Point_Temperatures'
    existing_reports = st.session_state.ipk.odesign.GetChildObject('Results').GetChildNames()
    omodule_report = st.session_state.ipk.odesign.GetModule('ReportSetup')
    if mon_point_table in existing_reports:
        omodule_report.DeleteReports(mon_point_table)
    st.session_state.ipk.post.oreportsetup.CreateReport(mon_point_table, "Monitor", "Data Table", sol_name, [], a, b)
    mon_point_temp_file = mon_point_table + '.csv'
    if os.path.exists(os.path.join(os.getcwd(), mon_point_temp_file)):
        os.remove(os.path.join(os.getcwd(), mon_point_temp_file))
    st.session_state.ipk.post.oreportsetup.ExportToFile(mon_point_table, os.path.join(os.getcwd(),
                                                                                      mon_point_temp_file), False)
    df = pd.read_csv(mon_point_temp_file)
    df.drop(columns=df.columns[0], axis=1, inplace=True)
    column_list = list(df.columns)
    renamed_columns = []
    for i in column_list:
        name = i.strip('point_')
        name = name.split(' ')[0]
        name = name.strip('.Temperature')
        renamed_columns.append(name)
    df.columns = renamed_columns
    df = df.transpose()
    df.columns = ['Temperature [C]']
    df.reset_index(inplace=True)
    df = df.rename(columns={'index': 'Point Name'})
    df.to_csv(mon_point_temp_file, index=False)
    df = pd.read_csv(mon_point_temp_file)
    return df


# Function to get junction temperature of network blocks
def get_network_junction_temperatures(sol_name):
    thermal_bcs = get_boundary_condition_type()
    if 'Network' in thermal_bcs:
        network_blocks = thermal_bcs['Network']
        mon_point_quant = []
        for i in network_blocks:
            x = str(i) + '.Internal.Temperature'
            mon_point_quant.append(x)
        a = ["X:=", ["All"]]
        b = ["X Component:=", "X", "Y Component:=", mon_point_quant]
        mon_point_table = 'Network_Junction_Temperatures'
        existing_reports = st.session_state.ipk.odesign.GetChildObject('Results').GetChildNames()
        omodule_report = st.session_state.ipk.odesign.GetModule('ReportSetup')
        if mon_point_table in existing_reports:
            omodule_report.DeleteReports(mon_point_table)
        st.session_state.ipk.post.oreportsetup.CreateReport(mon_point_table, "Monitor", "Data Table", sol_name, [],
                                                            a, b)
        mon_point_temp_file = mon_point_table + '.csv'
        if os.path.exists(os.path.join(os.getcwd(), mon_point_temp_file)):
            os.remove(os.path.join(os.getcwd(), mon_point_temp_file))
        st.session_state.ipk.post.oreportsetup.ExportToFile(mon_point_table,
                                                            os.path.join(os.getcwd(), mon_point_temp_file),
                                                            False)
        df = pd.read_csv(mon_point_temp_file)
        df.drop(columns=df.columns[0], axis=1, inplace=True)
        column_list = list(df.columns)
        renamed_columns = []
        for i in column_list:
            name = i.strip('.Internal.Temperature [cel]')
            renamed_columns.append(name)
        df.columns = renamed_columns
        df = df.transpose()
        df.columns = ['Temperature [C]']
        df.reset_index(inplace=True)
        df = df.rename(columns={'index': 'Network Junction'})
        df.to_csv(mon_point_temp_file, index=False)
        df = pd.read_csv(mon_point_temp_file)
        return df
    else:
        error_message = "No network blocks in the model!"
        return error_message


# Function to get maximum temperature of objects
def get_object_max_temperatures(sol_name):
    obj_list = st.session_state.ipk.modeler.model_objects
    if 'Region' in obj_list:
        try:
            obj_list.remove('Region')
        except RuntimeError:
            print('Region not present in obj_list')
    obj_bcs = get_boundary_condition_association()
    solid_blocks = []
    hollow_blocks = []
    for i in obj_bcs:
        for j in obj_bcs[i]:
            if j == 'Solid Block':
                for k in obj_bcs[i][j]:
                    solid_blocks.append(k)
            elif j == 'Hollow Block':
                for k in obj_bcs[i][j]:
                    hollow_blocks.append(k)
            else:
                pass
    if st.session_state.ipk.odesign.GetChildObject('3D Modeler').Get3DComponentDefinitionNames():
        comp3d_name = st.session_state.ipk.odesign.GetChildObject('3D Modeler').Get3DComponentDefinitionNames()[0]
        comp3d_instance_name = st.session_state.ipk.odesign.GetChildObject('3D Modeler'). \
            Get3DComponentInstanceNames(comp3d_name)[0]
        comp3d_part_names = list(
            st.session_state.ipk.odesign.GetChildObject('3D Modeler').Get3DComponentPartNames(comp3d_instance_name))
        solid_blocks = solid_blocks + comp3d_part_names
    calc_expr = []
    omodule = st.session_state.ipk.odesign.GetModule("FieldsReporter")
    omodule.CalcStack("clear")
    for i in solid_blocks:
        st.session_state.ipk.post.ofieldsreporter.EnterQty('Temp')
        st.session_state.ipk.post.ofieldsreporter.EnterVol(i)
        st.session_state.ipk.post.ofieldsreporter.CalcOp('Maximum')
        named_expr = i
        if omodule.DoesNamedExpressionExists(named_expr):
            omodule.DeleteNamedExpr(named_expr)
        st.session_state.ipk.post.ofieldsreporter.AddNamedExpression(named_expr, 'Fields')
        calc_expr.append(named_expr)
    for i in hollow_blocks:
        st.session_state.ipk.post.ofieldsreporter.EnterQty('Temp')
        st.session_state.ipk.post.ofieldsreporter.EnterSurf(i)
        st.session_state.ipk.post.ofieldsreporter.CalcOp('Maximum')
        named_expr = i
        if omodule.DoesNamedExpressionExists(named_expr):
            omodule.DeleteNamedExpr(named_expr)
        st.session_state.ipk.post.ofieldsreporter.AddNamedExpression(named_expr, 'Fields')
        calc_expr.append(named_expr)
    a = ["X:=", ["All"]]
    b = ["X Component:=", "X", "Y Component:=", calc_expr]
    obj_max_temp = 'Object_Max_Temperatures'

    existing_reports = st.session_state.ipk.odesign.GetChildObject('Results').GetChildNames()
    omodule_report = st.session_state.ipk.odesign.GetModule('ReportSetup')
    if obj_max_temp in existing_reports:
        omodule_report.DeleteReports(obj_max_temp)
    st.session_state.ipk.post.oreportsetup.CreateReport(obj_max_temp, "Fields", "Data Table", sol_name, [], a, b)

    obj_max_temp_file = obj_max_temp + '.csv'
    if os.path.exists(os.path.join(os.getcwd(), obj_max_temp_file)):
        os.remove(os.path.join(os.getcwd(), obj_max_temp_file))
    st.session_state.ipk.post.oreportsetup.ExportToFile(obj_max_temp, os.path.join(os.getcwd(), obj_max_temp_file),
                                                        False)
    df = pd.read_csv(obj_max_temp_file)
    df.drop(columns=df.columns[0], axis=1, inplace=True)
    column_list = list(df.columns)
    renamed_columns = []
    for i in column_list:
        name = i.strip('[]')
        renamed_columns.append(name)
    df.columns = renamed_columns
    df = df.transpose()
    df.columns = ['Temperature [C]']
    df.reset_index(inplace=True)
    df = df.rename(columns={'index': 'Object'})
    df.to_csv(obj_max_temp_file, index=False)
    df = pd.read_csv(obj_max_temp_file)
    return df


# Function to plot contours of temperature on PCB layers
def get_temperature_contours_on_pcb_layers(sol_name):
    pcb = st.session_state.ipk.modeler.primitives.user_defined_component_names
    pcb_layers = sorted(st.session_state.ipk.modeler.get_3d_component_object_list(pcb[0]))
    pcb_layer_temps = st.session_state.ipk.post.create_fieldplot_surface(pcb_layers, "Temperature", sol_name,
                                                                         plot_name="Temperature_on_PCB_layers")
    path_image = pcb_layer_temps.export_image(os.path.join(os.getcwd(), "Temperature_on_PCB_layers.png"))
    return path_image


# Function to plot contours of temperature on all objects in the model
def get_temperature_contours_on_all_objects(sol_name):
    model_objects = st.session_state.ipk.modeler.model_objects
    if 'Region' in model_objects:
        try:
            model_objects.remove('Region')
        except RuntimeError:
            print('Region not present in object list')
    temp_all_objs = st.session_state.ipk.post.create_fieldplot_surface(model_objects, "Temperature", sol_name,
                                                                       plot_name="Temperature_on_all_objects")
    path_image = temp_all_objs.export_image(os.path.join(os.getcwd(), "Temperature_on_all_objects.png"))
    return path_image


# Function to get board side heat flux for objects touching the PCB
def get_object_board_side_heat_flux(sol_name):
    model_objects = st.session_state.ipk.modeler.model_objects
    if 'Region' in model_objects:
        try:
            model_objects.remove('Region')
        except RuntimeError:
            print('Region not present in obj_list')
    pcb = st.session_state.ipk.modeler.primitives.user_defined_component_names
    pcb_layers = sorted(st.session_state.ipk.modeler.get_3d_component_object_list(pcb[0]))
    components = [x for x in model_objects if x not in pcb_layers]
    pcb_top_layer = st.session_state.ipk.modeler.get_object_from_name(pcb_layers[0])
    pcb_bot_layer = st.session_state.ipk.modeler.get_object_from_name(pcb_layers[-1])
    board_side_faces = {}
    for comp in components:
        obj_handle = st.session_state.ipk.modeler.get_object_from_name(comp)
        if obj_handle.get_touching_faces(pcb_top_layer):
            board_side_faces[comp] = str(obj_handle.get_touching_faces(pcb_top_layer)[0])
        else:
            board_side_faces[comp] = str(obj_handle.get_touching_faces(pcb_bot_layer)[0])
    board_side_face_list = []
    for key in board_side_faces:
        face_id = board_side_faces[key].split(' ')[1]
        face_name = key + '_board_side'
        st.session_state.ipk.modeler.create_face_list([face_id], name=face_name)
        board_side_face_list.append(face_name)
    calc_expr = []
    report_module = st.session_state.ipk.odesign.GetModule("FieldsReporter")
    report_module.CalcStack("clear")
    for i in board_side_face_list:
        report_module.EnterQty("Heat_Flux")
        report_module.EnterSurf(i)
        report_module.CalcOp("Integrate")
        named_expr = str(i) + "_heat_flux"
        if report_module.DoesNamedExpressionExists(named_expr):
            report_module.DeleteNamedExpr(named_expr)
        st.session_state.ipk.post.ofieldsreporter.AddNamedExpression(named_expr, 'Fields')
        calc_expr.append(named_expr)
    a = ["X:=", ["All"]]
    b = ["X Component:=", "X", "Y Component:=", calc_expr]
    obj_board_side_heat_flux = 'Object_Board_Side_Heat_Flux'
    existing_reports = st.session_state.ipk.odesign.GetChildObject('Results').GetChildNames()
    omodule_report = st.session_state.ipk.odesign.GetModule('ReportSetup')
    if obj_board_side_heat_flux in existing_reports:
        omodule_report.DeleteReports(obj_board_side_heat_flux)
    st.session_state.ipk.post.oreportsetup.CreateReport(obj_board_side_heat_flux, "Fields", "Data Table",
                                                        sol_name, [], a, b)
    obj_board_side_heat_flux_file = obj_board_side_heat_flux + '.csv'
    if os.path.exists(os.path.join(os.getcwd(), obj_board_side_heat_flux_file)):
        os.remove(os.path.join(os.getcwd(), obj_board_side_heat_flux_file))
    st.session_state.ipk.post.oreportsetup.ExportToFile(obj_board_side_heat_flux,
                                                        os.path.join(os.getcwd(), obj_board_side_heat_flux_file), False)
    df = pd.read_csv(obj_board_side_heat_flux_file)
    df.drop(columns=df.columns[0], axis=1, inplace=True)
    column_list = list(df.columns)
    renamed_columns = []
    for i in column_list:
        name = i.strip('_board_side_heat_flux []')
        renamed_columns.append(name)
    df.columns = renamed_columns
    df = df.transpose()
    df.columns = ['Heat Flow [W]']
    df.reset_index(inplace=True)
    df = df.rename(columns={'index': 'Object'})
    df.to_csv(obj_board_side_heat_flux_file, index=False)
    df = pd.read_csv(obj_board_side_heat_flux_file)
    return df


def quit_aedt():
    if st.session_state.desktop:
        st.session_state.ipk.save_project()
        pid = st.session_state.desktop.aedt_process_id
        os.kill(pid, signal.SIGTERM)
        file_list = os.listdir(os.getcwd())
        for file in file_list:
            if file.endswith('.lock'):
                os.remove(file)
    else:
        st.warning('⚠️ No active AEDT sessions open!')


if 'desktop' not in st.session_state:
    st.session_state.desktop = False
if 'ipk' not in st.session_state:
    st.session_state.ipk = False
if 'project' not in st.session_state:
    st.session_state.project = False
if 'launch_aedt' not in st.session_state:
    st.session_state.launch_aedt = False
if 'create_report' not in st.session_state:
    st.session_state.create_report = False
if 'post_quant' not in st.session_state:
    st.session_state.post_quant = False
if 'close_aedt' not in st.session_state:
    st.session_state.close_aedt = False
if 'workdir' not in st.session_state:
    st.session_state.workdir = False

c1, c2 = st.columns([1, 2])
aedt_version = c1.selectbox('Select AEDT Release:', ('2023 R1', '2023 R2'))
c2.write('Select AEDT project file (*.aedt)')
aedt_project_button = c2.button('Select AEDT Project')
if aedt_project_button:
    st.session_state.project = True
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    try:
        files = fd.askopenfilenames(parent=root, initialdir=os.getcwd(),
                                    filetypes=[('Ansys Electronics Desktop File', '*.aedt')])
        st.session_state.project = os.path.abspath(files[0])
    except RuntimeError:
        st.session_state.project = False

if st.session_state.project:
    st.session_state.workdir = os.path.dirname(os.path.abspath(st.session_state.project))
    os.chdir(st.session_state.workdir)

placeholder = st.empty()
if st.session_state.project:
    placeholder.markdown(f'''**Selected AEDT Project File:** ```{st.session_state.project}```''')

st.session_state.launch_aedt = st.button('Launch AEDT')

st.markdown('---')

aedt_release = re.sub(' R', '.', aedt_version)
post_tuple = ('Monitor Point Temperatures', 'Network Junction Temperatures', 'Object Temperatures',
              'Temperature Contours on PCB Layers', 'Temperature Contours on Entire Model',
              'Heat Flow Rates at Object-PCB Interfaces')
st.session_state.post_quant = st.selectbox('Postprocessing selection:', post_tuple)

st.session_state.create_report = st.button('Create Report')

if st.session_state.launch_aedt:
    if os.path.exists(os.path.join(os.getcwd(), st.session_state.project + ".lock")):
        os.remove(os.path.join(os.getcwd(), st.session_state.project + ".lock"))
    st.session_state.desktop = pyaedt.Desktop(aedt_release)
    st.session_state.ipk = pyaedt.Icepak(st.session_state.project)

if st.session_state.create_report and st.session_state.desktop:
    solution_name = get_solution_name()
    if st.session_state.post_quant == 'Monitor Point Temperatures':
        mon_df = get_monitor_point_temperatures(solution_name)
        st.dataframe(mon_df)
    elif st.session_state.post_quant == 'Network Junction Temperatures':
        net_df = get_network_junction_temperatures(solution_name)
        st.dataframe(net_df)
    elif st.session_state.post_quant == 'Object Temperatures':
        obj_temp_df = get_object_max_temperatures(solution_name)
        st.dataframe(obj_temp_df)
    elif st.session_state.post_quant == 'Temperature Contours on PCB Layers':
        if os.path.exists(os.path.abspath(os.path.join(st.session_state.workdir, 'Temperature_on_PCB_layers.png'))):
            image_path = os.path.abspath(os.path.join(st.session_state.workdir, 'Temperature_on_PCB_layers.png'))
        else:
            image_path = get_temperature_contours_on_pcb_layers(solution_name)
        image = Image.open(image_path)
        st.image(image, caption='Temperature Contours on PCB Layers')
    elif st.session_state.post_quant == 'Temperature Contours on Entire Model':
        if os.path.exists(os.path.abspath(os.path.join(st.session_state.workdir, 'Temperature_on_all_objects.png'))):
            image_path = os.path.abspath(os.path.join(st.session_state.workdir, 'Temperature_on_all_objects.png'))
        else:
            image_path = get_temperature_contours_on_all_objects(solution_name)
        image = Image.open(image_path)
        st.image(image, caption='Temperature Contours on Entire Model')
    elif st.session_state.post_quant == 'Heat Flow Rates at Object-PCB Interfaces':
        solution_name = get_solution_name()
        heat_flux_df = get_object_board_side_heat_flux(solution_name)
        st.dataframe(heat_flux_df)
    else:
        pass

st.session_state.close_aedt = st.button('Save and Quit AEDT')
if st.session_state.close_aedt:
    quit_aedt()
elif st.session_state.desktop:
    st.warning(f'AEDT session open @Port: {st.session_state.desktop.aedt_process_id}')
else:
    st.warning('⚠️ No active AEDT sessions open!')
