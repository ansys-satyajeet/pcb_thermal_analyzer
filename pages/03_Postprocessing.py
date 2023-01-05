import os
import signal
import pyaedt
import pandas as pd
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from ctypes import windll

st.set_page_config(layout="centered", page_icon="🌡️", page_title="PCB Thermal Analyzer")
st.title('📊Postprocessing')

# Fix blur issue in tkinter window panels
windll.shcore.SetProcessDpiAwareness(1)


def get_solution_name():
    sol_name = st.session_state.ipk.get_setups()[0] + ':' + st.session_state.ipk.post.post_solution_type
    return sol_name


def get_boundary_condition_association():
    list_bcs = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildNames()
    thermal_bc_types = {}
    for i in list_bcs:
        type_bc = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildObject(i).GetPropValue('Type')
        if type_bc in thermal_bc_types:
            if not isinstance(thermal_bc_types[type_bc], list):
                thermal_bc_types[type_bc] = [thermal_bc_types[type_bc]]
            thermal_bc_types[type_bc].append(i)
        else:
            thermal_bc_types[type_bc] = i
    return thermal_bc_types


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


def get_network_junction_temperatures(sol_name):
    list_bcs = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildNames()
    thermal_bc_types = {}
    for i in list_bcs:
        # l = []
        type_bc = st.session_state.ipk.odesign.GetChildObject('Thermal').GetChildObject(i).GetPropValue('Type')
        if type_bc in thermal_bc_types:
            if not isinstance(thermal_bc_types[type_bc], list):
                thermal_bc_types[type_bc] = [thermal_bc_types[type_bc]]
            thermal_bc_types[type_bc].append(i)
        else:
            thermal_bc_types[type_bc] = i
    mon_point_quant = []
    for i in thermal_bc_types['Network']:
        x = i + '.Internal.Temperature'
        mon_point_quant.append(x)
    a = ["X:=", ["All"]]
    b = ["X Component:=", "X", "Y Component:=", mon_point_quant]
    mon_point_table = 'Network_Junction_Temperatures'
    existing_reports = st.session_state.ipk.odesign.GetChildObject('Results').GetChildNames()
    omodule_report = st.session_state.ipk.odesign.GetModule('ReportSetup')
    if mon_point_table in existing_reports:
        omodule_report.DeleteReports(mon_point_table)
    st.session_state.ipk.post.oreportsetup.CreateReport(mon_point_table, "Monitor", "Data Table", sol_name, [], a, b)
    mon_point_temp_file = mon_point_table + '.csv'
    if os.path.exists(os.path.join(os.getcwd(), mon_point_temp_file)):
        os.remove(os.path.join(os.getcwd(), mon_point_temp_file))
    st.session_state.ipk.post.oreportsetup.ExportToFile(mon_point_table, os.path.join(os.getcwd(), mon_point_temp_file),
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


def get_object_max_temperatures(sol_name):
    obj_list = st.session_state.ipk.modeler.model_objects
    if 'Region' in obj_list:
        try:
            obj_list.remove('Region')
        except RuntimeError:
            print('Region not present in obj_list')
    thermal_bcs = get_boundary_condition_association()
    network_blocks = thermal_bcs['Network']
    hollow_blocks = thermal_bcs['Hollow Block']
    solid_blocks = [x for x in obj_list if x not in network_blocks]
    solid_blocks = [x for x in solid_blocks if x not in hollow_blocks]
    comp3d_name = st.session_state.ipk.odesign.GetChildObject('3D Modeler').Get3DComponentDefinitionNames()[0]
    comp3d_instance_name = st.session_state.ipk.odesign.GetChildObject('3D Modeler').\
        Get3DComponentInstanceNames(comp3d_name)[0]
    comp3d_part_names = list(
        st.session_state.ipk.odesign.GetChildObject('3D Modeler').Get3DComponentPartNames(comp3d_instance_name))
    solid_blocks = [x for x in solid_blocks if x not in comp3d_part_names]
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

c1, c2 = st.columns([1, 2])
aedt_version = c1.selectbox('Select AEDT Release:', ('2022 R2', '2023 R1'))
c2.write('Select AEDT project file (*.aedt)')
aedt_project_button = c2.button('Select AEDT Project')
if aedt_project_button:
    st.session_state.project = True
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    try:
        files = filedialog.askopenfilenames(master=root, filetypes=[('Ansys Electronics Desktop File', '*.aedt')])
        st.session_state.project = os.path.abspath(files[0])
    except RuntimeError:
        st.session_state.project = False

if st.session_state.project:
    workdir = os.path.dirname(os.path.abspath(st.session_state.project))
    os.chdir(workdir)

placeholder = st.empty()
if st.session_state.project:
    placeholder.markdown(f'''**Selected AEDT Project File:** ```{st.session_state.project}```''')

st.session_state.launch_aedt = st.button('Launch AEDT')

st.markdown('---')

if aedt_version == '2022 R2':
    aedt_release = '2022.2'
elif aedt_version == '2023 R1':
    aedt_release = '2023.1'
else:
    aedt_release = '2022.2'

post_tuple = ('Monitor Point Temperatures', 'Network Junction Temperatures', 'Object Temperatures',
              'Temperature Contours on PCB Layers', 'Temperature Contours on Entire Model',
              'Heat Flow Rates at Object-PCB Interfaces')
st.session_state.post_quant = st.selectbox('Postprocessing selection:', post_tuple)

st.session_state.create_report = st.button('Create Report')

if st.session_state.launch_aedt and st.session_state.desktop:
    if os.path.exists(os.path.join(os.getcwd(), st.session_state.project + ".lock")):
        os.remove(os.path.join(os.getcwd(), st.session_state.project + ".lock"))
    st.session_state.desktop = pyaedt.Desktop(aedt_release)
    st.session_state.ipk = pyaedt.Icepak(st.session_state.project)

if st.session_state.create_report and st.session_state.desktop:
    if st.session_state.post_quant == 'Monitor Point Temperatures':
        solution_name = get_solution_name()
        mon_df = get_monitor_point_temperatures(solution_name)
        st.dataframe(mon_df)
    elif st.session_state.post_quant == 'Network Junction Temperatures':
        solution_name = get_solution_name()
        net_df = get_network_junction_temperatures(solution_name)
        st.dataframe(net_df)
    elif st.session_state.post_quant == 'Object Temperatures':
        solution_name = get_solution_name()
        obj_temp_df = get_object_max_temperatures(solution_name)
        st.dataframe(obj_temp_df)
    else:
        pass

st.session_state.close_aedt = st.button('Save and Quit AEDT')
if st.session_state.close_aedt:
    quit_aedt()
else:
    st.warning('⚠️ No active AEDT sessions open!')
