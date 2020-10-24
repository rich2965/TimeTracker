from ctypes import wintypes, windll, create_unicode_buffer
from activity import *
import time
import datetime
import json
import sys
import PySimpleGUI as sg 
import threading
from os import system
import pandas as pd
from pandas.io.json import json_normalize


active_window_name = ""
active_window_subname= ""
activity_name = ""
activity_sub_name = ""

first_time = True
thread_1 = ""
stop_threads = False
timer_running, counter = False, 0

project_list = ProjectList([])
full_data_table = []
list_of_projects = []
header_list=["Project","Time"]



#Loading Previously Used Project Codes 
try:
    with open('list_of_projects.json', 'r') as json_file:
        data = json.load(json_file)
        list_of_projects = data["list_of_projects"]


        for project in list_of_projects:
            loaded_project = Project(project, "300",[])
            project_list.projects.append(loaded_project)
    


except Exception:
    print('Unable to load previous project codes')


#Data objects for reading recorded data 
data_table_columns = { 'project_id':[],
'sub_id':[],
'activity_name': [],
'sub_activity_name': [],
'time_seconds':[],
'time_start':[],
'time_end':[]
}

data_table_entry = { 'project_id':'',
'sub_id':'',
'activity_name': '',
'sub_activity_name': '',
'time_seconds':'',
'time_start':'',
'time_end':''
}


# default settings for GUI

sg.ChangeLookAndFeel('DarkBlue3') 



#Table Build out
tab1_layout = [
[sg.Table(values=[],
                  headings=header_list,
                  size = (500,500),
                  display_row_numbers=False,
                  auto_size_columns=False,
                  num_rows=10,
                  key='Table')]
    ]

tab2_layout = [
[sg.InputOptionMenu(list_of_projects,key="-PROJECT-"),sg.Button('Search', button_color=('white', 'grey'),focus=True, key='Search')],
[sg.Table(values=[],
                  headings=header_list,
                  size = (500,500),
                  col_widths= 50,
                  display_row_numbers=False,
                  auto_size_columns=False,
                  num_rows=10,
                  key='Table2')]

]

#Column 1 Build Out

col1 = sg.Column([
    [sg.Frame(layout=[[sg.Text(size=(10, 2),  justification='center',background_color = 'black',  key='-CURRENT-')]
    ,[sg.Listbox(values=list_of_projects, size=(20, 12), key='-LIST-', enable_events=True)]
    ],
              title='Current Project:')]
    
    
    ],vertical_alignment ='top'

)

#Column 2 Build Out
col2 = sg.Column([
    [ sg.Text(size=(15, 2),  justification='center')  ,sg.Frame(layout=[
    [ sg.Text('None',size=(10, 2),  justification='center',background_color = 'black',  key='-OUTPUT-')]  
    ],title='Work Time:')]
    ,
    [sg.Frame(layout=[
    [sg.TabGroup([[sg.Tab('Time Breakdown', tab1_layout), sg.Tab('Project Breakdown', tab2_layout)]])]  
    ],title='Data:')]
    ,
    [sg.Button('Refresh Data', button_color=('white', 'grey'),focus=True, key='Refresh')]

    
    
    ],vertical_alignment ='top'
    

)



layout: list = [
    [col1,col2],
    #[sg.Button('Load Previous Projects', button_color=('white', 'grey'),focus=True, key='Load')],
    [sg.Button('Add Project', button_color=('white', 'grey'),focus=True, key='Add'), sg.Button('Remove Project', button_color=('white', 'grey'),focus=True, key='Remove')],    
            [sg.Button('Start', button_color=('white', 'springgreen4'),focus=True, key='Start',disabled = True),      
            sg.Button('Pause', button_color=('white', 'black'), key='Pause', disabled=True),      
            #sg.Button('Resume', button_color=('white', 'firebrick3'), key='Resume'),      
            sg.Button('Exit', button_color=('white', 'firebrick3'), key='Exit')] 
                 
]

window: object = sg.Window('Time Tracker', layout=layout, resizable=True)





def add_project(project_id,project_sub_id):
    global project_list
    project = Project(project_id, project_sub_id,[])
    project_list.projects.append(project)


def load_projects_list():
    list_of_projects = []
    for project in project_list.projects:
        list_of_projects.append(project.id)
    
    return list_of_projects


#-----functions for tracking time
def get_Activity_Window():
    _active_window_name = None
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        hWnd = windll.user32.GetForegroundWindow()
        length = windll.user32.GetWindowTextLengthW(hWnd)
        buf = create_unicode_buffer(length+1)
        windll.user32.GetWindowTextW(hWnd, buf, length+1)
        _active_window_name = buf.value
    else:
        print("sys.platform={platform} is not supported.".format(platform=sys.platform))
        print(sys.version)
    return _active_window_name

def activity_name_splitter(activity_name):
    split_name = activity_name.split('-')
    if len(split_name) > 1:
        return split_name[-1].lstrip(), split_name[-2].rstrip().lstrip()
    else:
        return split_name[-1].lstrip(), None

def timetracker(project_id,project_sub_id,project_list,activity_name, activity_sub_name,first_time,start_time, new_window_name, new_window_subname ): 
    if not first_time:
        end_time = datetime.datetime.now()
        time_entry = TimeEntry(start_time,end_time,0)
        time_entry._get_specific_times()

        exists = False
        sub_activity_exists = False
        for project in project_list.projects:
            if project.id == project_id and project.sub_id == project_sub_id:
                #print("Project Found. ID: " + project_id)

                for activity in project.activities:
                    if activity.name == activity_name:  
                        exists = True
                        for sub_activity in activity.sub_names:
                            if  sub_activity.sub_name == activity_sub_name:
                                sub_activity.time_entries.append(time_entry)
                                sub_activity_exists = True

                        if not sub_activity_exists:
                            new_sub_activity = SubActivity(activity_sub_name,[time_entry])
                            activity.sub_names.append(new_sub_activity)
                                    

                if not exists:
                    sub_activity = SubActivity(activity_sub_name,[time_entry])
                    activity = Activity(activity_name,[sub_activity])
                    #print(activity.serialize())
                    project.activities.append(activity)
                with open('activities.json', 'w') as json_file:
                    json.dump(project_list.serialize(), json_file, indent =4, sort_keys =True)
                    start_time = datetime.datetime.now()

                
    first_time = False
    active_window_name = new_window_name 
    active_window_subname = new_window_subname
    return first_time, active_window_name, active_window_subname,start_time

def calculate_time_spent(data_table):
    sum_table = data_table.groupby(['project_id'], as_index = False)['time_seconds'].sum()
    sum_table['time_seconds'] = pd.to_datetime(sum_table['time_seconds'], unit ='s').dt.strftime('%H:%M:%S')
    sum_table.columns = ['Projects','Time']
    return (sum_table.values.tolist())

def calculate_time_spent_project(project_id):
    project_table = full_data_table.groupby(['project_id'], as_index = False).get_group(project_id).groupby('activity_name', as_index = False)['time_seconds'].sum()
    project_table["time_seconds"] = pd.to_datetime(project_table['time_seconds'], unit ='s').dt.strftime('%H:%M:%S')
    project_table.columns = ['Activities','Time']
    return (project_table.values.tolist())


def load_full_data_to_table():
    with open('activities.json', 'r') as f:
        data = json.load(f)
    data_table = pd.DataFrame(data_table_columns)

    for project in data['projects']:
            for activity in project['activities']:
                for sub_activity in activity['sub_names']:
                    for time in sub_activity['time_entries']:
                        data_table_entry['project_id'] =project['id']
                        data_table_entry['sub_id'] =project['sub_id']
                        data_table_entry['activity_name'] =activity['name']
                        data_table_entry['sub_activity_name'] =sub_activity['sub_name']
                        data_table_entry['time_seconds'] = time['seconds']
                        data_table_entry['time_start'] = time['start_time']
                        data_table_entry['time_end'] = time['end_time']
                        data_table = data_table.append(data_table_entry,ignore_index = True)

    return (data_table)



def main_tracker():
    global active_window_name 
    global active_window_subname
    global activity_name
    global activity_sub_name 
    global project_list
    global first_time 
    
    start_time = datetime.datetime.now()

    try:
        while True:
            new_window_name, new_window_subname = activity_name_splitter(get_Activity_Window()) 
            global stop_threads
            if  active_window_name != new_window_name or active_window_subname != new_window_subname: 
                activity_name = active_window_name
                activity_sub_name = active_window_subname
                first_time, active_window_name, active_window_subname,start_time = timetracker(values['-LIST-'][0], "300", project_list,activity_name, activity_sub_name, first_time, start_time,new_window_name,new_window_subname )    


            time.sleep(1)

            if stop_threads:
                timetracker(values['-LIST-'][0], "300", project_list,new_window_name, new_window_subname, first_time, start_time,new_window_name,new_window_subname )
                break
            


    except KeyboardInterrupt:
        with open('activities.json', 'w') as json_file:
            json.dump(project_list.serialize(), json_file, indent=4, sort_keys=True)


##-----MAIN EVENT LOOP-----------------------------------##
while True:
    event, values = window.read(timeout=1000)
    #print(event)
    #print(values['-PROJECT-'])
    if event in (sg.WIN_CLOSED, 'Exit'):             # if user closed the window using X or clicked Exit button
        stop_threads = True
        sg.popup_timed('Exiting...', auto_close_duration = 2)
        if list_of_projects:
            with open('list_of_projects.json', 'w') as json_file:
                saved_project_list = {"list_of_projects":list_of_projects}
                json.dump(saved_project_list, json_file, indent=4, sort_keys=True)
        break
    


    if event == '-LIST-':
        window['Start'].update(disabled=False)
        window['Remove'].update(disabled=False)

    

        

    if event == "Add":
        new_project = sg.popup_get_text('Add Project:', title='New Project')
        if new_project not in [None,''] and new_project not in list_of_projects:    
            add_project(new_project,"300")
            list_of_projects = load_projects_list()
            window['-LIST-'].update(values=list_of_projects)
        else:
            sg.popup("No new project added.")
        window['Start'].update(disabled=True)

    if event == "Remove":
       # current_selection_index = window.Element('-LIST-').Widget.curselection()[0] #This pulls the index number of the list item selected
        current_select_value = values['-LIST-'][0]
        for project in project_list.projects:
            if project.id == current_select_value and project.activities==[] :
                project_list.projects.remove(project)
                list_of_projects = load_projects_list()
                window['-LIST-'].update(values=list_of_projects)
            elif project.id == current_select_value and project.activities :
                sg.popup("Can't remove project as time has already been logged against it")
        window['Remove'].update(disabled=True)
        #list_of_projects.pop(current_selection)
        


    if event =='Start':
        timer_running = True
        stop_threads = False
        window['Start'].update(disabled=True)
        window['Pause'].update(disabled=False)
        window['-LIST-'].update(disabled=True)
        window['Add'].update(disabled=True)
        window['Remove'].update(disabled=True)
        window['Refresh'].update(disabled=True)
        #window['Load'].update(disabled=True)
        window['Start'].update('Resume')
        window['-CURRENT-'].update(values['-LIST-'][0])
        thread1 = threading.Thread(target= main_tracker)
        thread1.start()

    if timer_running:
        window['-OUTPUT-'].update('{:02d}:{:02d}:{:02d}'.format((counter // 3600) % 24, (counter // 60) % 60, counter % 60))
        counter += 1

    if event =='Refresh':
        full_data_table = load_full_data_to_table()
        project_stats = calculate_time_spent(full_data_table)
        window['Table'].update(project_stats)
        #window['Refresh'].update(disabled=True)

    if event =='Search':
    
        if values['-PROJECT-'] not in full_data_table['project_id'].values.tolist():
            sg.popup("No Activities Found for project: "+values['-PROJECT-'] )
        else:
            selected_project_stats = calculate_time_spent_project(values['-PROJECT-'])
            window['Table2'].update(selected_project_stats)
        #window['Refresh'].update(disabled=True)

    if event == 'Pause':
        counter -= 1
        window['Start'].update(disabled=False)
        window['Pause'].update(disabled=True)
        window['-LIST-'].update(disabled=False)
        window['Add'].update(disabled=False)
        window['Remove'].update(disabled=False)
        window['Refresh'].update(disabled=False)
        #window['Load'].update(disabled=False)
        window['-CURRENT-'].update('None')
        stop_threads = True
        timer_running = not timer_running
    

window.close()


        



