
import datetime
import json
from dateutil import parser



class ProjectList:
    def __init__(self, projects):
        self.projects = projects
    
    def initialize_me(self):
        project_list = ProjectList([])
        with open('activities.json', 'r') as f:
            data = json.load(f)
            project_list = ProjectList(
                activities = self.get_projects_from_json(data)
            )
        return project_list
    
    def get_projects_from_json(self,data):
        return_list=[]
        for project in data['projects']:
            return_list.append(
                Project(
                    id = project['id'],
                    sub_id = project['sub_id'],
                    activities = self.get_activities_from_json(project)
                )
            )


    def get_activities_from_json(self, data):
        return_list = []
        for activity in data['activities']:
            return_list.append(
                Activity(
                    name = activity['name'],
                    sub_names = self.get_sub_activities_from_json(activity),
                )
            )
        self.activities = return_list
        return return_list
    
    def get_sub_activities_from_json(self, data):
        return_list = []
        for sub_activity in data['sub_names']:
            return_list.append(
                SubActivity(
                    name = sub_activity['sub_name'],
                    time_entries = self.get_time_entires_from_json(sub_activity),
                )
            )
        self.sub_activities = return_list
        return return_list

    def get_time_entires_from_json(self, data):
        return_list = []
        for entry in data['time_entries']:
            return_list.append(
                TimeEntry(
                    start_time = parser.parse(entry['start_time']),
                    end_time = parser.parse(entry['end_time']),
                    days = entry['days'],
                    hours = entry['hours'],
                    minutes = entry['minutes'],
                    seconds = entry['seconds'],
                )
            )
        self.time_entries = return_list
        return return_list
    
    def serialize(self):
        return {
            'projects' : self.projects_to_json()
        }
    
    def projects_to_json(self):
        projects_ = []
        for project in self.projects:
            projects_.append(project.serialize())
        return projects_

class Project:
    def __init__(self, id, sub_id,activities):
        self.id = id
        self.sub_id = sub_id
        self.activities = activities

    def serialize(self):
        return {
            'id' : self.id,
            'sub_id': self.sub_id,
            "activities": self.make_activities_to_json()
            }
    def make_activities_to_json(self):
        activities = []
        for activity in self.activities:
            activities.append(activity.serialize())
        return activities

class ActivityList:
    def __init__(self, activities):
        self.activities = activities
    
    def initialize_me(self):
        activity_list = ActivityList([])
        with open('activities.json', 'r') as f:
            data = json.load(f)
            activity_list = ActivityList(
                activities = self.get_activities_from_json(data)
            )
        return activity_list
    
    def get_activities_from_json(self, data):
        return_list = []
        for activity in data['activities']:
            return_list.append(
                Activity(
                    name = activity['name'],
                    sub_names = self.get_sub_activities_from_json(activity),
                )
            )
        self.activities = return_list
        return return_list
    
    def get_sub_activities_from_json(self, data):
        return_list = []
        for sub_activity in data['sub_names']:
            return_list.append(
                SubActivity(
                    name = sub_activity['sub_name'],
                    time_entries = self.get_time_entires_from_json(sub_activity),
                )
            )
        self.sub_activities = return_list
        return return_list

    def get_time_entires_from_json(self, data):
        return_list = []
        for entry in data['time_entries']:
            return_list.append(
                TimeEntry(
                    start_time = parser.parse(entry['start_time']),
                    end_time = parser.parse(entry['end_time']),
                    #days = entry['days'],
                    #hours = entry['hours'],
                    #minutes = entry['minutes'],
                    seconds = entry['seconds'],
                )
            )
        self.time_entries = return_list
        return return_list
    
    def serialize(self):
        return {
            'activities' : self.activities_to_json()
        }
    
    def activities_to_json(self):
        activities_ = []
        for activity in self.activities:
            activities_.append(activity.serialize())
        return activities_

class Activity:
    def __init__(self, name,sub_names):
        self.name = name
        self.sub_names = sub_names

    def serialize(self):
        return {
            'name' : self.name,
            'sub_names': self.make_sub_activity_entires_to_json()
            }
    def make_sub_activity_entires_to_json(self):
        sub_activities = []
        for sub_activity in self.sub_names:
            sub_activities.append(sub_activity.serialize())
        return sub_activities

class SubActivity:
    def __init__(self, sub_name, time_entries):
        self.sub_name = sub_name
        self.time_entries = time_entries

    def serialize(self):
        return {
                'sub_name' : self.sub_name,
                'time_entries' : self.make_time_entires_to_json()
        }
        

    def make_time_entires_to_json(self):
        time_list = []
        for time in self.time_entries:
            time_list.append(time.serialize())
        return time_list



class TimeEntry:
    def __init__(self, start_time, end_time, seconds):
        self.start_time = start_time
        self.end_time = end_time
        self.total_time = end_time - start_time
        #self.days = days
        #self.hours = hours
        #self.minutes = minutes
        self.seconds = seconds
    
    def _get_specific_times(self):
        self.seconds =  self.total_time.seconds
        #self.hours = self.days * 24 + self.seconds // 3600
        #self.minutes = (self.seconds % 3600) // 60
        self.seconds = self.seconds

    def serialize(self):
        return {
            'start_time' : self.start_time.strftime("%Y-%m-%d %H:%M:%S"),
            'end_time' : self.end_time.strftime("%Y-%m-%d %H:%M:%S"),
            #'days' : self.days,
            #'hours' : self.hours,
            #'minutes' : self.minutes,
            'seconds' : self.seconds
        }