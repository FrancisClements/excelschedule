import schedule as s, json, os
from datetime import datetime as dt
from tkinter import *

#deals with the JSON data
class JSONImport:
    def __init__(self):
        self.filename = 'settings.json'
        print(__file__)
        #checks if the file exists (required)
        if not os.path.isfile(self.filename):
            raise Exception('JSON file is missing.')

        #opens and reads the file (which acts as the template)
        with open(self.filename, 'r') as f:
            self.data = json.load(f)

    def write_json(self):
        with open(self.filename, 'w') as f:
            #timestamp
            now = dt.today()
            str_format = '%m/%d/%Y at %I:%M%p'
            self.data['_metadata'] = now.strftime(str_format)

            json.dump(self.data, f, indent=3, sort_keys=True)
            f.truncate()

#program window
class Program:
    def __init__(self):
        self.s = JSONImport()
        self.s.write_json()
        print('UPDATESS')
        # root = Tk()

        # label = Label(root, text = 'Hello World')
        # label.pack()

#creates the schedule
# s.create_schedule()
p = Program()