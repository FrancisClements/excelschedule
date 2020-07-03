import schedule, json, os
from datetime import datetime as dt
from tkinter import *
import tkinter.ttk as ttk
from tkcolorpicker import askcolor

#deals with the JSON data
class JSONImport:
    def __init__(self):
        self.filename = 'settings.json'
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

            json.dump(self.data, f, indent=4, sort_keys=True)
            f.truncate()

#initializing the methods for tkinter
class Kinter:
    def __init__(self, root):
        self.root = root
        self.new()

    def new(self):
        #presets
        self.btn_size = {
            "default": {'x':30, 'y':10}
        }

    def label(self, pos, item):
        #assign l as blank if item is None
        def blank():
            l = Label(self.root, text ='         ')
            return l 

        #initializes the label
        l = blank() if item == None else Label(self.root, text = item)
        #prints the label
        l.grid(row = pos[1], column = pos[0])

    def button(self, pos, item, size = 'default', state_type = None, cmd = None):
        #initializes the button
        b = Button(self.root, text = item, state = state_type,
                padx = self.btn_size[size]['x'], pady = self.btn_size[size]['y'],
                command = cmd)

        #prints the button
        b.grid(row = pos[1], column = pos[0])

#program window
class Program:
    def __init__(self):
        self.j = JSONImport()
        self.root = Tk()
        self.root.title('Excel Schedule Maker')
        self.k = Kinter(self.root)
        
        self.show_main_menu()
        self.run()
        print('App Exit')

    def test_label(self):
        self.k.label((2,5), 'Button is clicked')    

    def show_main_menu(self): 
        self.k.label((1,0), 'This is grid1')
        self.k.label((1,1), 'This is grid2')
        self.k.label((1, 2), None)
        self.k.label((1, 3), 'This is grid3')
        self.k.button((2, 4), 'CLICK', cmd = self.test_label)
 
    def run(self):
        self.root.mainloop()
    
    def create_schedule(self):
        schedule.create_schedule()

if __name__ == "__main__":
    p = Program()
    # r = Tk()
    # r.geometry("800x400")

    # t = Text(r, height=20, width=40)
    # t.pack()

    # r.mainloop()

    # root = Tk()
    # style = ttk.Style(root)
    # style.theme_use('clam')
    # hex_code, RGB_code = askcolor((255, 255, 0), root) 
    # print(hex_code, RGB_code)
    # root.mainloop()