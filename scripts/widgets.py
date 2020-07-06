from tkinter import *
from datetime import datetime as dt
import tkinter.ttk as ttk, tkinter.font as tkfont, json, os
from tkcolorpicker import askcolor

#gathering global variables for tkinter
def load_settings():
    global data
    filename = 'settings.json'
    #checks if the file exists (required)
    if not os.path.isfile(filename):
        raise Exception('JSON file is missing.')

    #opens and reads the file (which acts as the template)
    with open(filename, 'r') as f:
        data = json.load(f)

    #makes widget dict settings into variables
    var_list = ['widget_settings', 'system_colors']

    for var in var_list:
        for key, val in zip(data[var], data[var].values()):
            globals()[key] = val

def write_data():
    filename = 'settings.json'
    with open(filename, 'w') as f:
        #timestamp
        now = dt.today()
        str_format = '%m/%d/%Y at %I:%M%p'
        data['_metadata'] = now.strftime(str_format)

        json.dump(data, f, indent=4, sort_keys=True)
        f.truncate()

load_settings()         

#initializing the methods for tkinter
'''
    Kinter class is basically a cheat-sheet for
    TKinter. This class helps reduce the line of code
    to the main.py script.

    Parameters:
    root - a frame where you can insert widgets such as labels, buttons
    widget_list - collects created widgets for easy rendering.
'''
class Kinter:
    def __init__(self, root, widget_list = None):
        self.root = root
        self.font = tkfont.Font()
        self.widget_list = widget_list

    def add_to_list(self, widget):
        #adds widgets to the list
        if self.widget_list != None:
            self.widget_list.append(widget)
        return widget

    def set_font(self, font_pack):
        #configures font
        self.font = tkfont.Font(family = font_pack[0], size = font_pack[1])    
        FONT_SIZE['default'] = font_pack[1]

    def label(self, item, font_size = 'default'):
        #sets size
        used_font = self.set_font_size(font_size)
                
        #initializes the label
        l = Label(self.root, text = item, font = used_font)
        return self.add_to_list(l)
        
    def button(self, item, size = 'default', font_size = 'default', state_type = None, cmd = None):
        #sets size
        used_font = self.set_font_size(font_size)

        #initializes the button
        b = Button(self.root, text = item, state = state_type, command = cmd, font = used_font)

        #sets button size
        if isinstance(size, str):
            b['padx'] = BUTTON_SIZE[size]['x']
            b['pady'] = BUTTON_SIZE[size]['y']
        else:
            b['padx'], b['pady'] = size
        return self.add_to_list(b)

    def entry(self, width, bg = None, read_only = 0, font_size = 'default'):
        #sets size
        used_font = self.set_font_size(font_size)

        #initializes entry
        e = Entry(self.root)
        e['width'] = width
        e['bg'] = bg

        #sets the state of the entry.
        if isinstance(read_only, int):
            e['state'] = DISABLED if read_only else NORMAL
        else:
            e['state'] = read_only

        return self.add_to_list(e)

    def labelframe(self, title, font_size = 'default', padding = [0,0]):
        #sets size
        used_font = self.set_font_size(font_size)

        #initializes the label frame
        lf = LabelFrame(self.root, text = title, 
            padx = padding[0], pady = padding[1], font = used_font)
        return self.add_to_list(lf)

    def checkbox(self, item, var = None, cmd = None, font_size = 'default'):
        #sets sizes
        used_font = self.set_font_size(font_size)

        cb = Checkbutton(self.root, text = item, 
            variable = var, font = used_font, command = cmd)
        return self.add_to_list(cb)

    #non-widget methods
    def widget_pack(self, widget, padding = [0,0], fill_wid = None, expand_wid = None, snap = None):
        #renders array of pack()
        widget.pack(padx = padding[0], pady = padding[1])
        widget.pack_configure(side = snap)
        widget.pack_configure(expand = expand_wid, fill = fill_wid)

    def widget_grid(self, widget, pos = [0,0], span = [1,1], padding = [0,0], snap = None):
        #renders array of grid()
        widget.grid(row = pos[1], column = pos[0])
        widget.grid_configure(padx = padding[0], pady = padding[1])
        widget.grid_configure(columnspan = span[0], rowspan = span[1])
        widget.grid_configure(sticky = snap)

    def set_font_size(self, font_size):
        #sets custom font size for widgets
        new_font = self.font.copy()

        if isinstance(font_size, str):
            new_font.config(size = FONT_SIZE[font_size])
        else:
            new_font.config(size = font_size)

        return new_font
