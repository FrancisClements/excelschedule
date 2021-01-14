from tkinter import *
from datetime import datetime as dt
from tkcolorpicker import askcolor  
import tkinter.ttk as ttk, tkinter.font as tkfont, json, os, re
import pandas as pd


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
        self.style = ttk.Style()
        self.widget_list = widget_list
        self.config_style()

    def config_style(self):
        #label formats: header, warning
        self.sub_font = 'Corbel'
        size = FONT_SIZE['header']
        self.style.configure('Header.TLabel', font = f'Corbel {size} bold')
        self.style.configure('Warning.TLabel', foreground = RED)

        #button formats: form
        #button modified: size
        padding = [str(x) for x in BUTTON_SIZE['default'].values()]
        padding = ' '.join(padding)
        self.style.configure('Form.TButton', padding = '0 0')
        self.style.configure('TButton', padding = f'{padding}')

        #entry modified: disabled
        self.style.layout('Disable.TEntry', [('Entry.plain.field', {'children': [(
                       'Entry.background', {'children': [(
                           'Entry.padding', {'children': [(
                               'Entry.textarea', {'sticky': 'nswe'})],
                      'sticky': 'nswe'})], 'sticky': 'nswe'})],
                      'border':'20', 'sticky': 'nswe'})] )

        self.style.map('Disable.TEntry',
            foreground = [('disabled', 'black')],
            fieldbackground = [('!disabled', 'red'), ('disabled', f'{LIGHT_GREY1}' )])

        #separator modified: background
        self.style.configure('TSeparator', background = LIGHT_GREY2)

    def add_to_list(self, widget):
        #adds widgets to the list
        if self.widget_list != None:
            self.widget_list.append(widget)
        return widget

    def label(self, item, theme = 'default', **kwargs):
        #initializes the label
        l = ttk.Label(self.root, text = item, **kwargs)

        if theme == 'header':
            l.configure(style = 'Header.TLabel')
        elif theme == 'warning':
            l.configure(style = 'Warning.TLabel')

        return self.add_to_list(l)
        
    def button(self, item, size = 'default', font_size = 'default', state_type = None, cmd = None):
        #initializes the button
        b = ttk.Button(self.root, text = item, state = state_type, command = cmd)#, font = used_font)

        #sets button size
        if size == 'form':
            b.config(style = 'Form.TButton')
        
        # if isinstance(size, str):
        #     b['padx'] = BUTTON_SIZE[size]['x']
        #     b['pady'] = BUTTON_SIZE[size]['y']
        # else:
        #     b['padx'], b['pady'] = size
        return self.add_to_list(b)

    def entry(self, width = 25, read_only = 0, limit = None, **kwargs):
        str_format = re.compile(r'[\\/:\:\?><\|\*]')
        #functions for validating data
        def validate_len(new_val):
            #limits amount of characters
            if len(new_val) > limit or re.search(str_format, new_val):
                self.notify()
                return False
            return True

        def validate(new_val):
            if re.search(str_format, new_val):
                self.notify()
                return False
            return True
                #initializes entry
        e = ttk.Entry(self.root, **kwargs)
        e['width'] = width

        #command input
        if limit == None:
            vcmd = (e.register(validate),"%P")
        else:
            vcmd = (e.register(validate_len),"%P")
        
        e.configure(validate = 'key', validatecommand = vcmd)

        #sets the state of the entry.
        if isinstance(read_only, int) or isinstance(read_only, bool):
            e['state'] = DISABLED if read_only else NORMAL
            e['style'] = 'Disable.TEntry' if read_only else None
        else:
            e['state'] = read_only

        return self.add_to_list(e)

    def labelframe(self, title, padding = [0,0], **kwargs):
        #initializes the label frame
        pad = ' '.join(map(str, padding))
        lf = ttk.LabelFrame(self.root, text = title, padding = pad, **kwargs)
        return self.add_to_list(lf)

    def checkbox(self, item, var = None, cmd = None):
        cb = ttk.Checkbutton(self.root, text = item,
            command = cmd, variable = var)
        return self.add_to_list(cb)

    def dropdown(self, item, val = 0, cmd = None, var = None, **kwargs):
        # automatically updates the variable
        def update_var(event):
            if var != None:
                var.set(drop.get())
            if cmd != None:
                cmd()

        drop = ttk.Combobox(self.root, value = item, **kwargs)
        drop.current(val)
        drop.bind('<<ComboboxSelected>>', update_var)

        #attempt to store a variable class (this only works once)
        if var != None:
            var.set(drop.get())

        return self.add_to_list(drop)
    
    def sep(self, **kwargs):
        #separator bar
        separate = ttk.Separator(self.root, **kwargs)
        return self.add_to_list(separate)

    def color_picker(self, var = None, color = 'white'):
        #sets default color if var is not filled
        fill = LIGHT_GREY1
        #asks a color picker window when it's clicked
        def clicked(event):
            fill = var.get()
            RGB_code, hex_code = askcolor(fill, self.root, 
                    title = 'Color Picker (Subject Color)')
            fill = hex_code

            if fill != None:
                var.set(fill) if var != None else ''
                picker.config(background = fill)

        #colorpicker uses Label, not ttk.Label. Thay are different
        picker = Label(self.root, text = 'A', fg = color, bg = fill, 
                width = 5, relief = 'solid', bd = 1, pady = 5)
        picker.bind('<Button-1>', clicked)
        return self.add_to_list(picker)

    #non-widget methods

    def widget_pack(self, widget, padding = [0,0], fill_wid = None, expand_wid = None, snap = None, **kwargs):
        #fill_wid = 'x', 'y', or 'both'
        #expand_wid = 0 or 1
        #renders array of pack()
        widget.pack(padx = padding[0], pady = padding[1])
        widget.pack_configure(side = snap)
        widget.pack_configure(expand = expand_wid, fill = fill_wid)
        widget.pack_configure(**kwargs)

    def widget_grid(self, widget, pos = [0,0], span = [1,1], padding = [0,0], snap = None):
        #renders array of grid()
        widget.grid(row = pos[1], column = pos[0])
        widget.grid_configure(padx = padding[0], pady = padding[1])
        widget.grid_configure(columnspan = span[0], rowspan = span[1])
        widget.grid_configure(sticky = snap)

    def grid_config(self, pos = [0,0]):
        if isinstance(pos[1], list) or isinstance(pos[1], tuple):
            for row in range(pos[1][0], pos[1][1]+1):
                Grid.rowconfigure(self.root, row, weight = 1)
        elif isinstance(pos[1], int):
            Grid.rowconfigure(self.root, pos[1], weight = 1)
            
        if isinstance(pos[0], list) or isinstance(pos[0], tuple):
            for col in range(pos[0][0], pos[0][1]+1):
                Grid.columnconfigure(self.root, col, weight = 1)
        elif isinstance(pos[0], int):
            Grid.columnconfigure(self.root, pos[0], weight = 1)

    def notify(self):
        self.root.bell()

'class that creates a tooltip when a widget is hovered'

class ToolTip(object):
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return

        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + self.widget.winfo_width() + 10
        y = y + cy + self.widget.winfo_rooty() + 3

        #height pattern (15n+10): 25, 40, 55, 70
        #3, 12, 18
        #correct pattern: 3, 11, 18.5, 26.5

        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

#function that accesses the class
def tooltip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()

    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

#gathering global variables for tkinter
def load_settings():
    global data
    filename = 'settings.json'
    #checks if the file exists (required)
    print(__file__)
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

def write_data(input_data = None):
    global data
    filename = 'settings.json'

    if input_data != None:
        data = input_data


    with open(filename, 'w') as f:
        #timestamp
        now = dt.today()
        str_format = '%m/%d/%Y at %I:%M%p'
        data['_metadata'] = now.strftime(str_format)

        json.dump(data, f, indent=4, sort_keys=True)
        f.truncate()
    load_settings()

load_settings()

#reads the excel file using pandas
def read_file(filename):
    df = pd.read_excel(filename)
    return df


if __name__ == "__main__":
    root = Tk()
    Kinter(root)