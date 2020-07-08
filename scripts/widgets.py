from tkinter import *
from datetime import datetime as dt
import tkinter.ttk as ttk, tkinter.font as tkfont, json, os
from tkcolorpicker import askcolor  

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
        self.style = ttk.Style()
        self.widget_list = widget_list
        self.config_style()

    def config_style(self):
        #label formats: header, warning
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
                      'sticky': 'nswe'})], 'sticky': 'nswe', 'border': '20'})],
                      'border':'20', 'sticky': 'nswe'})] )

        # self.style.configure("Disable.TEntry",
        #          background='red', 
        #          fieldbackground = f'{LIGHT_GREY1}', relief = FLAT)

        self.style.map('Disable.TEntry',
            foreground = [('disabled', 'black')],
            fieldbackground = [('!disabled', 'red'), ('disabled', f'{LIGHT_GREY1}' )])


    def add_to_list(self, widget):
        #adds widgets to the list
        if self.widget_list != None:
            self.widget_list.append(widget)
        return widget

    def label(self, item, theme = 'default'):
        #initializes the label
        l = ttk.Label(self.root, text = item)

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

    def entry(self, width = 20, read_only = 0):
        #initializes entry
        e = ttk.Entry(self.root)
        e['width'] = width

        #sets the state of the entry.
        if isinstance(read_only, int) or isinstance(read_only, bool):
            e['state'] = DISABLED if read_only else NORMAL
            e['style'] = 'Disable.TEntry' if read_only else None
        else:
            e['state'] = read_only

        return self.add_to_list(e)

    def labelframe(self, title, font_size = 'default', padding = [0,0]):
        #sets size
        used_font = self.set_font_size(font_size)

        #initializes the label frame
        lf = ttk.LabelFrame(self.root, text = title)#, 
            #padx = padding[0], pady = padding[1], font = used_font)
        return self.add_to_list(lf)

    def checkbox(self, item, var = None, cmd = None):
        cb = ttk.Checkbutton(self.root, text = item,
            command = cmd, variable = var)
        return self.add_to_list(cb)

    def dropdown(self, item, val = 0, cmd = None):
        drop = ttk.Combobox(self.root, value = item)
        drop.current(val)
        return self.add_to_list(drop)

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

'''class that creates a tooltip when a widget is hovered'''

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

if __name__ == "__main__":
    root = Tk()
    Kinter(root)