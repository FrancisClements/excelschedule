import schedule, re
from tkinter import *
from tkinter import filedialog
from widgets import *
from pathlib import Path
# import tkinter.ttk as ttk, tkinter.font as tkfont
# from tkcolorpicker import askcolor

#main menu
class MainMenu(Frame):
    def __init__(self, master = None):
        self.data = data
        self.master_cls = master
        self.main_k = Kinter(self)
        self.allow_next = 0
        super().__init__(self.master_cls.root)

        # self.configure(bg = 'blue')

    def render(self):
        #clears the data files from JSON
        data['files']['input_file'] = ''

        self.pack(fill = 'both', expand = 1, pady = 50, padx = 20)

        self.title()
        self.file_form()
        self.next_warning()

    def title(self):
        #frame of the widgets
        f = Frame(height = 100, width = 100, master = self)
        f.pack(expand = 1)
        widget_list = []
        title_k = Kinter(f, widget_list)

        #things to write
        title_k.label('Excel Schedule Maker', theme = 'header')
        title_k.label('Create a graphical version of your schedule using Excel')
        title_k.label('Insert your Excel schedule table and name the output file')

        for widget in widget_list:
            title_k.widget_pack(widget)

    def file_form(self):
        f = Frame(height = 100, width = 100, master = self)
        f.pack(expand = 1)
        widget_list = []
        title_k = Kinter(f, widget_list)

        #Input File
        input_wid_list = []
        input_frame = Kinter(title_k.labelframe('Input File', padding = [0, 5]), input_wid_list)
        input_txt = StringVar()

        input_form = input_frame.entry(width = 40, read_only = 1)
        input_form.config(textvariable = input_txt)#, readonlybackground = LIGHT_GREY1)

        i_btn = input_frame.button('Browse', size = 'form', 
                    cmd = lambda x = input_txt: self.browse(x))
        # i_btn.config(bg = LIGHT_GREY2)

        #Output File
        out_wid_list = []
        out_frame = Kinter(title_k.labelframe('Output Filename', padding = [0, 5]), out_wid_list)
        self.out_txt = StringVar()

        out_form = out_frame.entry(width = 40)
        out_form.config(textvariable = self.out_txt)#, bg = LIGHT_GREY1)
        out_frame.label('.xlsx')
        out_frame.label('The output file is at the same location as the input file.')

        #rendering
        for widget in widget_list: #the whole main menu frames
            title_k.widget_pack(widget, fill_wid = 'both')

        for index, widget in enumerate(input_wid_list): #input form
            input_frame.widget_grid(widget, pos = [index, 0], padding = [7, 7])

        #output form
        out_frame.widget_grid(out_wid_list[0], pos = [0,0])
        out_frame.widget_grid(out_wid_list[1], pos = [1,0])
        out_frame.widget_grid(out_wid_list[2], pos = [0,1], span=[2,1], padding = [4,0])

    def next_warning(self):
        #Next Button + Warning
        f = Frame(height = 100, width = 100, master = self, pady = 5)
        f.pack(expand = 1)

        last_wid_list = []
        last_frame = Kinter(f, last_wid_list)
        warning_txt = StringVar(value = '')
        warning_lbl = last_frame.label('', theme = 'warning')
        warning_lbl.configure(textvariable = warning_txt, width = 40,
                    anchor = W)
        
        next_btn = last_frame.button('Next',
                cmd = lambda x = warning_txt: self.check_form(x))
        # next_btn.config(bg = LIGHT_GREY2)

        #render
        last_frame.widget_grid(last_wid_list[0], pos = [0,0], span = [2,1],
                    padding = [7,5], snap = W)
        last_frame.widget_grid(last_wid_list[1], pos = [2,0], padding = [10,5])

    def check_form(self, txt_var):
        #checks if the forms are correctly inputted
        input_file = data['files']['input_file']
        output_file = self.out_txt.get()
        form_check = [len(x) == 0 for x in [input_file, output_file]]

        #checks if no data inputted
        if all(form_check):
            txt_var.set('Enter the Input File and Output Filename')
            self.bell()
            return
        elif any(form_check):
            print('one of them')
            warning_txt = 'Enter the '
            warning_txt += 'Input File' if form_check[0] else 'Output Filename'
            txt_var.set(warning_txt)
            self.bell()
            return

        #All conditions are met
        self.allow_next = 1
        input_list = input_file.split('/')
        input_list[-1] = output_file + '.xlsx'

        output_file = '/'.join(input_list)

        print(input_list)
        data['files']['output_file'] = output_file
        self.next_frame()
        print(input_file, output_file)

    def next_frame(self):
        write_data()
        self.master_cls.next_frame()

    def browse(self, str_val):
        base_dir = str(Path.home())
        type_list = (('Excel Workbooks', '*.xlsx'), 
                    ('All Files', '*.*'))
        filename = filedialog.askopenfile(title = 'Open', filetypes = type_list)

        if filename != None:
            filename = filename.name
            filename_str = filename.split('/')

            #writing up stuff
            str_val.set(filename_str[-1])
            data['files']['input_file'] = filename
            print('Input:', filename)

#options
class Options(Frame):
    def __init__(self, master = None):
        self.data = data
        self.master = master
        self.frames = []
        self.main_k = Kinter(self, self.frames)
        super().__init__(master.root)

        #Boolean States
        self.state = {}
        for option in ['hour_list', 'header', 'name']:
            #value to true on hour list and header
            self.state[option] = BooleanVar(value = True) if option != 'name' else BooleanVar()

        #Data Inputs (all empty)
        self.input_data = {}
        for data_vars in ['header', 'name', 'time_format', 'day_format', 'subject_key']:
            self.input_data[data_vars] = StringVar()

    def render(self):
        self.pack(fill = 'both', expand = 1, pady = 50, padx = 50)
        '''
        LAYOUT OF OPTIONS >>
                        V
                TITLE L | TITLE  R
                DESC  L | DESC   R
                --------+--------
                L     R | COLOR
                 FRAMES | SELECTION
                --------+--------
                    BUTTONS
                    
        '''
        #title/desc LEFT-----------------------------
        tl = self.main_k.label('Customize', theme = 'header')
        dl = self.main_k.label('Set up options to make your schedule look better')
        self.main_k.widget_grid(tl, pos = [0,0], span = [2,1], snap = W)
        self.main_k.widget_grid(dl, pos = [0,1], span = [2,1], padding = [10,0], snap = W)

        #left & right frames-------------------------
        self.main_k.widget_grid(self.left_frame(), pos = [0,2], snap = NW)
        self.main_k.widget_grid(self.right_frame(), pos = [1,2], snap = NW)

        #verical line--------------------------------
        sep = self.main_k.sep(orient = VERTICAL)
        self.main_k.widget_grid(sep, pos = [3,0], span = [1,3], snap = NS)

        #title/desc RIGHT----------------------------
        tr = self.main_k.label('Color Selection', theme = 'header')
        dr = self.main_k.label('Pick a color of your choice that corresponds to your subject\n'
                'If we got the wrong category, feel free to change it below')
        self.main_k.widget_grid(tr, pos = [4,0], padding = [5,0], snap = W)
        self.main_k.widget_grid(dr, pos = [4,1], padding = [15,0], snap = W)
                    
        #color selection----------------------------
        self.main_k.widget_grid(self.color_frame(), pos = [4,2], snap = NSEW)
        
    def get_data(self):
        print('\n\n\n--- STATES')
        for state in self.state:
            print(state,':',  self.state[state].get())

        print('\n--- INPUT DATA')
        for inputs in self.input_data:
            print(inputs, ':',  self.input_data[inputs].get())
    
    def left_frame(self):
        str_format = re.compile(r'[\\/:\:\?><\|\*]')
        widgets = []
        f = Frame(height = 100, width = 100, master = self, pady = 5)
        frame = Kinter(f, widgets)

        #Setting default value of header
        self.input_data['header'].set('My Schedule')

        #widgets
        frame.checkbox('Enable Hour List', var = self.state['hour_list'])
        frame.checkbox('Enable Header', var = self.state['header'])
        header_form = frame.entry(limit = 30, textvariable = self.input_data['header'])

        frame.checkbox('Include Name', var = self.state['name'])
        name_form = frame.entry(read_only = 1, limit = 30, 
                textvariable = self.input_data['name'])

        frame.button('Check', cmd = self.get_data)

        #tooltip descriptions
        desc = [
            'Adds a list of hour time in your schedule\n' '(adds more organization)',
            'Add a header or title',
            'Add your name to your schedule'
        ]

        #listed self.state keys that uses an entry
        entry_states = ['header', 'name']
        desc_index = 0
        for i, widget in enumerate(widgets):
            #if the widget is a checkbox
            if isinstance(widget, ttk.Checkbutton):
                #command that changes the state of the input form
                if desc_index >= 1:
                    widget.config(command = lambda x = self.state[entry_states[desc_index-1]],
                        y = widgets[i+1]: self.set_entry(x,y))
                    pad = [10,0]

                else:
                    pad = [10,5]
                #adds tooltip
                tooltip(widget, desc[desc_index])
                desc_index += 1
                #render for checkboxes
                frame.widget_grid(widget, pos= [0,i], snap = W, padding = pad)
            else:
                #render for entry
                frame.widget_grid(widget, pos = [0,i], snap = W, padding = [30,(0,5)])

        return f

    def right_frame(self):
        widgets = []
        f = Frame(height = 100, width = 100, master = self, pady = 5)
        frame = Kinter(f, widgets)

        #list of options/vars
        day_formats = ['Initial', 'Partial', 'Full']
        time_formats = ['12hr + AM/PM', '12hr + a/p', '24hr']


        #widgets
        frame.label('Day Format') #e.g. Mon, Monday, M
        frame.dropdown(day_formats, 1, var = self.input_data['day_format'], 
                state = 'readonly')

        frame.label('Time Format')
        frame.dropdown(time_formats, 0, var = self.input_data['time_format'],
                state = 'readonly')

        #tooltip descriptions
        desc = [
            'Sets the day formatting\n' '(e.g. M/Mon/Monday)',
            'Sets the time formatting\n' '(e.g. 1:00PM/1:00p/13:00)',
        ]
        desc_index = 0

        #fixing spacing of the widgets + render
        for i, widget in enumerate(widgets):
            if isinstance(widget, ttk.Label):
                frame.widget_grid(widget, pos = [0,i], padding = [10,(5,0)], snap = W)
                #tooltip
                tooltip(widget, desc[desc_index])
                desc_index += 1
            else:
                frame.widget_grid(widget, pos = [0,i], padding = [10,0], snap = W)

        return f

    def color_frame(self):
        widgets = []
        f = Frame(height = 100, width = 100, master = self, pady = 5, padx = 5)
        frame = Kinter(f, widgets)

        #excel data using pandas
        self.pd = read_file(data['files']['input_file'])
        headers = self.pd.keys().tolist()

        #dropdown menu------------------------
        dd_t = frame.label('Subject')
        dd = frame.dropdown(headers, 1, var = self.input_data['subject_key'], 
                state = 'readonly')

        #render for dropdown
        # frame.widget_grid(dd_t, pos = [0,1], padding = [10, (5,0)], snap = W)
        # frame.widget_grid(dd, pos = [0,2], padding = [10, 0], snap = W)
        frame.widget_pack(dd_t, padding = [10, (5,0)], expand_wid = 1, anchor = 'w')
        frame.widget_pack(dd, padding = [10,0], expand_wid = 1, anchor = 'w')
            
        #color selection----------------------
        color_widgets = []
        c = frame.labelframe('Colors', padding = [0,5])
        colors = Kinter(c, color_widgets)
        #creates all subjects
        for subject in self.pd[self.input_data['subject_key'].get()]:
            colors.label(subject)

        #renders all subjects inside frame
        colors.grid_config(pos = [[0,1],0])
        for i, widget in enumerate(color_widgets):
            colors.widget_grid(widget, pos = [i%2, i//2], snap = NSEW)

        #renders color selection frame
        # frame.widget_grid(c, pos = [0,3], padding = [10, (5,0)], snap = NSEW)
        frame.widget_pack(c, padding = [10, (5,0)], expand_wid = 1, fill_wid = 'x', anchor = 'w')

        return f

    

    def set_entry(self, var, widget):
        #if the button is unchecked, the input form will disable.
        #var = true/false, widget = entry widget
        if not var.get():
            widget.delete(0,END)
            widget.configure(style = 'Disable.TEntry', state = DISABLED)
        else:
            widget.configure(style = 'TEntry', state = NORMAL)

    def get_status(self, states):
        for key, val in zip(states, states.values()):
            print(key, val.get())

#program window
class Program:
    def __init__(self, mode = None):
        self.root = Tk()
        self.root.title('Excel Schedule Maker')
        self.root.geometry('900x550')
        self.new()
        self.run(mode)
        print('App Exit')

    def new(self):
        #puts all frames in a list
        self.frames = [0, [MainMenu(self), Options(self)]]

    def next_frame(self):
        #hides the current frame, shows the next frame
        self.frames[1][self.frames[0]].pack_forget()
        self.frames[0] += 1
        self.frames[1][self.frames[0]].render()

    def run(self, mode = None):
        #prints the mainmenu
        if mode == 'test':
            self.frames[1][1].render()
            pass
        else:
            self.frames[1][0].render()

        # print(self.root.pack_slaves())

        #main loop of the program
        self.root.mainloop() 

    def create_schedule(self):
        schedule.create_schedule()

if __name__ == "__main__":
    p = Program('test')

    # r = Tk()
    # r.geometry("800x400")
    # i = IntVar()
    # i.set('2')

    # l_frame = LabelFrame(r, padx = 20, pady = 30, text = 'Frame')

    # for x in range(1,3):
    #     t = Radiobutton(l_frame, text = 'Choice' + str(x), variable = i, value = x)
    #     t.pack()

    # l_frame.pack()
    # r.mainloop()

    # root = Tk()
    # style = ttk.Style(root)
    # style.theme_use('clam')
    # hex_code, RGB_code = askcolor((255, 255, 0), root) 
    # print(hex_code, RGB_code)
    # root.mainloop()