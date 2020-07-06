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

        #clears the data files from JSON
        data['files']['input_file'] = ''

        # self.configure(bg = 'blue')

    def render(self):
        self.pack(fill = 'both', expand = 1, pady = 50)

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
        title_k.label('Excel Schedule Maker', font_size = 'header')
        title_k.label('Create a graphical version of your schedule using Excel')
        title_k.label('Insert the schedule file and name the output file.')

        for widget in widget_list:
            title_k.widget_pack(widget, fill_wid = 'both')
 
    def file_form(self):
        f = Frame(height = 100, width = 100, master = self)
        f.pack(expand = 1)
        widget_list = []
        title_k = Kinter(f, widget_list)

        #Input File
        input_wid_list = []
        input_frame = Kinter(title_k.labelframe('Input File', padding = [0, 5]), input_wid_list)
        input_txt = StringVar()

        input_form = input_frame.entry(width = 40, read_only = 'readonly')
        input_form.config(textvariable = input_txt, readonlybackground = LIGHT_GREY1)

        i_btn = input_frame.button('Browse', size = 'form', 
                    cmd = lambda x = input_txt: self.browse(x))
        i_btn.config(bg = LIGHT_GREY2)

        #Output File
        out_wid_list = []
        out_frame = Kinter(title_k.labelframe('Output Filename', padding = [0, 5]), out_wid_list)
        self.out_txt = StringVar()

        out_form = out_frame.entry(width = 40)
        out_form.config(textvariable = self.out_txt, bg = LIGHT_GREY1)
        out_frame.label('.xlsx')
        out_frame.label('The output file is at the same location as the input file.')

        #rendering
        for widget in widget_list: #the whole main menu frames
            title_k.widget_pack(widget, fill_wid = 'both')

        for index, widget in enumerate(input_wid_list): #input form
            input_frame.widget_grid(widget, pos = [index, 0], padding = [7, 7])

        #output form
        out_frame.widget_grid(out_wid_list[0], pos = [0,0])
        out_frame.widget_grid(out_wid_list[1], pos = [1,0], snap=W)
        out_frame.widget_grid(out_wid_list[2], pos = [0,1], span=[2,1], padding = [4,0])

    def next_warning(self):
        #Next Button + Warning
        f = Frame(height = 100, width = 100, master = self, pady = 5)
        f.pack(expand = 1)

        last_wid_list = []
        last_frame = Kinter(f, last_wid_list)
        warning_txt = StringVar(value = '')
        warning_lbl = last_frame.label('')
        warning_lbl.configure(textvariable = warning_txt, width = 33,
                    anchor = W, fg = RED)
        
        next_btn = last_frame.button('Next',
                cmd = lambda x = warning_txt: self.check_form(x))
        next_btn.config(bg = LIGHT_GREY2)

        #render
        last_frame.widget_grid(last_wid_list[0], pos = [0,0], span = [2,1], padding = [7,5], snap = W)
        last_frame.widget_grid(last_wid_list[1], pos = [2,0], padding = [10,5])

    def check_form(self, txt_var):
        #characters that are not allowed: /\:*?"><|
        str_format = re.compile(r'[\\/:\:\?><\|\*]')
        input_file = data['files']['input_file']
        output_file = self.out_txt.get()
        form_check = [len(x) == 0 for x in [input_file, output_file]]

        #checks if no data inputted
        if all(form_check):
            txt_var.set('Enter the Input File and Output Filename')
            return
        elif any(form_check):
            print('one of them')
            warning_txt = 'Enter the '
            warning_txt += 'Input File' if form_check[0] else 'Output Filename'
            txt_var.set(warning_txt)
            return

        #checks the naming of the filename
        if re.search(str_format, output_file):
            txt_var.set('Filename is not allowed')
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
        self.main_k = Kinter(self)
        super().__init__(master.root, bg = LIGHT_GREY1)

    def render(self):
        self.pack(fill = 'both', expand = 1, pady = 50, padx = 10)
        f = Frame(height = 100, width = 100, master = self)
        f.pack(expand = 1)
        options_list = []
        options_frame = Kinter(f, options_list)

        #setting up option states
        state = {}
        for option in ['hour_list', 'header', 'name']:
            state[option] = BooleanVar()

        options_frame.label('Customize', font_size = 'header')
        options_frame.checkbox('Enable Hour List', var = state['hour_list'])
        options_frame.label('Day Format')
        options_frame.label('Time Format')
        options_frame.checkbox('Enable Header', var = state['header'])
        options_frame.checkbox('Add your name to the schedule', var = state['name'])
        # options_frame.button('Enter')

        for i, widget in enumerate(options_list):
            if i == 0:
                cell = [0,0]
            elif isinstance(widget, Button):
                cell = [1, (i-1)//2+2]
                options_frame.widget_grid(widget, pos = cell)
                continue

            else:
                cell = [(i-1)%2, (i-1)//2+1]

            options_frame.widget_grid(widget, pos = cell, snap = W, padding = [10,0])


    def get_state(self, var):
        print(var.get())

#program window
class Program:
    def __init__(self):
        self.root = Tk()
        self.root.title('Excel Schedule Maker')
        self.root.geometry('600x400')
        self.new()
        self.run('test')
        print('App Exit')

    def new(self):
        #puts all frames in a list
        self.frames = [0, [MainMenu(self), Options(self)]]
        print(self.frames)

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
    p = Program()

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