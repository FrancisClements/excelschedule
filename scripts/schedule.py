#Backend
import pandas as pd, xlsxwriter as xl
import os, re, json
from datetime import datetime
from tkinter import messagebox
from widgets import *

#debugging
import inspect, itertools

class SchedMaker:
    def __init__(self, json_data):
        #read the JSON
        self.data = json_data
        self.init_sched()

    def init_sched(self):
        #json init
        data_section = self.data['data']
        chosen_col = data_section['subject_key']
        chosen_time = [data_section['time_key_0'], data_section['time_key_1']]
        day_mode = self.data['data']['day_format']                                                    #PARAMETERS: 'FULL', 'PARTIAL', 'INITIAL' (/)

        #pandas init
        self.df = pd.read_excel(self.data['files']['input_file'], index_col = chosen_col)             #PARAMETERS: 'CODE' (header on excel) (/)
        
        #checks if it only has one time input
        if not self.data['options']['enable_time_twice']:
            time_list = []
            #loops through the cells and extracts the time range
            for time in self.df[chosen_time[0]]:
                time_list.extend(self.strip_time(time))

        else:
            time_list = self.df[chosen_time[0]].to_list() + self.df[chosen_time[1]].to_list()             #PARAMETERS: FROM TIME, TO TIME (headers on excel) (/)

        self.time_list = self.time_sort(time_list) #sorts the time
        self.day_list = self.get_day_list(day_mode.upper())

    def strip_time(self, time):
        #strips concatenated time
        try:
            time = time.replace(' ', '')
            time = time.split('-')
            return time
        except:
            return error(f'The time column has incorrect formatting\n'
                    'Please check the file and try again.')

    def get_day_list(self, mode):
        day_enabled = self.data['options']['enable_day']

        #gets only the day_list if it's enabled
        if day_enabled:
            day_key = self.data['data']['day_key']
            self.week_list = {
                'FULL': ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
                'INITIAL': ['M', 'T', 'W', 'TH', 'F', 'S']
            }
            self.week_list['PARTIAL'] = [x[:3].upper() for x in self.week_list['FULL']] #Tuesday = Tue
            word = ''.join(self.df[day_key])                                                               #PARAMETER: DAY (header) (/)
            chosen_list = self.week_list[mode]

            #removes any excess days
            for i in range(len(self.week_list['FULL'])):
                re_filter = self.regex_day(i, word)
                if not re_filter:
                    del chosen_list[i]

            return chosen_list
        return ['Mon - Fri']

    def regex_day(self, index, word):
        #returns list of subjects that is on that day.
        day = self.week_list['FULL'][index]
        if day == 'Thursday':
            day_filter = re.compile(r'(?i:(t(?=[^u])))')
        elif day == 'Tuesday':
            day_filter = re.compile(r'(?i:(t\b)|(t(?=[^h])))')
        else:
            day_filter = re.compile(rf'(?i:({day[0]}({day[1:3]}({day[3:]})?)?))') #e.g.(M(on(day)?)?)

        return day_filter.search(word)

    def time_sort(self, t):
        #removes duplicates
        t = list(set(t))

        t = self.str_to_time(t)

        #converts into datetime object and sort
        t.sort()

        #converts into string
        t = self.time_to_str(t)
        
        return t

    def str_to_time(self, t):
        '''
            These are all of the available time formats to recognize the time
            format of the input file. These are the list of time formats:

            1.  1:00PM   ->  %I:%M%p
            2.  1:00p    ->  %I:%M%p
                -just add 'm' to recognize the format
            3.  13:00    ->  %H:%M

            RegEx pseudocode:
            1. 1:00p and 1:00PM
                hour:minute + letter (one or two) 
            2. 13:00
                hour (1-24) :minute (no AM/PM)
        '''
        time_formats = [
            [re.compile(r'(?i:[pam]$)'), '%I:%M%p'], #Hi Pamela Figueroa
            [re.compile(r'(?i:\d+:\d+[^p^a^m]$)'), '%H:%M']
        ]
        format_matched = 0

        #finds the correct time format that's written on the schedule file
        for i, t_format in enumerate(time_formats):
            if t_format[0].search(t[0]):
                if (i == 1 and t[0][:2] <= '24') or i == 0:
                    self.str_format = t_format[1]                                                           #PARAMTER: Time Format; FORMAT EXAMPLE: 4:45pm (/)
                    format_matched = 1
                    break

        #if nothing matched on all of the patterns
        if not format_matched:
            return error(f'The time column has incorrect formatting\n'
                            'Please check the file and try again.')


        for i, time_in in enumerate(t):
            #adding 'm' for time formats (1:00p)
            if t[i][-1] not in ['m', 'M'] and self.str_format == time_formats[0][1]:
                time_in += 'm'                                                                        #PARAMETER: added 'm' (8:00a -> 8:am) HARDCODED (/)
            t[i] = datetime.strptime(time_in, self.str_format)

        return t

    def time_to_str(self, t):
        time_formats = [
            [re.compile(r'(?i:[pam]$)'), '%I:%M%p'], #Hi Pamela Figueroa
            [re.compile(r'(?i:[0-9]?[a-zA-Z]$)'), '%H:%M']
        ]
        '''
            Data came from JSON
            Output_format possible values:
            1. '12hr + AM/PM'
            2. '12hr + a/p'
            3. '24hr'
        '''
        data_desc = self.data['data']['time_format']
        self.output_str_format = time_formats[0][1] if data_desc[:2] == '12' else time_formats[1][1]

        for i, time_in in enumerate(t):
            t[i] = time_in.strftime(self.output_str_format)                                             #Set to lower case PM -> pm (/)
            if data_desc[-1] == 'p': #selected time format is 12hr + a/p
                t[i] = t[i][:-1].lower()

        return t


class ExcelWriter:
    def __init__(self, schedule, json_data):
        self.data = json_data
        self.schedule = schedule
        if os.path.isfile(self.data['files']['output_file']):                                         #PARAMETER: filename (HARDCODED) (/)
            os.remove(self.data['files']['output_file'])

        self.book = xl.Workbook(self.data['files']['output_file'])
        self.sheet = self.book.add_worksheet()
        self.offset = [0,0]
        self.row = self.offset[0]
        self.col = self.offset[1]
        self.new()
        self.write()
        self.book.close()

    def write(self):
        self.write_title()
        self.write_name()
        self.write_day()
        self.write_time()
        self.write_subject()

    def new(self):
        self.format_list = {
            'BOLD': {'bold':True},
            'CENTER': {'align': 'center', 'valign': 'vcenter'},
            'BORDER': {'border': 1},
            'WRAP': {'text_wrap':True}
        }
        self.color = self.data['system_colors']
        self.subj_colors = self.data['data']['colors']
        self.subj_font_color = self.data['data']['font_color'].upper()
        self.preset = {
            'header': [['COLOR', 'ACCENT_LIGHT'], 'BOLD', 'CENTER', 'BORDER'],
            'name': [['COLOR', 'DARK_GREY'], ['FONTCOLOR', 'WHITE'], 'CENTER', 'BORDER'],
            'title': [['COLOR', 'BLACK'], ['FONTCOLOR', 'WHITE'], 'BOLD', 'CENTER', 'BORDER'],
            'subject': ['BOLD', ['FONTCOLOR', self.subj_font_color]],
            'room': [['SIZE', 9], ['FONTCOLOR', self.subj_font_color]]
            }
        self.cell_height = 16
        self.cell_width = 13

        #sets a starting point when writing
        #assuming that there are no offset values, these are the values
        self.state = self.data['options']
        self.input_data = self.data['data']

        #column sizes (x)
        time_col = int(self.state['enable_hour_list']) + 1

        #row sizes (y)
        title_row = int(self.state['enable_header'] and self.input_data['header'] != '')
        name_row = int(self.state['enable_name'] and self.input_data['name'] != '')
        day_row = int(self.state['enable_day'])

        self.start_cell = {
            #name: y = title
            "name": [0, title_row],
            #days: x = time, y = title + name
            "days": [time_col, title_row + name_row],
            #time y = day+title+name
            "time": [0, day_row + title_row + name_row],
            #subject: x = time+1, y = day+title+name
            "subject": [time_col, day_row + title_row + name_row]
        }
    
    def set_col(self, a, b, width):
        self.sheet.set_column(a, b ,width)

    def set_row(self, pos, width):
        self.sheet.set_row(pos, width)

    def border_format(self, flag_list):
        if len(flag_list) == 0:
            return {'border': 1}
        else:
            format_dict = {'border': 1} #cell border
            for flag in flag_list:
                if flag in ['top', 'bottom', 'left', 'right']:
                    format_dict[flag] = 0
                elif re.search(r'^#', flag):
                    format['color'] = flag
        
        return format_dict

    def cell_format(self, order_list):
        #format_arr = list
        #changes format into str
        format_list = order_list.copy()
        for index, x in enumerate(format_list):
            if isinstance(x, list):
                if x[0] == 'COLOR':
                    format_list[index] = {'fg_color': x[1]} if re.search(r'^#', x[1]) else {'fg_color': self.color[x[1]]}
                elif x[0] == 'BORDER':
                    format_list[index] = self.border_format(x[1:]) #activates self.border_format if it's inputted as list
                elif x[0] == 'FONTCOLOR':
                    format_list[index] = {'font_color': x[1]} if re.search(r'^#', x[1]) else {'font_color': self.color[x[1]]}
                elif x[0] == 'SIZE':
                    format_list[index] = {'font_size': x[1]}
            else:
                format_list[index] = self.format_list[x] #searches through self.format list dict

        x = {}
        for format_dict in format_list: #format_list = [{}, {}]
            for key in format_dict: #format_dict = collection of keys in a dict
                x[key] = format_dict[key] #unpacks values of format_arr

        return self.book.add_format(x)
    
    def write_title(self):
        self.col = self.offset[0]
        self.row = self.offset[1]
        title = self.input_data['header']
        col_end = self.col + len(self.schedule.day_list) + self.start_cell['days'][0] - 1

        #print if the title has content and it's enabled
        if title != '' and self.state['enable_header']:
            self.sheet.merge_range(self.row, self.col, self.row, 
                    col_end, title, self.cell_format(self.preset['title']))

            self.set_row(self.row, self.cell_height) # sets row height

    def write_name(self):
        self.col = self.offset[0]
        self.row = self.start_cell['name'][1] + self.offset[1]
        name = self.input_data['name']
        enable_name = self.state['enable_name']
        col_end = self.col + len(self.schedule.day_list) + self.start_cell['days'][0] - 1

        #print the name if it's enabled and has content
        if name != '' and enable_name:
            name = 'Schedule by: ' + self.data['data']['name']

            self.sheet.merge_range(self.row, self.col, self.row, 
                col_end, name, self.cell_format(self.preset['name']))

    def write_day(self):
        #return immediately if day is disabled
        # if not self.state['enable_day']:
        #     return

        hour_enabled = self.state['enable_hour_list']
        self.col = self.start_cell['days'][0] + self.offset[0]
        self.row = self.start_cell['days'][1] + self.offset[1]

        for i, day in enumerate(self.schedule.day_list):
            self.sheet.write(self.row, self.col + i, day, self.cell_format(self.preset['header']))
        else:
            self.set_col(self.col-1, self.col+i, self.cell_width)

    def write_time(self):
        self.col = self.start_cell['time'][0] + self.offset[0]
        self.row = self.start_cell['time'][1] + self.offset[1]

        #counts the number of cells to merge
        merge_n = 1
        start_cell = [0,0]

        #cell format
        hour_enabled = self.state['enable_hour_list']
        hour_format = self.cell_format([['COLOR', 'ACCENT_LIGHT'], 'BOLD', 'CENTER', ['BORDER', 'right']])
        even_color = self.cell_format([['COLOR', 'ACCENT_LIGHT'], 'CENTER', 'BORDER'])
        odd_color = self.cell_format([['COLOR', 'ACCENT_DARK'], 'CENTER', 'BORDER'])
        cell_switch = True #makes a coloring switch regardless of its position

        self.col += int(hour_enabled)

        for i, time in enumerate(self.schedule.time_list):
            #writes the time
            self.sheet.write(self.row, self.col, time, hour_format)

            #do all the code below if the hour_list is true
            if hour_enabled:
                cell_color = even_color if cell_switch else odd_color

                #write the hours
                #detects if last hour matches current hr
                if i != 0 and time[:2] == last_hr:
                    if merge_n == 1: #sets starting cell
                        start_cell = [self.row-1, self.col-1]
                    merge_n +=1
                #writes hour time & switches the cell_color
                elif i != 0 and merge_n == 1:
                    self.sheet.write_number(self.row-1, self.col-1, int(last_hr), cell_color)
                    cell_switch = not cell_switch
                elif merge_n != 1:
                    self.sheet.merge_range(start_cell[0], start_cell[1], self.row-1, self.col-1, int(last_hr), cell_color)
                    cell_switch = not cell_switch
                    merge_n = 1
                last_hr = time[:2]       
                
            self.set_row(self.row, self.cell_height) # sets row height
            self.row += 1
        else:
            #write the time
            cell_color = even_color if cell_switch else odd_color
            if hour_enabled:
                if merge_n != 1:
                    self.sheet.merge_range(start_cell[0], start_cell[1], self.row-1, self.col-1, int(last_hr), cell_color)
                else:
                    self.sheet.write_number(self.row-1, self.col-1, int(last_hr), cell_color)

    def write_subject(self):
        self.col = self.start_cell['subject'][0] + self.offset[0]
        self.row = self.start_cell['subject'][1] + self.offset[1]
        add_room = self.state['enable_add_classroom']

        #yields subject, col, row_start, row_end
        subject_list = self.get_subject(self.schedule.df.index)
        for subject in subject_list:
            for cell_index in subject:
                col = self.col + cell_index[1]
                row_start = self.row + cell_index[2]
                row_end = self.row + cell_index[3]

                #text
                text = [self.cell_format(self.preset['subject']), cell_index[0]]
                room_txt = [" "*15, self.cell_format(self.preset['room']), cell_index[4]]

                if row_start != row_end:
                    self.sheet.merge_range(row_start, col, row_end, col, '', self.cell_format(['CENTER', 'WRAP',
                            ['COLOR', self.subj_colors[str(cell_index[0])]]]))

                if add_room:
                    text.extend(room_txt)
                    self.sheet.write_rich_string(row_start, col, *text, self.cell_format(['CENTER', 'WRAP',
                            ['COLOR', self.subj_colors[str(cell_index[0])]]]))
                else:
                    self.sheet.write(row_start, col, text[1], self.cell_format(['BOLD', 'CENTER', 'WRAP',
                            ['COLOR', self.subj_colors[str(cell_index[0])]], 
                            ['FONTCOLOR', self.subj_font_color]]))
 
    def get_subject(self, subjects):
        time_index = 0
        printed_sub = []

        #process: nested for loops (subject nests time)
        for subject in subjects: #list of subjects
            if subject not in printed_sub: #skip if duplicate
                printed_sub.append(subject)   
            else:
                continue

            subj_time, subj_day, subj_room = self.get_time_day_room(subject) #gets time and day

            if isinstance(subj_time[0], list):
                #if the subject has more than one time range.
                #activates get_cell_coords on time range times
                for index, (t, d, r) in enumerate(zip(subj_time, subj_day, subj_room), start = 1):
                    yield self.get_cell_coords(subject, t, d, r)
            else:
                #if the subject only has one time range
                yield self.get_cell_coords(subject, subj_time, subj_day, subj_room)

    #revise this. the code is repeating. you should make it into a single loop, calling the coords. instead having 2 loops.
    #this is redundant.
    def get_cell_coords(self, subject, time_range, day_list, room):
        print('Arguments:', subject, time_range, day_list, room)
        day_enabled = self.state['enable_day']

        #loops through days in a week
        for col_index, day in enumerate(self.schedule.day_list):
            #activates if the day list is within that day (e.g. Monday in MWF)
            if not day_enabled or self.schedule.regex_day(col_index, day_list):

                #converts time from Excel into selected time format
                time_range = self.schedule.str_to_time(time_range)
                time_range = self.schedule.time_to_str(time_range)

                #finds the time's correct cell coordinates by using list.index()
                #it automatically sets index to 0 if the day_list is disabled or it's not found
                row_start = self.schedule.time_list.index(time_range[0])
                row_end = self.schedule.time_list.index(time_range[1])
                print('--',col_index, (row_start, row_end), subject, time_range)

                yield subject, col_index, row_start, row_end, room

        print()

    def get_time_day_room(self, subject):
        #time/day keys
        time_keys = [self.input_data['time_key_0'], self.input_data['time_key_1']]
        day_key = self.input_data['day_key']

        #sees if days writing is enabled
        day_enabled = self.state['enable_day']
        
        #sees if the 2nd time_key is included
        twice_enabled = self.state['enable_time_twice']
        allowed = time_keys if twice_enabled else time_keys[:1] 

        room_key = self.input_data['room_key']
        room_enabled = self.state['enable_add_classroom']

        subj_time = [self.schedule.df.loc[subject][key] for key in allowed]                        #PARAMETERS: 'FROM TIME', 'TO TIME' (headers) (/)
        
        subj_day = self.schedule.df.loc[subject][day_key] if day_enabled else ' ' #gets class days #PARAMETER: DAYS (header on Excel) (/)
        subj_room = self.schedule.df.loc[subject][room_key] if room_enabled else '' #gets the classroom

        
        #activates if the subject has duplicate due to having different time/day/classroom
        if day_enabled and isinstance(subj_day, pd.core.series.Series): 
            subj_day = subj_day.to_list()
            subj_room = subj_room.to_list() if room_enabled else ' '
            if twice_enabled:
                subj_time = [[x,y] for x,y in zip(subj_time[0], subj_time[1])]
            else:
                subj_time = subj_time[0].to_list()
                subj_time = [[s] for s in subj_time]

        #checks if the time range is concatenated (in - out)
        if not twice_enabled:
            for i, time_range in enumerate(subj_time):
                if isinstance(time_range, list):
                    print('activated')
                    for sub_range in time_range:
                        subj_time[i] = self.schedule.strip_time(sub_range)

                else:
                    subj_time = self.schedule.strip_time(time_range)
        return subj_time, subj_day, subj_room

def error(message = None):
    print('ERROR:', message)
    if message == None:
        messagebox.showwarning('An error occured', 'There is something wrong with the file.\n'
                                                    'File is not created in the process')
    else:
        messagebox.showwarning('An error occured', message)

def create_schedule():
    sched = SchedMaker(data)
    e = ExcelWriter(sched, data)

if __name__ == "__main__":
    create_schedule()