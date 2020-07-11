#Backend
import pandas as pd, xlsxwriter as xl
import os, re, json
from datetime import datetime
from tkinter import messagebox

filename = 'settings.json'
#checks if the file exists (required)
if not os.path.isfile(filename):
    raise Exception('JSON file is missing.')

#loads JSON
with open(filename, 'r') as f:
    json_data = json.load(f)

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
                try:
                    time = time.strip()
                    time = time.split('-')
                    time_list.extend(time)
                except:
                    return error(f'The time column ({chosen_time[0]}) has incorrect formatting\n'
                            'Please check the file and try again.')

        else:
            time_list = self.df[chosen_time[0]].to_list() + self.df[chosen_time[1]].to_list()             #PARAMETERS: FROM TIME, TO TIME (headers on excel) (/)

        self.time_list = self.time_sort(time_list) #sorts the time
        self.day_list = self.get_day_list(day_mode.upper())
        return print(self.day_list, self.time_list, sep='\n')

    def get_day_list(self, mode):
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
            [re.compile(r'(?i:[0-9]?[a-zA-Z]$)'), '%H:%M']
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
            if t[i][-1] != 'm':
                time_in += 'm'                                                                            #PARAMETER: added 'm' (8:00a -> 8:am) HARDCODED (/)
            t[i] = datetime.strptime(time_in, self.str_format)

        #converts into datetime object and sort
        t.sort()
        min_t, max_t = t[0].hour, t[-1].hour
        t_range = max_t - min_t

        #converts into string
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
        self.write_day()
        self.write_time()
        self.write_subject()

    def new(self):
        self.format_list = {
            'BOLD': {'bold':True},
            'CENTER': {'align': 'center', 'valign': 'vcenter'},
            'BORDER': {'border': 1}
        }
        self.color = self.data['system_colors']
        self.subj_colors = self.data['data']['colors']
        self.subj_font_color = self.data['data']['font_color'].upper()
        self.preset = {'header': [['COLOR', 'ACCENT_LIGHT'], 'BOLD', 'CENTER', 'BORDER']}
        self.cell_height = 16
        self.cell_width = 12
    
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
            print(x)
            if isinstance(x, list):
                if x[0] == 'COLOR':
                    format_list[index] = {'fg_color': x[1]} if re.search(r'^#', x[1]) else {'fg_color': self.color[x[1]]}
                elif x[0] == 'BORDER':
                    format_list[index] = self.border_format(x[1:]) #activates self.border_format if it's inputted as list
                elif x[0] == 'FONTCOLOR':
                    print(x[1])
                    format_list[index] = {'font_color': x[1]} if re.search(r'^#', x[1]) else {'font_color': self.color[x[1]]}
            else:
                format_list[index] = self.format_list[x] #searches through self.format list dict

        x = {}
        for format_dict in format_list: #format_list = [{}, {}]
            for key in format_dict: #format_dict = collection of keys in a dict
                x[key] = format_dict[key] #unpacks values of format_arr

        return self.book.add_format(x)
    
    def write_day(self):
        self.col += 2
        for i, day in enumerate(self.schedule.day_list):
            self.sheet.write(self.row, self.col + i, day, self.cell_format(self.preset['header']))
        else:
            self.set_col(self.col-1, self.col+i, self.cell_width)

        self.row += 1
        self.col -= 2

    def write_time(self):
        merge_n = 1
        start_cell = [0,0]
        #cell format
        hour_format = self.cell_format([['COLOR', 'ACCENT_LIGHT'], 'BOLD', 'CENTER', ['BORDER', 'right']])
        even_color = self.cell_format([['COLOR', 'ACCENT_LIGHT'], 'CENTER', 'BORDER'])
        odd_color = self.cell_format([['COLOR', 'ACCENT_DARK'], 'CENTER', 'BORDER'])
        cell_switch = True #makes a coloring switch regardless of its position

        for i, time in enumerate(self.schedule.time_list):
            #writes the time
            self.sheet.write(self.row, self.col+1, time, hour_format)

            cell_color = even_color if cell_switch else odd_color
            #write the hours
            #detects if last hour matches current hr
            if i != 0 and time[:2] == last_hr:
                if merge_n == 1: #sets starting cell
                    start_cell = [self.row-1, self.col]
                merge_n +=1

            #writes hour time & switches the cell_color
            elif i != 0 and merge_n == 1:
                self.sheet.write_number(self.row-1, self.col, int(last_hr), cell_color)
                cell_switch = not cell_switch
            elif merge_n != 1:
                self.sheet.merge_range(start_cell[0], start_cell[1], self.row-1, self.col, int(last_hr), cell_color)
                cell_switch = not cell_switch
                merge_n = 1

            last_hr = time[:2]
            self.set_row(self.row, self.cell_height) # sets row width
            self.row += 1
        else:
            #write the time
            if merge_n != 1:
                self.sheet.merge_range(start_cell[0], start_cell[1], self.row-1, self.col, int(last_hr), cell_color)
            else:
                self.sheet.write_number(self.row-1, self.col, int(last_hr), cell_color)

    def write_subject(self):
        self.col += 2
        self.row -= len(self.schedule.time_list)
        #yields subject, col, row_start, row_end
        subject_list = self.get_subject(self.schedule.df.index)
        for subject in subject_list:
            for cell_index in subject:
                col = self.col + cell_index[1]
                row_start = self.row + cell_index[2]
                row_end = self.row + cell_index[3]
                self.sheet.merge_range(row_start, col, row_end, col, cell_index[0], 
                    self.cell_format(['CENTER', ['COLOR', self.subj_colors[cell_index[0]]], ['FONTCOLOR', self.subj_font_color]]))
 
    def get_subject(self, subjects):
        time_index = 0
        printed_sub = []

        #process: nested for loops (subject nests time)
        for subject in subjects: #list of subjects
            if subject not in printed_sub: #skip if duplicate
                printed_sub.append(subject)   
            else:
                continue

            subj_time, subj_day = self.get_time_day(subject) #gets time and day

            if isinstance(subj_time[0], list):
                #if the subject has more than one time range.
                #activates get_cell_coords on time range times
                for index, time_range in enumerate(subj_time, start = 1):
                    yield self.get_cell_coords(time_range, subj_day, subject, index)
            else:
                #if the subject only has one time range
                yield self.get_cell_coords(subj_time, subj_day, subject)

        #formatting the columns
        print('CURRENT POS (col, row)', end=': ')
        print(self.col, self.row)

    def get_cell_coords(self, time_range, day_list, subject, restrict_val = None):
        #restrict control value
        val = 0
        #loops through days in a week
        for col_index, day in enumerate(self.schedule.day_list):
            #activates if the day list is within that day (e.g. Monday in MWF)
            if self.schedule.regex_day(col_index, day_list):
                val += 1
                #adds 'm' if it ends on 'a/p'
                if self.schedule.str_format == '%I:%M%p' and time_range[0][-1] in ['a', 'p']:
                    time_range = [(time + 'm').upper() for time in time_range]

                #improvised xor value
                if restrict_val == None or (restrict_val != None and restrict_val == val):
                    #finds the time's correct cell coordinates by using list.index()                  
                    row_start = self.schedule.time_list.index(time_range[0])
                    row_end = self.schedule.time_list.index(time_range[1])
                    yield subject, col_index, row_start, row_end     

    def get_time_day(self, subject):
        subj_time = [self.schedule.df.loc[subject][x + ' TIME'] for x in ['FROM', 'TO']]          #PARAMETERS: 'FROM TIME', 'TO TIME' (headers)
        subj_day = self.schedule.df.loc[subject]['DAYS'] #gets class days                         #PARAMETER: DAYS (header on Excel)

        #activates if the subject has duplicate due to having different time/day/classroom
        if isinstance(subj_day, pd.core.series.Series): 
            subj_day = ''.join(subj_day)
            subj_time = [[x,y] for x,y in zip(subj_time[0], subj_time[1])]

        return subj_time, subj_day

def error(message = None):
    if message == None:
        messagebox.showwarning('File not created', 'There is something wrong with the file.\n'
                                                    'File is not created in the process')
    else:
        messagebox.showwarning('File not created', message)

def create_schedule():
    sched = SchedMaker(json_data)
    e = ExcelWriter(sched, json_data)

if __name__ == "__main__":
    create_schedule()