#Backend
import pandas as pd, xlsxwriter as xl
import os, re
from datetime import datetime


class SchedMaker:
    def __init__(self):
        self.df = pd.read_excel('sched_tbl.xlsx', index_col = 'CODE')                                 #PARAMETERS: 'CODE' (header on excel)
        time_list = self.df['FROM TIME'].to_list() + self.df['TO TIME'].to_list()                     #PARAMETERS: FROM TIME, TO TIME (headers on excel)
        day_mode = 'PARTIAL'                                                                          #PARAMETERS: 'FULL', 'PARTIAL', 'INITIAL'

        self.time_list = self.time_sort(time_list) #sorts the time
        self.day_list = self.get_day_list(day_mode)
        # print(self.hr_list, self.time_list, sep='\n')

    def get_day_list(self, mode):
        self.week_list = {
            'FULL': ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
            'INITIAL': ['M', 'T', 'W', 'TH', 'F', 'S']
        }
        self.week_list['PARTIAL'] = [x[:3].upper() for x in self.week_list['FULL']] #Tuesday = Tue
        word = ''.join(self.df['DAYS'])
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
        t = list(set(t))
        str_format = '%I:%M%p'                                                                        #FORMAT EXAMPLE: 4:45pm
        #converts into datetime object and sort
        for i, time_in in enumerate(t):
            time_in += 'm'
            t[i] = datetime.strptime(time_in, str_format)

        t.sort()
        min_t, max_t = t[0].hour, t[-1].hour
        t_range = max_t - min_t

        #covnerts into string
        for i, time_in in enumerate(t):
            t[i] = time_in.strftime(str_format).lower()                                               #Set to lower case PM -> pm
        
        return t


class ExcelWriter:
    def __init__(self, schedule):
        self.schedule = schedule
        if os.path.isfile('HAUsched.xlsx'):
            os.remove('HAUsched.xlsx')

        self.book = xl.Workbook('HAUsched.xlsx')
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
            'CENTER': {'align': 'center', 'valign': 'vcenter'}
        }
        self.color = {
            'ACCENT_LIGHT': '#d9d9d9',
            'ACCENT_DARK': '#bfbfbf'
        }
        self.preset = {'header': [['COLOR', 'ACCENT_LIGHT'], 'BOLD', 'CENTER']}
        self.cell_height = 16
        self.cell_width = 12
    
    def set_col(self, a, b, width):
        self.sheet.set_column(a, b ,width)

    def set_row(self, pos, width):
        self.sheet.set_row(pos, width)

    def cell_format(self, order_list):
        #format_arr = list
        #changes format into str
        format_list = order_list.copy()
        for index, x in enumerate(format_list):
            if isinstance(x, list):
                format_list[index] = {'fg_color': self.color[x[1]]}
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
        even_color = self.cell_format([['COLOR', 'ACCENT_LIGHT'], 'CENTER'])
        odd_color = self.cell_format([['COLOR', 'ACCENT_DARK'], 'CENTER'])
        cell_switch = True #makes a coloring switch regardless of its position

        for i, time in enumerate(self.schedule.time_list):
            #writes the time
            self.sheet.write(self.row, self.col+1, time, self.cell_format([['COLOR', 'ACCENT_LIGHT'], 'BOLD', 'CENTER']))
            cell_color = even_color if cell_switch else odd_color
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
        self.sheet.write(self.row, self.col, 'koko ni')
        time_index = 0
        printed_sub = []

        #process: nested for loops (subject nests time)
        for subject in self.schedule.df.index: #list of subjects
            if subject not in printed_sub: #skip if duplicate
                printed_sub.append(subject)   
                print(subject, end=': ')
            else:
                continue

            subj_time = [self.schedule.df.loc[subject][x + ' TIME'] for x in ['FROM', 'TO']]          #PARAMETERS: 'FROM TIME', 'TO TIME' (headers)
            subj_day = self.schedule.df.loc[subject]['DAYS'] #gets class days                         #PARAMETER: DAYS (header on Excel)

            #activates if the subject has duplicate due to having different time/day/classroom
            if isinstance(subj_day, pd.core.series.Series): 
                time_index = 0
                subj_day = ''.join(subj_day)
                subj_time = [[x,y] for x,y in zip(subj_time[0], subj_time[1])]

            for index, day in enumerate(self.schedule.day_list): #list of days in a week
                if self.schedule.regex_day(index, subj_day):
                    if isinstance(subj_time[0], list):
                            col_index = self.schedule.time_list.index(subj_time[time_index][0] + 'm') #PARAMETER: Added 'm' for am/pm (HARDCODED)
                            row_index = self.schedule.time_list.index(subj_time[time_index][1] + 'm') #PARAMETER: Added 'm' for am/pm (HARDCODED)
                            print(day, index, col_index, row_index)   
                    else:
                        col_index = self.schedule.time_list.index(subj_time[0] + 'm')                 #PARAMETER: Added 'm' for am/pm (HARDCODED)
                        row_index = self.schedule.time_list.index(subj_time[1] + 'm')                 #PARAMETER: Added 'm' for am/pm (HARDCODED)
                        print(day, index, col_index, row_index)                
                    
            time_index += 1
            print(subj_time, 'day-mark:', subj_day)
            print('')       


        #formatting the columns
        print('CURRENT POS (col, row)')
        print(self.col, self.row)

e = ExcelWriter(SchedMaker())
