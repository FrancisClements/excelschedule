#Backend
import pandas as pd, xlsxwriter as xl
import os
from datetime import datetime


class SchedMaker:
    def __init__(self):
        self.df = pd.read_excel('sched_tbl.xlsx', index_col = 'CODE')
        time_list = self.df['FROM TIME'].to_list() + self.df['TO TIME'].to_list()                     #PARAMETERS: FROM TIME, TO TIME (headers on excel)
        day_mode = 'PARTIAL'                                                                          #PARAMETERS: 'FULL', 'PARTIAL', 'INITIAL'

        self.get_time_list()
        self.time_list = self.time_sort(time_list) #sorts the time
        self.day_list = self.get_day_list(day_mode)
        # print(self.hr_list, self.time_list, sep='\n')

    def get_day_list(self, mode):
        week_list = {
            'FULL': ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
            'INITIAL': ['M', 'T', 'W', 'TH', 'F', 'S']
        }
        week_list['PARTIAL'] = [x[:3].upper() for x in week_list['FULL']]
        print(week_list[mode])
        return week_list[mode]

    def get_time_list(self):#converts dataframe into a list
        self.df_list = []
        for sub in self.df.iterrows():
            self.df_list.append(sub)

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
            t[i] = time_in.strftime(str_format)
        
        return t


class ExcelWriter:
    def __init__(self, schedule):
        self.schedule = schedule
        print(self.schedule.time_list)
        if os.path.isfile('HAUsched.xlsx'):
            os.remove('HAUsched.xlsx')

        self.book = xl.Workbook('HAUsched.xlsx')
        self.sheet = self.book.add_worksheet()
        self.offset = [0,0]
        self.row = self.offset[0]
        self.col = self.offset[1]
        self.new()
        self.write_time()
        self.book.close()

    def new(self):
        self.format_list = {
            'BOLD': {'bold':True},
            'CENTER': {'align': 'center', 'valign': 'vcenter'}
        }
        self.color = {
            'ACCENT_LIGHT': '#d9d9d9',
            'ACCENT_DARK': '#bfbfbf'
        }
    
    def cell_format(self, format_arr):
        #changes format into str
        for index, x in enumerate(format_arr):
            format_arr[index] = self.format_list[x]

        x = {}
        for base in format_arr:
            for nest in base:
                x[nest] = base[nest]

        return self.book.add_format(x)
    
    def write_day(self):
        pass



    def write_time(self):
        # writes the time
        merge_n = 1
        start_cell = [0,0]
        for i, time in enumerate(self.schedule.time_list):
            self.sheet.write(self.row, self.col+1, time)

            #detects if last hour matches current hr
            if i != 0 and time[:2] == last_hr:
                if merge_n == 1: #sets starting cell
                    start_cell = [self.row-1, self.col]
                merge_n +=1

            #writes hour time
            elif i != 0 and merge_n == 1:
                self.sheet.write_number(self.row-1, self.col, int(last_hr))
                # print('--------printed', last_hr)

            #merges the cell
            elif merge_n != 1:
                self.sheet.merge_range(start_cell[0], start_cell[1], self.row-1, self.col, int(last_hr))
                # print('combo', merge_n)
                # print('--------printed', last_hr)
                merge_n = 1

            # print(merge_n, time)
            last_hr = time[:2]
            self.row += 1
        else:
            #write the time
            if merge_n != 1:
                self.sheet.merge_range(start_cell[0], start_cell[1], self.row-1, self.col, int(last_hr))
                # print('combo', merge_n)
            else:
                self.sheet.write_number(self.row-1, self.col, int(last_hr))
            # print('--------printed', last_hr)

        #formatting the columns
        self.sheet.set_column(self.col+1, self.col+1, 11, self.cell_format(['BOLD', 'CENTER'])) 
        self.sheet.set_column(self.col, self.col, None, self.cell_format(['BOLD', 'CENTER']))
                    


e = ExcelWriter(SchedMaker())
