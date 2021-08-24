#!/usr/bin/python 
# -*- coding: utf-8 -*-

import os
import xlwings as xw
import pandas as pd
import datetime

class Employee:
    def __init__(self, name = '', attandence_date_number = 0, legal_work_hours = 0, basic_wage = 0, paid_leave = 0, attandence_rate = 0, full_wages = 0): #
        self.name = name # employee's name
        self.attandence_date_number = attandence_date_number # attandance date this month
        self.legal_work_hours = legal_work_hours # legal work hours (in day)this month
        self.besic_wage = basic_wage # basic wages this month
        self.paid_leave = paid_leave # paid leave that used this month
        self.attandence_rate = attandence_rate
        self.full_wages = full_wages

    def att_rate(self): #calculate the attancence rate
        attendence = 0
        attendence = (self.attandence_date_number + self.paid_leave) / self.legal_work_hours
        self.attandence_rate = attendence
        return attendence

    def basic_wages(self): # calculate basic wages 
        basic_wage_this_month = self.attandence_rate * self.basic_wage
        return basic_wage_this_month

    def bonus(self):
        bonus_this_month = (self.full_wages - self.basic_wage) * self.attandence_rate
        return bonus_this_month

def read_attandence_copy_to_newSheet(path):
    app = xw.App(visible=False, add_book = False)
    workbook = app.books.open(path)
    worksheet = workbook.sheets[6]
    print(worksheet)

    cell_area = worksheet.range('A46', 'y46').expand('down').value
    sheet = workbook.sheets[4]
    sheet.range('A1', 'y1').value = cell_area
    workbook.save()
    workbook.close()
    app.quit()
def cal_att(worksheet):
    att_rate = []
    real_att = worksheet.range('A2:A25').expand('right').value
    for i in range(0, len(real_att),2):
        personal_att = real_att[i][1] + real_att[i+1][1]
        if i + 1 > len(real_att):
            break
        if personal_att == 0:
            break
        #att = real_att[i][1] + real_att[i+1][1]
        att_rate.append(personal_att)
    
    return att_rate

'''def cal_att_rate():
    att_rate = []
    
    workbook.sheets[3].range('A2:A12').expand('right').value
    for i in range(0,len(workbook.sheets[3].range('A2:A12').expand('right').value)):
        rate = workbook.sheets[3].range('A2:A12').expand('right').value[2]/workbook.sheets[3].range('A2:A12').expand('right').value[3]
'''
def take_attdence(worksheet1, date_year, date_month): #search the input date and month and copy data for that month and return a list
    values = worksheet1.range('A1', 'Y1').expand('down').options(pd.DataFrame).value
    date = worksheet1.range('A1').expand('down').value
    tarindex = 0
    for i in date:
        if i == datetime.datetime(date_year, date_month, 1, 0, 0):
            tarindex = date.index(i)
    return values.iloc[tarindex - 1]

def copy_selected_line(sau_book, tar_book, selected_year, selected_month):# copy the list to target xslx sheet
    data_copy_from = sau_book.range('A1:Y1').value
    tar_book.range('A1').expand('down').value = take_attdence(sau_book, selected_year, selected_month)

xlsx_tar_path = input('input target xlsx')
basedata_path = input('input base_path')

app = xw.App(visible=False, add_book = False)
workbook = app.books.open(xlsx_tar_path)
basedata = app.books.open(basedata_path)

read_attandence_copy_to_newSheet(xlsx_tar_path)# copy the whole useful sheet block to a new sheet 出勤加班
name_raw = basedata.sheets[0].range('A2:A12').value
name = [] # copy name of employees
for i in range(0, len(name_raw)):
    name.append(name_raw[i])
print(name)

workbook.sheets[3].range('A2').options(transpose = True).value = name

in_year = int(input('输入 年'))# get the values of date and month
in_month = int(input('输入 月'))
in_date = [str(in_year), str(in_month)]

paid_leave = workbook.sheets[3].range('B2').options(transpose = True).value # input paid leave for every employee
'''
for i in range(0, len(name)):
    print('输入 ', name[i], ' 的带薪假')
    paid_leave_in = float(input())
    paid_leave.append(paid_leave_in)
print(paid_leave)
workbook.sheets[3].range('B2').options(transpose = True).value = paid_leave
'''
legal_work_date = [] # fill in the legal work day
workday = basedata.sheets[1].range('A2:B2').expand('down').value
splited = []

for i in workday:
    splited.append(i[0].split('/'))
legal_work_date = workday[splited.index(in_date)][1]
workbook.sheets[3].range('E2').options(transpose = True).value = legal_work_date

copy_selected_line(workbook.sheets[4], workbook.sheets[5], in_year, in_month) # copy selected month to sheet'selected_month'

# fill in attandence(c2, expand down)
values = workbook.sheets[5].range('A2:A23').expand('right').value
print('values  ',values)
att = []
att_rate = []

# calculate att and att_rate
for i in range(0, len(values)-1, 2):
    if i+1 >= len(values):
        break
    if values[i][1] + values[i+1][1] == 0:
        att.append(0)
    att.append(values[i][1]+ values[i+1][1])
    #att_rate.append(a/(values[i][1]+ values[i+1][1]))
for i in range(0, len(att)):
    if att[i] == 0:
        att[i] = 0
    else:
        att_rate.append(att[i]/float(legal_work_date))
#fill in att and att_rate
workbook.sheets[3].range('C2').options(transpose = True).value = att
workbook.sheets[3].range('D2').options(transpose = True).value = att_rate

temp_att_rate = workbook.sheets[3].range('D2').expand('down').value
temp_base_wage = basedata.sheets[0].range('C2').expand('down').value
basic_wage = []
for i in range(0, len(name)): # fill in basic wage
    temp_basic_wage = temp_att_rate[i] * temp_base_wage[i]
    basic_wage.append(temp_basic_wage)

workbook.sheets[3].range('F2').options(transpose = True).value = basic_wage


solid_basic_wage = basedata.sheets[0].range('C2').expand('down').value
full_wage = basedata.sheets[0].range('D2').expand('down').value
temp_att_rate = workbook.sheets[3].range('D2').expand('down').value
bonus = []
print(full_wage)
print(solid_basic_wage)
print(att_rate)

for i in range(0, len(name)):# fill in bonus
    print(i)
    bon = (full_wage[i] - solid_basic_wage[i]) * temp_att_rate[i]
    bonus.append(bon)
print(bonus)
workbook.sheets[3].range('G2').options(transpose = True).value = bonus

#fill in total monthly paid

tot_mon = []
for i in range(0, len(bonus)):
    tot = basic_wage[i] + bonus[i]
    tot_mon.append(tot)
workbook.sheets[3].range('H2').options(transpose = True).value = tot_mon

print("github version")

workbook.save()
workbook.close()
app.quit()