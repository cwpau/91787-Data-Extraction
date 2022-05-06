import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
import numpy as np
from tkinter import *
import tkinter.filedialog as fd
import random
import re
import xlsxwriter
import os
import datetime

def getmonth(filenameforoutput):
    a = filenameforoutput.lower()
    print(a)
    months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    if any(month in a for month in months):
        results = re.findall(r'_[a-z]{3}_', a)  #find month

        try:
            results = results[0].strip("_")
            datetime_object = datetime.datetime.strptime(results, "%b")
            month_number = datetime_object.month
        except ValueError:
            results = results[1].strip("_")
            datetime_object = datetime.datetime.strptime(results, "%b")
            month_number = datetime_object.month
            print("Ensure month abbreviation is given e.g. _Jan_ and no other underscore+3 characters+underscore combination is in filename")
        return month_number
    else:
        print("month info is not found in filename")

def gethour(filenameforoutput):
    a = filenameforoutput.lower()
    print(a)

    if "hr" in a:
        a = a.partition('hr')[1] + a.partition('hr')[2]
        print(a)
        results = re.findall(r'hr_[0-9]{1,2}', a)       #accepts 1 or 2 digit
        hourdigit = ""
        for c in results[0]:            #re.findall returns a list
            if c.isdigit():
                hourdigit = hourdigit + c
        print(hourdigit)
        return hourdigit
    else:
        print("Hour info is not found in filename")

def getweather(df_imported, numberofspeed):
    print(df_imported.iat[0,0])
    T = [df_imported.iat[1+numberofspeed, 4]]
    T = T[0].partition(':')[2].strip()

    RH = [df_imported.iat[1+numberofspeed, 6]]
    RH = RH[0].partition(':')[2].strip()
    return T, RH

def exportfunc(df_result):
    global first
    if first:
        try:
            os.remove("output.xlsx")
        except:
            print("File not existing, OK to proceed")

        df_export = pd.DataFrame(data=df_result)
        writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        df_export.to_excel(writer, sheet_name='running', merge_cells=True, index=False, freeze_panes=(3, 0))

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['running']
        for i in [1,2]:
            for j in [0,1,2,3]:
                 worksheet.write(i, j, None)

        # Add a header format.
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1})
        # Write the column headers with the defined format.
        for col_num, value in enumerate(df_result.columns.values):
            worksheet.write(0, col_num , value, header_format)

        #merge headers with same value  #hard-code
        # merge_format = workbook.add_format({'align': 'center'})
        # worksheet.merge_range('F1:W1', merge_format)
        # worksheet.merge_range('X1:AO1')
        # worksheet.merge_range('AP1:BG1')
        # worksheet.merge_range('BH1:BY')
        # worksheet.merge_range('BZ1:CQ1')

        writer.save()
        first = False
    else:
        pass


root = tk.Tk()
root.withdraw()
filesopened = fd.askopenfilenames(parent=root, title='Choose the files')
root.destroy()

print(filesopened)

target_columns = ["Speed", "PC.4", "TAXI.4", "LGV3.4", "LGV4.4", "LGV6.4", "HGV7.4", "HGV8.4",
                  "PLB.4", "PV4.4", "PV5.4",  "NFB6.4","NFB7.4", "NFB8.4",  "FBSD.4", "FBDD.4",
                  "MC.4","HGV9.4","NFB9.4"]     # "ALL", "ALL.1", "ALL.2", "ALL.3", "ALL.4"

target_pollutants = ["Pollutant Name: Oxides of Nitrogen", "Pollutant Name: PM30", "Pollutant Name: PM10",
                     "Pollutant Name: PM2.5", "Pollutant Name: Nitrogen Dioxide"]

sep = "*****************************************************************************************"
df_toexport = pd.DataFrame()
first = True
first_merge = True
for filesname in filesopened:
    # with open(filesname) as file:
    #     lines = file.readlines()
    #     lines = [line.rstrip() for line in lines]

    # print(lines)
    filenameforoutput = filesname.rsplit("/",1)[1]

    # print(filesname)
    print(filenameforoutput)

    month = getmonth(filenameforoutput)
    # print(type(month))
    hour = gethour(filenameforoutput)
    # print(type(hour))

    df = pd.read_csv(filesname, sep=',',  skiprows = 16)    #import raw output file
    df_import = df.copy()


    df = df[target_columns]                                 #go to interested columns only
    #slice df to running only
    endrow = df[df['Speed'] == sep ].index[0]
    # print(endrow)
    df = df[:(endrow-6)]                                    #running only

    #count unique speeds in running

    itemsinspeed = df['Speed'].unique().tolist()
    list = [i.strip() for i in itemsinspeed]
    list = [s for s in list if s.isdigit()]
    numberofspeed = len(list)                                  #no. of speed count
    # print(numberofspeed)

    T = getweather(df_import, numberofspeed)[0]
    RH = getweather(df_import, numberofspeed)[1]
    print(T, RH)

    # print(df)
    df_result = df.iloc[:, 0]
    df_result.loc[-1] = 'Speed'  # adding a row
    df_result.index = df_result.index + 1  # shifting index
    df_result.sort_index(inplace=True)
    df_result.rename('Speed')
    df_result = df_result[0:(numberofspeed+2)]              #creates the series
    df_result = df_result.to_frame()
    # print(df_result)
    df_result.columns = ['Speed']
    # print(df_result)
    # pd.concat(df_result, axis=1)
    # df_temp.insert(0, 'Hour', hour)    #insert hour
    # df_temp.insert(0, 'Month', month)    #insert month
    df_result = df_result[2:]
    # print(df_result)
    # print(df_result.info())

    #identify and go to specific pollutants

    row_list = df.loc[df['Speed'].isin(target_pollutants)].index.values
    for i in row_list:
        df_temp = df.iloc[i:(i+3+numberofspeed)]    #go to corresponding data of pollutant
        print(df_temp)
        pol_colname = df_temp.loc[:,'Speed'].iat[0].split(": ")[1]

        # #swapping column names
        df2 = df_temp.iloc[1:, 1:]      # only select pollutant data
        df2.columns = [pol_colname]*len(df2.columns)        #change column names to pollutant name

        # df2.columns = pd.MultiIndex.from_product([[pol_colname], df2.iloc[0]])
        df2.insert(0, 'Speed', df_temp.iloc[:, 0])  # insert speed back to df
        print(df2)

        if first_merge:
            df_result = df_result.merge(df2, how = 'right', on = 'Speed')

        else:
            df3 = df2[2:]
            df_result = df_result.merge(df3, how = 'left', on = 'Speed')
    first_merge = False

    print(df_result)
    df_result.insert(0, 'RH', RH)    #insert RH
    df_result.insert(0, 'Temperature', T)    #insert Temp
    df_result.insert(0, 'Hour', hour)    #insert hour
    df_result.insert(0, 'Month', month)    #insert month
    print(df_result)
    #export
    df_toexport = pd.concat([df_toexport, df_result])

exportfunc(df_toexport)
