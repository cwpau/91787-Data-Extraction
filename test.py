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

def getmonth(filenameforoutput):
    a = filenameforoutput.lower()
    print(a)
    months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    month_no = [1,2,3,4,5,6,7,8,9,10,11,12]
    zip_iterator = zip(month_no , months)
    month_dict = dict(zip_iterator)
    print(month_dict)
    if any(month in a for month in months):
        results = re.findall(r'_[a-z]{3}_', a)  #find month
        results = results[0].replace("_", "")
        month_in_number = find_key(month_dict, results)
        print(results)
        print(month_in_number)
        return month_in_number
    else:
        print("month info is not found in filename")

def find_key(input_dict, value):
    return next((k for k, v in input_dict.items() if v == value), None)

def gethour(filenameforoutput):
    a = filenameforoutput.lower()
    print(a)

    if "hr" in a:
        a = a.partition('hr')[1] + a.partition('hr')[2]
        results = re.findall(r'hr_[0-9]', a)
        results = int(results[0].replace("hr_", ""))
        return results
    else:
        print("Hour info is not found in filename")

def getweather(df_imported):
    T = df_imported.iat[15,5]
    print(T)
    RH = df_imported.iat[15,7]
    print(RH)
    return T, RH

def exportfunc(df_result):
    global first
    if first:
        try:
            os.remove("output.xlsx")
        except:
            print("File not existing, OK to proceed")

        df_export = pd.DataFrame(data=df_result)
        df_export.to_excel("output.xlsx", engine='openpyxl', sheet_name='running', merge_cells=True)
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
                  "MC.4","HGV9.4","NFB9.4", "ALL", "ALL.1", "ALL.2", "ALL.3", "ALL.4"]

target_pollutants = ["Pollutant Name: Oxides of Nitrogen", "Pollutant Name: PM30", "Pollutant Name: PM10",
                     "Pollutant Name: PM2.5", "Pollutant Name: Nitrogen Dioxide"]

sep = "*****************************************************************************************"
df_toexport = pd.DataFrame()
first = True
for filesname in filesopened:
    # with open(filesname) as file:
    #     lines = file.readlines()
    #     lines = [line.rstrip() for line in lines]

    # print(lines)
    filenameforoutput = filesname.rsplit("/",1)[1]

    # print(filesname)
    print(filenameforoutput)

    month = getmonth(filenameforoutput)             #not a list
    print(type(month))
    hour = gethour(filenameforoutput)           #not a list
    print(type(hour))

    df = pd.read_csv(filesname, sep=',',  skiprows = 16)    #import raw output file
    T = getweather(df)[0]
    RH = getweather(df)[1]

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

    print(df)
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
        df_temp = df.iloc[i:(i+3+numberofspeed)]
        # print(df_temp)

        pol_colname = df_temp.loc[:,'Speed'].iat[0].split(": ")[1]
        # print("pollutant is:" ,pol_colname)
        #
        # ffill_rows = [i]                                                    #fill row with pollutant as heading
        # df_temp.loc[ffill_rows] = df_temp.loc[ffill_rows].ffill(axis=1)
        # #swapping column names
        df2 = df_temp.iloc[1:, 1:]

        df2.columns = [pol_colname]*len(df2.columns)
        df2.iloc[0] = df2.iloc[0]+","+df2.iloc[1]
        df2.columns = pd.MultiIndex.from_product([[pol_colname], df2.iloc[0]])
        df2.insert(0, 'Speed', df_temp.iloc[:, 0])  # insert speed back to df
        df3 = df2[2:]
        # print(df_result)
        # print(df3)
        # print("DF2**", df2)



        # df2.columns.values[[0]] = ['Speed']
        #
        # print(df2.columns)

        #
        # print("dfresult", df_result)
        # print("DF3", df3)
        # df_result = df_result.merge(df3, how = 'left', left_on ='Speed', right_on ='Speed')
        df_result = df_result.merge(df2, on = 'Speed')
        # print(df_result)
    print(df_result)
    df_result.insert(0, 'Temperature', T)
    df_result.insert(0, 'RH', RH)
    df_result.insert(0, 'Hour', hour)    #insert hour
    df_result.insert(0, 'Month', month)    #insert month
    #export
    df_toexport = pd.concat([df_toexport, df_result])

exportfunc(df_toexport)
