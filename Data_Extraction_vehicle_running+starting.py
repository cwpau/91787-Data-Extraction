import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
import numpy as np
from tkinter import *
import tkinter.filedialog as fd
import re
import xlsxwriter
import os
import datetime

def getmonth(filenameforoutput):
    a = filenameforoutput.lower()
    # print(a)
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
            print("Error: Ensure month abbreviation is given e.g. _Jan_ and no other underscore+3 characters+underscore combination is in filename")
        return month_number
    else:
        print("Error: month info is not found in filename")

def gethour(filenameforoutput):
    a = filenameforoutput.lower()
    # print(a)

    if "hr" in a:
        a = a.partition('hr')[1] + a.partition('hr')[2]
        # print(a)
        results = re.findall(r'hr_[0-9]{1,2}', a)       #accepts 1 or 2 digit
        hourdigit = ""
        for c in results[0]:            #re.findall returns a list
            if c.isdigit():
                hourdigit = hourdigit + c
        # print(hourdigit)
        return hourdigit
    else:
        print("Error: Hour info is not found in filename")

def getweather(df_imported, numberofspeed):
    # print(df_imported.iat[0,0])
    T = [df_imported.iat[1+numberofspeed, 4]]
    T = T[0].partition(':')[2].strip()

    RH = [df_imported.iat[1+numberofspeed, 6]]
    RH = RH[0].partition(':')[2].strip()
    return T, RH

def get_numberofspeed(df_running):
    itemsinspeed = df_running['Speed'].unique().tolist()
    list = [i.strip() for i in itemsinspeed]
    list = [s for s in list if s.isdigit()]
    return len(list)                                  #no. of speed count

def get_numberoftime(df_starting):
    itemsintime = df_starting['Speed'].unique().tolist()           #still speed, later change to time
    list = [i.strip() for i in itemsintime]
    list = [s for s in list if s.isdigit()]
    return len(list)                                  #no. of time count

#create constant for running data
def create_constantdf(df, colname, length):
    df_result = df.iloc[:, 0]
    df_result.loc[-1] = colname  # adding a row
    df_result.index = df_result.index + 1  # shifting index
    df_result.sort_index(inplace=True)
    df_result.rename(colname)
    df_result = df_result[0:(length+2)]              #creates the series
    df_result = df_result.to_frame()
    df_result.columns = [colname]
    return df_result[2:]

#identify and go to specific pollutants (RUNNING)
def searchpollutant(df, target_pollutants, length, df_result, df_toexport, colname):
    global first_merge
    global first_merge_starting
    df.reset_index()
    row_list = df.loc[df['Speed'].isin(target_pollutants)].index.values     #both swtarting and running still using SPEED as colname
    # print(row_list)
    for i in row_list:  #loop through each pollutant row range

        if colname == 'Speed':
            df_temp = df.iloc[i:(i+3+length)]    #go to corresponding data of pollutant
            pol_colname = df_temp.loc[:, 'Speed'].iat[0].split(": ")[1]
        elif colname == 'Time':
            df_temp = df.iloc[i-endrow-7: (i-endrow-7+3+length)]
            pol_colname = df_temp.loc[:, 'Speed'].iat[0].split(": ")[1]
            # print(df_temp)


        # #swapping column names
        df2 = df_temp.iloc[1:, 1:]      # only select pollutant data
        df2.columns = [pol_colname]*len(df2.columns)        #change column names to pollutant name

        df2.insert(0, colname, df_temp.iloc[:, 0])  # insert Speed/Time back to df
        # print(df2)

        if colname == 'Speed':
            if first_merge: #
                df_result = df_result.merge(df2, how = 'right', on = colname)
                # print("FIRST_MERGE",first_merge)

            else:
                df3 = df2[2:]
                df_result = df_result.merge(df3, how = 'left', on = colname)
                # print("FIRST_MERGE",first_merge)

        else:
            if first_merge_starting: #
                df_result = df_result.merge(df2, how = 'right', on = colname)
                # print("FIRST_MERGE",first_merge)

            else:
                df3 = df2[2:]
                df_result = df_result.merge(df3, how = 'left', on = colname)
                # print("FIRST_MERGE",first_merge)


    if colname == 'Speed':
        first_merge = False
    elif colname == 'Time':
        first_merge_starting = False

    df_result.insert(0, 'RH', RH if colname == 'Speed' else 'ALL')     #insert RH
    df_result.insert(0, 'Temperature', T)    #insert Temp
    df_result.insert(0, 'Hour', hour)    #insert hour
    df_result.insert(0, 'Month', month)    #insert month
    # print(df_result)

    #export
    return pd.concat([df_toexport, df_result])

def exportfunc(df_result, df_result_ws2):
    global first
    if first:
        try:
            os.remove("output.xlsx")
        except:
            print("output.xlsx File not existing, OK to proceed without removing file")

        #write first DF to first WS
        df_export = pd.DataFrame(data=df_result)
        writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        df_export.to_excel(writer, sheet_name='running', merge_cells=True, index=False, freeze_panes=(3, 0))

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['running']
        for i in [1,2]:
            for j in [0,1,2,3]:
                 worksheet.write(i, j, " ")

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

        #WS2
        df_export2 = pd.DataFrame(data=df_result_ws2)
        df_export2.to_excel(writer, sheet_name='starting', merge_cells=True, index=False, freeze_panes=(3, 0))

        worksheet = writer.sheets['starting']
        for i in [1,2]:
            for j in [0,1,2,3]:
                 worksheet.write(i, j, " ")

        # Add a header format.
        # Write the column headers with the defined format.
        for col_num, value in enumerate(df_result_ws2.columns.values):
            worksheet.write(0, col_num , value, header_format)

        writer.save()
        first = False
    else:
        pass

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    filesopened = fd.askopenfilenames(parent=root, title='Choose the files')
    root.destroy()

    print("files imported:", filesopened)

    target_columns = ["Speed", "PC.4", "TAXI.4", "LGV3.4", "LGV4.4", "LGV6.4", "HGV7.4", "HGV8.4",
                      "PLB.4", "PV4.4", "PV5.4",  "NFB6.4","NFB7.4", "NFB8.4",  "FBSD.4", "FBDD.4",
                      "MC.4","HGV9.4","NFB9.4"]     # "ALL", "ALL.1", "ALL.2", "ALL.3", "ALL.4"

    target_pollutants = ["Pollutant Name: Oxides of Nitrogen", "Pollutant Name: PM30", "Pollutant Name: PM10",
                         "Pollutant Name: PM2.5", "Pollutant Name: Nitrogen Dioxide"]

    sep = "*****************************************************************************************"



    df_toexport = pd.DataFrame()
    df_toexport2 = pd.DataFrame()
    first = True
    first_merge = True
    first_merge_starting = True
    for filesname in filesopened:

        # with open(filesname) as file:
        #     lines = file.readlines()
        #     lines = [line.rstrip() for line in lines]

        # print(lines)
        filenameforoutput = filesname.rsplit("/",1)[1]

        # print(filesname)
        print(filenameforoutput, "extracting...")

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

        starting_endrow = df[df['Speed'] == sep ].index[1]
        # print(starting_endrow)
        df_running = df[:(endrow-6)]                                    #running only
        # print(df_running)
        df_starting = df[(endrow+7):(starting_endrow-6)]               #starting only
        # print("STARTING", df_starting)

        ### #count unique speeds in running
        numberofspeed = get_numberofspeed(df_running)

        #count unique timese in starting
        numberoftime = get_numberoftime(df_starting)

        T = getweather(df_import, numberofspeed)[0]
        RH = getweather(df_import, numberofspeed)[1]
        # print(T, RH)

        df_result = create_constantdf(df_running, 'Speed', numberofspeed)

        #create constant for starting data
        df_result_ws2 = create_constantdf(df_starting, 'Time', numberoftime)
        # print(df_result_ws2)

        df_toexport = searchpollutant(df_running, target_pollutants, numberofspeed, df_result, df_toexport, 'Speed')

        df_toexport2 = searchpollutant(df_starting, target_pollutants, numberoftime, df_result_ws2, df_toexport2, 'Time')
        df_toexport2 = df_toexport2.drop_duplicates()


    #export to excel
    exportfunc(df_toexport, df_toexport2)
    print("ALL files successfully extracted and exported to output.xlsx")