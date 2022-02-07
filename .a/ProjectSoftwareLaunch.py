#!/usr/bin/python3
from tkinter import *
from tkinter import filedialog
from tkinter import Tk, Frame, Label
import tkinter as tk
import os  # for interacting with OS
import time  # for pausing code
from datetime import date
import xlsxwriter  # for writing to excel doc
import xlrd  # for reading xlcel doc
from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth


# =====menu1=======================================================================================
def main():
    mainwindow = Tk()  # build window
    # titles = ['project', 'title', 'date created', 'status', 'age\n(how old info)',
    #          'confidence level\n(1-10)', 'sales rep', 'source', 'proposal date',
    #          '  pca  ', ' contract ']
    titles = ['project', 'title', 'proposal due\n(yy/mm/dd)', 'date created\n(yy/mm/dd)', 'confidence level',
              'sales rep', 'status', 'project age\n-days', 'pca', 'contract', 'awarded', 'PM',
              'est completion\ndate'
              ]  # 'source',
    # status = ['design', 'contracting', 'building', 'warranty']
    numOfRows = 1  # data rows
    entry1 = []
    # ---window layout---------
    mainframe = Frame(mainwindow, bg="black")
    mainframe.pack(fill="both", expand=True)
    mainwindow.title("Kraus Project Program")
    # --------------
    label = Label(mainframe, text="Project Information", bg="black", fg="white", padx=5, pady=5)
    label.config(font=("Arial", 12))
    label.pack(fill="both")
    # --------------
    hor_frame1 = Frame(mainframe)
    btn1 = Button(hor_frame1, text="Import local file",
                  command=lambda: importClientInfo(grid_frame, label2, titles, entry1))
    btn1.grid(column=0, row=0, columnspan=2)
    btn2 = Button(hor_frame1, text="Choose Import File", command=lambda: importFile(grid_frame, label2, titles, entry1))
    btn2.grid(column=2, row=0, columnspan=2)
    btn3 = Button(hor_frame1, text="Save File", command=lambda: exportClientInfo(titles, entry1))
    btn3.grid(column=4, row=0, columnspan=2)
    hor_frame1.pack(fill="x")
    hor_frame2 = Frame(mainframe)
    btn4 = Button(hor_frame2, text="Export to Cloud", command=lambda: exportToCloud(titles, entry1))
    btn4.grid(column=0, row=0, columnspan=2)
    btn5 = Button(hor_frame2, text="Import from Cloud",
                  command=lambda: importFromCloud(grid_frame, label2, titles, entry1))
    btn5.grid(column=2, row=0, columnspan=2)
    hor_frame2.pack(fill="x")
    # -----------------
    hor_frame3 = Frame(mainframe)
    btn5 = Button(hor_frame3, text="Add Row", command=lambda: plusRow(grid_frame, label2, titles, entry1))
    btn5.grid(column=0, row=0, columnspan=2)
    btn6 = Button(hor_frame3, text="Delete Row", command=lambda: deleteRow(label2, titles, entry1))
    btn6.grid(column=2, row=0, columnspan=2)
    hor_frame3.pack(fill="x")
    # ------------------
    row = 0
    x = 0
    grid_frame = Frame(mainframe)
    while row <= numOfRows:
        col = 0
        while col < len(titles):
            if row == 0:
                lbl = Label(grid_frame, text=titles[col], bg="grey", fg="white", pady=5)
                lbl.grid(column=col, row=row, padx=2, pady=5, sticky="nsew")
                grid_frame.grid_columnconfigure(col, weight=1)
            else:
                if titles[col] == 'contract' or titles[col] == 'pca':
                    v = IntVar()
                    entry1.append(tk.Checkbutton(grid_frame, variable=v, width=5))
                    entry1[x].configure(state='disabled', disabledforeground='black')
                    entry1[x].var = v
                    entry1[x].grid(column=col, row=row)
                    grid_frame.grid_columnconfigure(col, weight=0)

                elif titles[col] == 'project':  # button
                    butText = StringVar()
                    entry1.append(tk.Button(grid_frame, textvariable=butText, width=5,
                                            command=lambda: popupMenu(butText, entry1, titles, grid_frame, label2)))
                    butText.set('project #')
                    entry1[x].var = butText
                    entry1[x].grid(column=col, row=row, sticky=NSEW)
                    grid_frame.grid_columnconfigure(col, weight=0)

                elif titles[col] == 'status':
                    ddOptions = StringVar()
                    ddOptions.set('-----')
                    entry1.append(tk.OptionMenu(grid_frame, ddOptions, '---', 'design', 'bidding',
                                                'contracting', 'building', 'warranty', 'closed'))
                    entry1[x].var = ddOptions
                    entry1[x].grid(column=col, row=row)
                    grid_frame.grid_columnconfigure(col, weight=0)

                elif titles[col] == 'confidence level':
                    entry1.append(tk.Entry(grid_frame, width=5))
                    entry1[x].configure(state='disabled', disabledforeground='black')
                    entry1[x].grid(column=col, row=row)
                    grid_frame.grid_columnconfigure(col, weight=0)

                else:
                    entry1.append(tk.Entry(grid_frame))
                    # print(titles[col])
                    if titles[col] == 'proposal due\n(yy/mm/dd)' or titles[col] == 'est completion\ndate':
                        entry1[x].configure(state='disabled', disabledforeground='black', width=14)
                    elif titles[col] == 'date created\n(yy/mm/dd)' or titles[col] == 'project age\n(days)':
                        entry1[x].configure(width=12)
                    elif titles[col] == 'project age\n-days' or titles[col] == 'awarded':  #
                        entry1[x].configure(state='disabled', disabledforeground='black', width=6)
                    elif titles[col] == 'sales rep' or titles[col] == 'PM' or titles[col] == 'title':
                        entry1[x].configure(state='disabled', disabledforeground='black', width=10)
                    entry1[x].grid(column=col, row=row)
                    grid_frame.grid_columnconfigure(col, weight=2)
                x += 1
            col += 1
        row += 1
    grid_frame.pack(fill="both")
    # print('lenth: ', len(entry1))
    # ----------------------
    label2 = Label(mainframe, text="Project X", bg="black", fg="white", padx=5, pady=5)
    label2.config(font=("Arial", 12))
    label2.pack(fill="x")
    # ---main loop------------------
    mainwindow.mainloop()  # display screen


def importFile(grid_frame1, label2, titles, entry1):  # WORKs
    # allow user to select .rec file----
    # print('select .xlsx file to open')
    # time.sleep(0.5)
    root = Tk()
    root.withdraw()
    loc = filedialog.askopenfilename(initialdir=os.curdir + '/data/')
    sheetnum = 0
    try:
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(sheetnum)
        numOfRows = int(sheet.cell_value(0, 1))  # 3
        row = 2  # start below titles (row3)
        x = 0
        custInfo = []
        while row <= (numOfRows + 1):
            col = 0
            while col < len(titles):
                try:
                    if titles[col] == 'contract' or titles[col] == 'pca' or titles[col] == 'awarded':
                        custInfo.append(sheet.cell_value(row, col))
                        if custInfo[x] == 1 and entry1[x].var.get() == 0:
                            entry1[x].toggle()
                        if custInfo[x] == 0 and entry1[x].var.get() == 1:
                            entry1[x].toggle()

                    elif titles[col] == 'project':
                        rowCheck = entry1[x].cget('text')
                        custInfo.append(sheet.cell_value(row, col))
                        entry1[x].var.set(custInfo[x])

                    elif titles[col] == 'status':
                        custInfo.append(sheet.cell_value(row, col))
                        entry1[x].var.set(custInfo[x])

                    else:
                        if titles[col] == 'proposal due\n(yy/mm/dd)' or titles[col] == 'sales rep' \
                                or titles[col] == 'PM' or titles[col] == 'est completion\ndate' or \
                                titles[col] == 'confidence level' or titles[col] == 'project age\n-days' \
                                or titles[col] == 'awarded':
                            entry1[x].configure(state='normal')
                        entry1[x].delete(0, END)
                        custInfo.append(sheet.cell_value(row, col))  # cell_value(row,col)
                        entry1[x].insert(0, custInfo[x])
                        if titles[col] == 'proposal due\n(yy/mm/dd)' or titles[col] == 'sales rep' \
                                or titles[col] == 'PM' or titles[col] == 'est completion\ndate' or \
                                titles[col] == 'confidence level' or titles[col] == 'project age\n-days' \
                                or titles[col] == 'awarded':
                            entry1[x].configure(state='disabled', disabledforeground='black')
                    col += 1
                    x += 1
                except:
                    print('addrow')
                    plusRow(grid_frame1, label2, titles, entry1)
            row += 1
    except:
        print('Error importing data selected')


def importClientInfo(grid_frame1, label2, titles, entry1):  # read in ws
    # print('importing xlsx info')
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    sheetnum = 0
    try:
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(sheetnum)
        numOfRows = int(sheet.cell_value(0, 1))
        row = 2  # start below titles (row3)
        x = 0
        custInfo = []
        while row <= (numOfRows + 1):
            col = 0
            while col < len(titles):
                try:
                    if titles[col] == 'contract' or titles[col] == 'pca':
                        custInfo.append(sheet.cell_value(row, col))
                        if custInfo[x] == 1 and entry1[x].var.get() == 0:
                            entry1[x].toggle()
                        if custInfo[x] == 0 and entry1[x].var.get() == 1:
                            entry1[x].toggle()
                    elif titles[col] == 'project':
                        rowCheck = entry1[x].cget('text')  # needed
                        custInfo.append(sheet.cell_value(row, col))
                        entry1[x].var.set(custInfo[x])
                    elif titles[col] == titles[col] == 'status':
                        custInfo.append(sheet.cell_value(row, col))
                        entry1[x].var.set(custInfo[x])
                    else:
                        # print('row ',row,' col ',col)
                        if titles[col] == 'proposal due\n(yy/mm/dd)' or titles[col] == 'sales rep' \
                                or titles[col] == 'PM' or titles[col] == 'est completion\ndate' or \
                                titles[col] == 'confidence level' or titles[col] == 'project age\n-days' \
                                or titles[col] == 'awarded' or titles[col] == 'title':
                            entry1[x].configure(state='normal')
                        entry1[x].delete(0, END)
                        custInfo.append(sheet.cell_value(row, col))  # cell_value(row,col)
                        entry1[x].insert(0, custInfo[x])
                        if titles[col] == 'proposal due\n(yy/mm/dd)' or titles[col] == 'sales rep' \
                                or titles[col] == 'PM' or titles[col] == 'est completion\ndate' or \
                                titles[col] == 'confidence level' or titles[col] == 'project age\n-days' \
                                or titles[col] == 'awarded' or titles[col] == 'title':
                            entry1[x].configure(state='disabled', disabledforeground='black')
                    col += 1
                    x += 1
                except:
                    print('addrow')
                    plusRow(grid_frame1, label2, titles, entry1)
            row += 1

    except:
        print('Error importing information menu1')


def exportClientInfo(titles, entry1):  # save main page
    # export grid info to xlsx doc
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    numOfRows = int(len(entry1) / len(titles))
    numOfCol1 = int(len(titles))
    numOfCol2 = 0
    numOfCol3 = 0
    numOfCol4 = 0
    numOfdataRows = numOfRows
    cells = []
    # try to read current ws-----------------------
    try:
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        numOfdataRows = int(sheet.cell_value(0, 1))  # 1
        numOfCol1 = int(sheet.cell_value(0, 2))  # 13
        numOfCol2 = int(sheet.cell_value(0, 3))  #
        numOfCol3 = int(sheet.cell_value(0, 4))  # 20
        numOfCol4 = int(sheet.cell_value(0, 5))  # 20
        datacols = numOfCol2 + numOfCol1 + numOfCol3 + numOfCol4
        for row in range(numOfdataRows + 2):
            cells.append([])
            for col in range(datacols):
                cells[row].append(sheet.cell_value(row, col))
        print('ready')
        # print(cells)
    except:
        print('no current wb2')
    # read main menu ----------------------------------
    if numOfRows > numOfdataRows:
        numOfdataRows = numOfRows
    newData = []
    x = 0
    row = 0
    while row < numOfRows + 1:
        col = 0
        newData.append([])
        # print('row:', row, ' col:', col)
        while col < numOfCol1:
            # print('row:', row, ' col:', col)
            if row == 0:
                newData[row].append(titles[col])
                # ws.write(row, col, titles[col])  # worksheet(row,col,stuff)
            else:
                # print('row:', row, ' col:', col, ' x ',x)
                # check button
                if titles[col] == 'contract' or titles[col] == 'pca':
                    newData[row].append(int(entry1[x].var.get()))
                # button
                elif titles[col] == 'project' or titles[col] == 'status':
                    newData[row].append(entry1[x].var.get())
                else:
                    # print('row:', row, ' col:', col)
                    newData[row].append(entry1[x].get())
                x += 1
            col += 1
        row += 1
    print('read main menu')
    # ----------------------------------------------------
    workbookName = loc
    wb = xlsxwriter.Workbook(workbookName)
    ws = wb.add_worksheet()
    row = 0  # write data
    # write old data-----------------------------------------
    if len(cells) > 0:
        for dataRow in cells:
            col = 0
            for oldData in dataRow:
                ws.write(row, col, oldData)
                col += 1
            row += 1
        print('wrote old data to ws')

    # write new data-------------------------
    ws.write(0, 0, 'key')
    ws.write(0, 1, numOfdataRows)  # 1
    ws.write(0, 2, numOfCol1)
    ws.write(0, 3, numOfCol2)
    ws.write(0, 4, numOfCol3)
    ws.write(0, 5, numOfCol4)
    row = 1
    for newDataRow in newData:
        col = 0
        for data in newDataRow:
            ws.write(row, col, data)
            col += 1
        row += 1
    wb.close()  # save info
    print('wrote new data to ws')


def exportToCloud(titles, entry1):
    curdir = os.getcwd()
    cred = curdir + '/a/creds.txt'
    exportClientInfo(titles, entry1)
    # Below code does the authentication
    # part of the code
    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile(cred)
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile(cred)
    drive = GoogleDrive(gauth)
    # save file to cloud
    path = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (path + '/data/' + workbookName)
    f = drive.CreateFile({'title': workbookName, 'parents': [{"id": 'your-id'}]})
    f.SetContentFile(loc)
    f.Upload()
    f = None
    print('done')


def importFromCloud(grid_frame1, label2, titles, entry1):
    curdir = os.getcwd()
    cred = curdir + '/a/creds.txt'
    # Below code does the authentication
    # part of the code
    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile(cred)
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile(cred)
    drive = GoogleDrive(gauth)
    file_list = drive.ListFile(
        {"q": "'your-id' in parents and trashed=false", 'maxResults': 10}).GetList()
    # print(file_list)
    # --name file--
    path = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (path + '/data/' + workbookName)
    # from file_list download [0] and save it to loc
    file_list[0].GetContentFile(loc)
    print('downloaded')
    # write to UI
    importClientInfo(grid_frame1, label2, titles, entry1)


def plusRow(grid_frame, label2, titles, entry1):  # WORKS

    label2.pack_forget()
    numOfRows = int((len(entry1) / len(titles)))  # 1
    x = len(entry1)  # 12
    row = numOfRows + 1  # 2
    if numOfRows > 10:  # limit rows
        row = 100
        numOfRows -= 1
    while row <= numOfRows + 1:
        col = 0
        while col < len(titles):
            if titles[col] == 'contract' or titles[col] == 'pca':
                v = IntVar()
                entry1.append(tk.Checkbutton(grid_frame, variable=v, width=5))
                entry1[x].configure(state='disabled', disabledforeground='black')
                entry1[x].var = v
                entry1[x].grid(column=col, row=row)
                grid_frame.grid_columnconfigure(col, weight=0)

            elif titles[col] == 'project':  # button
                butText = StringVar()
                entry1.append(tk.Button(grid_frame, textvariable=butText, width=5,
                                        command=lambda: popupMenu(butText, entry1, titles, grid_frame, label2)))
                butText.set('project #')
                entry1[x].var = butText
                entry1[x].grid(column=col, row=row, sticky=NSEW)
                grid_frame.grid_columnconfigure(col, weight=0)

            elif titles[col] == 'status':
                ddOptions = StringVar()
                ddOptions.set('-----')
                entry1.append(tk.OptionMenu(grid_frame, ddOptions, '---', 'design', 'bidding',
                                            'contracting', 'building', 'warranty', 'closed'))
                entry1[x].var = ddOptions
                entry1[x].grid(column=col, row=row)
                grid_frame.grid_columnconfigure(col, weight=0)

            elif titles[col] == 'confidence level':
                entry1.append(tk.Entry(grid_frame, width=5))
                entry1[x].configure(state='disabled', disabledforeground='black')
                entry1[x].grid(column=col, row=row)
                grid_frame.grid_columnconfigure(col, weight=0)

            else:
                entry1.append(tk.Entry(grid_frame))
                # print(titles[col])
                if titles[col] == 'proposal due\n(yy/mm/dd)' or titles[col] == 'est completion\ndate':
                    entry1[x].configure(state='disabled', disabledforeground='black', width=14)
                elif titles[col] == 'date created\n(yy/mm/dd)' or titles[col] == 'project age\n(days)':
                    entry1[x].configure(width=12)
                elif titles[col] == 'project age\n-days' or titles[col] == 'awarded':
                    entry1[x].configure(state='disabled', disabledforeground='black', width=6)
                elif titles[col] == 'sales rep' or titles[col] == 'PM' or titles[col] == 'title':
                    entry1[x].configure(state='disabled', disabledforeground='black', width=10)
                entry1[x].grid(column=col, row=row)
                grid_frame.grid_columnconfigure(col, weight=2)
            col += 1
            x += 1
        row += 1
    grid_frame.pack(fill="both")
    numOfRows += 1
    # print('lenth: ', len(entry1))
    label2.pack(fill='x')
    if numOfRows > 11:  # safety to kill errors
        quit()


def deleteRow(label2, titles, entry1):  # WORKS
    label2.pack_forget()
    x = 0
    numOfCol = int(len(entry1) / len(titles))
    lenEntry = len(entry1) - 1
    if (numOfCol > 0):
        while x < len(titles):
            entry1[lenEntry - x].destroy()
            entry1.pop(lenEntry - x)  # adjust entry1 size
            x += 1
    label2.pack(fill='x')
    # print(numOfCol-1)


# ====popmenu1========================================================================================
def popupMenu(butText, entry1, titles1, grid_frame, label2):
    # rows=int(len(entry1)/12)
    # print(butText)
    rowNum = 0
    x = 0
    row = 1
    # find row number of project
    while row <= len(entry1) / len(titles1):
        # print(str(entry1[x].var).split('R')[1])
        if str(entry1[x].var).split('R')[1] == str(butText).split('R')[1]:
            rowNum = row + 1  # counting from 0
            break
        row += 1
        x += len(titles1)
    # titles-----------------
    # contact info
    titles2 = ['name', 'address', 'phone #']
    titles3 = ['source', 'primary email', 'secondary email']
    # project info
    titles4 = ['title', 'project address', 'city', 'state', 'ZIP']
    titles5 = ['sales person', 'confidence level', 'project type', 'pca', 'contract']
    titles6 = ['proposal due', 'project manager', 'notes']
    titles7 = ['awarded', 'construction start date', 'est. completion date', 'days remaining']
    titles8 = ['contract price', 'change orders', 'deposits', 'payments', 'balance remaining']
    titlesX = titles1 + titles2 + titles3 + titles4 + titles5 + titles6 + titles7 + titles8
    popup = tk.Toplevel()
    popframe = Frame(popup, bg="black")
    popframe.pack(fill="both", expand=True)
    popup.title('Project Info')
    # --------------
    # hor_frame1 = Frame(popframe)
    # label = Label(popframe, text="Project Information X", bg="black", fg="white", padx=5, pady=5)
    # label.config(font=("Arial", 12))
    # label.pack(fill="x")

    hor_frame = Frame(popframe)
    b1 = Button(hor_frame, text="Submit",
                command=lambda: submit(rowNum, butText, titlesX, entry2, entry1, popup, grid_frame, label2))
    b1.grid(column=0, row=0, columnspan=2)
    b2 = Button(hor_frame, text="Close", command=popup.destroy)
    b2.grid(column=2, row=0, columnspan=2)
    hor_frame.pack(fill='x')
    # ---------
    label4 = Label(popframe, text="Contact Information", bg="black", fg="white", padx=5, pady=5)
    label4.config(font=("Arial", 12))
    label4.pack(fill="both")
    # customer info rows
    row = 0
    x = 0
    entry2 = []
    grid_frame1 = Frame(popframe)
    while row < 5:  # rows
        col = 0
        while col < 7:  # 6 columns
            if row == 0:
                if col == 0:  # label
                    lbl = Label(grid_frame1, text='Project #', bg="grey", fg="white", borderwidth=8)
                    lbl.grid(column=col, row=row, padx=5, pady=5)  # sticky="nsew") #padx, height, width
                    grid_frame1.grid_columnconfigure(col, weight=1)

                if col == 1:  # project #
                    entry2.append(tk.Entry(grid_frame1, borderwidth=5))
                    entry2[x].insert(0, butText.get())
                    entry2[x].grid(column=col, row=row)
                    grid_frame1.grid_columnconfigure(col, weight=1)
                    grid_frame1.grid_rowconfigure(row, minsize=20)
                    x += 1

            elif row == 1:
                lbl = Label(grid_frame1, text=titles2[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame1.grid_columnconfigure(col, weight=1)
                if col == len(titles2) - 1:
                    col = 10

            elif row == 2:
                entry2.append(tk.Entry(grid_frame1))
                entry2[x].grid(column=col, row=row)
                grid_frame1.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles2) - 1:
                    col = 10

            elif row == 3:
                lbl = Label(grid_frame1, text=titles3[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame1.grid_columnconfigure(col, weight=1)
                if col == len(titles3) - 1:
                    col = 10

            elif row == 4:
                entry2.append(tk.Entry(grid_frame1))
                entry2[x].grid(column=col, row=row)
                grid_frame1.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles3) - 1:
                    col = 10
            else:
                print('too many rows1: ', row)
            col += 1
        row += 1
    grid_frame1.pack(fill="both")

    label3 = Label(popframe, text="Project Information", bg="black", fg="white", padx=5, pady=5)
    label3.config(font=("Arial", 12))
    label3.pack(fill="both")
    # ---------------------
    # project info rows
    row = 0
    grid_frame2 = Frame(popframe)  # create frame
    while row < 10:  # rows
        col = 0
        while col < 7:

            if row == 0:
                lbl = Label(grid_frame2, text=titles4[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles4) - 1:
                    col = 10

            elif row == 1:
                entry2.append(tk.Entry(grid_frame2))
                entry2[x].grid(column=col, row=row)
                grid_frame2.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles4) - 1:
                    col = 10

            elif row == 2:
                lbl = Label(grid_frame2, text=titles5[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles5) - 1:
                    col = 10

            elif row == 3:
                # print('row ', row, ' col ', col)
                if titles5[col] == 'sales person':
                    # print('row ',row,' col ',col)
                    ddOptions = StringVar()
                    ddOptions.set('----')
                    entry2.append(tk.OptionMenu(grid_frame2, ddOptions, 'n/a', 'Alex K', 'Kate K', 'Jeremy M'))
                    entry2[x].var = ddOptions
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)

                elif titles5[col] == 'confidence level':
                    ddOptions = StringVar()
                    ddOptions.set('----')
                    entry2.append(tk.OptionMenu(grid_frame2, ddOptions,
                                                'n/a', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'))
                    entry2[x].var = ddOptions
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)

                elif titles5[col] == 'project type':
                    ddOptions = StringVar()
                    ddOptions.set('----')
                    entry2.append(
                        tk.OptionMenu(grid_frame2, ddOptions, 'new', 'remodel', 'addition', 'small', 'design'))
                    entry2[x].var = ddOptions
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)

                elif titles5[col] == 'pca' or titles5[col] == 'contract':
                    v = IntVar()
                    entry2.append(tk.Checkbutton(grid_frame2, variable=v))
                    entry2[x].var = v
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)

                else:
                    entry2.append(tk.Entry(grid_frame2))
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=2)

                x += 1
                if col == len(titles5) - 1:
                    col = 10

            elif row == 4:
                lbl = Label(grid_frame2, text=titles6[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles6) - 1:
                    col = 10

            elif row == 5:

                if titles6[col] == 'proposal due':
                    ddOptions = StringVar()
                    ddOptions.set('----')
                    entry2.append(
                        tk.OptionMenu(grid_frame2, ddOptions, '+5', '+10', '+15', '+30', '+45', '+60'))
                    entry2[x].var = ddOptions
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)

                elif titles6[col] == 'project manager':
                    ddOptions = StringVar()
                    ddOptions.set('----')
                    entry2.append(
                        tk.OptionMenu(grid_frame2, ddOptions, 'n/a', 'Alex K', 'Kate K', 'Jeremy M'))
                    entry2[x].var = ddOptions
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)

                elif titles6[col] == 'notes':
                    entry2.append(tk.Entry(grid_frame2))
                    entry2[x].grid(column=col, row=row, columnspan=3, sticky='nsew')
                    grid_frame2.grid_columnconfigure(col, weight=2)

                else:
                    entry2.append(tk.Entry(grid_frame2))
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles6) - 1:
                    col = 10

            elif row == 6:
                lbl = Label(grid_frame2, text=titles7[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles7) - 1:
                    col = 10

            elif row == 7:

                if titles7[col] == 'awarded':
                    ddOptions = StringVar()
                    ddOptions.set('----')
                    entry2.append(
                        tk.OptionMenu(grid_frame2, ddOptions, 'yes', 'no'))
                    entry2[x].var = ddOptions
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=0)
                else:
                    entry2.append(tk.Entry(grid_frame2))
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles7) - 1:
                    col = 10

            elif row == 8:
                lbl = Label(grid_frame2, text=titles8[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles8) - 1:
                    col = 10

            elif row == 9:
                if titles8[col] == 'change orders':
                    butText1 = StringVar()
                    entry2.append(tk.Button(grid_frame2, textvariable=butText1,
                                            command=lambda: popupMenu2(rowNum, entry2, entry1, titlesX)))
                    butText1.set('change order')
                    entry2[x].grid(column=col, row=row)
                    entry2[x].var = butText1
                    grid_frame2.grid_columnconfigure(col, weight=2)
                elif titles8[col] == 'payments':
                    butText2 = StringVar()
                    entry2.append(tk.Button(grid_frame2, textvariable=butText2,
                                            command=lambda: popupMenu3(rowNum, entry2, entry1, titlesX)))
                    butText2.set('payments')
                    entry2[x].grid(column=col, row=row)
                    entry2[x].var = butText2
                    grid_frame2.grid_columnconfigure(col, weight=2)
                else:
                    entry2.append(tk.Entry(grid_frame2))
                    entry2[x].grid(column=col, row=row)
                    grid_frame2.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles8) - 1:
                    col = 10

            else:
                print('too many rows', row)
            col += 1
        row += 1
    grid_frame2.pack(fill="both")
    # ----------------
    label2 = Label(popframe, text="Project X popup", bg="black", fg="white", padx=5, pady=5)
    label2.config(font=("Arial", 12))
    label2.pack(fill="both")
    # populate popup menu:
    importClientInfo2(rowNum, entry2, entry1, titlesX)


def dateCalc(year, month, day):
    # 31 day months
    if month == 1 or month == 3 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12:
        # print('31 days')
        if day > 31:
            month += 1
            day -= 31
    # 30 day months
    elif month == 4 or month == 6 or month == 9 or month == 11:
        # print('30 days')
        if day > 30:
            month += 1
            day -= 30
    # 29 day months
    else:
        # print('29 days')
        if day > 29:
            month += 1
            day -= 29
    if month > 12:
        year += 1
        month -= 12

    # print(day)
    # print(month)
    # print(year)

    dateX = str(year) + '/' + str(month) + '/' + str(day)
    return dateX


def prodDue(dc, pd1):
    # date created + production due
    pd = '---'
    try:
        dc = dc.split('/')
        day = int(dc[2])
        month = int(dc[1])
        year = int(dc[0])
        day = day + int(pd1[1:])
        pd = dateCalc(year, month, day)

    except:
        print('production due calc error')

    return pd


def projAge(dc):
    # print('calc age')
    # print(dc)
    age = '----'
    try:
        # todays date - date created
        today = str(date.today()).split('-')
        # print(dc)
        # print(today)
        year1 = int(today[0][2:])
        month1 = int(today[1])
        day1 = int(today[2])
        dc = dc.split('/')
        year2 = int(dc[0])
        month2 = int(dc[1])
        day2 = int(dc[2])
        year = year1 - year2
        month = month1 - month2
        day = day1 - day2
        if day < 0:
            month -= 1
            day += 30
        age = str(year * 365 + month * 30 + day).split('.')[0]
    except:
        print('calc age error')
    return age


def submit(rowNum, butText, titlesX, entry2, entry1, popup, grid_frame, label2):
    numOfCol3 = 0
    numOfCol4 = 0
    dc = ''
    pd1 = ''
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    cells = []
    titles = titlesX[:(len(titlesX) - len(entry2) + 1)]
    numOfRows = 0
    y = 0
    # read old data from wb---------------------------
    while y == 0:
        try:
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(0)
            numOfRows = int(sheet.cell_value(0, 1))  #
            numOfCol1 = int(sheet.cell_value(0, 2))  #
            numOfCol2 = int(sheet.cell_value(0, 3))  #
            numOfCol3 = int(sheet.cell_value(0, 4))
            numOfCol4 = int(sheet.cell_value(0, 5))
            cols = numOfCol2 + numOfCol1 + numOfCol3 + numOfCol4
            # write existing data to cells
            for row in range(numOfRows + 2):
                cells.append([])
                for col in range(cols):
                    cells[row].append(sheet.cell_value(row, col))
            y = 1
        except:
            print('no current wb1')
            exportClientInfo(titles, entry1)
    # ------------------------------------
    # rowNum counts from 0
    # NumOfRows counts from 1
    if (rowNum - 1) > numOfRows:  # logic to check which is longer
        numOfRows = rowNum - 1
    # read popup menu data---------------------------------------------------
    newData = []
    offset = len(titlesX) - len(entry2)
    for row in range(2):
        newData.append([])
        # print(row)
        for col in range(len(entry2)):
            if row == 0:
                # offset = len(titlesX) - len(entry2)
                if col == 0:
                    newData[row].append('project')
                else:
                    newData[row].append(titlesX[col + offset])
            else:
                # check box
                if newData[0][col] == 'pca' or newData[0][col] == 'contract':
                    newData[row].append(int(entry2[col].var.get()))
                # button
                elif newData[0][col] == 'change orders' or newData[0][col] == 'payments':
                    newData[row].append(entry2[col].var.get())
                # drop down
                elif newData[0][col] == 'sales person' or newData[0][col] == 'confidence level' \
                        or newData[0][col] == 'project type' or newData[0][col] == 'proposal due' \
                        or newData[0][col] == 'project manager' or newData[0][col] == 'awarded':
                    newData[row].append(entry2[col].var.get())
                # entry
                else:
                    newData[row].append(entry2[col].get())
    # -------------------------------------------------------
    # write old data
    workbookName = loc
    wb = xlsxwriter.Workbook(workbookName)
    ws = wb.add_worksheet()
    row = 0
    for dataRow in cells:
        col = 0
        for oldData in dataRow:
            ws.write(row, col, oldData)
            col += 1
        row += 1
    print('wrote old data to ws')
    # ----get menu1 variables----------------------------------------
    pca = newData[1][newData[0].index('pca')]
    cont = newData[1][newData[0].index('contract')]
    sp = newData[1][newData[0].index('sales person')]
    cl = newData[1][newData[0].index('confidence level')]
    pd1 = newData[1][newData[0].index('proposal due')]
    award = newData[1][newData[0].index('awarded')]
    pm = newData[1][newData[0].index('project manager')]
    title = newData[1][newData[0].index('title')]
    est = newData[1][newData[0].index('est. completion date')]
    butText.set(newData[1][0])  # change project id
    # ---------------------------------------------------
    try: dc = cells[rowNum][titlesX.index('date created\n(yy/mm/dd)')]
    except: print('no date created')
    # calculate proposal due
    pd = prodDue(dc, pd1)
    # calculate age of project
    age = projAge(dc)
    # write new data-------------------------------------------------------
    ws.write(0, 0, 'key')
    ws.write(0, 1, numOfRows)
    ws.write(0, 2, int(len(titlesX) - len(entry2) + 1))  #
    ws.write(0, 3, int(len(entry2) - 1))
    ws.write(0, 4, numOfCol3)
    ws.write(0, 5, numOfCol4)
    # write to view only menu1------------------------
    ws.write(rowNum, titlesX.index('pca'), pca)
    ws.write(rowNum, titlesX.index('title'), title)
    ws.write(rowNum, titlesX.index('confidence level'), cl)
    ws.write(rowNum, titlesX.index('sales rep'), sp)
    ws.write(rowNum, titlesX.index('project age\n-days'), age)
    ws.write(rowNum, titlesX.index('proposal due\n(yy/mm/dd)'), pd)
    ws.write(rowNum, titlesX.index('contract'), cont)
    ws.write(rowNum, titlesX.index('awarded'), award)
    ws.write(rowNum, titlesX.index('PM'), pm)
    ws.write(rowNum, titlesX.index('est completion\ndate'), est)
    # -----write new data--------------------------------------
    row = 0
    for dataRow in newData:
        col = offset
        for data in dataRow:
            if row == 0 and col > offset:
                ws.write(1, col, data)
            else:
                if col == offset:
                    ws.write(rowNum, 0, data)  # write project id
                else:
                    ws.write(rowNum, col, data)
            col += 1
        row += 1
    print('wrote new data')
    # --------------------
    wb.close()  # save info
    # popup.destroy()
    print('saved popmenu1')
    importClientInfo(grid_frame, label2, titles, entry1)
    importClientInfo2(rowNum, entry2, entry1, titlesX)


def importClientInfo2(rowNum, entry2, entry1, titlesX):
    # print('importing xlsx info')
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    # print(rowNum)
    try:
        cp=0
        coNum =0
        dep =0
        pay=0
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        # numOfRows = int(sheet.cell_value(0, 1))
        numofcol1 = int(sheet.cell_value(0, 2))
        numofcol2 = int(sheet.cell_value(0, 3))
        cols = numofcol1 + numofcol2
        row = rowNum  # start below titles (row3)
        x = 1
        custInfo = []
        while row == rowNum:  # only for specific row
            col = numofcol1  # start at second half
            while col < cols:
                # check box----------------------------
                if titlesX[col] == 'pca' or titlesX[col] == 'contract':
                    custInfo.append(sheet.cell_value(row, col))
                    if custInfo[x - 1] == 1 and entry2[x].var.get() == 0:
                        entry2[x].toggle()
                    if custInfo[x - 1] == 0 and entry2[x].var.get() == 1:
                        entry2[x].toggle()
                # button--------------------------------
                elif titlesX[col] == 'change orders':
                    custInfo.append(sheet.cell_value(row, col))
                    try: coNum = float(custInfo[x-1])
                    except: coNum = 0
                    entry2[x].var.set(custInfo[x - 1])
                elif titlesX[col] == 'payments':
                    custInfo.append(sheet.cell_value(row,col))
                    try: pay = float(custInfo[x-1])
                    except: pay = 0
                    entry2[x].var.set(custInfo[x-1])
                # drop down--------------------------------
                elif titlesX[col] == 'sales person' or titlesX[col] == 'confidence level' \
                        or titlesX[col] == 'project type' or titlesX[col] == 'proposal due' \
                        or titlesX[col] == 'project manager' or titlesX[col] == 'awarded':
                    custInfo.append(sheet.cell_value(row, col))
                    entry2[x].var.set(custInfo[x - 1])
                # entry boxes----------------------------
                else:
                    custInfo.append(sheet.cell_value(row, col))  # cell_value(row,col)
                    if titlesX[col] == 'construction start date':
                        csd = custInfo[x - 1]
                    elif titlesX[col] == 'est. completion date':
                        ecd = custInfo[x - 1]
                    elif titlesX[col] == 'contract price':
                        try: cp = float(custInfo[x-1])
                        except: cp = 0
                    elif titlesX[col] == 'deposits':
                        try: dep = float(custInfo[x-1])
                        except: dep = 0
                    #----------------------------
                    if titlesX[col] == 'days remaining':
                        # print('calc days remain')
                        try:
                            # print(csd)
                            # print(ecd)
                            # days remain = ecd - csd
                            # csd, ecd = (yy/mm/dd)
                            ecd = ecd.split('/')
                            day1 = int(ecd[2])
                            month1 = int(ecd[1])
                            year1 = int(ecd[0])
                            csd = csd.split('/')
                            day2 = int(csd[2])
                            month2 = int(csd[1])
                            year2 = int(csd[0])
                            day = day1 - day2
                            month = month1 - month2
                            year = year1 - year2
                            if day < 0:
                                month -= 1
                                day += 30
                            daysremain = year * 365 + month * 30 + day
                            entry2[x].delete(0, END)
                            entry2[x].insert(0, daysremain)
                        except:
                            print('error calc days remain')
                            entry2[x].delete(0, END)
                            entry2[x].insert(0, 0)

                    elif titlesX[col] == 'balance remaining':
                        # calc: contract price + change orders - deposits -payments
                        #cant int 13.00, can float or decimal
                        balRemain = round(cp + coNum - dep - pay,2)
                        entry2[x].delete(0, END)
                        entry2[x].insert(0, balRemain)
                    else:
                        entry2[x].delete(0, END)
                        entry2[x].insert(0, custInfo[x - 1])
                col += 1
                x += 1
            row += 1
    except: print('Error importing information popmenu1')


# ====change orders==================================================================================
def popupMenu2(rowNum, entry2, entry1, titlesX):
    # print('change orders')
    titles10 = ['change order #', 'description', 'amount $']
    popup2 = tk.Toplevel()
    popframe2 = Frame(popup2, bg="black")
    popframe2.pack(fill="both", expand=True)
    popup2.title('Change Orders Menu')
    # --------------
    # hor_frame1 = Frame(popframe2)
    label1 = Label(popframe2, text="Project X Change Orders", bg="black", fg="white", padx=5, pady=5)
    label1.config(font=("Arial", 12))
    label1.pack(fill="x")
    # -----
    hor_frame2 = Frame(popframe2)
    b1 = Button(hor_frame2, text="Submit",
                command=lambda: submit2(popup2, entry3, titles10, rowNum, entry2, entry1, titlesX))
    b1.grid(column=0, row=0, columnspan=2)
    b2 = Button(hor_frame2, text="Close", command=popup2.destroy)
    b2.grid(column=2, row=0, columnspan=2)
    b3 = Button(hor_frame2, text="Add Row", command=lambda: addRow2(label9, entry3, titles10, grid_frame2, popup2))
    b3.grid(column=4, row=0, columnspan=2)
    hor_frame2.pack(fill='x')
    row = 0
    x = 0
    entry3 = []
    grid_frame2 = Frame(popframe2)
    while row < 2:  # rows
        col = 0
        while col < 3:  # 6 columns
            if row == 0:
                lbl = Label(grid_frame2, text=titles10[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles10) - 1:
                    col = 10
            elif row == 1:
                entry3.append(tk.Entry(grid_frame2))
                entry3[x].grid(column=col, row=row)
                grid_frame2.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles10) - 1:
                    col = 10
            #else: print('too many rows2: ', row)
            col += 1
        row += 1
    grid_frame2.pack(fill="both")

    label9 = Label(popframe2, text="Project Change Orders", bg="black", fg="white", padx=5, pady=5)
    label9.config(font=("Arial", 12))
    label9.pack(fill="both")
    # import saved data----------------------------------
    importClientInfo3(rowNum, entry3, label9, titles10, grid_frame2, popup2)


def submit2(popup2, entry3, titles10, rowNum, entry2, entry1, titlesX):
    # read popup menu 2 data------------------------------------
    numOfDataRows = len(entry3) / len(titles10)
    numOfCol1 = 0
    numOfCol2 = 0
    rollup = 0
    data = []
    x = 0
    row = 0
    while row < numOfDataRows:
        data.append([])
        col = 0
        while col < len(titles10):
            # print(row, col)
            data[row].append(entry3[x].get())
            if col == (len(titles10) - 1):
                try:
                    rollup += int(entry3[x].get())
                except:
                    print('error: not a number')
            x += 1
            col += 1
        row += 1
    # print('data read')
    # print(data)
    # read ws data
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    cells = []
    # read old data----------------------------------------------
    try:
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        numOfRows = int(sheet.cell_value(0, 1))  #
        numOfCol1 = int(sheet.cell_value(0, 2))  #
        numOfCol2 = int(sheet.cell_value(0, 3))
        numOfCol3 = int(sheet.cell_value(0, 4))
        numOfCol4 = int(sheet.cell_value(0, 5))
        # numOfCol3 = int(sheet.cell_value(0, 4))
        cols = numOfCol2 + numOfCol1 + numOfCol3 + numOfCol4
        # write existing data to cells
        for row in range(numOfRows + 2):
            cells.append([])
            for col in range(cols):
                cells[row].append(sheet.cell_value(row, col))
    except:
        print('no current wb1')
    # write old data-----------------------------------------
    if len(cells) > 0:  # write old stuff
        workbookName = loc
        wb = xlsxwriter.Workbook(workbookName)
        ws = wb.add_worksheet()
        row = 0
        for dataRow in cells:
            col = 0
            for oldData in dataRow:
                ws.write(row, col, oldData)
                col += 1
            row += 1
        print('wrote old data')
        # write new data--------------------------------------------
        # cannot write popmenu2 if menu1 and popmenu1 doesnt exist
        endOfoldData = numOfCol1 + numOfCol2
        ws.write(0, 4, 20)  # 20 col buffer+++++++++++++++++++++++++++++++++++++++++++
        col = endOfoldData
        for dataRow in data:
            ws.write(rowNum, col, str(dataRow))
            col += 1
        while (col - endOfoldData) < 20:
            ws.write(rowNum, col, 'x')
            col += 1
        ws.write(rowNum, titlesX.index('change orders'), rollup)  # 37 fix++++++++++++++++++++++++++++++++++++++++++++++++
        print('wrote new data')
        # save wb
        wb.close()

        # update popmenu1
        importClientInfo2(rowNum, entry2, entry1, titlesX)
    # popup2.destroy()
    # end-------------


def addRow2(label9, entry3, titles10, grid_frame2, popup2):
    numOfRows = int((len(entry3) / len(titles10)))  # 1
    if numOfRows < 20:
        label9.pack_forget()
        x = len(entry3)  # 12
        row = numOfRows + 1  # 2
        # if numOfRows > 10: #limit rows
        #    row = 100
        #    numOfRows -= 1
        while row <= numOfRows + 1:
            col = 0
            while col < len(titles10):
                entry3.append(tk.Entry(grid_frame2))
                entry3[x].grid(column=col, row=row)
                grid_frame2.grid_columnconfigure(col, weight=2)
                col += 1
                x += 1
            row += 1
        grid_frame2.pack(fill="both")
        numOfRows += 1
        label9.pack(fill='x')
    # print('lenth: ', len(entry1))
    if numOfRows > 20:  # safety to kill errors
        popup2.destroy()


def importClientInfo3(rowNum, entry3, label9, titles10, grid_frame2, popup2):
    # print('importing xlsx info')
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    # print(rowNum)
    # try:
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    numOfRows = int(sheet.cell_value(0, 1))
    numofcol1 = int(sheet.cell_value(0, 2))
    numofcol2 = int(sheet.cell_value(0, 3))
    numofcol3 = int(sheet.cell_value(0, 4))
    #numofcol4 = int(sheet.cell_value(0, 5))
    offset = numofcol1 + numofcol2
    if rowNum > numOfRows + 2:
        # no data exists
        y = 0
    else:
        y = 1
    if y == 1:
        row = rowNum  # rowNum is definite
        x = 0
        COdata = []  # change order data
        col = offset  # start at offset
        while col < offset + numofcol3:
            # entry boxes
            COdata.append(sheet.cell_value(row, col))  # cell_value(row,col)
            # print(COdata)
            if COdata[int(x / 3)] == 'x' or COdata[int(x / 3)] == '':
                col = 200  # break out of loop
            else:
                rowAvailable = True
                while rowAvailable == True:
                    data = COdata[int(x / 3)].split(',')
                    # print(data)
                    try:
                        z = 0
                        for dat in data:
                            entry3[x + z].insert(0, dat.strip(" \'[]"))
                            z += 1
                        col += 1
                        x += 3
                        rowAvailable = False
                    except:
                        # print('addrow2')
                        addRow2(label9, entry3, titles10, grid_frame2, popup2)
            # col += 1
            # x += 3

    else:
        print('no menu3 data exists')
    # except: print('Error importing information menu3')


# ====payments==================================================================================
def popupMenu3(rowNum, entry2, entry1, titlesX):
    # print('change orders')
    titles10 = ['invoice #', 'date paid\n(yy/mm/dd)', 'amount $']
    popup3 = tk.Toplevel()
    popframe3 = Frame(popup3, bg="black")
    popframe3.pack(fill="both", expand=True)
    popup3.title('Payments Menu')
    # --------------
    # hor_frame1 = Frame(popframe2)
    label1 = Label(popframe3, text="Project X Payments", bg="black", fg="white", padx=5, pady=5)
    label1.config(font=("Arial", 12))
    label1.pack(fill="x")
    # -----
    hor_frame2 = Frame(popframe3)
    b1 = Button(hor_frame2, text="Submit",
                command=lambda: submit3(popup3, entry3, titles10, rowNum, entry2, entry1, titlesX))
    b1.grid(column=0, row=0, columnspan=2)
    b2 = Button(hor_frame2, text="Close", command=popup3.destroy)
    b2.grid(column=2, row=0, columnspan=2)
    b3 = Button(hor_frame2, text="Add Row", command=lambda: addRow3(label9, entry3, titles10, grid_frame2, popup3))
    b3.grid(column=4, row=0, columnspan=2)
    hor_frame2.pack(fill='x')
    row = 0
    x = 0
    entry3 = []
    grid_frame2 = Frame(popframe3)
    while row < 2:  # rows
        col = 0
        while col < 3:  # 6 columns
            if row == 0:
                lbl = Label(grid_frame2, text=titles10[col], bg="grey", fg="white")
                lbl.grid(column=col, row=row, padx=5, pady=5, sticky="nsew")
                grid_frame2.grid_columnconfigure(col, weight=1)
                if col == len(titles10) - 1:
                    col = 10
            elif row == 1:
                entry3.append(tk.Entry(grid_frame2))
                entry3[x].grid(column=col, row=row)
                grid_frame2.grid_columnconfigure(col, weight=2)
                x += 1
                if col == len(titles10) - 1:
                    col = 10
            #else: print('too many rows2: ', row)
            col += 1
        row += 1
    grid_frame2.pack(fill="both")

    label9 = Label(popframe3, text="Project Payments", bg="black", fg="white", padx=5, pady=5)
    label9.config(font=("Arial", 12))
    label9.pack(fill="both")
    # import saved data----------------------------------
    importClientInfo4(rowNum, entry3, label9, titles10, grid_frame2, popup3)


def submit3(popup2, entry3, titles10, rowNum, entry2, entry1, titlesX):
    # read popup menu 2 data------------------------------------
    numOfDataRows = len(entry3) / len(titles10)
    numOfCol1 = 0
    numOfCol2 = 0
    numOfCol3 = 0
    rollup = 0
    data = []
    x = 0
    row = 0
    while row < numOfDataRows:
        data.append([])
        col = 0
        while col < len(titles10):
            # print(row, col)
            data[row].append(entry3[x].get())
            if col == (len(titles10) - 1):
                try:
                    rollup += float(entry3[x].get())
                except:
                    print('error: not a number')
            x += 1
            col += 1
        row += 1
    # print('data read')
    # print(data)
    # read ws data
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    cells = []
    # read old data----------------------------------------------
    try:
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        numOfRows = int(sheet.cell_value(0, 1))
        numOfCol1 = int(sheet.cell_value(0, 2))
        numOfCol2 = int(sheet.cell_value(0, 3))
        numOfCol3 = int(sheet.cell_value(0, 4))
        numOfCol4 = int(sheet.cell_value(0, 5))
        cols = numOfCol2 + numOfCol1 + numOfCol3 + numOfCol4
        # write existing data to cells
        for row in range(numOfRows + 2):
            cells.append([])
            for col in range(cols):
                cells[row].append(sheet.cell_value(row, col))
    except:
        print('no current wb1')
    # write old data-----------------------------------------
    if len(cells) > 0:  # write old stuff
        workbookName = loc
        wb = xlsxwriter.Workbook(workbookName)
        ws = wb.add_worksheet()
        row = 0
        for dataRow in cells:
            col = 0
            for oldData in dataRow:
                ws.write(row, col, oldData)
                col += 1
            row += 1
        print('wrote old data')
        # write new data--------------------------------------------
        # cannot write popmenu2 if menu1 and popmenu1 doesnt exist
        endOfoldData = numOfCol1 + numOfCol2 + numOfCol3
        ws.write(0, 5, 20)  # 20 col buffer+++++++++++++++++++++++++++++++++++++++++++
        col = endOfoldData
        for dataRow in data:
            ws.write(rowNum, col, str(dataRow))
            col += 1

        while (col - endOfoldData) < 20:
            ws.write(rowNum, col, 'x')
            col += 1
        # print(rollup)
        ws.write(rowNum, titlesX.index('payments'), rollup)  # ++++++++++++++++++++++++++++++++++++++++++++++++
        print('wrote new data')
        # save wb
        wb.close()

        # update popmenu1
        importClientInfo2(rowNum, entry2, entry1, titlesX)
    # popup2.destroy()
    # end-------------


def addRow3(label9, entry3, titles10, grid_frame2, popup2):
    numOfRows = int((len(entry3) / len(titles10)))  # 1
    if numOfRows < 20:
        label9.pack_forget()
        x = len(entry3)  # 12
        row = numOfRows + 1  # 2
        # if numOfRows > 10: #limit rows
        #    row = 100
        #    numOfRows -= 1
        while row <= numOfRows + 1:
            col = 0
            while col < len(titles10):
                entry3.append(tk.Entry(grid_frame2))
                entry3[x].grid(column=col, row=row)
                grid_frame2.grid_columnconfigure(col, weight=2)
                col += 1
                x += 1
            row += 1
        grid_frame2.pack(fill="both")
        numOfRows += 1
        label9.pack(fill='x')
    # print('lenth: ', len(entry1))
    if numOfRows > 20:  # safety to kill errors
        popup2.destroy()


def importClientInfo4(rowNum, entry3, label9, titles10, grid_frame2, popup2):
    # print('importing xlsx info')
    curdir = os.getcwd()
    today = date.today()
    workbookName = str(today) + " project info.xlsx"
    loc = (curdir + '/data/' + workbookName)
    # print(rowNum)
    # try:
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    numOfRows = int(sheet.cell_value(0, 1))
    numofcol1 = int(sheet.cell_value(0, 2))
    numofcol2 = int(sheet.cell_value(0, 3))
    numofcol3 = int(sheet.cell_value(0, 4))
    numofcol4 = int(sheet.cell_value(0, 5))
    offset = numofcol1 + numofcol2 + numofcol3
    if rowNum > numOfRows + 2:
        # no data exists
        y = 0
    else:
        y = 1
    if y == 1:
        row = rowNum  # rowNum is definite
        x = 0
        paydata = []  # change order data
        col = offset  # start at offset
        while col < offset + numofcol4:
            # entry boxes
            paydata.append(sheet.cell_value(row, col))  # cell_value(row,col)
            # print(COdata)
            if paydata[int(x / 3)] == 'x' or paydata[int(x / 3)] == '':
                col = 200  # break out of loop
            else:
                rowAvailable = True
                while rowAvailable == True:
                    data = paydata[int(x / 3)].split(',')
                    # print(data)
                    try:
                        z = 0
                        for dat in data:
                            entry3[x + z].insert(0, dat.strip(" \'[]"))
                            z += 1
                        col += 1
                        x += 3
                        rowAvailable = False
                    except: addRow3(label9, entry3, titles10, grid_frame2, popup2)
            # col += 1
            # x += 3

    else: print('no popmenu3 data exists')


# ------------------------------------------------------------
# ============================================================
if __name__ == '__main__':
    main()

quit()
