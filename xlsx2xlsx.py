# coding: utf-8
from __future__ import unicode_literals
from io import StringIO
from openpyxl import load_workbook, Workbook

import tkFileDialog
import ttk
from Tkinter import *

root = Tk()

def openfile(filename):
    wb = load_workbook(filename=filename, read_only=True)
    ws = wb.worksheets[0]
    temp = StringIO()

    for row in ws.rows:
        if row[0].value:
            temp.writelines(str(row[0].value) + '\n')
    return temp

def handle(fd):
    treedata = []
    fd.seek(0)
    
    title = False
    
    line = fd.readline().split(',')
    while True:
        key, value, sched = [], [], []
        if line[0] == 'RELEASE HEADER SECTION\n':
            key.extend(fd.readline().split(','))
            value.extend(fd.readline().split(','))

            line = fd.readline().split(',')
            key.extend(fd.readline().split(','))
            value.extend(fd.readline().split(','))

            line = fd.readline().split(',')
            key.extend(fd.readline().split(','))
            value.extend(fd.readline().split(','))

            line = fd.readline().split(',')
            key.extend(fd.readline().split(','))
            while True:
                line = fd.readline().split(',')
                if not line[0] or line[0] == 'RELEASE HEADER SECTION\n':
                    break
                sched.append(line)

        if not title:
            treedata.append(key)
            title = True
        for i in range(len(sched)):
            row = []
            row.extend(value)
            row.extend(sched[i])
            if i > 0:
                row.append('DITTO')
            treedata.append(row)

        if not line[0]:
            break
    return treedata

def output(filename, treedata):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    for row in treedata:
        ws.append(row)
    wb.save(out_filename)

def update_table(treedata):
    tree = ttk.Treeview(root, columns=treedata[0], show="headings")

    for key in treedata[0]:
        tree.column(key, width=100, anchor='center')
        tree.heading(key, text=key)

    tree.pack()
    for row in treedata[1:]:
        tree.insert('', 1, values=row)

file_opt = {
    "defaultextension": ".xlsx",
    "filetypes": [('xlsx files', '.xlsx')],
    "parent": root
}

def askopenfilename():
    filename = tkFileDialog.askopenfilename(**file_opt) 
    if filename:
        fd = openfile(filename)
        treedata = handle(fd)
        update_table(treedata)

menubar = Menu(root)
menubar.add_command(label="打开xlsx", command=askopenfilename)
menubar.add_command(label="导出xlsx", command=update_table)

root.config(menu=menubar)

root.mainloop()
