# coding: utf-8
from __future__ import unicode_literals
from io import StringIO
from openpyxl import load_workbook, Workbook
from helper import XlsxHelper

import tkFileDialog
import tkMessageBox
import ttk
from Tkinter import *


class MainHandler(object):
    def __init__(self):
        self.helper = XlsxHelper()
        self.file_opt = {
            "defaultextension": ".xlsx",
            "filetypes": [('xlsx files', '.xlsx')],
            "parent": root
        }

    def update_table(self):
        sy = Scrollbar(root)
        sy.pack(side=RIGHT, fill=Y)
        tree = ttk.Treeview(root, columns=self.helper.treedata[0], show="headings", height=20)
        tree.configure(yscroll=sy.set)
        sy.config(command=tree.yview)

        for key in self.helper.treedata[0]:
            tree.column(key, width=10, anchor='center')
            tree.heading(key, text=key)

        tree.pack()
        for row in self.helper.treedata[1:]:
            tree.insert('', 1, values=row)

    def askopenfilename(self):
        if self.helper.treedata:
            tkMessageBox.showinfo("错误", '已打开过 需重启')
        else:
            filename = tkFileDialog.askopenfilename(**self.file_opt)
            if filename:
                try:
                    self.helper.openfile(filename)
                    self.helper.handle()
                except Exception, e:
                    tkMessageBox.showinfo("错误", unicode(e))
                if self.helper.treedata:
                    self.update_table()

    def asksavefilename(self):
        if not self.helper.treedata:
            tkMessageBox.showinfo("错误", '先打开需要处理的xlsx')
        else:
            filename = tkFileDialog.asksaveasfilename(**self.file_opt)
            if filename:
                self.helper.save(filename)

root = Tk()
root.geometry('500x300+400+300')
root.resizable(False, False)

handler = MainHandler()

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="打开xlsx", command=handler.askopenfilename)
filemenu.add_command(label="导出xlsx", command=handler.asksavefilename)
menubar.add_cascade(label="文件", menu=filemenu)
root.config(menu=menubar)

root.mainloop()
