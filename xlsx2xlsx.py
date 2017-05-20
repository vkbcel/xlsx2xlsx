# coding: utf-8
from __future__ import unicode_literals

import tkFileDialog
import tkMessageBox
import ttk
from Tkinter import *

from helper import XlsxHelper

class MainHandler(object):
    def __init__(self):
        self.helper = XlsxHelper()
        self.file_opt = {
            "defaultextension": ".xlsx",
            "filetypes": [('xlsx files', '.xlsx')],
            "parent": root
        }

    def update_table(self):
        treedata = self.helper.as_table()
        
        tree = ttk.Treeview(root, columns=treedata[0], show="headings", height=20)
        ysb = Scrollbar(root, orient='vertical', command=tree.yview)
        ysb.pack(side=RIGHT, fill=Y)
        xsb = Scrollbar(root, orient='horizontal', command=tree.xview)
        xsb.pack(side=BOTTOM, fill=X)
        tree.configure(yscroll=ysb.set, xscroll=xsb.set)

        for key in treedata[0]:
            tree.column(key, width=100, anchor='center')
            tree.heading(key, text=key)

        tree.pack()
        for row in treedata[1:]:
            tree.insert('', 100, values=row)

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
                tkMessageBox.showinfo("成功", '导出成功')
                root.quit()

root = Tk()
root.geometry('800x300+200+100')
root.resizable(True, False)

handler = MainHandler()

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="打开xlsx", command=handler.askopenfilename)
filemenu.add_command(label="导出xlsx", command=handler.asksavefilename)
menubar.add_cascade(label="文件", menu=filemenu)
root.config(menu=menubar)

root.mainloop()
