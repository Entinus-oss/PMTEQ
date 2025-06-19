# -*- coding: utf-8 -*-
"""
Created on Thu Apr 11 16:36:08 2024

@author: manip pico
"""



# Import Required Module
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
 
# Create Object
root = ThemedTk(theme='equilux')
root.configure(background="#464646") #'#1A1110')
i=True
def change_style():
    global i
    if i:
        style.configure('W.TButton', font=('calibri', 10), background="red", foreground="yellow")
        i=False
    else:
        style.configure('W.TButton', font=('calibri', 10), background="black", foreground="white")
        i=True

# Set geometry (widthxheight)
root.geometry('1000x1000')
 
# This will create style object
style = ttk.Style()
 
# This will be adding style, and 
# naming that style variable as 
# W.Tbutton (TButton is used for ttk.Button).
style.configure('W.TButton', font =
               ('calibri', 10, 'bold', 'underline'),
                foreground = 'red')

# This will create style object
style2 = ttk.Style()
 
# This will be adding style, and 
# naming that style variable as 
# W.Tbutton (TButton is used for ttk.Button).
style2.configure('TButton', font =
               ('calibri', 10, 'bold', 'underline'),
               background='#1A1110',
               foreground = 'white')
 
# Changes will be reflected
# by the movement of mouse.
style2.map('TButton', foreground = [('active', '!disabled', 'green')],
                     background = [('active', 'blue')])
 
# Style will be reflected only on 
# this button because we are providing
# style only on this Button.
''' Button 1'''
btn1 = ttk.Button(root, text = 'Quit !',
                  style='W.TButton',
                  command = change_style)
btn1.grid(row = 0, column = 3, padx = 100)

Entry=ttk.Entry(root, width = 5)
Entry.grid(row=2,column=1)
Label=ttk.Label(root,text='abcdefg')
Label.grid(row=3,column=1)
 
''' Button 2'''
 
btn2 = ttk.Button(root, text = 'Click me !', style='TButton', command = None)
btn2.grid(row = 1, column = 3, pady = 10, padx = 100)
 
# Execute Tkinter
root.mainloop()
