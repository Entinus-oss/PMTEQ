# -*- coding: utf-8 -*-
"""
Created on Thu May 15 11:10:49 2025

@author: Manip
"""
import tkinter as tk

root = tk.Tk()

# specify size of window.
root.geometry("250x170")

console_log = tk.Text(root, bg='black', fg='green2', height=1)

console_log.configure(state='normal') # set the textbox as writable
console_log.insert(tk.END, "Welcome to the program allowing you to take IV curve and spectral datas. Please follow the following steps carefully :")
console_log.configure(state='disabled') # set the textbox as writable

console_log.pack()
tk.mainloop()