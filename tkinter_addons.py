# -*- coding: utf-8 -*-
"""
Created on Wed May 14 16:14:58 2025

@author: Manip
"""

import tkinter as tk
from tkinter import ttk as ttk

class PlaceholderEntry(ttk.Entry):
    '''
    Custom modern Placeholder Entry box, takes positional argument master and placeholder along with\n
    textcolor(default being black) and placeholdercolor(default being grey).\n
    Use acquire() for getting output from entry widget\n
    Use shove() for inserting into entry widget\n
    Use remove() for deleting from entry widget\n
    Use length() for getting the length of text in the widget\n
    BUG 1: Possible bugs with binding to this class\n
    BUG 2: Anomalous behaviour with config or configure method
    '''
    def __init__(self, master, placeholder,textcolor='white',placeholdercolor='grey', **kwargs):
        self.text = placeholder
        self.__has_placeholder = False # placeholder flag
        self.placeholdercolor = placeholdercolor
        self.textcolor = textcolor

        # style for ttk widget
        self.s = ttk.Style()

        # init entry box
        ttk.Entry.__init__(self, master, style='my.TEntry', **kwargs)
        self.s.configure('my.TEntry',forground=self.placeholdercolor)

        # add placeholder if box empty
        self._add()

        # bindings of the widget
        self.bind('<FocusIn>', self._clear)
        self.bind('<FocusOut>', self._add)
        self.bind_all('<Key>', self._normal)
        self.bind_all('<Button-1>', self._cursor)

    def _clear(self, *args): # method to remove the placeholder
        if self.get() == self.text and self.__has_placeholder:  # remove placeholder when focus gain
            self.delete(0, tk.END)
            self.s.configure('my.TEntry', foreground='white',
                             font=(0, 0, 'normal'))
            self.__has_placeholder = False #set flag to false

    def _add(self, *args): # method to add placeholder
        if self.get() == '' and not self.__has_placeholder:  # if no text add placeholder
            self.s.configure('my.TEntry', foreground=self.placeholdercolor,
                             font=(0, 0, 'bold'))
            self.insert(0, self.text)  # insert placeholder
            self.icursor(0)  # move insertion cursor to start of entrybox
            self.__has_placeholder = True #set flag to true

    def _normal(self, *args): #method to set the text to normal properties
        self._add()  # if empty add placeholder
        if self.get() == self.text and self.__has_placeholder:  # clear the placeholder if starts typing
            self.bind('<Key>', self._clear)
            self.icursor(-1)  # keep insertion cursor to the end
        else:
            self.s.configure('my.TEntry', foreground=self.textcolor,
                         font=(0, 0, 'normal'))  # set normal font

    def acquire(self):  
        """Custom method to get the text"""
        if self.get() == self.text and self.__has_placeholder:
            return 'None'
        else:
            return self.get()

    def shove(self, index, string):  
        """Custom method to insert text into entry"""
        self._clear()
        self.insert(index, string)

    def remove(self, first, last):  
        """Custom method to remove text from entry"""
        if self.get() != self.text:
            self.delete(first, last)
            self._add()
        elif self.acquire() == self.text and not self.__has_placeholder:
            self.delete(first, last)
            self._add()

    def length(self):
        """Custom method to get the length of text in the entry widget"""
        if self.get() == self.text and self.__has_placeholder:
            return 0
        else:
            return len(self.get())

    def _cursor(self, *args):  # method to not allow user to move cursor when placeholder exists
        if self.get() == self.text and self.__has_placeholder:
            self.icursor(0)


class PrintLogger(tk.Text): # create file like object

    def __init__(self, textbox): # pass reference to text widget
        self.textbox = textbox # keep ref

    def write(self, text, *args, autoupdate=True):
        self.textbox.configure(state='normal') # set the textbox as writable
        self.textbox.insert(tk.END, text + " ") # write text to textbox
        for subtext in args:   
            self.textbox.insert(tk.END, str(subtext) + " ")
        self.textbox.insert(tk.END, "\n")
        self.textbox.configure(state='disabled') # set the textbox as read only
            # could also scroll to end of textbox here to make sure always visible
        if autoupdate:
            self.textbox.update()
        self.textbox.see(tk.END)
        
    def flush(self): # needed for file like object
        pass

