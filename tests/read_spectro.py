# -*- coding: utf-8 -*-
"""
Created on Wed May 14 11:53:27 2025

@author: Manip
"""

import u3
import time


try:
    d = u3.U3()

    while True:
        ainValue = d.getAIN(0)#Lecture de la valeur de la sortie du symphony. Quand elle sera à 0, cela voudra dire que la CCD est en train detre lue
        print(ainValue)
    
except:
    raise "Error"