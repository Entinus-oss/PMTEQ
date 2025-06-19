# -*- coding: utf-8 -*-
"""
Created on Wed May 14 11:55:23 2025

@author: Manip
"""


import u3
import time

n = 4.5

try:
    d = u3.U3()
    
    DAC0_VALUE = d.voltageToDACBits(n, dacNumber = 0, is16Bits = False)
    d.getFeedback(u3.DAC0_8(DAC0_VALUE))
    ainValue = d.getAIN(0)#Lecture de la valeur de la sortie du symphony. Quand elle sera à 0, cela voudra dire que la CCD est en train detre lue
    print(ainValue)
    
except:
    raise "Error"
    
