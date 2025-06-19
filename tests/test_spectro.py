# -*- coding: utf-8 -*-
"""
Created on Wed May 14 11:31:36 2025

@author: Manip
"""

import u3
import time


try:
    d = u3.U3()
    
    tic = time.time()
    #Trigger IN sends a 100ms pulse 
    DAC0_VALUE = d.voltageToDACBits(4.5, dacNumber = 0, is16Bits = False)
    d.getFeedback(u3.DAC0_8(DAC0_VALUE))
    time.sleep(0.1)
    DAC0_VALUE = d.voltageToDACBits(0, dacNumber = 0, is16Bits = False)
    d.getFeedback(u3.DAC0_8(DAC0_VALUE))
    
    # while True:
    #     TLLOutput2 = d.getAIN(0)#Lecture de la valeur de la sortie du symphony. Quand elle sera à 0, cela voudra dire que la CCD est en train detre lue
    #     # print(TLLOutput2)
    #     time.sleep(0.01)
    #     tac = time.time()
    #     print(TLLOutput2, tac-tic)
    #     if tac-tic > 3:
    #         break
        
except:
    raise "Error"