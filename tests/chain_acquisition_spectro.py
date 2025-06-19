# -*- coding: utf-8 -*-
"""
Created on Wed May 14 11:55:23 2025

@author: Manip
"""


import u3
import time

active_high_threshold = 4.5 # V
unactive_high_threshold = 0.1 # V

acquisition_number = 20

try:
    d = u3.U3()
    
    time.sleep(1)
    
    for i in range(acquisition_number):
        
        #Trigg IN pulse
        DAC0_VALUE = d.voltageToDACBits(4.5, dacNumber = 0, is16Bits = False)
        d.getFeedback(u3.DAC0_8(DAC0_VALUE))
        
        tic = time.time()
        
        time.sleep(0.1)
        DAC0_VALUE = d.voltageToDACBits(0, dacNumber = 0, is16Bits = False)
        d.getFeedback(u3.DAC0_8(DAC0_VALUE))
        
        #Read TLLOutput2
        is_exposing = True
        is_readingout = False
        print("Exposure State")
        
        while is_exposing or is_readingout:
            TLLOutput2 = d.getAIN(0) # Reads TLLOutput2 state (active = 5V, unactive = 0V)
            # print(TLLOutput2)
            time.sleep(0.1) # Reads TLLOutput2 every 1 ms
            
            if (TLLOutput2 > active_high_threshold) and (is_readingout == False): # Enters reading out state
                is_readingout = True
                is_exposing = False
                tic_exp = time.time()
                print("Exposure Time :", tic_exp-tic)
                print("Reading Out State")
            elif (TLLOutput2 < unactive_high_threshold) and (is_readingout == True): # Leaves reading out state
                is_readingout = False
                print("Leaving Reading Out State")
        
        time.sleep(0.1)
        # Now the CCD is neither is the exposure or the reading out state, we can Trigg IN again
    
except:
    raise "Error"
    
