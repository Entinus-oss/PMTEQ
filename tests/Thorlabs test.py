# -*- coding: utf-8 -*-
"""
Created on Wed Apr 17 14:14:56 2024

@author: manip pico
"""
import tkinter as tk
from tkinter import ttk, filedialog
from ttkthemes import ThemedTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.widgets
import numpy as np
import pylablib as pll
import sys
import asyncio
import time
import pyvisa as visa
import csv
from os.path import exists
import pandas as pd
from ctypes import cdll,c_long, c_ulong, c_uint32,byref,create_string_buffer,c_bool,c_char_p,c_int,c_int16,c_double, sizeof, c_voidp
from TLPM import TLPM #comes from the TLPM file in the same directory !! move it if you change this program's directory !!! There was an issue with the TLPM_64.dll not loading correctly I had to specify its adress but couldnt do it in the TLPM file in Programs x86


import os
from datetime import date

from pylablib.devices import Andor
from pylablib.devices import Thorlabs

# test1=Thorlabs.kinesis.KinesisPiezoMotor('97101451')

LambdaOverTwo=Thorlabs.kinesis.KinesisMotor('83835088',is_rack_system=True)
L_Pos=LambdaOverTwo.get_position()
# LambdaOverTwo.move_to(0)
# time.sleep(10/200000*L_Pos+5)
LambdaOverTwo.setup_velocity(acceleration=100000, max_velocity=100000, scale=True)
tlpm=TLPM()
resourceName = create_string_buffer(b'USB0::0x1313::0x8072::P2003849::INSTR')
tlpm.open(resourceName, c_bool(True), c_bool(True))
power=c_double()

def get_min_max_powers():
    global min_power
    global max_power
    global Liste_powers
    global counter_powers
    global power

    L_Pos=LambdaOverTwo.get_position()
    print(0.25/10000*L_Pos+1.5)
    LambdaOverTwo.move_to(0)
    time.sleep(0.25/10000*L_Pos+1.5) #time for the thing to get to 0 otherwise it will add the commands to the previous one
    print(LambdaOverTwo.get_velocity_parameters(scale=True))
    tlpm.measPower(byref(power))
    Power=power.value
    L_Pos=LambdaOverTwo.get_position()
    Liste_pos=np.array([L_Pos])
    Liste_Powers=np.array([Power])
    counter_powers=0
    while counter_powers<10:
        LambdaOverTwo.move_by(50000)
        time.sleep(2.5)
        tlpm.measPower(byref(power))
        Power=power.value
        L_Pos=LambdaOverTwo.get_position()
        Liste_pos = np.append(Liste_pos,L_Pos)
        Liste_Powers=np.append(Liste_Powers,Power)
        counter_powers+=1
        
    min_power=np.min(Liste_Powers)
    max_power=np.max(Liste_Powers)
    plt.figure(1)
    plt.plot(Liste_pos)
    # plt.figure(2)
    # plt.plot(Liste_Powers)
    
get_min_max_powers()
    
tlpm.close()
LambdaOverTwo.close()