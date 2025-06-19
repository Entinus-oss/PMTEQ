# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 15:39:00 2024
https://pylablib.readthedocs.io/en/latest/.apidoc/pylablib.devices.Andor.html

@author: manip pico
"""
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import scrolledtext
from ttkthemes import ThemedTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.widgets
import numpy as np
import sys
import time
import pyvisa as visa
import csv
from os.path import exists
import pandas as pd
from ThorlabsPM100 import ThorlabsPM100

import os
import math
from datetime import date

from pylablib.devices import Thorlabs

import u3

from tkinter_addons import PlaceholderEntry, PrintLogger

plt.style.use('dark_background')

#%%Instruments Parameters
KEITHLEY_NAME = "GPIB0::26::INSTR"
KEITHLEY_CHANNEL = 'smua' # Channel A
KEITHLEY_COMPLIANCE = '500E-6'
KEITHLEY_LOG_START = "Keithley " + KEITHLEY_NAME + " CH " + KEITHLEY_CHANNEL[-1] + " :"

PM100_NAME = "USB0::0x1313::0x8072::P2008402::INSTR"
PM100_NAME_BYTES = b"USB0::0x1313::0x8072::P2008402::INSTR"
PM100_LOG_START = "PM100 " + PM100_NAME + "  :"

ACTIVE_HIGH_THRESHOLD = 4.5 # V
UNACTIVE_HIGH_THRESHOLD = 0.1 # V

#%%Time

today_date=date.today()
today_date_str=today_date.strftime("%Y%m%d")

if os.path.exists("C:/Users/Manip/Documents/"+today_date_str)==False:
    os.makedirs("C:/Users/Manip/Documents/"+today_date_str)

#%%Functions
def browse_directory():
    global directory
    directory=filedialog.askdirectory()
    Directory_entry.delete(0, tk.END)
    Directory_entry.insert(0,directory)
    

#%%Keithley related

def Turn_Keithley_Link_On():
    global Keithley
    try:
        rm = visa.ResourceManager()
        Keithley=rm.open_resource(KEITHLEY_NAME)
    
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv=0") #sets voltage to 0
        Keithley.write(KEITHLEY_CHANNEL+".source.output="+ KEITHLEY_CHANNEL+".OUTPUT_ON") #sets ouput to on
        Keithley.write(KEITHLEY_CHANNEL+".source.limiti="+KEITHLEY_COMPLIANCE)
        
        console.write(KEITHLEY_LOG_START, "Connected  Output ON")
    except:
        console.write(KEITHLEY_LOG_START, "Is not powered on.")
    
def stop_Keithley():
    try:
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv=0") #sets voltage to 0
        Keithley.write(KEITHLEY_CHANNEL+".source.output="+ KEITHLEY_CHANNEL+".OUTPUT_OFF") #sets ouput to on  # Turns off output
        Keithley.close()
        console.write(KEITHLEY_LOG_START, "0 V  Output OFF  Closed")
    except:
        console.write(KEITHLEY_LOG_START, "Could not close properly.")
    
def Set_Keithley_tension():
    try:
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv="+str(V_entry.get())) #sets voltage to V_entry
        
        Keithley.write('print('+KEITHLEY_CHANNEL+".measure.i())") #Measure current
        Q = Keithley.read()  #Measurement is a table, first is voltage, second current
        cur_amp=float((Q.replace('\n','')))
        I_read_label.config(text='{:.3f} uA'.format(cur_amp*1e6)) #showing the current flowing through
        
        Keithley.write('print('+KEITHLEY_CHANNEL+".measure.v())") #Measure current
        Q = Keithley.read()  #Measurement is a table, first is voltage, second current
        cur_volt=float((Q.replace('\n','')))
        V_read_label.config(text='{:.3f} V'.format(cur_volt))
        
        console.write(KEITHLEY_LOG_START, str(V_entry.get()), "V  ", str(cur_amp), "A")
    except:
        console.write(KEITHLEY_LOG_START, "Has not been initialized, please power the device then click 'Turn Keithley ON'.")

def set_tension(v):
    try:
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv="+str(v)) #sets voltage to V_entry
        Keithley.write('print('+KEITHLEY_CHANNEL+".measure.i())") #Measure current
        Q = Keithley.read()  #Measurement is a table, first is voltage, second current
        cur_amp=float((Q.replace('\n','')))
    
        Keithley.write('print('+KEITHLEY_CHANNEL+".measure.v())") #Measure current
        Q = Keithley.read()  #Measurement is a table, first is voltage, second current
        cur_volt=float((Q.replace('\n','')))
        return cur_volt, cur_amp
    except:
        console.write("Could not set tension. Did you turned on the Keithley beforehand ?")
    
def Current_evolution():
    try:
        #function to measure current over time for capacitive effects
        console.write("Starting I(t)...")
        style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
        directory=str(Directory_entry.get())
        Start_voltage=float(V_entry.get())
        if os.path.exists(directory+"/Current evolution")==False:
            os.makedirs(directory+"/Current evolution")
        if os.path.exists(directory+"/Current evolution plots")==False:
            os.makedirs(directory+"/Current evolution plots")
        filename=str(IV_entry.get())
        filename=directory+'/Current evolution/'+filename
        filename_new=filename
        
        res_amp = []  # Vector that will contain current
        res_time = [] # Vector that will contain time
        
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv=0") #sets voltage to 0
        Keithley.write(KEITHLEY_CHANNEL+".source.output="+ KEITHLEY_CHANNEL+".OUTPUT_ON") #sets ouput to on
        time.sleep(1)
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv="+str(Start_voltage)) #sets voltage to 0
        time.sleep(1)
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv=0") #sets voltage to 0
        tic = time.time()
        for k in range(1000):
            console.write(str(k+1)+"/1000")
            time.sleep(0.1)
            toc = time.time()
            Keithley.write('print('+KEITHLEY_CHANNEL+".measure.i())") #Measure current
            Q = Keithley.read()  # La mesure est un tableau, la premiere case est la tension, la seconde le courant
            A=float((Q.replace('\n','')))
            cur_amp=A
            res_amp.append(cur_amp)
            res_time.append(toc-tic)
        tac = time.time()
        console.write("I(t) done in", str(tac-tic), "s")
        console.write("Creating csv file...")
        df = pd.DataFrame(list(zip(res_time, res_amp)), columns=['Time (s)', 'Intensity (A)'])
    
        n = 1
        while os.path.exists(filename_new + '.csv') == True:  # Boucle pour ne pas effacer de donnees
            filename_new = filename + '_' + str(n)
            n += 1
        file_name = filename_new + '.csv'
        df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
        console.write("csv file created at", directory+'/Current evolution/')
        style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
    except:
        console.write(KEITHLEY_LOG_START, "Has not been initialized, please power the device then click 'Turn Keithley ON'.")
    
def Tension_series():
    try:
        console.write("Starting I(V)...")
        style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
        directory=str(Directory_entry.get())
        if os.path.exists(directory+"/I-V")==False:
            os.makedirs(directory+"/I-V")
        if os.path.exists(directory+"/I-V plots")==False:
            os.makedirs(directory+"/I-V plots")
        filename=str(IV_entry.get())
        filename=directory+'/I-V/'+filename
        filename_new=filename
        
        volt = np.arange(float(V_start_entry.get()), float(V_finish_entry.get())+float(Step_entry.get()), float(Step_entry.get()))  # Definition de toutes les tensions a explorer
        res_volt = []  # Vecteur qui contiendra la tension appliquee
        res_amp = []  # Vecteur qui contiendra le courant mesure
    
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv=0") #sets voltage to 0
        Keithley.write(KEITHLEY_CHANNEL+".source.output="+ KEITHLEY_CHANNEL+".OUTPUT_ON") #sets ouput to on
        tic = time.time()
        for i in range(len(volt)):  # Boucle pour les mesures
            console.write(str(i+1)+"/"+str(len(volt)))
            Keithley.write(KEITHLEY_CHANNEL+".source.levelv=" + str(volt[i]))  # Application de la tension
            time.sleep(0.05)  # Waiting for a permanent regime
            
            Keithley.write('print('+KEITHLEY_CHANNEL+".measure.i())") #Measure current
            Q = Keithley.read()  # La mesure est un tableau, la premiere case est la tension, la seconde le courant
            cur_amp=float((Q.replace('\n','')))
            res_amp.append(cur_amp)
            
            Keithley.write('print('+KEITHLEY_CHANNEL+".measure.v())") #Measure Voltage
            Q=Keithley.read() #La mesure est un tableau, la premiere case est la tension, la seconde le courant
            cur_volt=float((Q.replace('\n','')))
            res_volt.append(cur_volt)
        tac = time.time()
        console.write("I(V) done in", str(tac-tic), "s")
        console.write("Creating csv file...")
        Keithley.write(KEITHLEY_CHANNEL+".source.levelv=0") #sets voltage to 0
        
        df = pd.DataFrame(list(zip(res_volt, res_amp)), columns=['Tension (V)', 'Intensity (A)'])
    
        n = 1
        while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
            filename_new = filename + '_' + str(n)
            n += 1
        file_name = filename_new + '.csv'
        df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
        print("csv file created at", directory+'/I-V/')
        style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
    except:
        print(KEITHLEY_LOG_START, "Has not been initialized, please power the device then click 'Turn Keithley ON'.")
    
#%%PM100 related

def start_PM():
    global PM100
    global PM_refresh
    console.write("Starting PM100...")
    rm = visa.ResourceManager()
    inst=rm.open_resource(PM100_NAME)
    PM100 = ThorlabsPM100(inst=inst)
    console.write(PM100_LOG_START, "Started successfully")
    PM_refresh=int(PM_exp_entry.get())
    console.write(PM100_LOG_START, "Refresh rate set to", str(PM_refresh))
    Power_button.config(text= 'set PM refresh', command=set_PM_refresh)
    root.after(200,measure_power)
    
def stop_PM100():
    try :
        PM100.abort()
        console.write(PM100_LOG_START, "Aborted.")
    except:
        console.write(PM100_LOG_START, "Could not close properly.")
    
def measure_power():
    global PM100
    global Power_value
    global PM_refresh
    Power_value.config(text='{:.0f} uW'.format(PM100.read*1e6))
    root.after(PM_refresh,measure_power)
    
def set_PM_refresh():
    global PM_refresh
    PM_refresh=int(PM_exp_entry.get())
    console.write(PM100_LOG_START, "Refresh rate set to", str(PM_refresh))
    
#%%Motors for Power and Polar measurements

def start_Power_Control():
    global LambdaOverTwo
    LambdaOverTwo = Thorlabs.kinesis.KinesisMotor('83828438',is_rack_system=True) #83835088 is the one with the glass inside
    LambdaOverTwo.setup_velocity(acceleration=100000, max_velocity=100000, scale=True)
    L_Pos=LambdaOverTwo.get_position()
    LambdaOverTwo.move_to(0)
    time.sleep(12/200000*np.abs(L_Pos)+5)
    LambdaOverTwo.set_position_reference()
    Get_min_max_powers_button.config(text="Get min/max powers", command=get_min_max_powers)
    Start_Power_Series_button.config(command=Power_series)
    
def start_Polar_Control():
    global LambdaOverTwoPolar
    LambdaOverTwoPolar = Thorlabs.kinesis.KinesisMotor('83835088',is_rack_system=True) #83828438 is the one with the glass screwed on top
    LambdaOverTwoPolar.setup_velocity(acceleration=100000, max_velocity=100000, scale=True)
    L_Pos=LambdaOverTwoPolar.get_position()
    LambdaOverTwoPolar.move_to(0)
    time.sleep(12/200000*np.abs(L_Pos)+5)
    LambdaOverTwoPolar.set_position_reference()
    Start_Polar_Control_button.config(text="Go to angle", command=go_to_angle)
    Start_Polar_Series_button.config(command=Polar_Series)
    
def get_min_max_powers():
    global min_power
    global max_power
    global Liste_Powers
    global counter_powers
    
    style_buttons.configure('TButton', font=font, background="#464646", foreground="red")
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    
    def loading():
        global counter_powers
        global Liste_Powers
        global min_power
        global max_power
        
        if counter_powers<20:
            tlpm.measPower(byref(power))
            Power=power.value
            Liste_Powers=np.append(Liste_Powers,Power)
            LambdaOverTwo.move_by(10000)
            counter_powers+=1
            root.after(1500,loading)
        else:
            tlpm.measPower(byref(power))
            Power=power.value
            Liste_Powers=np.append(Liste_Powers,Power)
            min_power=np.min(Liste_Powers)
            max_power=np.max(Liste_Powers)
            min_power_label.config(text="Min Power = {:.0f}uW".format(min_power*1e6))
            max_power_label.config(text="Max Power = {:.0f}uW".format(max_power*1e6))
            if cam_started==False:
                style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
            style_buttons.configure('TButton', font=font, background="#464646", foreground="#a6a6a6")
            Go_To_Power_button.config(command=Go_To_Power)
            
    Liste_Powers=np.array([])
    L_Pos=LambdaOverTwo.get_position()
    LambdaOverTwo.move_to(0)
    time.sleep(0.25/10000*np.abs(L_Pos)+1.5) #time for the thing to get to 0 otherwise it will add the commands to the previous one
    counter_powers=0

    root.after(200,loading)

    
def Go_To_Power():
    global sign_of_slope
    global check_begin
    global power_threshold
        
    def check_positive_negative_slope():
        global check_begin
        global Power1
        global Power2
        if check_begin:
            tlpm.measPower(byref(power))
            Power1=power.value
            LambdaOverTwo.move_by(10000)
            check_begin=False
            root.after(1500,check_positive_negative_slope)
        else:
            tlpm.measPower(byref(power))
            Power2=power.value
            LambdaOverTwo.move_by(-10000)
            if Power1>Power2:
                sign_of_slope=-1
            else:
                sign_of_slope=+1
            root.after(1500,loading)
            
    def loading():
        global power_threshold
        global max_power
        global min_power
        
        tlpm.measPower(byref(power))
        Power=power.value
        if abs(Power_wanted-Power)>power_threshold:
            steps_moved=int(sign_of_slope*50000/(max_power-min_power)*(Power_wanted-Power))
            LambdaOverTwo.move_by(steps_moved)
            root.after(int(250/10000*steps_moved+1500),loading)
        else:
            style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
            style_buttons.configure('TButton', font=font, background="#464646", foreground="#a6a6a6")

    power_threshold = 5e-6 #when closer than 5uW, stops moving
    
    Power_wanted = float(Power_wanted_entry.get())*1e-6
    
    if Power_wanted<max_power and Power_wanted>min_power :
        
        style_buttons.configure('TButton', font=font, background="#464646", foreground="red")
        style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    
        check_begin=True
        sign_of_slope=1
        
        root.after(200,check_positive_negative_slope)
    
def Power_series():
    global power_series_counter
    global Final_array_Power_Series
    global res_power
    global power
    global filename_new
    global cam_started
    global exposure
    
    start=time.time()
    
    cam.stop_acquisition()
    cam_started=False
    
    style_buttons.configure('TButton', font=font, background="#464646", foreground="red")
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    
    def Go_To_Power_min():
        global sign_of_slope
        global check_begin
        global power_threshold
            
        def check_positive_negative_slope():
            global check_begin
            global Power1
            global Power2
            if check_begin:
                tlpm.measPower(byref(power))
                Power1=power.value
                LambdaOverTwo.move_by(10000)
                check_begin=False
                root.after(1500,check_positive_negative_slope)
            else:
                tlpm.measPower(byref(power))
                Power2=power.value
                LambdaOverTwo.move_by(-10000)
                if Power1>Power2:
                    sign_of_slope=-1
                else:
                    sign_of_slope=+1
                root.after(1500,loading)
                
        def loading():
            global power_threshold
            global max_power
            global min_power
            
            tlpm.measPower(byref(power))
            Power=power.value
            if abs(Power_wanted-Power)>power_threshold:
                steps_moved=int(sign_of_slope*50000/(max_power-min_power)*(Power_wanted-Power))
                LambdaOverTwo.move_by(steps_moved)
                root.after(int(250/10000*steps_moved+1500),loading)

        power_threshold = 5e-6 #when closer than 5uW, stops moving
        
        Power_wanted = min_power
        
        check_begin=True
        sign_of_slope=1
        
        root.after(200,check_positive_negative_slope)
    
    def loading_Power_Series():
        global Final_array_Power_Series
        global power_series_counter
        global filename_new
        global res_power
        global power
        
        percent=float(power_series_counter/nb_power_steps*100)
        time_left=total_time*(100-percent)/100
        wait_time_label.config(text='t left = {:.1f}s'.format(time_left))
        percent_label.config(text='{:.1f}%'.format(percent))
        cam.set_acquisition_mode('cont')
        cam.set_exposure(exposure)
        cam.start_acquisition()
        time.sleep(exposure+1)
        
        tlpm.measPower(byref(power))
        res_power.append(power.value)
        
        B=cam.read_newest_image(peek=True)
        Final_array_Power_Series=np.append(Final_array_Power_Series,B,axis=0)
        
        cam.stop_acquisition()
        power_series_counter+=1
        percent=float(power_series_counter/nb_power_steps*100)
        time_left=total_time*(100-percent)/100
        hours_left=time_left//3600
        minutes_left=(time_left %3600)//60
        seconds_left=time_left % 60
        wait_time_label.config(text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left))
        percent_label.config(text='{:.1f}%'.format(percent))
        
        LambdaOverTwo.move_by(steps_moved)
        if percent<100:
            parent.after(int(250/10000*steps_moved+1500),loading_Power_Series)
        elif percent>=100:
            parent.destroy()
            style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
            style_buttons.configure('TButton', font=font, background="#464646", foreground="#a6a6a6")
        
            res_power.insert(0,0)
            res_power.insert(0,0)
            
            Final_array_Power_Series = np.transpose(Final_array_Power_Series)
            Final_array_Power_Series = np.insert(Final_array_Power_Series,obj=0,values=res_power,axis=0)
            
            Final_arra_Power_Seriesy = Final_array_Power_Series.astype('float64')
            
            #final array will be V/I 2 first lines WL/E 2 first columns and data below (2x2 0s in the left uppermost part of the array)
            df = pd.DataFrame(Final_array_Power_Series)
            n = 1
            while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
                filename_new = filename + '_' + str(n)
                n += 1
            file_name = filename_new + '.csv'
            df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
            
            end=time.time()
            print('time for power series = '+str(end-start))
    
    
    directory=str(Directory_entry.get())
    if os.path.exists(directory+"/Power_series")==False:
        os.makedirs(directory+"/Power_series")
    if os.path.exists(directory+"/Power plots")==False:
        os.makedirs(directory+"/Power plots")
    filename=str(Power_Series_Entry.get())
    filename=directory+'/Power_series/'+filename
    filename_new=filename
    
    root.after(200, Go_To_Power_min)
    
    nb_power_steps=int(Nb_Power_Steps_Entry.get())
    
    res_power=[]
    
    power_series_counter=0
    
    exposure=float(Exposure_entry.get())
    getWL=spec.get_calibration()
    WL=getWL*1e9
    E=6.63e-34*3e8/1.6e-19/getWL
    
    Final_array_Power_Series = np.array([WL,E]) #array contenant les data qui seront save directement
    
    parent=tk.Toplevel(root) #configuration of the loading window that will display percentage of the measurement and time left
    parent.configure(background="#464646")
    parent.title('Waiting window')
    parent.geometry('350x100+1000+400')
    
    steps_moved=int(90000/(nb_power_steps-1))
    
    total_time=(exposure+1)*nb_power_steps+(0.25/10000*steps_moved+1.5)*(nb_power_steps-1) #en s

    percent=float(0)
    time_left=total_time+20 #add in the time for go_to_power_min
    hours_left=time_left//3600
    minutes_left=(time_left %3600)//60
    seconds_left=time_left % 60
    
    percent_label=ttk.Label(parent,text='{:.1f}%'.format(percent), font=('American typewriter',20))
    percent_label.pack(side = 'bottom')
    wait_time_label=ttk.Label(parent,text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left),font=('American typewriter',20))
    wait_time_label.pack(side = 'top')
    parent.after(20000,loading_Power_Series)
    
    parent.mainloop() #loading window runs until the end of measurement
    
def go_to_angle():
    angle_wanted=float(Angle_Wanted_Entry.get())
    #this motor moves 360k steps to move 180deg
    #reference is gonna be zero steps = 0deg
    pos=LambdaOverTwoPolar.get_position()
    pos_wanted=int(angle_wanted/360*720000)
    steps_moved=abs(pos_wanted-pos)
    LambdaOverTwoPolar.move_to(pos_wanted)
    #time.sleep(0.250/10000*steps_moved+1.5)
    
def Polar_Series():
    global polar_series_counter
    global Final_array_Polar_Series
    global res_polar
    global filename_new
    global cam_started
    global exposure
    
    start=time.time()
    
    cam.stop_acquisition()
    cam_started=False
    
    style_buttons.configure('TButton', font=font, background="#464646", foreground="red")
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    
    def loading_Polar_Series():
        global Final_array_Polar_Series
        global polar_series_counter
        global filename_new
        global res_polar
        global exposure
        
        percent=float(polar_series_counter/nb_polar_steps*100)
        time_left=total_time*(100-percent)/100
        wait_time_label.config(text='t left = {:.1f}s'.format(time_left))
        percent_label.config(text='{:.1f}%'.format(percent))
        cam.set_acquisition_mode('cont')
        cam.set_exposure(float(exposure))
        cam.start_acquisition()
        time.sleep(exposure+1)
        
        step_value=LambdaOverTwoPolar.get_position()
        res_polar.append(step_value)
        
        B=cam.read_newest_image(peek=True)
        Final_array_Polar_Series=np.append(Final_array_Polar_Series,B,axis=0)
        
        cam.stop_acquisition()
        polar_series_counter+=1
        percent=float(polar_series_counter/nb_polar_steps*100)
        time_left=total_time*(100-percent)/100
        hours_left=time_left//3600
        minutes_left=(time_left %3600)//60
        seconds_left=time_left % 60
        wait_time_label.config(text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left))
        percent_label.config(text='{:.1f}%'.format(percent))
        
        LambdaOverTwoPolar.move_by(steps_moved)
        if percent<100:
            parent.after(int(250/10000*steps_moved+1500),loading_Polar_Series)
        elif percent>=100:
            parent.destroy()
            style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
            style_buttons.configure('TButton', font=font, background="#464646", foreground="#a6a6a6")
        
            res_polar.insert(0,0)
            res_polar.insert(0,0)
            
            Final_array_Polar_Series = np.transpose(Final_array_Polar_Series)
            Final_array_Polar_Series = np.insert(Final_array_Polar_Series,obj=0,values=res_polar,axis=0)
            
            Final_array_Polar_Series = Final_array_Polar_Series.astype('float64')
            
            #final array will be V/I 2 first lines WL/E 2 first columns and data below (2x2 0s in the left uppermost part of the array)
            df = pd.DataFrame(Final_array_Polar_Series)
            n = 1
            while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
                filename_new = filename + '_' + str(n)
                n += 1
            file_name = filename_new + '.csv'
            df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
            
            end=time.time()
            print('time for polar series = '+str(end-start))
    
    
    directory=str(Directory_entry.get())
    if os.path.exists(directory+"/Polar_series")==False:
        os.makedirs(directory+"/Polar_series")
    if os.path.exists(directory+"/Polar plots")==False:
        os.makedirs(directory+"/Polar plots")
    filename=str(Polar_Series_Entry.get())
    filename=directory+'/Polar_series/'+filename
    filename_new=filename
    
    L_Pos=LambdaOverTwoPolar.get_position()
    LambdaOverTwoPolar.move_to(0)
    time.sleep(12/200000*np.abs(L_Pos)+5)
    
    nb_polar_steps=int(Nb_Polar_Steps_Entry.get())
    
    res_polar=[]
    polar_series_counter=0
    
    exposure=float(Exposure_entry.get())
    getWL=spec.get_calibration()
    WL=getWL*1e9
    E=6.63e-34*3e8/1.6e-19/getWL
    
    Final_array_Polar_Series = np.array([WL,E]) #array contenant les data qui seront save directement
    
    parent=tk.Toplevel(root) #configuration of the loading window that will display percentage of the measurement and time left
    parent.configure(background="#464646")
    parent.title('Waiting window')
    parent.geometry('350x100+1000+400')
    
    steps_moved=int(360000/(nb_polar_steps-1))
    
    total_time=(exposure+1)*nb_polar_steps+(0.25/10000*steps_moved+1.5)*(nb_polar_steps-1) #en s

    percent=float(0)
    time_left=total_time+20 #add in the time for go_to_power_min
    hours_left=time_left//3600
    minutes_left=(time_left %3600)//60
    seconds_left=time_left % 60
    
    percent_label=ttk.Label(parent,text='{:.1f}%'.format(percent), font=('American typewriter',20))
    percent_label.pack(side = 'bottom')
    wait_time_label=ttk.Label(parent,text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left),font=('American typewriter',20))
    wait_time_label.pack(side = 'top')
    parent.after(30000,loading_Polar_Series)
    
    parent.mainloop() #loading window runs until the end of measurement
    
#%% iHR550 and Symphony Measurements

def test_write_console():
    console.write('Hello world', 'Konnichiwa', 6, 0.12313)

def cancel_current_operation():
    """ Doesn't actually do anything because there is no need to, just here to
        ensure our dear user that nothing has gone wrong """
       
    console.write('Current Operation CANCELLED')

def prep_e_v_map():
        # Number of Acq
        V_start = float(V_start_entry.get())
        V_end = float(V_finish_entry.get())
        nb_step = float(Step_entry.get())
        
        default_nb_acq = math.ceil((V_end-V_start)/nb_step) + 1
        nb_acq_entry.delete(0, tk.END)
        nb_acq_entry.insert(0, default_nb_acq)
        nb_acq_entry.update()
        
        console.write("Welcome to the program allowing you to take IV curve and spectral datas. Please follow the following steps carefully :")
        console.write("\n 1 - Open Synergy, then under 'Collect' in the top bar select 'Experiment Setup'. The first time you do this the Hardware should initialize.")
        console.write("\n 2 - Select the IV_Curve_Labjack.xml Experiment File. Select the number of spectra you wish to take under 'Accumulations', the exposition time and the center wavelength.")
        console.write("\n 3 - Under 'Triggers' at the bottom of the left pannel, select the following :")
        console.write("        Input Trigger : Symphony > Trigger Input > Each - For Each Acq > TTL Rising Edge")
        console.write("        Output Trigger : Symphony > Trigger Output 2 > Chip Readout > TTL Active High")
        console.write("WARNING : If the exposure time does not match the one you set in SynerJY, do the 3rd step even though the settings were pre-filled correctly.")
        console.write("\nFinally, return to 'General' click on RUN then press START in this console to start the measurements")
        
        # Activate Start button
        StartAcq_button.configure(command=e_v_map)
        
def unprep_e_v_map():
    """Before Start button is initalized by prep_e_v_map()"""
    
    console.write("To start an EV map, first press Start E-V.")
    
def e_v_map():
    
    StartAcq_button.configure(command=unprep_e_v_map)
        
    try:
        d = u3.U3()
            
        # # Number of Acq
        # V_start = float(V_start_entry.get())
        # V_end = float(V_finish_entry.get())
        # nb_step = float(Step_entry.get())
        
        # default_nb_acq = math.ceil((V_end-V_start)/nb_step)
        # nb_acq_entry.delete(0, tk.END)
        # nb_acq_entry.insert(0, default_nb_acq)
        # nb_acq_entry.update()
        
        # console.write("Welcome to the program allowing you to take IV curve and spectral datas. Please follow the following steps carefully :")
        # console.write("\n 1 - Open Synergy, then under 'Collect' in the top bar select 'Experiment Setup'. The first time you do this the Hardware should initialize.")
        # console.write("\n 2 - Select the IV_Curve_Labjack.xml Experiment File. Select the number of spectra you wish to take under 'Accumulations', the exposition time and the center wavelength.")
        # console.write("\n 3 - Under 'Triggers' at the bottom of the left pannel, select the following :")
        # console.write("        Input Trigger : Symphony > Trigger Input > Each - For Each Acq > TTL Rising Edge")
        # console.write("        Output Trigger : Symphony > Trigger Output 2 > Chip Readout > TTL Active High")
        # console.write("WARNING : If the exposure time does not match the one you set in SynerJY, do the 3rd step even though the settings were pre-filled correctly.")
        # input("\nFinally, return to 'General' click on RUN then press START in this console to start the measurements")
        # console.write("\nFinally, return to 'General' click on RUN then press START in this console to start the measurements")
        
        time.sleep(1)
            
        # Number of Acq in case they changed during the input
        V_start = float(V_start_entry.get())
        V_end = float(V_finish_entry.get())
        nb_step = float(Step_entry.get())
        
        acquisition_number = int(nb_acq_entry.get()) + 1

        console.write("\nNumber of acquisition :", acquisition_number)
        
        # Array used to store IV curve
        v_array = np.empty(acquisition_number)
        i_array = np.empty(acquisition_number)
        
        # Array of V that will be set to the Keithley
        v_toset = np.linspace(V_start, V_end, acquisition_number)
        
        for i in range(acquisition_number):
            
            v_array[i], i_array[i] = set_tension(v_toset[i])
            
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
            console.write("Exposure State")
            
            while is_exposing or is_readingout:
                TLLOutput2 = d.getAIN(0) # Reads TLLOutput2 state (active = 5V, unactive = 0V)
                # print(TLLOutput2)
                time.sleep(0.001) # Reads TLLOutput2 every 1 ms
                
                if (TLLOutput2 > ACTIVE_HIGH_THRESHOLD) and (is_readingout == False): # Enters reading out state
                    is_readingout = True
                    is_exposing = False
                    tic_exp = time.time()
                    console.write("Exposure Time :", tic_exp-tic, "s")
                    console.write("Reading Out State")
                elif (TLLOutput2 < UNACTIVE_HIGH_THRESHOLD) and (is_readingout == True): # Leaves reading out state
                    is_readingout = False
                    console.write("Leaving Reading Out State")
            
            time.sleep(0.1)
            # Now the CCD is neither is the exposure or the reading out state, we can Trigg IN again
        
        console.write("Measurements finished.")
        IV_data = {"Tension (V)" : v_array,
                   "Courant (A)" : i_array}
        
        df = pd.DataFrame(IV_data)
        
        # Create the savings directory if it does not already exist
        # Saves the data to this directory
        directory=str(Directory_entry.get())
        if os.path.exists(directory+"/Energy-Tension")==False:
            os.makedirs(directory+"/Energy-Tension")
        if os.path.exists(directory+"/2D plots")==False:
            os.makedirs(directory+"/2D plots")
        filename=str(TwoD_entry.get())
        filename=directory+'/Energy-Tension/'+filename
        filename_new=filename
        
        # add _n to the file in case of duplicates
        n = 1
        while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
            filename_new = filename + '_' + str(n)
            n += 1
        file_name = filename_new + '.csv'
        df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
        
        console.write("IV-Curve file created at", directory+'/Energy-Tension/')
        
    except:
        console.write("Error : Something went wrong. Note that calling u3.U3() returns the first object found, there might be conflicts if there are several devices connected to the PC.")
        
#%%Create tkinter window and styling for buttons, labels and entries

root = ThemedTk(theme='equilux')
root.geometry("1100x500+0+0")
root.configure(background="#464646")

font=('American typewriter', 12)

#Style
style = ttk.Style()
style.configure('Switch.TButton',background="#464646",borderwidth=0)

style_entries=ttk.Style()
style_entries.configure('TEntry', font=font, background="#000000", foreground="white")

style_buttons=ttk.Style()
style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
style_buttons.map('W.TButton', foreground = [('active', '!disabled', 'green')],
                     background = [('active', 'blue')])

style_all_buttons=ttk.Style()
style_all_buttons.configure('TButton', font=font, background="#464646", foreground="#a6a6a6")
style_all_buttons.map('TButton', foreground = [('active', '!disabled', 'green')],
                     background = [('active', 'blue')])

persistent_frame= ttk.Frame(root)
persistent_frame.grid(row=0,column=0,sticky='nsew')

cam_started=False
flipper_state='direct'

# class PrintLogger(tk.Text): # create file like object

#     def __init__(self, textbox): # pass reference to text widget
#         self.textbox = textbox # keep ref

#     def write(self, text, autoupdate=True):
#         self.textbox.configure(state='normal') # set the textbox as writable
#         self.textbox.insert(tk.END, text) # write text to textbox
#         self.textbox.configure(state='disabled') # set the textbox as read only
#             # could also scroll to end of textbox here to make sure always visible
#         if autoupdate:
#             self.textbox.update()
#         self.textbox.see(tk.END)
        
#     def flush(self): # needed for file like object
#         pass


# Console-like textbox for acquisition
console_log = scrolledtext.ScrolledText(persistent_frame, bg='black', fg='green2', height=1)
console_log.grid(row=9, column=0, rowspan=13, columnspan=4, padx=0, pady=0, sticky='nwse')
console = PrintLogger(console_log)

#%% Buttons and stuff

#Button to start acq
StartAcq_button= ttk.Button(persistent_frame,text='Start', style='W.TButton', command=unprep_e_v_map)
StartAcq_button.grid(row=30, column=0, padx=0, pady=0, sticky='w')

#Button to start acq
cancel_button= ttk.Button(persistent_frame,text='CANCEL', style='W.TButton', command=cancel_current_operation)
cancel_button.grid(row=30, column=1, padx=0, pady=0, sticky='w')

#%%Keithley related
turn_on_keithley_button = ttk.Button(persistent_frame,text='Turn Keithley ON',command=Turn_Keithley_Link_On)
turn_on_keithley_button.grid(row=2, column=0, padx=0, pady=0, sticky='w')

#Set V
set_tension_button = ttk.Button(persistent_frame,text='Set V')
set_tension_button.grid(row=2, column=1, padx=0, pady=0)
V_label = ttk.Label(persistent_frame,text="Tension (V)", font=font)
V_label.grid(row=2, column=2, padx=0, pady=0, sticky="w")
V_entry = ttk.Entry(persistent_frame, width=5, font=font)
V_entry.insert(0, "0")
V_entry.grid(row=2, column=2, padx=0, pady=0, sticky="e")
I_read_label = ttk.Label(persistent_frame, text='I= N/A', font=font)
I_read_label.grid(row=4, column=1, padx=0, pady=0)
V_read_label = ttk.Label(persistent_frame, text='V= N/A', font=font)
V_read_label.grid(row=3, column=1, padx=0, pady=0)

#Tension series
check_current_evolution_button = ttk.Button(persistent_frame,text="Start I-t",style='W.TButton')
check_current_evolution_button.grid(row=8, column=2, padx=0, pady=0)

tension_series_button = ttk.Button(persistent_frame,text='Start I-V',command=Tension_series,style='W.TButton')
tension_series_button.grid(row=6, column=2, padx=0, pady=0)

E_V_button = ttk.Button(persistent_frame,text='Start E-V',command=prep_e_v_map, style='W.TButton')
E_V_button.grid(row=7, column=2, padx=0, pady=0)

V_start_label = ttk.Label(persistent_frame,text="Start (V)", font=font)
V_start_label.grid(row=3, column=2, padx=0, pady=0, sticky="w")
V_start_entry = ttk.Entry(persistent_frame, width=5, font=font)
V_start_entry.insert(0, "0")
V_start_entry.grid(row=3, column=2, padx=0, pady=0, sticky="e")
V_finish_label = ttk.Label(persistent_frame,text="Finish (V)", font=font)
V_finish_label.grid(row=4, column=2, padx=0, pady=0, sticky="w")
V_finish_entry = ttk.Entry(persistent_frame, width=5, font=font)
V_finish_entry.insert(0, "0")
V_finish_entry.grid(row=4, column=2, padx=0, pady=0, sticky="e")
Step_label = ttk.Label(persistent_frame,text="Step", font=font)
Step_label.grid(row=5, column=2, padx=0, pady=0, sticky="w")
Step_entry = ttk.Entry(persistent_frame, width=5, font=font)
Step_entry.insert(0, "0.05")
Step_entry.grid(row=5, column=2, padx=0, pady=0, sticky="e")

# Number of Acq
default_nb_acq = math.ceil(abs(float(V_finish_entry.get())-float(V_start_entry.get())/float(Step_entry.get())))
nb_acq_entry = ttk.Entry(persistent_frame, width=10, font=font)
nb_acq_entry.insert(0, "0")
nb_acq_entry.grid(row=30, column=2, padx=0, pady=0)

#Button Keithley 
set_tension_button.config(command=Set_Keithley_tension)
tension_series_button.config(command=Tension_series)
# E_V_button.config(command=Map_Tension_Energy)
check_current_evolution_button.config(command=Current_evolution)
        
#%%Power control
Power_button = ttk.Button(persistent_frame,text='Start PM100',command=start_PM)
Power_button.grid(row=2,column=4,padx=0,pady=0)
Power_value = ttk.Label(persistent_frame,text='Power = N/A', font=font)
Power_value.grid(row=3, column=4, padx=0, pady=0)
Power_wanted_entry = ttk.Entry(persistent_frame, width=5, font=font)
Power_wanted_entry.insert(0, "50")
Power_wanted_entry.grid(row=7, column=4, padx=0, pady=0)
Get_min_max_powers_button = ttk.Button(persistent_frame, text='Start Power Control',command=start_Power_Control)
Get_min_max_powers_button.grid(row=2, column=5, padx=0, pady=0)
Go_To_Power_button = ttk.Button(persistent_frame, text="Go to power (uw)")
Go_To_Power_button.grid(row=6, column=4,padx=0,pady=0)
min_power_label=ttk.Label(persistent_frame,text="Min power = N/A", font=font)
min_power_label.grid(row=3, column=5, padx=0, pady=0)
max_power_label=ttk.Label(persistent_frame,text="Max power = N/A", font=font)
max_power_label.grid(row=4, column=5, padx=0, pady=0)
PM_exp_entry = ttk.Entry(persistent_frame,width=5, font=font)
PM_exp_entry.insert(0, "100")
PM_exp_entry.grid(row=5,column=4,padx=0,pady=0)
PM_exp_label =ttk.Label(persistent_frame,text="PM refresh (ms)", font=font)
PM_exp_label.grid(row=4, column=4, padx=0, pady=0)
Start_Power_Series_button = ttk.Button(persistent_frame,text="Power series",style='W.TButton')
Start_Power_Series_button.grid(row=7,column=5,padx=0,pady=0)
Nb_Power_Steps_Label = ttk.Label(persistent_frame,text='Nb power steps', font=font)
Nb_Power_Steps_Label.grid(row=5,column=5,padx=0,pady=0)
Nb_Power_Steps_Entry = ttk.Entry(persistent_frame,width=5, font=font)
Nb_Power_Steps_Entry.insert(0, "2")
Nb_Power_Steps_Entry.grid(row=6,column=5,padx=0,pady=0)
Power_Series_Label = ttk.Label(persistent_frame,text='Power series name', font=font)
Power_Series_Label.grid(row=8,column=4,padx=0,pady=0)
Power_Series_Entry = ttk.Entry(persistent_frame,width=40, font=font)
Power_Series_Entry.grid(row=9,column=4,columnspan=3,padx=0,pady=0)

#%%Polar Control
Start_Polar_Control_button = ttk.Button(persistent_frame,text='Start Polar Control',command=start_Polar_Control)
Start_Polar_Control_button.grid(row=14,column=4,padx=0,pady=0)
Angle_Wanted_Label= ttk.Label(persistent_frame,text="Angle wanted (deg)",font=font)
Angle_Wanted_Label.grid(row=15,column=4,padx=0,pady=0)
Angle_Wanted_Entry=ttk.Entry(persistent_frame,width=5,font=font)
Angle_Wanted_Entry.insert(0,"0")
Angle_Wanted_Entry.grid(row=16,column=4,padx=0,pady=0)
Start_Polar_Series_button = ttk.Button(persistent_frame,text='Start Polar Series',style='W.TButton')
Start_Polar_Series_button.grid(row=16,column=5,padx=0,pady=0)

Nb_Polar_Steps_Label= ttk.Label(persistent_frame,text="Nb polar steps",font=font)
Nb_Polar_Steps_Label.grid(row=14,column=5,padx=0,pady=0)
Nb_Polar_Steps_Entry= ttk.Entry(persistent_frame,width=5, font=font)
Nb_Polar_Steps_Entry.insert(0,"10")
Nb_Polar_Steps_Entry.grid(row=15,column=5,padx=0,pady=0)
Polar_Series_Label= ttk.Label(persistent_frame,text='Polar series name', font=font)
Polar_Series_Label.grid(row=17,column=4,padx=0,pady=0)
Polar_Series_Entry= ttk.Entry(persistent_frame,width=40, font=font)
Polar_Series_Entry.grid(row=18,column=4,columnspan=2,padx=0,pady=0)

#%%Saving part

IV_entry = PlaceholderEntry(persistent_frame, width=20, font=font, placeholder="I-V name")
IV_entry.grid(row=6, column=0, columnspan=2, padx=0, pady=0, sticky="we")

TwoD_entry = PlaceholderEntry(persistent_frame, width=20, font=font, placeholder="E-V name")
TwoD_entry.grid(row=7, column=0, columnspan=2, padx=0, pady=0, sticky="we")

It_entry = PlaceholderEntry(persistent_frame, width=20, font=font, placeholder="I-t name")
It_entry.grid(row=8, column=0, columnspan=2, padx=0, pady=0, sticky="we")

Directory_label = ttk.Label(persistent_frame,text="Directory path",font=font)
Directory_label.grid(row=0, column=0, padx=0, pady=0,sticky='w')
Directory_entry = ttk.Entry(persistent_frame, width=75,font=font)
Directory_entry.insert(0, "C:/Users/Manip/Documents/"+today_date_str)
Directory_entry.grid(row=1, column=0, columnspan=4, padx=0, pady=0, sticky='w')
browse_button = ttk.Button(persistent_frame, text='Browse', command=browse_directory)
browse_button.grid(row=1,column=4, padx=0, pady=0, sticky='e')

#%% Main
persistent_frame.mainloop()


stop_Keithley()
stop_PM100()
    
try:
    LambdaOverTwo.close()
except:
    a=1
    
try:
    LambdaOverTwoPolar.close()
except:
    a=1
    
sys.exit()