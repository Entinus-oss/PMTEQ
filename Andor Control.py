# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 15:39:00 2024
https://pylablib.readthedocs.io/en/latest/.apidoc/pylablib.devices.Andor.html

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

# from qmi.instruments.picoquant.picoharp import PicoQuant_PicoHarp300
# from qmi.core.context import QMI_Context

import os
from datetime import date

# from pylablib.devices import Andor
from pylablib.devices import Thorlabs
# from pylablib.devices import M2

# pll.par["devices/dlls/andor_shamrock"] = "C:/Program Files/Andor SDK/Shamrock64/ShamrockCIF.dll"
# pll.par["devices/dlls/andor_sdk2"] = "path/to/sdk2/dlls"

plt.style.use('dark_background')

#%%Time

today_date=date.today()
today_date_str=today_date.strftime("%Y%m%d")

if os.path.exists("C:/Users/Manip/Documents/"+today_date_str)==False:
    os.makedirs("C:/Users/Manip/Documents/"+today_date_str)

#%%Functions
def Nothing():
    e=1

def start_Andor_Shamrock():
    global cam
    global spec
    cam = Andor.AndorSDK2Camera()
    spec = Andor.ShamrockSpectrograph()
    
    cam.open()
    spec.open()
    spec.set_flipper_port('output','direct')
    cam.set_read_mode('fvb')
    spec.setup_pixels_from_camera(cam)
    
    #all buttons associated to andor camera get their command affected
    cooler_button.configure(command = start_cooler)

    StartAcq_button.configure(command = start_camera_acquisition)

    StopAcq_button.configure(command = stop_camera_acquisition)

    WL_button.configure(command = set_WL)

    T_measure_button.configure(command = read_temperature)
    
    WL_measure_button.configure(command = read_wavelength)
    
    Flipper_port_button.configure(command=flip_output_port)
    
    Long_spectrum_button.configure(command = Long_spectrum)
    
    savefile_button.configure(command=save_file)

def start_cooler():
    T=float(T_entry.get())
    cam.set_temperature(T,enable_cooler= True)
    
def read_temperature():
    T_read_label.config(text='T_camera={:.1f} C'.format(cam.get_temperature()))
    
def read_wavelength():
    WL_read_label.config(text='Central WL={:.4f} nm'.format(spec.get_wavelength()*1e9))
    
def set_WL():
    WL=float(WL_entry.get())*1e-9
    spec.set_wavelength(WL)
    
def start_camera_acquisition():
    global cam_started
    
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    cam.set_acquisition_mode('cont')
    cam.set_exposure(float(Exposure_entry.get()))
    cam.start_acquisition()
    cam_started=True
    
def stop_camera_acquisition():
    global cam_started
    
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
    cam.stop_acquisition()
    cam_started=False
    
def flip_output_port():
    #function to decide wether the PL hits the CCD camera (direct) or the outside fiber, probably linked to an APD (side)
    global flipper_state
    if flipper_state=='direct':
        spec.set_flipper_port('output','side')
    else:
        spec.set_flipper_port('output','direct')
    flipper_state=spec.get_flipper_port('output')
    Flipper_port_label.config(text=flipper_state)
    
def switch_changed():
    #visual effect for flipping the switches for eV/nm or autoscale
    if switch_var.get():
        switch_nm_eV.config(image=switch_on_image)
    else:
        switch_nm_eV.config(image=switch_off_image)
        
def browse_directory():
    global directory
    directory=filedialog.askdirectory()
    Directory_entry.delete(0, tk.END)
    Directory_entry.insert(0,directory)
    
def save_file():
    global Power_value
    DataArray=np.array((E,WL,C[0]))
    DataArray=np.transpose(DataArray)
    
    directory=str(Directory_entry.get())
    
    Tension=str(V_entry.get())
    Excitation_wavelength=str(ExcWL_entry.get())
    Excitation_power=str(Power_value)
    Central_wavelength=str(WL_entry.get())
    Exposure=str(Exposure_entry.get())
    
    filename='Exc'+Excitation_wavelength+'_P'+Excitation_power+'_V'+Tension+'_WL'+Central_wavelength+'_Exp'+Exposure+'.txt'
    np.savetxt(str(directory)+'/'+str(filename),DataArray)

#%%Keithley related

def Turn_Keithley_Link_On():
    global Keithley
    rm = visa.ResourceManager()
    Keithley=rm.open_resource('GPIB2::24::INSTR')
    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0") #sets voltage to 0
    Keithley.write(":OUTP ON") #sets ouput to on
    Keithley.write(":SENS:CURR:PROT 500E-6") #sets compliance limit on I
    
    #buttons associated to Keithley get their command assigned
    set_tension_button.config(command=Set_Keithley_tension)
    tension_series_button.config(command=Tension_series)
    E_V_button.config(command=Map_Tension_Energy)
    check_current_evolution_button.config(command=Current_evolution)
    
def close_Keithley():
    try:
        Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0") #sets voltage to 0
        Keithley.write(":OUTP OFF")  # Turns off output
        Keithley.close()
    except:
        a=1
    
def Set_Keithley_tension():
    global Keithley
    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL "+str(V_entry.get())) #sets voltage to required entry
    Keithley.write(":MEAS:CURR:DC?")  # Measure current
    Q = Keithley.read()  #Measurement is a table, first is voltage, second current
    A = np.array(Q.split(','))
    A = A.astype('float64')
    cur_volt = A[0]  # Data storage
    cur_amp = A[1]
    I_read_label.config(text='I={:.5f} uA'.format(cur_amp*1e6)) #showing the current flowing through
    
def Current_evolution():
    #function to measure current over time for capacitive effects
    global Keithley
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
    
    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0") #making sure
    Keithley.write(":OUTP ON")
    time.sleep(1)
    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL "+str(Start_voltage))
    time.sleep(1)
    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0")
    tic = time.time()
    for k in range(1000):
        time.sleep(0.1)
        toc = time.time()
        Keithley.write(":MEAS:CURR:DC?")  # Commande de mesure
        Q = Keithley.read()  # La mesure est un tableau, la premiere case est la tension, la seconde le courant
        A = np.array(Q.split(','))
        A = A.astype('float64')
        current = A[1]
        res_amp.append(current)
        res_time.append(toc-tic)
        
    df = pd.DataFrame(list(zip(res_time, res_amp)), columns=['Time (s)', 'Intensity (A)'])

    n = 1
    while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
        filename_new = filename + '_' + str(n)
        n += 1
    file_name = filename_new + '.csv'
    df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
        
    
def Tension_series():
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

    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0")  # Mise a 0V de la tension pour etre sur
    Keithley.write(":OUTP ON")  # Allumage de loutput

    for i in range(len(volt)):  # Boucle pour les mesures
        Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL " + str(volt[i]))  # Application de la tension
        time.sleep(0.05)  # Waiting for a permanent regime
        Keithley.write(":MEAS:CURR:DC?")  # Commande de mesure
        Q = Keithley.read()  # La mesure est un tableau, la premiere case est la tension, la seconde le courant
        A = np.array(Q.split(','))
        A = A.astype('float64')
        cur_volt = A[0]  # Stockage des donnees
        cur_amp = A[1]
        res_volt.append(cur_volt)  # Stockage des donnees
        res_amp.append(cur_amp)

    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0")
    
    df = pd.DataFrame(list(zip(res_volt, res_amp)), columns=['Tension (V)', 'Intensity (A)'])

    n = 1
    while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
        filename_new = filename + '_' + str(n)
        n += 1
    file_name = filename_new + '.csv'
    df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
    
def Map_Tension_Energy():
    global i
    global percent
    global time_left
    global wait_time_label
    global percent_label
    global parent
    global Final_array
    global filename_new
    global cam_started
    global exposure
    global New_Table
    global colorbar
    
    def loading_2D_map():
        #function to be done each time a new measurement in the map is finished (adds measurement to the table + actualizes percentage and live feed)
        global Final_array
        global i
        global filename_new
        global New_Table
        global colorbar
        percent=float(i/len(volt)*100)
        time_left=total_time*(100-percent)/100
        wait_time_label.config(text='t left = {:.1f}s'.format(time_left))
        percent_label.config(text='{:.1f}%'.format(percent))
        cam.set_acquisition_mode('cont')
        cam.set_exposure(exposure)
        cam.start_acquisition()
        Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL " + str(volt[i]))  # Application de la tension
        time.sleep(exposure+1)  # Waiting for a permanent regime
        Keithley.write(":MEAS:CURR:DC?")  # Commande de mesure
        Q = Keithley.read()  # La mesure est un tableau, la premiere case est la tension, la seconde le courant
        A = np.array(Q.split(','))
        A = A.astype('float64')
        cur_volt = A[0]  # Stockage des donnees
        cur_amp = A[1]
        res_volt.append(cur_volt)  # Stockage des donnees
        res_amp.append(cur_amp)
        
        B=cam.read_newest_image(peek=True) #reads Andor latest measurement without destroying it (peek=True)
        Final_array=np.append(Final_array,B,axis=0) #add measurement to the others in a 2D map
        
        cam.stop_acquisition()
        
        New_Table=np.transpose(New_Table)
        New_Table[i]=B
        New_Table=np.transpose(New_Table)
        
        i+=1 #writing percentage of elapsed measurement
        percent=float(i/len(volt)*100)
        time_left=total_time*(100-percent)/100
        hours_left=time_left//3600
        minutes_left=(time_left %3600)//60
        seconds_left=time_left % 60
        wait_time_label.config(text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left))
        percent_label.config(text='{:.1f}%'.format(percent))
        
        #plotting data live
        extent=[volt[0],volt[-1],WL[-1],WL[0]]
        ax.clear()
        ax.set_title('Acquiring...')
        ax.set_aspect((extent[1]-extent[0])/(extent[2]-extent[3]))
        
        colors='viridis'
        vmin=np.min([[New_Table[i][j]for i in range(len(New_Table))] for j in range(i)])
        vmax=np.partition(New_Table.flatten(),-10)[-10]
        vmax=(vmax+vmin)/2
        pos=ax.imshow(New_Table, cmap=colors,vmin=vmin, vmax=vmax, interpolation='none', extent=extent)
        ax.set_aspect((extent[1]-extent[0])/(extent[2]-extent[3]))
        plt.xlabel('Bias voltage (V)')
        plt.ylabel('Wavelength (nm)')
        
        plt.ticklabel_format(style='sci',scilimits=(0,3))
        
        def WL_to_E(x):
            #function to create a secondary axis on the map
            near_zero=np.isclose(x,0)
            x2=np.array(x,float)
            x2[near_zero]=np.inf
            x2[~near_zero]=6.63e-34*3e8/1.6e-19/(x[~near_zero]*1e-9)
            return x2
        
        E_to_WL=WL_to_E
        
        secax = ax.secondary_yaxis('right', functions=(WL_to_E, E_to_WL))
        secax.set_ylabel('Energy (eV)')
        
        Actualized_map.draw()
        
        if percent<100:
            parent.after(200,loading_2D_map)
        elif percent>=100:
            parent.destroy()
            style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
            res_volt.insert(0,0)
            res_volt.insert(0,0)
            res_amp.insert(0,0)
            res_amp.insert(0,0)
            
            Final_array = np.transpose(Final_array)
            Final_array = np.insert(Final_array,obj=0,values=res_amp,axis=0)
            Final_array = np.insert(Final_array,obj=0,values=res_volt,axis=0)
            
            Final_array = Final_array.astype('float64')
            
            #final array will be V/I 2 first lines WL/E 2 first columns and data below (2x2 0s in the left uppermost part of the array)
            df = pd.DataFrame(Final_array)
            n = 1
            while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
                filename_new = filename + '_' + str(n)
                n += 1
            file_name = filename_new + '.csv'
            df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')

            Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0")
    
    
    cam.stop_acquisition()
    cam_started=False
    
    directory=str(Directory_entry.get())
    if os.path.exists(directory+"/Energy-Tension")==False:
        os.makedirs(directory+"/Energy-Tension")
    if os.path.exists(directory+"/2D plots")==False:
        os.makedirs(directory+"/2D plots")
    filename=str(TwoD_entry.get())
    filename=directory+'/Energy-Tension/'+filename
    filename_new=filename
    
    step=float(Step_entry.get())
    volt = np.arange(float(V_start_entry.get()), float(V_finish_entry.get())+step/2, step)  # Definition de toutes les tensions a explorer
    res_volt = []  # Vecteur qui contiendra la tension appliquee
    res_amp = []  # Vecteur qui contiendra le courant mesure

    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0")  # Mise a 0V de la tension pour etre sur
    Keithley.write(":OUTP ON")  # Allumage de loutput
    
    exposure=float(Exposure_entry.get())
    getWL=spec.get_calibration()
    WL=getWL*1e9
    E=6.63e-34*3e8/1.6e-19/getWL
    
    Final_array = np.array([WL,E]) #array contenant les data qui seront save directement
    
    parent=tk.Toplevel(root) #configuration of the loading window that will display percentage of the measurement and time left
    parent.configure(background="#464646")
    parent.title('Waiting window')
    parent.geometry('900x900+850+100')
    
    total_time=(exposure+1)*len(volt) #en s
    i=0
    percent=float(0)
    time_left=total_time
    hours_left=time_left//3600
    minutes_left=(time_left %3600)//60
    seconds_left=time_left % 60
    
    percent_label=ttk.Label(parent,text='{:.1f}%'.format(percent), font=('American typewriter',20))
    percent_label.pack(side = 'top')
    wait_time_label=ttk.Label(parent,text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left),font=('American typewriter',20))
    wait_time_label.pack(side = 'top')
    
    fig,ax=plt.subplots(figsize=(15,12))
    Actualized_map=FigureCanvasTkAgg(fig,parent)
    Actualized_map.get_tk_widget().pack(side='bottom')
    
    extent=[volt[0],volt[-1],WL[-1],WL[0]]
    
    ax.set_title('Acquiring...')
    ax.set_aspect((extent[1]-extent[0])/(extent[2]-extent[3]))
    
    colors='viridis'
    vmin=700
    vmax=1100
    New_Table=np.zeros((len(WL),len(volt)))
    pos=ax.imshow(New_Table, cmap=colors,vmin=vmin, vmax=vmax, interpolation='none', extent=extent)
    ax.set_aspect((extent[1]-extent[0])/(extent[2]-extent[3]))
    plt.xlabel('Bias voltage (V)')
    plt.ylabel('Wavelength (nm)')
    
    plt.ticklabel_format(style='sci',scilimits=(0,3))
    
    def WL_to_E(x):
        near_zero=np.isclose(x,0)
        x2=np.array(x,float)
        x2[near_zero]=np.inf
        x2[~near_zero]=6.63e-34*3e8/1.6e-19/(x[~near_zero]*1e-9)
        return x2
    
    E_to_WL=WL_to_E
    
    secax = ax.secondary_yaxis('right', functions=(WL_to_E, E_to_WL))
    secax.set_ylabel('Energy (eV)')
    parent.after(200,loading_2D_map)
    
    parent.mainloop() #loading window runs until the end of measurement

def Long_spectrum():
    global cam_started
    global power
    
    directory=str(Directory_entry.get())
    if os.path.exists(directory+"/Long_spectrum")==False:
        os.makedirs(directory+"/Long_spectrum")
    if os.path.exists(directory+"/Long_spectrum_images")==False:
        os.makedirs(directory+"/Long_spectrum_images")
    
    #measure power once to put it inside title
    power =  c_double()
    tlpm.measPower(byref(power))
    Power_value=int(power.value*1e6)
    
    cam.stop_acquisition()
    cam_started=False
    
    #visually buttons turn red
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    
    #get WL start and stop from entries
    WL_start=float(WL_start_entry.get())
    WL_stop=float(WL_stop_entry.get())
    
    exposure=float(Exposure_entry.get())
    cam.set_acquisition_mode('cont')
    cam.set_exposure(exposure)
    cam.start_acquisition()
    
    Final_spectrum=np.array([[]])
    
    central_WL=WL_start+19
    while central_WL<=WL_stop-19: #do a measurement each 33nm (not 40 because we remove sides of the CCD)
        spec.set_wavelength(central_WL*1e-9)
        getWL=spec.get_calibration()
        WL=getWL*1e9
        E=6.63e-34*3e8/1.6e-19/getWL
    
        time.sleep(2*exposure)
        B=cam.read_newest_image(peek=True)
    
        if np.shape(Final_spectrum)==(1,0):
            Final_spectrum=np.append(Final_spectrum,WL[50:1950])
            Final_spectrum=np.append([Final_spectrum],np.array([E[50:1950],B[0,50:1950]]), axis=0)
        else :
            # Base_level=np.mean(Final_spectrum[2,-100:-1])
            # New_base_level=np.mean(B[0,1900:2000])
            # Base_level_erasure=New_base_level-Base_level
            
            To_append = np.array([WL[50:1950],E[50:1950],B[0,50:1950]])#-Base_level_erasure])
            Final_spectrum=np.append(Final_spectrum,To_append, axis=1)
        
        central_WL+=33
        
    spec.set_wavelength((WL_stop-19)*1e-9)
    getWL=spec.get_calibration()
    WL=getWL*1e9
    E=6.63e-34*3e8/1.6e-19/getWL
    
    time.sleep(2*exposure)
    B=cam.read_newest_image(peek=True)
    
    pix = min(range(len(WL)), key=lambda i: abs(WL[i]-Final_spectrum[0,-1])) #on recherche le recouvrement avec le dernier spectre
    
    # Base_level=np.mean(Final_spectrum[2,-2000+pix:-1])
    # New_base_level=np.mean(B[0,pix:2000])
    # Base_level_erasure=New_base_level-Base_level
    To_append = np.array([WL[pix-50:1950],E[pix-50:1950],B[0,pix-50:1950]])#-Base_level_erasure])
    
    Final_spectrum=np.append(Final_spectrum,To_append, axis=1)
    
    cam.stop_acquisition()
    
    
    
    df = pd.DataFrame(Final_spectrum)
    
    Tension=str(V_entry.get())
    Excitation_wavelength=str(ExcWL_entry.get())
    Excitation_power=str(Power_value)
    Exposure=str(Exposure_entry.get())
    
    filename=directory+'\Long_spectrum\Exc'+Excitation_wavelength+'_P'+Excitation_power+'uW_V'+Tension+'_WL_'+str(WL_start)+'to'+str(WL_stop)+'_Exp'+Exposure
    filename_new=filename
    n = 1
    while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
        filename_new = filename + '_' + str(n)
        n += 1
    file_name = filename_new + '.csv'
    df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
   
#%%PM100 related

def start_PM():
    global tlpm
    global PM_refresh
    tlpm=TLPM()
    resourceName = create_string_buffer(b'USB0::0x1313::0x8072::P2003849::INSTR')
    tlpm.open(resourceName, c_bool(True), c_bool(True))
    PM_refresh=int(PM_exp_entry.get())
    Power_button.config(text= 'set PM refresh', command=set_PM_refresh)
    root.after(200,measure_power)
    
def measure_power():
    global tlpm
    global Power_value
    global PM_refresh
    global power
    power =  c_double()
    tlpm.measPower(byref(power))
    Power_value.config(text='{:.0f} uW'.format(power.value*1e6))
    root.after(PM_refresh,measure_power)
    
def set_PM_refresh():
    global PM_refresh
    PM_refresh=int(PM_exp_entry.get())
    
def close_PM100():
    try:
        tlpm.close()
    except:
        a=1
    
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
    
#%%M2 related
    
def turn_on_M2():
    global Msquared
    Msquared = M2.solstis.Solstis('192.168.1.222', 39933, timeout=5.0) #automatically connects wvmeter if Solstis and Wvmeter windows are open
    
    Set_M2_WL_Button.config(command=Set_M2_WL)
    PLE_start_button.config(command=PLE)
    
    tab_PLE.after(1000,measure_M2_WL)
    
def Set_M2_WL():
    
    wanted_WL=float(Set_M2_WL_Entry.get())
    
    Msquared.lock_etalon()
    Msquared.fine_tune_wavelength(wanted_WL/1e9, sync=False, timeout=0.1)
    
    Real_WL=Msquared.get_fine_wavelength()*1e9
    
    status=Msquared.get_fine_tuning_status()
    
    while abs(Real_WL-wanted_WL)>0.01 and status=="tuning":
        time.sleep(0.01)
        Real_WL=Msquared.get_fine_wavelength()*1e9
        status=Msquared.get_fine_tuning_status()
    
    Msquared.stop_fine_tuning()
    Msquared.unlock_etalon()
    time.sleep(0.1)
    Final_WL=Msquared.get_fine_wavelength()*1e9
    
    Real_M2_WL_Label.config(text='{:.4f} nm'.format(Final_WL))
    
def measure_M2_WL():
    Real_WL=Msquared.get_fine_wavelength()*1e9
    
    Real_M2_WL_Label.config(text='{:.4f} nm'.format(Real_WL))
    
    tab_PLE.after(100,measure_M2_WL)
    
def go_to_WL(wanted_WL):
    Msquared.lock_etalon()
    Msquared.fine_tune_wavelength(wanted_WL/1e9, sync=False, timeout=0.1)
    
    Real_WL=Msquared.get_fine_wavelength()*1e9
    
    status=Msquared.get_fine_tuning_status()
    
    while abs(Real_WL-wanted_WL)>0.01 and status=="tuning":
        time.sleep(0.01)
        Real_WL=Msquared.get_fine_wavelength()*1e9
        status=Msquared.get_fine_tuning_status()
    
    Msquared.stop_fine_tuning()
    Msquared.unlock_etalon()
    time.sleep(0.1)
    Final_WL=Msquared.get_fine_wavelength()*1e9
    
    return Final_WL

def PLE():
    global Msquared
    global PLE_counter
    global Final_array_PLE_Series
    global filename_new
    global res_WL
    global cam_started
    global exposure
    
    def loading_PLE_Series():
        global Final_array_PLE_Series
        global PLE_counter
        global filename_new
        global res_WL
        global exposure

        
        percent=float(PLE_counter/nb_PLE_steps*100)
        time_left=total_time*(100-percent)/100
        hours_left=time_left//3600
        minutes_left=(time_left %3600)//60
        seconds_left=time_left % 60
        wait_time_label.config(text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left),font=('American typewriter',20))
        percent_label.config(text='{:.1f}%'.format(percent))

        WL=goal_WLs[PLE_counter]
        Real_WL=go_to_WL(WL)

        #3 times because if it fails, let's try again but not an infinite amount of time, sometimes lambda is just not possible to reach with a sufficiently low error
        if abs(Real_WL-WL)>0.05:
            Real_WL=go_to_WL(WL)
            
        if abs(Real_WL-WL)>0.05:
            Real_WL=go_to_WL(WL)
            
        if abs(Real_WL-WL)>0.05:
            Real_WL=go_to_WL(WL)
        
        Final_WLs.append(Real_WL)
        
        time.sleep(1)

        cam.set_acquisition_mode('cont')
        cam.set_exposure(float(exposure))
        cam.start_acquisition()
        time.sleep(exposure+1)
        
        B=cam.read_newest_image(peek=True)
        Final_array_PLE_Series=np.append(Final_array_PLE_Series,B,axis=0)
        
        cam.stop_acquisition()
        
        power =  c_double()
        tlpm.measPower(byref(power))
        Power=power.value*1e6
        Powers.append(Power)
        
        PLE_counter+=1
        percent=float(PLE_counter/nb_PLE_steps*100)
        time_left=total_time*(100-percent)/100
        hours_left=time_left//3600
        minutes_left=(time_left %3600)//60
        seconds_left=time_left % 60
        wait_time_label.config(text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left))
        percent_label.config(text='{:.1f}%'.format(percent))

        if percent<100:
            parent.after(2000,loading_PLE_Series)
        elif percent>=100:
            parent.destroy()
            style_buttons.configure('W.TButton', font=font, background="#464646", foreground="white")
            style_buttons.configure('TButton', font=font, background="#464646", foreground="#a6a6a6")
        
            Final_WLs.insert(0,0)
            Final_WLs.insert(0,0)

            Powers.insert(0,0)
            Powers.insert(0,0)
            
            Final_array_PLE_Series = np.transpose(Final_array_PLE_Series)

            Final_array_PLE_Series = np.insert(Final_array_PLE_Series,obj=0,values=Powers,axis=0)
            Final_array_PLE_Series = np.insert(Final_array_PLE_Series,obj=0,values=Final_WLs,axis=0)
            
            Final_array_PLE_Series = Final_array_PLE_Series.astype('float64')
            
            #final array will be laser WL/Power 2 first lines WL/E 2 first columns and data below (2x2 0s in the left uppermost part of the array)
            df = pd.DataFrame(Final_array_PLE_Series)
            n = 1
            while exists(filename_new+ '.csv') == True:  # Boucle pour ne pas effacer de donnees
                filename_new = filename + '_' + str(n)
                n += 1
            file_name = filename_new + '.csv'
            df.to_csv(file_name, sep='\t', index=False, encoding='utf-8')
            
            end=time.time()
            print('time for PLE series = '+str(end-start))
    
    
    directory=str(Directory_entry.get())
    if os.path.exists(directory+"/PLE_series")==False:
        os.makedirs(directory+"/PLE_series")
    if os.path.exists(directory+"/PLE plots")==False:
        os.makedirs(directory+"/PLE plots")
    filename=str(PLE_Series_Entry.get())
    filename=directory+'/PLE_series/'+filename
    filename_new=filename
    
    start=time.time()
    
    cam.stop_acquisition()
    cam_started=False
    
    style_buttons.configure('TButton', font=font, background="#464646", foreground="red")
    style_buttons.configure('W.TButton', font=font, background="#464646", foreground="red")
    
    begin=float(PLE_begin_WL_Entry.get())
    end=float(PLE_end_WL_Entry.get())
    step=float(PLE_step_Entry.get())
    
    Final_WLs=[]
    Powers=[]
    
    goal_WLs=np.arange(begin,end,step)
    
    nb_PLE_steps=len(goal_WLs)
    
    PLE_counter=0
    
    exposure=float(Exposure_entry.get())
    getWL=spec.get_calibration()
    WL=getWL*1e9
    E=6.63e-34*3e8/1.6e-19/getWL
    
    Final_array_PLE_Series = np.array([WL,E]) #array contenant les data qui seront save directement
    
    parent=tk.Toplevel(tab_PLE) #configuration of the loading window that will display percentage of the measurement and time left
    parent.configure(background="#464646")
    parent.title('Waiting window')
    parent.geometry('350x100+1000+400')
    
    total_time=(exposure+1+3*5)*nb_PLE_steps #en s

    percent=float(0)
    time_left=total_time
    hours_left=time_left//3600
    minutes_left=(time_left%3600)//60
    seconds_left=time_left%60
    
    percent_label=ttk.Label(parent,text='{:.1f}%'.format(percent), font=('American typewriter',20))
    percent_label.pack(side = 'bottom')
    wait_time_label=ttk.Label(parent,text='Time left = {:.0f}h'.format(hours_left)+'{:.0f}m'.format(minutes_left)+'{:.1f}s'.format(seconds_left),font=('American typewriter',20))
    wait_time_label.pack(side = 'top')
    
    parent.after(5000,loading_PLE_Series)
    
    parent.mainloop() #loading window runs until the end of measurement
    
#%%PicoHarp
def turn_on_PicoHarp():
    global picoharp
    global PicoHarp_mode
    global qmi_context
    global PICOHARP
    qmi_context = QMI_Context("pico_measurements")
    qmi_context.start()

    picoharp = qmi_context.make_instrument('ph300', PicoQuant_PicoHarp300, '1013201')

    picoharp.open()
    picoharp.initialize('HIST')
    PicoHarp_mode='HIST'
    
    time.sleep(0.1)
    picoharp.set_binning(0)
    
    #all other buttons' commands
    Switch_PH_mode_button.configure(command=change_PicoHarp_mode)
    Change_Resol_button.configure(command=set_resol)
    Syncdiv_button.configure(command=set_syncdiv)
    # Ch0_CFD_button.configure(command=Nothing)
    # Ch1_CFD_button.configure(command=Nothing)
    PicoHarp_start_button.configure(command=Nothing)
    
    tab_PicoHarp.after(1000,read_channels_count_rate)
    
def change_PicoHarp_mode():
    global PicoHarp_mode
    if PicoHarp_mode=='HIST':
        picoharp.initialize('T2')
        PicoHarp_mode='T2'
    elif PicoHarp_mode=='T2':
        picoharp.initialize('T3')
        PicoHarp_mode='T3'
    else:
        picoharp.initialize('HIST')
        PicoHarp_mode='HIST'
    PH_mode_label.config(text=PicoHarp_mode)
        
def read_channels_count_rate():
    Ch0_rate=picoharp.get_count_rate(0)
    Ch1_rate=picoharp.get_count_rate(1)
    
    Ch0_count_label2.config(text=str(Ch0_rate))
    Ch1_count_label2.config(text=str(Ch1_rate))
    #wait at least 100ms for a new measurement
    if picoharp.is_open():
        tab_PicoHarp.after(1000,read_channels_count_rate)
    
def read_measurement_time():
    elapsed_time=int(picoharp.get_elapsed_measurement_time()/1000)
    
    Time_elapsed_label1.config(text=str(elapsed_time))
    
    if picoharp.is_open():
        tab_PicoHarp.after(1000,read_measurement_time)
    
# def set_ch0_CFD():
#     CFD=int(Ch0_CFD_entry.get())
#     picoharp.set_input_cfd(0,CFD,10)
    
# def set_ch1_CFD():
#     CFD=int(Ch1_CFD_entry.get())
#     picoharp.set_input_cfd(1,CFD,10)
    
def set_syncdiv():
    Syncdiv=int(Syncdiv_label1.get())
    picoharp.set_sync_divider(Syncdiv)
        
def set_resol():
    #0= 1 x base res (base res = 10MHz)
    #1= 2 x base res
    #2= 4 x base res
    #3= 8 x base res
    Resol=int(Resol_label1.get())
    picoharp.set_binning(Resol)
    
def close_PicoHarp():
    try:
        picoharp.close()
        qmi_context.stop()
    except:
        a=1
    Ch0_count_label2.configure(text='N/A')
    Ch1_count_label2.configure(text='N/A')
    #all other buttons' commands
    Switch_PH_mode_button.configure(command=Nothing)
    Change_Resol_button.configure(command=Nothing)
    Syncdiv_button.configure(command=Nothing)
    # Ch0_CFD_button.configure(command=Nothing)
    # Ch1_CFD_button.configure(command=Nothing)
    
def Spectre_APD():
    #works for signal on ch1 only !!!
    Wavelength_start=int(WL_Start_APD_entry.get())
    Wavelength_end=int(WL_End_APD_entry.get())
    Increment=int(Increment_APD_entry.get())
    
    Wavelengths=np.arange(Wavelength_start,Wavelength_end,Increment)
    
    #preventing moving the spectro too far out of its working area, preventing starting measurements that are very long
    if Wavelength_start<650 or Wavelength_start>1000:
        return
    if Wavelength_end<650 or Wavelength_end>1000:
        return
    if len(Wavelengths)<1 or len(Wavelengths)>100:
        return
    #move spectro to WL_start
    spec.set_wavelength(Wavelength_start)
    time.sleep(30) #wait in case grating moves in descending wavelengths
    
    Means=np.zeros(len(Wavelengths))
    for i in range(len(Wavelengths)):
        spec.set_wavelength(Wavelengths[i])
        Mean=0
        for j in range(10):
            Mean+=picoharp.get_count_rate(1)/10
            time.sleep(0.15)
        Means[i]=Mean
    ax2.plot(Wavelengths,Means,color='black')
    ax2.set_xlabel('Wavelength (nm)')
    ax2.set_ylabel('Countrate')
    Spectre_WL_APD.draw()

    
#%%Create tkinter window and styling for buttons, labels and entries

root = ThemedTk(theme='equilux')
root.geometry("1900x1000+0+0")
root.configure(background="#464646")

font=('American typewriter', 12)


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

notebook_frame = ttk.Frame(root)
notebook_frame.grid(row=0,column=1,sticky='nsew')

notebook = ttk.Notebook(notebook_frame)

tab_plot=ttk.Frame(notebook)
tab_PLE=ttk.Frame(notebook)
tab_FTIR=ttk.Frame(notebook)
tab_PicoHarp=ttk.Frame(notebook)

notebook.add(tab_plot, text='Spectrum')
notebook.add(tab_PLE, text='PLE (M2)')
notebook.add(tab_FTIR,text='FTIR')
notebook.add(tab_PicoHarp,text='PicoHarp')

notebook.pack(expand=1, fill='both')

persistent_frame= ttk.Frame(root)
persistent_frame.grid(row=0,column=0,sticky='nsew')

cam_started=False
flipper_state='direct'

#%% Buttons and stuff

#Start cam and spectro
Start_Andor_button = ttk.Button(persistent_frame,text="Start cam")
Start_Andor_button.grid(row=3,column=1,padx=0,pady=0)

#Set cooler temperature
T_label = ttk.Label(persistent_frame,text="T wanted (C)", font=font)
T_label.grid(row=0, column=0, padx=0, pady=0)
T_entry = ttk.Entry(persistent_frame, width = 5, font=font)
T_entry.insert(0, "-90")
T_entry.grid(row=1, column=0, padx=0, pady=0)

#Button to set temperature
cooler_button= ttk.Button(persistent_frame,text='Set temperature', style='W.TButton')
cooler_button.grid(row=2,column=0, padx=0, pady=0)

#Button to start acq
StartAcq_button= ttk.Button(persistent_frame,text='Start', style='W.TButton')
StartAcq_button.grid(row=35, column=0, padx=0, pady=0)

#Button to stop acq
StopAcq_button= ttk.Button(persistent_frame,text='Stop')
StopAcq_button.grid(row=36, column=0, padx=0, pady=0)

#Button to set WL
WL_button= ttk.Button(persistent_frame,text='Set WL')
WL_button.grid(row=6,column=0, padx=0, pady=0)

#output T box
T_measure_button = ttk.Button(persistent_frame,text='Measure T', style='W.TButton')
T_measure_button.grid(row=2, column=1, padx=0, pady=0)
T_read_label = ttk.Label(persistent_frame, text='T_camera= N/A', font=font)
T_read_label.grid(row=1, column=1, padx=0, pady=0)

#Set spectro WL
WL_label = ttk.Label(persistent_frame,text="WL wanted (nm)", font=font)
WL_label.grid(row=4, column=0, padx=0, pady=0)
WL_entry = ttk.Entry(persistent_frame, width = 10, font=font)
WL_entry.insert(0, "720")
WL_entry.grid(row=5, column=0, padx=0, pady=0)

#output WL box
WL_measure_button = ttk.Button(persistent_frame,text='Measure WL')
WL_measure_button.grid(row=6, column=1, padx=0, pady=0)
WL_read_label = ttk.Label(persistent_frame, text='Central WL= N/A', font=font)
WL_read_label.grid(row=5, column=1, padx=0, pady=0)

#Exposure entry
Exposure_label = ttk.Label(persistent_frame,text="Exposure (s)", font=font)
Exposure_label.grid(row=8, column=0, padx=0, pady=0)
Exposure_entry = ttk.Entry(persistent_frame, width=5, font=font)
Exposure_entry.insert(0, "1")
Exposure_entry.grid(row=9, column=0, padx=0, pady=0)

#Flip output port
Flipper_port_button=ttk.Button(persistent_frame,text='Flip port')
Flipper_port_button.grid(row=9, column=2, padx=0, pady=0)
Flipper_port_label=ttk.Label(persistent_frame,text='direct',font=font)
Flipper_port_label.grid(row=10,column=2, padx=0, pady=0)

#Long spectrum button
Long_spectrum_button = ttk.Button(persistent_frame,text='Wide spectrum', style='W.TButton')
Long_spectrum_button.grid(row=37,column=3, padx=0, pady=0)
WL_start_label = ttk.Label(persistent_frame,text='Start WL', font=font)
WL_start_label.grid(row=35,column=2, padx=0, pady=0)
WL_start_entry = ttk.Entry(persistent_frame,width=5, font=font)
WL_start_entry.insert(0,"900")
WL_start_entry.grid(row=36,column=2, padx=0, pady=0)
WL_stop_label = ttk.Label(persistent_frame,text='Stop WL', font=font)
WL_stop_label.grid(row=35,column=3, padx=0, pady=0)
WL_stop_entry = ttk.Entry(persistent_frame,width=5, font=font)
WL_stop_entry.insert(0,"940")
WL_stop_entry.grid(row=36,column=3, padx=0, pady=0)

#Load switch images
switch_off_image = tk.PhotoImage(file='C:/Users/Manip/Pictures/Saved Pictures/switch left dark bg small.png',master=persistent_frame)
switch_on_image = tk.PhotoImage(file='C:/Users/Manip/Pictures/Saved Pictures/switch right dark bg small.png',master=persistent_frame)

#Switch between plot in nm and eV
switch_var=tk.BooleanVar()
switch_nm_eV=ttk.Checkbutton(persistent_frame,variable=switch_var,command=switch_changed,image=switch_off_image,style='Switch.TButton',compound='center',width=1)
switch_nm_eV.grid(row=15,column=1, padx=0, pady=0)
nm_label =ttk.Label(persistent_frame,text='nm', font=font)
nm_label.grid(row=15, column=0, padx=0, pady=0, sticky='e')
eV_label =ttk.Label(persistent_frame,text='eV', font=font)
eV_label.grid(row=15, column=2, padx=0, pady=0, sticky='w')

#Style
style = ttk.Style()
style.configure('Switch.TButton',background="#464646",borderwidth=0)

#X and Y min and max labels and input boxes
XminWL_label = ttk.Label(tab_plot,text="Xmin (nm)", font=font)
XminWL_label.grid(row=81, column=10, padx=0, pady=0)
XminWL_entry = ttk.Entry(tab_plot, width=5, font=font)
XminWL_entry.insert(0, "700")
XminWL_entry.grid(row=82, column=10, padx=0, pady=0)

XmaxWL_label = ttk.Label(tab_plot,text="Xmax (nm)", font=font)
XmaxWL_label.grid(row=81, column=11, padx=0, pady=0)
XmaxWL_entry = ttk.Entry(tab_plot, width=5, font=font)
XmaxWL_entry.insert(0, "740")
XmaxWL_entry.grid(row=82, column=11, padx=0, pady=0)

XminE_label = ttk.Label(tab_plot,text="Xmin (eV)", font=font)
XminE_label.grid(row=81, column=12, padx=0, pady=0)
XminE_entry = ttk.Entry(tab_plot, width=5, font=font)
XminE_entry.insert(0, "1.68")
XminE_entry.grid(row=82, column=12, padx=0, pady=0)

XmaxE_label = ttk.Label(tab_plot,text="Xmax (eV)", font=font)
XmaxE_label.grid(row=81, column=13, padx=0, pady=0)
XmaxE_entry = ttk.Entry(tab_plot, width=5, font=font)
XmaxE_entry.insert(0, "1.78")
XmaxE_entry.grid(row=82, column=13, padx=0, pady=0)

Ymin_label = ttk.Label(tab_plot,text="Ymin", font=font)
Ymin_label.grid(row=81, column=14, padx=0, pady=0)
Ymin_entry = ttk.Entry(tab_plot, width=5, font=font)
Ymin_entry.insert(0, "750")
Ymin_entry.grid(row=82, column=14, padx=0, pady=0)

Ymax_label = ttk.Label(tab_plot,text="Ymax", font=font)
Ymax_label.grid(row=81, column=15, padx=0, pady=0)
Ymax_entry = ttk.Entry(tab_plot, width=5, font=font)
Ymax_entry.insert(0, "1100")
Ymax_entry.grid(row=82, column=15, padx=0, pady=0)

#%%Keithley related
turn_on_keithley_button = ttk.Button(persistent_frame,text='Turn Keithley ON',command=Turn_Keithley_Link_On)
turn_on_keithley_button.grid(row=60, column=1, padx=0, pady=0)

#Set V
set_tension_button = ttk.Button(persistent_frame,text='Set V')
set_tension_button.grid(row=63, column=2, padx=0, pady=0)
V_label = ttk.Label(persistent_frame,text="Tension (V)", font=font)
V_label.grid(row=61, column=1, padx=0, pady=0)
V_entry = ttk.Entry(persistent_frame, width=5, font=font)
V_entry.insert(0, "0")
V_entry.grid(row=62, column=1, padx=0, pady=0)
I_read_label = ttk.Label(persistent_frame, text='I= N/A', font=font)
I_read_label.grid(row=62, column=2, padx=0, pady=0)

#Tension series
check_current_evolution_button = ttk.Button(persistent_frame,text="Start I(t)",style='W.TButton')
check_current_evolution_button.grid(row=66, column=1, padx=0, pady=0)
tension_series_button = ttk.Button(persistent_frame,text='Start V series',command=Tension_series,style='W.TButton')
tension_series_button.grid(row=66, column=2, padx=0, pady=0)
E_V_button = ttk.Button(persistent_frame,text='Start E-V map',command=Map_Tension_Energy,style='W.TButton')
E_V_button.grid(row=66, column=3, padx=0, pady=0)
V_start_label = ttk.Label(persistent_frame,text="Start (V)", font=font)
V_start_label.grid(row=64, column=1, padx=0, pady=0)
V_start_entry = ttk.Entry(persistent_frame, width=5, font=font)
V_start_entry.insert(0, "0")
V_start_entry.grid(row=65, column=1, padx=0, pady=0)
V_finish_label = ttk.Label(persistent_frame,text="Finish (V)", font=font)
V_finish_label.grid(row=64, column=2, padx=0, pady=0)
V_finish_entry = ttk.Entry(persistent_frame, width=5, font=font)
V_finish_entry.insert(0, "0")
V_finish_entry.grid(row=65, column=2, padx=0, pady=0)
Step_label = ttk.Label(persistent_frame,text="Step", font=font)
Step_label.grid(row=64, column=3, padx=0, pady=0)
Step_entry = ttk.Entry(persistent_frame, width=5, font=font)
Step_entry.insert(0, "0.05")
Step_entry.grid(row=65, column=3, padx=0, pady=0)

#%%Power control
Power_button = ttk.Button(persistent_frame,text='Start PM100',command=start_PM)
Power_button.grid(row=0,column=3,padx=0,pady=0)
Power_value = ttk.Label(persistent_frame,text='Power = N/A', font=font)
Power_value.grid(row=1, column=3, padx=0, pady=0)
Power_wanted_entry = ttk.Entry(persistent_frame, width=5, font=font)
Power_wanted_entry.insert(0, "50")
Power_wanted_entry.grid(row=5, column=3, padx=0, pady=0)
Get_min_max_powers_button = ttk.Button(persistent_frame, text='Start Power Control',command=start_Power_Control)
Get_min_max_powers_button.grid(row=0, column=4, padx=0, pady=0)
Go_To_Power_button = ttk.Button(persistent_frame, text="Go to power (uw)")
Go_To_Power_button.grid(row=4, column=3,padx=0,pady=0)
min_power_label=ttk.Label(persistent_frame,text="Min power = N/A", font=font)
min_power_label.grid(row=1, column=4, padx=0, pady=0)
max_power_label=ttk.Label(persistent_frame,text="Max power = N/A", font=font)
max_power_label.grid(row=2, column=4, padx=0, pady=0)
PM_exp_entry = ttk.Entry(persistent_frame,width=5, font=font)
PM_exp_entry.insert(0, "100")
PM_exp_entry.grid(row=3,column=3,padx=0,pady=0)
PM_exp_label =ttk.Label(persistent_frame,text="PM refresh (ms)", font=font)
PM_exp_label.grid(row=2, column=3, padx=0, pady=0)
Start_Power_Series_button = ttk.Button(persistent_frame,text="Power series",style='W.TButton')
Start_Power_Series_button.grid(row=5,column=4,padx=0,pady=0)
Nb_Power_Steps_Label = ttk.Label(persistent_frame,text='Nb power steps', font=font)
Nb_Power_Steps_Label.grid(row=3,column=4,padx=0,pady=0)
Nb_Power_Steps_Entry = ttk.Entry(persistent_frame,width=5, font=font)
Nb_Power_Steps_Entry.insert(0, "2")
Nb_Power_Steps_Entry.grid(row=4,column=4,padx=0,pady=0)
Power_Series_Label = ttk.Label(persistent_frame,text='Power series name', font=font)
Power_Series_Label.grid(row=6,column=3,padx=0,pady=0)
Power_Series_Entry = ttk.Entry(persistent_frame,width=40, font=font)
Power_Series_Entry.grid(row=7,column=3,columnspan=3,padx=0,pady=0)

#%%Polar Control
Start_Polar_Control_button = ttk.Button(persistent_frame,text='Start Polar Control',command=start_Polar_Control)
Start_Polar_Control_button.grid(row=12,column=3,padx=0,pady=0)
Angle_Wanted_Label= ttk.Label(persistent_frame,text="Angle wanted (deg)",font=font)
Angle_Wanted_Label.grid(row=13,column=3,padx=0,pady=0)
Angle_Wanted_Entry=ttk.Entry(persistent_frame,width=5,font=font)
Angle_Wanted_Entry.insert(0,"0")
Angle_Wanted_Entry.grid(row=14,column=3,padx=0,pady=0)
Start_Polar_Series_button = ttk.Button(persistent_frame,text='Start Polar Series',style='W.TButton')
Start_Polar_Series_button.grid(row=14,column=4,padx=0,pady=0)

Nb_Polar_Steps_Label= ttk.Label(persistent_frame,text="Nb polar steps",font=font)
Nb_Polar_Steps_Label.grid(row=12,column=4,padx=0,pady=0)
Nb_Polar_Steps_Entry= ttk.Entry(persistent_frame,width=5, font=font)
Nb_Polar_Steps_Entry.insert(0,"10")
Nb_Polar_Steps_Entry.grid(row=13,column=4,padx=0,pady=0)
Polar_Series_Label= ttk.Label(persistent_frame,text='Polar series name', font=font)
Polar_Series_Label.grid(row=15,column=3,padx=0,pady=0)
Polar_Series_Entry= ttk.Entry(persistent_frame,width=40, font=font)
Polar_Series_Entry.grid(row=16,column=3,columnspan=2,padx=0,pady=0)

#%%Saving part

IV_label = ttk.Label(persistent_frame,text="IV name",font=font)
IV_label.grid(row=67, column=0, padx=0, pady=0)
IV_entry = ttk.Entry(persistent_frame, width=50,font=font)
IV_entry.grid(row=68, column=0, columnspan=3, padx=0, pady=0)

TwoD_label = ttk.Label(persistent_frame,text="2D map name",font=font)
TwoD_label.grid(row=69, column=0, padx=0, pady=0)
TwoD_entry = ttk.Entry(persistent_frame, width=50,font=font)
TwoD_entry.grid(row=70, column=0, columnspan=3, padx=0, pady=0)

ExcWL_label = ttk.Label(persistent_frame,text="Excitation WL (nm)",font=font)
ExcWL_label.grid(row=8, column=1, padx=0, pady=0)
ExcWL_entry = ttk.Entry(persistent_frame, width=5,font=font)
ExcWL_entry.insert(0, "532")
ExcWL_entry.grid(row=9, column=1, padx=0, pady=0)

Directory_label = ttk.Label(persistent_frame,text="Directory path",font=font)
Directory_label.grid(row=30, column=0, padx=0, pady=0)
Directory_entry = ttk.Entry(persistent_frame, width=75,font=font)
Directory_entry.insert(0, "C:/Users/manip pico/Documents/"+today_date_str)
Directory_entry.grid(row=31, column=0, columnspan=4, padx=0, pady=0)

browse_button = ttk.Button(persistent_frame, text='Browse', command=browse_directory)
browse_button.grid(row=31,column=4, padx=0, pady=0, sticky='w')

savefile_button = ttk.Button(persistent_frame, text='Save single file')
savefile_button.grid(row=32,column=2, padx=0, pady=0)

#%%PLE related

Start_M2_button=ttk.Button(tab_PLE,text="Link M2",command=turn_on_M2)
Start_M2_button.grid(row=0,column=0,padx=0,pady=0)

Set_M2_WL_Label=ttk.Label(tab_PLE,text="Wanted M2 WL (nm)",font=font)
Set_M2_WL_Label.grid(row=0,column=1,padx=0,pady=0)

Set_M2_WL_Entry=ttk.Entry(tab_PLE,width=5,font=font)
Set_M2_WL_Entry.insert(0, "710")
Set_M2_WL_Entry.grid(row=1,column=1,padx=0,pady=0)

Set_M2_WL_Button=ttk.Button(tab_PLE,text="Set M2 WL")
Set_M2_WL_Button.grid(row=2,column=1,padx=0,pady=0)

Real_M2_WL_title=ttk.Label(tab_PLE,text="Real M2 WL",font=font)
Real_M2_WL_title.grid(row=0,column=2,padx=0,pady=0)

Real_M2_WL_Label=ttk.Label(tab_PLE,text="N/A nm",font=font)
Real_M2_WL_Label.grid(row=1,column=2,padx=0,pady=0)

PLE_Series_Entry=ttk.Entry(tab_PLE,width=50,font=font)
PLE_Series_Entry.grid(row=3,column=0,columnspan=3,padx=0,pady=0)

PLE_begin_WL_Label=ttk.Label(tab_PLE,text='Start WL (nm)',font=font)
PLE_begin_WL_Label.grid(row=5,column=0,padx=0,pady=0)
PLE_begin_WL_Entry=ttk.Entry(tab_PLE,width=5,font=font)
PLE_begin_WL_Entry.insert(0,'705')
PLE_begin_WL_Entry.grid(row=6,column=0,padx=0,pady=0)

PLE_end_WL_Label=ttk.Label(tab_PLE,text='Finish WL (nm)',font=font)
PLE_end_WL_Label.grid(row=5,column=1,padx=0,pady=0)
PLE_end_WL_Entry=ttk.Entry(tab_PLE,width=5,font=font)
PLE_end_WL_Entry.insert(0,'720')
PLE_end_WL_Entry.grid(row=6,column=1,padx=0,pady=0)

PLE_step_Label=ttk.Label(tab_PLE,text='WL step (nm)',font=font)
PLE_step_Label.grid(row=5,column=2,padx=0,pady=0)
PLE_step_Entry=ttk.Entry(tab_PLE,width=5,font=font)
PLE_step_Entry.insert(0,'0.25')
PLE_step_Entry.grid(row=6,column=2,padx=0,pady=0)

PLE_start_button=ttk.Button(tab_PLE,text="Start PLE")
PLE_start_button.grid(row=7,column=2,padx=0,pady=0)

#%%PicoHarp related
PicoHarp_start_button=ttk.Button(tab_PicoHarp,text='Start PicoHarp',command=turn_on_PicoHarp)
PicoHarp_start_button.grid(row=0, column=0, padx=0, pady=0)

Switch_PH_mode_button=ttk.Button(tab_PicoHarp,text='Switch Mode')
Switch_PH_mode_button.grid(row=1, column=0, padx=0, pady=0)
PH_mode_label=ttk.Label(tab_PicoHarp, text='HIST',font=font)
PH_mode_label.grid(row=2,column=0,padx=0,pady=0)

Ch0_count_label1=ttk.Label(tab_PicoHarp,text='Ch. 0',font=font)
Ch0_count_label1.grid(row=2,column=4,padx=0,pady=0)
Ch0_count_label2=ttk.Label(tab_PicoHarp,text='N/A',font=font)
Ch0_count_label2.grid(row=3,column=4,padx=0,pady=0)

Ch1_count_label1=ttk.Label(tab_PicoHarp,text='Ch. 1',font=font)
Ch1_count_label1.grid(row=2,column=6,padx=0,pady=0)
Ch1_count_label2=ttk.Label(tab_PicoHarp,text='N/A',font=font)
Ch1_count_label2.grid(row=3,column=6,padx=0,pady=0)

Time_elapsed_label0=ttk.Label(tab_PicoHarp,text='Elapsed time (s)',font=font)
Time_elapsed_label0.grid(row=21,column=0,padx=0,pady=0)
Time_elapsed_label1=ttk.Label(tab_PicoHarp,text='0',font=font)
Time_elapsed_label1.grid(row=22,column=0,padx=0,pady=0)

Syncdiv_label0=ttk.Label(tab_PicoHarp,text='Syncdiv',font=font)
Syncdiv_label0.grid(row=15,column=0,padx=0,pady=0)
Syncdiv_label1=ttk.Combobox(tab_PicoHarp,state='readonly',values=['1','2','4','8'],width=5)
Syncdiv_label1.set('1')
Syncdiv_label1.grid(row=16,column=0,padx=0,pady=0)
Syncdiv_button=ttk.Button(tab_PicoHarp,text='Set Syncdiv')
Syncdiv_button.grid(row=17,column=0,padx=0,pady=0)

# Ch0_CFD_label0=ttk.Label(tab_PicoHarp,text='Ch0 CFD (mV)',font=font)
# Ch0_CFD_label0.grid(row=18,column=0,padx=0,pady=0)
# Ch0_CFD_entry=ttk.Spinbox(tab_PicoHarp,width=5,from_=0,to=800,increment=10,font=font)
# Ch0_CFD_entry.insert(0,'100') #between 0 and 800mV
# Ch0_CFD_entry.grid(row=19,column=0,padx=0,pady=0)
# Ch0_CFD_button=ttk.Button(tab_PicoHarp,text='Set Ch0 CFD')
# Ch0_CFD_button.grid(row=20,column=0,padx=0,pady=0)

# Ch1_CFD_label0=ttk.Label(tab_PicoHarp,text='Ch1 CFD (mV)',font=font)
# Ch1_CFD_label0.grid(row=18,column=1,padx=0,pady=0)
# Ch1_CFD_entry=ttk.Spinbox(tab_PicoHarp,width=5,from_=0,to=800,increment=10,font=font)
# Ch1_CFD_entry.insert(0,'100') #between 0 and 800mV
# Ch1_CFD_entry.grid(row=19,column=1,padx=0,pady=0)
# Ch1_CFD_button=ttk.Button(tab_PicoHarp,text='Set Ch1 CFD')
# Ch1_CFD_button.grid(row=20,column=1,padx=0,pady=0)

Resol_label0=ttk.Label(tab_PicoHarp,text='Resolution',font=font)
Resol_label0.grid(row=5,column=4,padx=0,pady=0)
Resol_label1=ttk.Combobox(tab_PicoHarp,state='readonly',values=['0','1','2','3'],width=5)
Resol_label1.set('0')
Resol_label1.grid(row=6,column=4,padx=0,pady=0)
Change_Resol_button=ttk.Button(tab_PicoHarp,text='Change Res')
Change_Resol_button.grid(row=7,column=4,padx=0,pady=0)

WL_Start_APD_label=ttk.Label(tab_PicoHarp,text='WL start (nm)')
WL_Start_APD_label.grid(row=10,column=5,padx=0,pady=0)
WL_Start_APD_entry=ttk.Entry(tab_PicoHarp,width=5,font=font)
WL_Start_APD_entry.insert(0,'900')
WL_Start_APD_entry.grid(row=11,column=5,padx=0,pady=0)

WL_End_APD_label=ttk.Label(tab_PicoHarp,text='WL end (nm)')
WL_End_APD_label.grid(row=10,column=6,padx=0,pady=0)
WL_End_APD_entry=ttk.Entry(tab_PicoHarp,width=5,font=font)
WL_End_APD_entry.insert(0,'901')
WL_End_APD_entry.grid(row=11,column=6,padx=0,pady=0)

Increment_APD_label=ttk.Label(tab_PicoHarp,text='Step (nm)')
Increment_APD_label.grid(row=10,column=7,padx=0,pady=0)
Increment_APD_entry=ttk.Entry(tab_PicoHarp,width=5,font=font)
Increment_APD_entry.insert(0,'1')
Increment_APD_entry.grid(row=11,column=7,padx=0,pady=0)

Start_spectrum_WL_APD_button=ttk.Button(tab_PicoHarp,text='Start spectrum')
Start_spectrum_WL_APD_button.grid(row=12,column=7,padx=0,pady=0)

fig2,ax2=plt.subplots(figsize=(10,5))
Spectre_WL_APD=FigureCanvasTkAgg(fig2,tab_PicoHarp)
Spectre_WL_APD.get_tk_widget().grid(row=100,column=0,rowspan=10,columnspan=10,padx=5,pady=5)

# Close_PicoHarp_button=ttk.Button(tab_PicoHarp,text='Close PicoHarp',command=close_PicoHarp)
# Close_PicoHarp_button.grid(row=0,column=1,padx=0,pady=0)

#%%Cursors and plot

class CursorPlot:
    def __init__(self, master):
        global C, exposure
        self.master = master
        #self.master.title("Cursor Plot")

        # Create a figure and axis for plotting
        self.fig, self.ax = plt.subplots(figsize=(10,8))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.master)
        self.canvas.get_tk_widget().grid(row=0, column=10, rowspan=80, columnspan=10, padx=5, pady=5)
        
        self.updateXY_button = ttk.Button(tab_plot,text='Update x and y limits', command=self.update_xylim)
        self.updateXY_button.grid(row=83,column=10, padx=0, pady=0)
        
        # Initialize cursor positions
        self.cursor_x_1WL = 720
        self.cursor_x_1E = 1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_1WL
        self.cursor_y_1 = 0
        self.cursor_x_2WL = 720
        self.cursor_x_2E = 1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_2WL
        self.cursor_y_2 = 0
        self.cursor_marker_y = 1100

        #Switch between autoscale and set manual limits
        self.switch_autoscale_var=tk.BooleanVar()
        self.switch_autoscale=ttk.Checkbutton(tab_plot,variable=self.switch_autoscale_var,command=self.switch_autoscale_changed,image=switch_off_image,style='Switch.TButton',compound='center',width=1)
        self.switch_autoscale.grid(row=83,column=12, padx=0, pady=0)
        self.nm_label =ttk.Label(tab_plot,text='Autoscale ON',font=font)
        self.nm_label.grid(row=83, column=11, padx=0, pady=0, sticky='e')
        self.eV_label =ttk.Label(tab_plot,text='OFF',font=font)
        self.eV_label.grid(row=83, column=13, padx=0, pady=0, sticky='w')
        
        # Set the limits at first
        self.xminWL_manual=float(XminWL_entry.get())
        self.xmaxWL_manual=float(XmaxWL_entry.get())
        self.xminE_manual=float(XminE_entry.get())
        self.xmaxE_manual=float(XmaxE_entry.get())
        self.ymin_manual=float(Ymin_entry.get())
        self.ymax_manual=float(Ymax_entry.get())
        
        self.xminWL=float(XminWL_entry.get())
        self.xmaxWL=float(XmaxWL_entry.get())
        self.xminE=float(XminE_entry.get())
        self.xmaxE=float(XmaxE_entry.get())
        self.ymin=float(Ymin_entry.get())
        self.ymax=float(Ymax_entry.get())
        
        # Cursor colors
        self.color1='palevioletred'
        self.color2='aqua'
        self.color_marker='yellow'
        
        # Add vertical cursor lines
        self.cursor_line_x_1, = self.ax.plot([self.cursor_x_1WL, self.cursor_x_1WL], [self.ymin, self.ymax], color=self.color1)
        self.cursor_line_x_2, = self.ax.plot([self.cursor_x_2WL, self.cursor_x_2WL], [self.ymin, self.ymax], color=self.color2)

        # Add horizontal cursor lines
        self.cursor_line_y_1, = self.ax.plot([self.xminWL,self.xmaxWL], [self.cursor_y_1, self.cursor_y_1], color=self.color1)
        self.cursor_line_y_2, = self.ax.plot([self.xminWL,self.xmaxWL], [self.cursor_y_2, self.cursor_y_2], color=self.color2)

        # Add on horizontal line for a marker
        self.cursor_line_marker_y, = self.ax.plot([self.xminWL,self.xmaxWL], [self.cursor_marker_y, self.cursor_marker_y], color=self.color_marker)
        
        # Create a label to show cursor positions
        self.cursor_label_1 = ttk.Label(master=self.master, text="",font=font)
        self.cursor_label_1.grid(row=81, column=16, padx=0, pady=0)
        self.cursor_label_2 = ttk.Label(master=self.master, text="",font=font)
        self.cursor_label_2.grid(row=82, column=16, padx=0, pady=0)
        self.cursor_label_delta_x = ttk.Label(master=self.master, text="",font=font)
        self.cursor_label_delta_x.grid(row=83, column=16, padx=0, pady=0)
        self.cursor_label_delta_y = ttk.Label(master=self.master, text="",font=font)
        self.cursor_label_delta_y.grid(row=84, column=16, padx=0, pady=0)
            
        
        self.cursor_center_button = ttk.Button(master=self.master,text='Center markers',command=self.center_cursors)
        self.cursor_center_button.grid(row=83,column=15,padx=0,pady=0)

        # Bind mouse events
        self.fig.canvas.mpl_connect('button_press_event', self.on_press)
        self.fig.canvas.mpl_connect('button_release_event', self.on_release)
        self.fig.canvas.mpl_connect('motion_notify_event', self.on_motion)

        # Cursor dragging flags
        self.dragging_cursor_1 = False
        self.dragging_cursor_2 = False
        self.dragging_cursor_marker = False

        # Cursor drag threshold (adjust as needed)
        self.xspanWL=abs(self.xmaxWL-self.xminWL)
        self.xspanE=abs(self.xmaxE-self.xminE)
        self.yspan=abs(self.ymax-self.ymin)
        
        self.drag_thresholdWL = self.xspanWL/400
        self.drag_thresholdE = self.xspanE/400
        self.drag_threshold_y = self.yspan/400
        
        C=np.zeros((1,2000))

        
        self.update_plot()
        
    def center_cursors(self):
        self.x_middleWL=(self.xmaxWL+self.xminWL)/2
        self.x_middleE=(self.xmaxE+self.xminE)/2
        
        self.cursor_x_1WL=self.x_middleWL
        self.cursor_x_2WL=self.x_middleWL
        self.cursor_x_1E=self.x_middleE
        self.cursor_x_2E=self.x_middleE
    
    def update_plot(self):
        global E,WL,C
        A=cam.read_newest_image(peek=True)

        if np.shape(A)==(1,2000):
            C=A
            
        #plots a zero line when A=None
        getWL=spec.get_calibration()
        WL=getWL*1e9
        E=6.63e-34*3e8/1.6e-19/getWL
        
        self.ax.clear()
        if self.switch_autoscale_var.get()==False:
            self.ymin=np.min(C[0])
            self.ymax=np.max(C[0])+1 #to make sure that ymax>ymin and its never both 0
            self.xminE=E[-1]
            self.xmaxE=E[0]
            self.xminWL=WL[0]
            self.xmaxWL=WL[-1]
        else:
            self.xminWL=self.xminWL_manual
            self.xmaxWL=self.xmaxWL_manual
            self.xminE=self.xminE_manual
            self.xmaxE=self.xmaxE_manual
            self.ymin=self.ymin_manual
            self.ymax=self.ymax_manual
            
        if switch_var.get():
            self.cursor_x_1=self.cursor_x_1E
            self.cursor_x_2=self.cursor_x_2E
            self.xmin=self.xminE
            self.xmax=self.xmaxE
            self.xlabel='Energy (eV)'
            self.X=E
        else:
            self.cursor_x_1=self.cursor_x_1WL
            self.cursor_x_2=self.cursor_x_2WL
            self.xmin=self.xminWL
            self.xmax=self.xmaxWL
            self.xlabel='Wavelength (nm)'
            self.X=WL
  

        self.ax.plot(self.X,C[0],color='darkorange')
        self.ax.set_xlabel(self.xlabel)
        self.ax.set_ylabel('Counts')
        self.ax.set_xlim([self.xmin,self.xmax])
        self.ax.set_ylim([self.ymin,self.ymax])


        self.update_cursors()
        self.redraw_cursors()
        self.ax.set_title('')
        
        self.ax.grid(True)
        self.canvas.draw()
            
        self.master.after(100,self.update_plot)

    def on_press(self, event):
        if event.inaxes == self.ax:
            x_clicked = event.xdata
            y_clicked = event.ydata
        if switch_var.get():
            # Check if the click is within the drag threshold of either cursor
            if abs(x_clicked - self.cursor_x_1E) < self.drag_thresholdE:
                self.dragging_cursor_1 = True
            elif abs(x_clicked - self.cursor_x_2E) < self.drag_thresholdE:
                self.dragging_cursor_2 = True
            elif abs(y_clicked - self.cursor_marker_y) < self.drag_threshold_y:
                self.dragging_cursor_marker = True
        else:
            # Check if the click is within the drag threshold of either cursor
            if abs(x_clicked - self.cursor_x_1WL) < self.drag_thresholdWL:
                self.dragging_cursor_1 = True
            elif abs(x_clicked - self.cursor_x_2WL) < self.drag_thresholdWL:
                self.dragging_cursor_2 = True
            elif abs(y_clicked - self.cursor_marker_y) < self.drag_threshold_y:
                self.dragging_cursor_marker = True

        # Update cursor positions
        self.update_cursors()
        self.redraw_cursors()
        self.fig.canvas.draw_idle()

    def on_release(self, event):
        self.dragging_cursor_1 = False
        self.dragging_cursor_2 = False
        
        self.dragging_cursor_marker = False

    def on_motion(self, event):
        if event.inaxes == self.ax:
            if switch_var.get():
                if self.dragging_cursor_1:
                    self.cursor_x_1E = event.xdata
                    self.cursor_x_1WL = 1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_1E
                    self.update_cursors()
                    self.redraw_cursors()
                elif self.dragging_cursor_2:
                    self.cursor_x_2E = event.xdata
                    self.cursor_x_2WL = 1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_2E
                    self.update_cursors()
                    self.redraw_cursors()
                elif self.dragging_cursor_marker:
                    self.cursor_marker_y = event.ydata
                    self.update_cursors()
                    self.redraw_cursors()
            else:
                if self.dragging_cursor_1:
                    self.cursor_x_1WL = event.xdata
                    self.cursor_x_1E = 1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_1WL
                    self.update_cursors()
                    self.redraw_cursors()
                elif self.dragging_cursor_2:
                    self.cursor_x_2WL = event.xdata
                    self.cursor_x_2E = 1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_2WL
                    self.update_cursors()
                    self.redraw_cursors()
                elif self.dragging_cursor_marker:
                    self.cursor_marker_y = event.ydata
                    self.update_cursors()
                    self.redraw_cursors()
                
    def redraw_cursors(self):
        if self.switch_autoscale_var.get()==False:
            self.ymin=np.min(C[0])
            self.ymax=np.max(C[0])
            self.xminE=E[-1]
            self.xmaxE=E[0]
            self.xminWL=WL[0]
            self.xmaxWL=WL[-1]
        else:
            self.xminWL=self.xminWL_manual
            self.xmaxWL=self.xmaxWL_manual
            self.xminE=self.xminE_manual
            self.xmaxE=self.xmaxE_manual
            self.ymin=self.ymin_manual
            self.ymax=self.ymax_manual
            
        if switch_var.get():
            self.cursor_x_1=self.cursor_x_1E
            self.cursor_x_2=self.cursor_x_2E
            self.xmin=self.xminE
            self.xmax=self.xmaxE
        else:
            self.cursor_x_1=self.cursor_x_1WL
            self.cursor_x_2=self.cursor_x_2WL
            self.xmin=self.xminWL
            self.xmax=self.xmaxWL
  
        # Add vertical cursor lines
        self.cursor_line_x_1, = self.ax.plot([self.cursor_x_1, self.cursor_x_1], [self.ymin, self.ymax], color=self.color1)
        self.cursor_line_x_2, = self.ax.plot([self.cursor_x_2, self.cursor_x_2], [self.ymin, self.ymax], color=self.color2)

        # Add horizontal cursor lines
        self.cursor_line_y_1, = self.ax.plot([self.xmin, self.xmax], [self.cursor_y_1, self.cursor_y_1], color=self.color1)
        self.cursor_line_y_2, = self.ax.plot([self.xmin, self.xmax], [self.cursor_y_2, self.cursor_y_2], color=self.color2)
        
        # Add horizontal line for marker
        self.cursor_line_marker_y, = self.ax.plot([self.xmin, self.xmax], [self.cursor_marker_y, self.cursor_marker_y], color=self.color_marker)
  
    def update_cursors(self):
        if self.switch_autoscale_var.get()==False:
            self.ymin=np.min(C[0])
            self.ymax=np.max(C[0])
            self.xminE=E[-1]
            self.xmaxE=E[0]
            self.xminWL=WL[0]
            self.xmaxWL=WL[-1]
        else:
            self.xminWL=self.xminWL_manual
            self.xmaxWL=self.xmaxWL_manual
            self.xminE=self.xminE_manual
            self.xmaxE=self.xmaxE_manual
            self.ymin=self.ymin_manual
            self.ymax=self.ymax_manual
            
        if switch_var.get():
            self.cursor_x_1=self.cursor_x_1E
            self.cursor_x_2=self.cursor_x_2E
            self.xmin=self.xminE
            self.xmax=self.xmaxE
        else:
            self.cursor_x_1=self.cursor_x_1WL
            self.cursor_x_2=self.cursor_x_2WL
            self.xmin=self.xminWL
            self.xmax=self.xmaxWL
            
        # Update cursor 1
        self.cursor_line_x_1.set_data([self.cursor_x_1, self.cursor_x_1], [self.ymin, self.ymax])
        self.cursor_y_1 = self.get_y_at_x(self.cursor_x_1)
        self.cursor_line_y_1.set_data([self.xmin, self.xmax], [self.cursor_y_1, self.cursor_y_1])

        # Update cursor 2
        self.cursor_line_x_2.set_data([self.cursor_x_2, self.cursor_x_2], [self.ymin, self.ymax])
        self.cursor_y_2 = self.get_y_at_x(self.cursor_x_2)
        self.cursor_line_y_2.set_data([self.xmin, self.xmax], [self.cursor_y_2, self.cursor_y_2])
        
        # Update marker
        self.cursor_line_marker_y.set_data([self.xmin, self.xmax], [self.cursor_marker_y, self.cursor_marker_y])

        if switch_var.get():
            self.cursor_label_1.config(text=f"Cursor 1 position: x={self.cursor_x_1:.8f} eV, y={self.cursor_y_1:.1f}")
            self.cursor_label_2.config(text=f"Cursor 2 position: x={self.cursor_x_2:.8f} eV, y={self.cursor_y_2:.1f}")
            self.delta_cursors_x=abs(self.cursor_x_1-self.cursor_x_2)*1e3 #put it in meV
        else:
            self.cursor_label_1.config(text=f"Cursor 1 position: x={self.cursor_x_1:.6f} nm, y={self.cursor_y_1:.1f}")
            self.cursor_label_2.config(text=f"Cursor 2 position: x={self.cursor_x_2:.6f} nm, y={self.cursor_y_2:.1f}")
            self.delta_cursors_x=abs(self.cursor_x_1-self.cursor_x_2)*1e9*6.63e-34*3e8/1.6e-19/self.cursor_x_1**2*1e3 #put it in meV

        self.delta_cursors_y=abs(self.cursor_y_1-self.cursor_y_2)
        self.cursor_label_delta_x.config(text=f"Energy delta={self.delta_cursors_x:.8f} meV")
        self.cursor_label_delta_y.config(text=f"Energy delta={self.delta_cursors_y:.1f} counts")
            
            
    def get_y_at_x(self, x):
        # Example function to get y-value from the plotted curve at the given x-coordinate
        # Here, we'll use interpolation to find the y-value
        x_data, y_data = self.ax.lines[0].get_data()
        if switch_var.get():
            x_data= np.flip(x_data)
            y_data= np.flip(y_data)
        y_interp = np.interp(x, x_data, y_data)
        return y_interp
    
    def update_xylim(self):
        self.xminWL_manual=float(XminWL_entry.get())
        self.xmaxWL_manual=float(XmaxWL_entry.get())
        self.xminE_manual=float(XminE_entry.get())
        self.xmaxE_manual=float(XmaxE_entry.get())
        self.ymin_manual=float(Ymin_entry.get())
        self.ymax_manual=float(Ymax_entry.get())
        
        # Cursor drag threshold (adjust as needed)
        self.xspanWL=abs(self.xmaxWL-self.xminWL)
        self.xspanE=abs(self.xmaxE-self.xminE)
        self.yspan=abs(self.ymax-self.ymin)
        
        self.drag_thresholdWL = self.xspanWL/400
        self.drag_thresholdE = self.xspanE/400
        self.drag_threshold_y = self.yspan/400
        
    def switch_autoscale_changed(self):
        if self.switch_autoscale_var.get():
            self.switch_autoscale.config(image=switch_on_image)
        else:
            self.switch_autoscale.config(image=switch_off_image)


#%%initialize everything
# start_Andor_Shamrock()

# app = CursorPlot(tab_plot)

persistent_frame.mainloop()

try:
    Keithley.write(":SOUR:VOLT:LEV:IMM:AMPL 0")
    Keithley.write(":OUTP OFF")  # Eteignage de loutput
    Keithley.close()
except:
    a=1
    
try:
    tlpm.close()
except:
    a=1
    
try:
    LambdaOverTwo.close()
except:
    a=1
    
try:
    LambdaOverTwoPolar.close()
except:
    a=1
    
try:
    Msquared.close()
except:
    a=1
    
try:
    picoharp.close()
    qmi_context.stop()
except:
    a=1

cam.close()
spec.close()
sys.exit()