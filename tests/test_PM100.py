# -*- coding: utf-8 -*-
"""
Created on Tue May 13 12:04:15 2025

@author: Manip
"""
from ThorlabsPM100 import ThorlabsPM100
from ctypes import create_string_buffer, c_bool
import pyvisa as visa

PM100_NAME = "USB0::0x1313::0x8072::P2008402::INSTR"
PM100_NAME_BYTES = b"USB0::0x1313::0x8072::P2008402::INSTR"

print("Starting PM100...")
rm = visa.ResourceManager()
inst=rm.open_resource(PM100_NAME)
PM100 = ThorlabsPM100(inst=inst)
print("Started successfully")

print(PM100.read, end="\r")
