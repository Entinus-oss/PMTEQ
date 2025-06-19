# -*- coding: utf-8 -*-
"""
Created on Tue May 13 15:30:58 2025

@author: Manip
"""

import pyvisa as visa

DG645_ADDRESS = "GPIB0::15::INSTR"

rm = visa.ResourceManager()
DG645=rm.open_resource(DG645_ADDRESS)
DG645.open()
DG645.read()
DG645.close()

