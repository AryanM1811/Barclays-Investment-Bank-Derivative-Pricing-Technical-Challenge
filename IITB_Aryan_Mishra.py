# -*- coding: utf-8 -*-
"""
Created on Mon Sep 5 03:01:17 2020

@author: Aryan Mishra
Roll Number: 180110015
IITB, 3rd Year UG
"""

"""
Importing Libraries/Packages
"""

import numpy as np
import xlwings as xw
import sys 
import os
import math

"""
Defining Cumulative Nomral Distribution Function
"""
def normal_cdf(x):
    "cdf for standard normal"
    q = math.erf(x / math.sqrt(2.0))
    return (1.0 + q) / 2.0


"""
Assigning Values to Variables From Shell Command
"""

S = float(sys.argv[1])
K = float(sys.argv[2])
T = float(sys.argv[3])
sig = float(sys.argv[4])
sig=sig/100
o = sys.argv[5]
rate = float(sys.argv[6])
rate=rate/100
delta = float(sys.argv[7])
delta=delta/100
filepath = sys.argv[8]
filename = sys.argv[9]
filesheet = sys.argv[10]

# S = 100
# K = 110
# T = 2
# sig = 0.1
# o = 'call'
# rate = 0.04
# delta = 0.01

"""
Calculating Paramters d1 and d2 for normal cdf
"""

d1=(np.log(S/K)+(rate-delta+0.5*(sig)**2)*T)/(sig*np.sqrt(T))
d2=(np.log(S/K)+(rate-delta-0.5*(sig)**2)*T)/(sig*np.sqrt(T))


"""
Calculating Output Premium
"""

if o.lower().strip() =="call":
    res=S*np.exp(-1*delta*T)*normal_cdf(d1) - K*np.exp(-1*rate*T)*normal_cdf(d2)
    # print(res)
elif o.lower().strip() == "put":
    res=K*np.exp(-1*rate*T)*normal_cdf(-1*d2) - S*np.exp(-1*delta*T)*normal_cdf(-1*d1)
    # print(res)
else:
    res="Please Enter Call/Put"
    # print("Please Enter Call/Put")

"""
Printing Output in Excel
"""
output_file= os.path.join(filepath,filename)

wb=xw.Book(output_file)
sht1= wb.sheets[filesheet]
sht1.range('E14').value=res