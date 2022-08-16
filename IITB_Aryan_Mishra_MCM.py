# -*- coding: utf-8 -*-
"""
Created on Mon Sep 6 10:41:23 2020

@author: Aryan Mishra
Roll Number: 180110015
IITB, 3rd Year UG
"""

"""
Importing Libraries/Packages
"""

from random import gauss
import xlwings as xw
import sys 
import numpy as np
import os

"""
Defining Functions for Asset Price, Call Pay Off, Put Pay Off
"""

def generate_asset_price(S,v,r,T):
    return S * np.exp((r - 0.5 * v**2) * T + v * np.sqrt(T) * gauss(0,1.0))

def call_payoff(S_T,K):
    return max(0.0,S_T-K)

def put_payoff(S_T,K):
    return max(0.0,K-S_T)

"""
Assigning Values to Variables From Shell Command
"""


S = float(sys.argv[1])
K = float(sys.argv[2])
T = float(sys.argv[3])
v = float(sys.argv[4])
v=v/100
o = sys.argv[5]
r = float(sys.argv[6])
r=r/100
delta = float(sys.argv[7])
delta=delta/100
filepath = sys.argv[8]
filename = sys.argv[9]
filesheet = sys.argv[10]


# # S = 100
# # K = 110
# # T = 2
# # v = 0.1
# # r = 0.04
# o = 'call'

"""
Calculating Output Premium
"""

simulations = 90000
payoffs = []
discount_factor = np.exp(-r * T)

for i in range(simulations):
    
    S_T = generate_asset_price(S,v,r,T)
    
    if o.lower().strip() =="call":
        payoffs.append(
            call_payoff(S_T, K)
            )
    elif o.lower().strip() == "put":
        payoffs.append(
            put_payoff(S_T, K)
            )
    else:
        res="Please Enter Call/Put"
        print("Please Enter Call/Put")
    
    
res = discount_factor * (sum(payoffs) / float(simulations))
# print ('Price: %.4f' % res)

"""
Printing Output in Excel
"""

output_file= os.path.join(filepath,filename)

wb=xw.Book(output_file)
sht1= wb.sheets[filesheet]
sht1.range('E14').value=res