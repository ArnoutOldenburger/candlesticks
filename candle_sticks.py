#!/usr/bin/env python
# -*- coding: utf-8 -*-
# ==============================================================================>
#
#   Copyright (C) KPN-ITNS / Glashart Media
#   All Rights Reserved
#
#   <Module>:=      candle_sticks.py]
#   Author:         H.A. Oldenburger
#   Date:           March 2019
#   Parameters:     
#   Purpose:        
#   Python version: v2.7.6
#
#   Amendment history:
#   20190617    AO  Initial version
#
#   20200625: https://github.com/freqtrade/technical/tree/master/technical
# ==============================================================================>

#    debug_message = "iRow= {row} - rico_open_maxline= {rico_open_maxline}".format(row=str(iRow),rico_open_maxline=str(rico_open_maxline))
#    win32api.MessageBox(xw.apps.active.api.Hwnd, debug_message, 'Info',win32con.MB_ICONINFORMATION)

import xlwings as xw
import win32api
import win32con
import numpy as np

import codecs
import glob, os
import sys
import time
import datetime
import getopt
import argparse
import xlrd
import xlwt
import csv
import re
import collections
import operator

import json 
import numpy as np
import math

from datetime import *
from datetime import date, timedelta

from xlrd import open_workbook, XLRDError
from xlwt import Workbook, easyxf, Borders

from tkinter import messagebox
from tkinter import ttk
from tkinter import *

# ==============================================================================>
# Constants / Defines.
# ==============================================================================>
CONSTANT_TRADES_TIME = "B"
CONSTANT_TRADES_OPEN = "C"
CONSTANT_TRADES_CLOSE = "D"
CONSTANT_TRADES_HIGH = "E"
CONSTANT_TRADES_LOW = "F"
CONSTANT_TRADES_VOLUMEFROM = "G"
CONSTANT_TRADES_VOLUMETO = "H"

CONSTANT_TRADES_TREND_RICO_OPEN_MAXLINE = "O"
CONSTANT_TRADES_TREND_RICO_OPEN_MINLINE = "Q"
CONSTANT_TRADES_TREND_RICO_CLOSE_MAXLINE = "S"
CONSTANT_TRADES_TREND_RICO_CLOSE_MINLINE = "U"
CONSTANT_TRADES_TREND_RICO_SUPPORT = "W"
CONSTANT_TRADES_TREND_RICO_RESISTANCE = "Y"
CONSTANT_TRADES_NR_POINTS_BETWEEN_BANDS = "AA"
CONSTANT_TRADES_RATIO_BETWEEN_BANDS = "AB"
CONSTANT_TRADES_NR_POINTS_BETWEEN_BANDS_2ND = "AC"

CONSTANT_DATA_TIME = 1
CONSTANT_DATA_OPEN_MAX_LINE = 2
CONSTANT_DATA_OPEN_MIN_LINE = 3
CONSTANT_DATA_CLOSE_MAX_LINE = 4
CONSTANT_DATA_CLOSE_MIN_LINE = 5
CONSTANT_DATA_SUPPORT_LINE = 6
CONSTANT_DATA_RESISTENCE_LINE = 7
CONSTANT_DATA_RICO_OPEN_MAXLINE = 8
CONSTANT_DATA_RICO_OPEN_MINLINE = 9
CONSTANT_DATA_RICO_CLOSE_MAXLINE = 10
CONSTANT_DATA_RICO_CLOSE_MINLINE = 11
CONSTANT_DATA_RICO_SUPPORT = 12
CONSTANT_DATA_RICO_RESISTANCE = 13
CONSTANT_DATA_NR_POINTS_BETWEEN_BANDS = 14
CONSTANT_DATA_RATIO_BETWEEN_BANDS = 15
CONSTANT_DATA_NR_POINTS_BETWEEN_BANDS_2ND = 16

# ==============================================================================>
# function object : candle_sticks_trends
# parameters      : iRowStart, iRowEnd
# return value    :
# description     :
# ==============================================================================>
def candle_sticks_trends(iRowStart, iRowEnd):
		
    trades_sht = xw.Book.caller().sheets['Trades']
    data_sht = xw.Book.caller().sheets['Data']

    str_open_values    = ""
    str_high_values    = ""
    str_low_values     = ""
    str_close_values   = ""

    iRowStart = int(iRowStart)
    iRowEnd = int(iRowEnd)
    iRow = iRowStart
		
    while iRow <= iRowEnd:

        str_cell = CONSTANT_TRADES_OPEN + str(iRow)
        open_value = str(trades_sht.range(str_cell).value)

        if str_open_values:
            str_open_values = "%s,%s" % (str_open_values, open_value)
        else:
            str_open_values = "%s" % (open_value)
        
        str_cell = CONSTANT_TRADES_CLOSE + str(iRow)
        high_value = str(trades_sht.range(str_cell).value)

        if str_high_values:
            str_high_values = "%s,%s" % (str_high_values, high_value)
        else:
            str_high_values = "%s" % (high_value)

        str_cell = CONSTANT_TRADES_HIGH + str(iRow)
        low_value = str(trades_sht.range(str_cell).value)

        if str_low_values:
            str_low_values = "%s,%s" % (str_low_values, low_value)
        else:
            str_low_values = "%s" % (low_value)
        
        str_cell = CONSTANT_TRADES_LOW + str(iRow)
        close_value = str(trades_sht.range(str_cell).value)

        if str_close_values:
            str_close_values = "%s,%s" % (str_close_values, close_value)
        else:
            str_close_values = "%s" % (close_value)
        
        iRow = iRow + 1
        
    h = np.fromstring(str_high_values, dtype=float, sep=',')
    l = np.fromstring(str_low_values, dtype=float, sep=',')
    o = np.fromstring(str_open_values, dtype=float, sep=',')
    c = np.fromstring(str_close_values, dtype=float, sep=',')
		
    maxline, minline = segtrends(o, segments = 2) 
    open_maxline = maxline.tolist()
    open_minline = minline.tolist()
	 
    maxline, minline = segtrends(c, segments = 2) 
    close_maxline = maxline.tolist()
    close_minline = minline.tolist()

    support_and_resistance = SupportAndResistance(h, l, c)
    
    support_line = support_and_resistance["support"]
    resistance_line = support_and_resistance["resistance"]

    nr_points_between_bands = support_and_resistance["nr_points_between_bands"]
    ratio_between_bands = support_and_resistance["ratio_between_bands"]
    nr_points_between_bands_2nd = support_and_resistance["nr_points_between_bands_2nd"]
    
    iSessionDays = (iRowEnd - iRowStart) + 1
    iFirstDay = iRowStart - 3 
		
    iOffset = (iSessionDays * iFirstDay) + 2
    
    str_cell = CONSTANT_TRADES_TIME + str(iRowEnd)
    str_time = str(trades_sht.range(str_cell).value)

    data_sht.cells(iOffset,CONSTANT_DATA_TIME).value = str_time
    data_sht.cells(iOffset,CONSTANT_DATA_OPEN_MAX_LINE).value = str(open_maxline)
    data_sht.cells(iOffset,CONSTANT_DATA_OPEN_MAX_LINE).value = str(open_maxline)
    data_sht.cells(iOffset,CONSTANT_DATA_OPEN_MIN_LINE).value = str(open_minline)
    data_sht.cells(iOffset,CONSTANT_DATA_CLOSE_MAX_LINE).value = str(close_maxline)
    data_sht.cells(iOffset,CONSTANT_DATA_CLOSE_MIN_LINE).value = str(close_minline)
    data_sht.cells(iOffset,CONSTANT_DATA_SUPPORT_LINE).value = str(support_line)
    data_sht.cells(iOffset,CONSTANT_DATA_RESISTENCE_LINE).value = str(resistance_line)
		
    rico_open_maxline = rico(open_maxline)
    rico_open_minline = rico(open_minline)
    rico_close_maxline = rico(close_maxline)
    rico_close_minline = rico(close_minline)
    rico_support = rico(support_and_resistance["support"])
    rico_resistance = rico(support_and_resistance["resistance"])
	 
    #iRow = iRow - 1
    
    str_cell = CONSTANT_TRADES_TREND_RICO_OPEN_MAXLINE + str(iRow)
    trades_sht.range(str_cell).value = rico_open_maxline

    str_cell = CONSTANT_TRADES_TREND_RICO_OPEN_MINLINE + str(iRow)
    trades_sht.range(str_cell).value = rico_open_minline
    
    str_cell = CONSTANT_TRADES_TREND_RICO_CLOSE_MAXLINE + str(iRow)
    trades_sht.range(str_cell).value = rico_close_maxline
    
    str_cell = CONSTANT_TRADES_TREND_RICO_CLOSE_MINLINE + str(iRow)
    trades_sht.range(str_cell).value = rico_close_minline
    
    str_cell = CONSTANT_TRADES_TREND_RICO_SUPPORT + str(iRow)
    trades_sht.range(str_cell).value = rico_support

    str_cell = CONSTANT_TRADES_TREND_RICO_RESISTANCE + str(iRow)
    trades_sht.range(str_cell).value = rico_resistance

    str_cell = CONSTANT_TRADES_NR_POINTS_BETWEEN_BANDS + str(iRow)
    trades_sht.range(str_cell).value = nr_points_between_bands
    
    str_cell = CONSTANT_TRADES_RATIO_BETWEEN_BANDS + str(iRow)
    trades_sht.range(str_cell).value = ratio_between_bands

    str_cell = CONSTANT_TRADES_NR_POINTS_BETWEEN_BANDS_2ND + str(iRow)
    trades_sht.range(str_cell).value = nr_points_between_bands_2nd
    
#==============================================================================>
# function object : segtrends
# parameters      : x, segments=2
# return value    :
# description     : 
#==============================================================================>
def segtrends( x, segments=2):
    """
    Turn minitrends to iterative process more easily adaptable to
    implementation in simple trading systems; allows backtesting functionality.

    :param x: One-dimensional data set
    :param window: How long the trendlines should be. If window < 1, then it
                   will be taken as a percentage of the size of the data
    :param charts: Boolean value saying whether to print chart to screen
    """

    y = np.array(x)

    # Implement trendlines
    segments = int(segments)
    maxima = np.ones(segments)
    minima = np.ones(segments)
    segsize = int(len(y)/segments)

    for i in range(1, segments+1):
        ind2 = i*segsize
        ind1 = ind2 - segsize
        maxima[i-1] = max(y[ind1:ind2])
        minima[i-1] = min(y[ind1:ind2])

    # Find the indexes of these maxima in the data
    x_maxima = np.ones(segments)
    x_minima = np.ones(segments)

    for i in range(0, segments):
        x_maxima[i] = np.where(y == maxima[i])[0][0]
				
        x_minima[i] = np.where(y == minima[i])[0][0]
				
    for i in range(0, segments-1):
        
        if x_maxima[i+1] - x_maxima[i] != 0:
            maxslope = (maxima[i+1] - maxima[i]) / (x_maxima[i+1] - x_maxima[i])
        else:
            maxslope =  0
						
        a_max = maxima[i] - (maxslope * x_maxima[i])
				
        b_max = maxima[i] + (maxslope * (len(y) - x_maxima[i]))
				
        maxline = np.linspace(a_max, b_max, len(y))

        if x_minima[i+1] - x_minima[i] != 0:
            minslope = (minima[i+1] - minima[i]) / (x_minima[i+1] - x_minima[i])
        else:
            minslope =  0
				
        a_min = minima[i] - (minslope * x_minima[i])
				
        b_min = minima[i] + (minslope * (len(y) - x_minima[i]))
				
        minline = np.linspace(a_min, b_min, len(y))

    return maxline, minline

#==============================================================================>
# function object : rico
# parameters      : numpy_array
# return value    : rico
# description     : rico = (y2-y1)/(x2-x1)
#==============================================================================>
def rico(numpy_array):
    l = len(numpy_array)

    z = [i for i in range(0,len(numpy_array))]
    x = np.array(z)
    y = numpy_array

    y1 = y[0]
    y2 = y[l-1]

    x1 = x[0]
    x2 = x[l-1]

    rico = (y2-y1)/(x2-x1)
    return rico     

#==============================================================================>
# function object : slope
# parameters      : numpy_array
# return value    : slope
# description     : slope = ((len(x)*sum(x*y)) - (sum(x)*sum(y)))/(len(x)*(sum(x**2))-(sum(x)**2))
#==============================================================================>
def slope(numpy_array):
    z = [i for i in range(0,len(numpy_array))]
    x = np.array(z)
    y = numpy_array

    first_part  = len(x) * sum(x * y)
    second_part = sum(x) * sum(y) 
    third_part  = first_part - second_part
    
    fourth_part = len(x) * sum(np.power(x, 2))
    fifth_part  = math.pow(sum(x), 2)
    sixth_part  = fourth_part - fifth_part
    
    slope       = third_part / sixth_part
    return slope     

#==============================================================================>
# function object : SupportAndResistance
# parameters      : h, l, c
# return value    :
# description     : 
#==============================================================================>
def SupportAndResistance(h, l, c):
    return_value = {}

    pivots = (h + l + c) / 3

    t = np.arange(len(c))
    sa, sb = fit_line(t, pivots - (h - l)) 
    ra, rb = fit_line(t, pivots + (h - l))

    support = sa * t + sb
    resistance = ra * t + rb 

    condition                   = (c > support) & (c < resistance)
    between_bands               = np.where(condition) 

    nr_points_between_bands     = len(np.ravel(between_bands))
    ratio_between_bands         = float(nr_points_between_bands)/len(c) 

    tomorrows_support           = sa * (t[-1] + 1) + sb
    tomorrows_resistance        = ra * (t[-1] + 1) + rb

    a1 = c[c > support]
    a2 = c[c < resistance]
    nr_points_between_bands_2nd = len(np.intersect1d(a1, a2))

    return_value["support"] = support.tolist()
    return_value["resistance"] = resistance.tolist()
    return_value["nr_points_between_bands"] = nr_points_between_bands
    return_value["ratio_between_bands"] = ratio_between_bands
    return_value["tomorrows_support"] = tomorrows_support
    return_value["tomorrows_resistance"] = tomorrows_resistance
    return_value["nr_points_between_bands_2nd"] = nr_points_between_bands_2nd

    return return_value

#==============================================================================>
# function object : fit_line
# parameters      : t, y
# return value    :
# description     : 
#==============================================================================>
def fit_line(t, y):
   A = np.vstack([t, np.ones_like(t)]).T

   #return np.linalg.lstsq(A, y)[0]
   return np.linalg.lstsq(A, y, rcond=-1)[0]
   
#==============================================================================>
# function object : moving_average
# parameters      : a, n
# return value    :
# description     : 
#==============================================================================>
   
def moving_average(a, n=3) :
    ret = np.cumsum(a, dtype=float)
    ret[n:] = ret[n:] - ret[:-n]
    return ret[n - 1:] / n   

    
    

