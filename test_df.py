#!/usr/bin/env python
# coding: utf-8

# In[1]:

import encodings
from re import X
import pandas as pd
import numpy as np
import networkx as nx
import matplotlib.pyplot as plt
import graphviz
import pygraphviz
import math
import simpy
import random
from collections import namedtuple
from contextlib import ExitStack
from numpy import inf
import itertools

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib
# get_ipython().run_line_magic('matplotlib', 'inline')
from matplotlib.gridspec import GridSpec
matplotlib.pyplot.style.use('seaborn-darkgrid')
import tkinter as tk
from tkinter import ttk
import datetime
from datetime import timedelta as td
from datetime import time as dtt
import warnings
warnings.filterwarnings("ignore")
import time
import mplcursors

# In[2]:

# Global Variables

# Import the Line Details from Base File and Create Base Data
# Nhập dữ liệu cơ sở và tạo dữ liệu cơ sở
file_path = r'Line Balancing.xlsm'

# For Capturing the Production During Simulation
# Bắt giữ các sản phẩm trong quá trình mô phỏng
throughput = 0

# Que as Raw Material in Production
# Que là nguyên liệu thô trong sản xuất
Que = namedtuple('Que','RM_id, Task, Task_Time, End_Time, Next_Task')

# List of All Global Variables
# Danh sách các biến số toàn cầu
production = globals()

# Get Color List for Network
# Lấy danh sách các màu cho mạng
node_colors = pd.read_excel(file_path, sheet_name='Colors',usecols='A,C')
node_colors = node_colors['Hex'].to_dict()


# Get Input Numbers for Line Balancing and Simulation
# Lấy số liệu cân bằng và mô phỏng

input_data = pd.read_excel(file_path, sheet_name='cap_ultra_test',skiprows=3,usecols='B:H')
cycle_time = max(input_data['ST (Minutes)']/input_data['No of Operators'])
workstations = (input_data['ST (Minutes)']).count()

# In[3]:

# Functions for Line Balancing
# Hàm cân bằng

def import_data(file_path):
    df = pd.read_excel(file_path,sheet_name='cap_ultra_test',skiprows=3,usecols='B:H')

    # Manipulate the Line Details data to split multiple Predecessors to individual rows 
    # Thao tác dữ liệu chi tiết đường dây để chia trước cho từng hàng 

    temp = pd.DataFrame(columns=['Task Number', 'Precedence'])
    
    for i, d in df.iterrows():
        for j in str(d[3]).split(','):
            rows = pd.DataFrame({'Task Number':d[0], 'Precedence': [int(j)]})
            temp = temp.append(rows)

    temp = temp[['Precedence','Task Number']]
    # print(temp) 
    temp.columns = ['Task Number','Next Task']
    temp['Task Number'] = temp['Task Number'].astype(int)

    # Append the Last Task
    last = pd.DataFrame({'Task Number': [max(temp['Task Number'])+ 1], 'Next Task': ['END']})
    temp = temp.append(last)

    # Create the Final Data for drawing precedence graph
    
    final_df = temp.merge(df[['Task Number','Task Description','Resource','ST (Minutes)' , 'No of Operators', 'Preparation Stage']],on='Task Number',how='left')
    final_df = final_df[['Task Number','Task Description', 'Resource','ST (Minutes)','Next Task' , 'No of Operators','Preparation Stage']]
    final_df['ST (Minutes)'] = final_df['ST (Minutes)'].fillna(0)
    final_df['Preparation Stage'] = final_df['Preparation Stage'].fillna(1)
    final_df['Task Description'] = final_df['Task Description'].fillna('START')
    final_df['Next Task'] = final_df['Next Task'].loc[0:(final_df.shape[0]-2)].astype(int)
    final_df['Task Number'] = final_df['Task Number'].loc[0:(final_df.shape[0]-1)].astype(int)

    # final_df = final_df.append(last)

    # counter = 1
    # for i, d in final_df.iterrows():
    #     if d[0] == 0:
    #         final_df.iloc[i,d[0]] = 'S_%d' %counter
    #         counter+=1

    for i, d in final_df.iterrows():
        if d[6] == 0:
            for j , k in final_df.iterrows():
                if (k[4] == d[0]):
                    print(k[4])
                    # final_df = final_df.replace(k[4] , d[4])
                    final_df.iloc[j,k[4]] = d[4]
            # final_df = final_df.drop(i, axis=0, inplace=False)

    print(final_df.info())
    return final_df


data = import_data(file_path)

# df = pd.read_excel(file_path,sheet_name='cap_ultra_test',skiprows=3,usecols='B:H')
# for i, d in df.iterrows():
#     print(df.iloc[i,3])
