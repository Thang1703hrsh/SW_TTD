#!/usr/bin/env python
# coding: utf-8

# In[1]:


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

input_data = pd.read_excel(file_path, sheet_name='hat4',skiprows=3,usecols='B:F')
cycle_time = max(input_data['ST (Minutes)'])
workstations = (input_data['ST (Minutes)']).count()

# In[3]:

# Functions for Line Balancing
# Hàm cân bằng

def import_data(file_path):
    df = pd.read_excel(file_path,sheet_name='hat4',skiprows=3,usecols='B:F')

    # Manipulate the Line Details data to split multiple Predecessors to individual rows
    temp = pd.DataFrame(columns=['Task Number', 'Precedence'])

    for i, d in df.iterrows():
        for j in str(d[3]).split(','):
            rows = pd.DataFrame({'Task Number':d[0], 'Precedence': [int(j)]})
            temp = temp.append(rows)

    temp = temp[['Precedence','Task Number']]
    temp.columns = ['Task Number','Next Task']
    temp['Task Number'] = temp['Task Number'].astype(int)

    # Append the Last Task
    last = pd.DataFrame({'Task Number': [max(temp['Task Number'])+ 1], 'Next Task': ['END']})
    temp = temp.append(last)

    # Create the Final Data for drawing precedence graph
    # Tạo dữ liệu cuối cùng cho biểu đồ thứ bậc
    
    final_df = temp.merge(df[['Task Number','Task Description','Resource','ST (Minutes)']],on='Task Number',how='left')
    final_df = final_df[['Task Number','Task Description', 'Resource','ST (Minutes)','Next Task']]
    final_df['ST (Minutes)'] = final_df['ST (Minutes)'].fillna(0)
    final_df['Task Description'] = final_df['Task Description'].fillna('START')

    counter = 1
    for i, d in final_df.iterrows():
        if d[0] == 0:
            final_df.iloc[i,d[0]] = 'S_%d' %counter
            counter+=1
    return final_df


# In[4]:


# Function to Create Precedence Graph and Create Precedence Graph for Visualizing the Line
# Tạo đồ thị để hình dung

def precedence_graph(data_set):
    g = nx.DiGraph()
    fig, ax = plt.subplots(1, 1, figsize=(25, 10))
    
    for i, d in data_set.iterrows():
        g.add_node(d[0],time = d[3])
        g.add_edge(d[0],d[4])

    labels = nx.get_node_attributes(g,'time')
    nx.draw(g,with_labels=True,node_size=700, node_color="skyblue",
           pos=nx.drawing.nx_agraph.graphviz_layout(
            g,
            prog='dot',
            args='-Grankdir=LR'
        ),ax=ax)
    
    # plt.draw()
    # plt.show()
    return g


# In[5]:

##=======================================
# Create Data Table for Line Balancing, Workstations And Allocation
# Tạo dữ liệu cho cân bằng dòng, các trạm và phân bổ
##=======================================


def create_LB_Table(data_set,g):
    line_balance = pd.DataFrame(columns=['Task Number','Number of Following Task'])
    end = data_set.iloc[-1][4]
    nodes = list(set(data_set['Task Number']))
    for i in nodes:
        unique_nodes = []
        for paths in nx.all_simple_paths(g,i,end):
            for path in paths:
                unique_nodes.append(path)
        successor_length = len(list(set(unique_nodes)))-2
        line_balance = line_balance.append({'Task Number':i, 'Number of Following Task': successor_length},ignore_index=True)
    line_balance.sort_values(by=['Number of Following Task'],ascending=False,inplace=True)    

    # Arrange the Data and return the final table
    # Sắp xếp dữ liệu và trả lại bảng cuối cùng

    final = pd.merge(line_balance,data_set[['Task Number','Task Description','Resource','ST (Minutes)','Next Task']],on='Task Number',how='left')
    final = final.drop_duplicates()
    final.sort_values(by=['Number of Following Task','ST (Minutes)'],ascending=[False,False],inplace=True)    
    final['Workstation'] = 0
    final['Allocated'] = 'No'
    final = final.reset_index()
    final = final[['Task Number','Task Description','Resource','Number of Following Task','ST (Minutes)','Workstation','Allocated','Next Task']]
    final.loc[final['ST (Minutes)']==0,'Allocated'] = 'Yes'
    final.loc[final['ST (Minutes)']==0,'Workstation'] = 1
    return final


# In[6]:

##=======================================
# Create Function for Workstation Allocation based on Largest Following Task Hueristic Algorithm
# Tạo hàm phân bổ trạm máy bằng thuật toán Largest Following Task Hueristic
##=======================================


def find_feasable_allocation(base_data, allocation_table, cycle_time, workstations):
    
    counter = [0] * workstations
    
    current_station = 1
    
    stations = {}
    
    for i in range(1,workstations + 1):
        stations[i] = 'open'
    
    count_station = 0
    # i là index, d là data trong mỗi hàng
    for i, d in allocation_table.iterrows():
        if d[1] != 'START':
            
            current_task = d[0] #Task hiện hành
            
            current_task_allocated = allocation_table[allocation_table['Task Number']==d[0]].Allocated.tolist()[0] #

            current_task_time = base_data[base_data['Task Number']==d[0]]['ST (Minutes)'].tolist()[0]
            
            previous_task = base_data[base_data['Next Task']== d[0]]['Task Number'].tolist()

            previous_task_list = []
            
            previous_stations_list = []
            for pt in previous_task:
                previous_task_list.append(allocation_table[allocation_table['Task Number']==pt].Allocated.tolist()[0])
                previous_stations_list.append(allocation_table[allocation_table['Task Number']==pt].Workstation.tolist()[0])

            count_allocations = sum(map(lambda x : x=='Yes',previous_task_list)) # Tính số lượng các task trước đó
            len_allocations = len(previous_task_list)

            if count_allocations == len_allocations:
                previous_task_allocated = 'Yes'
            else:
                previous_task_allocated = 'No'

            station_cut_off = max(previous_stations_list)

            if (previous_task_allocated == 'Yes') & (current_task_allocated == 'No') & (current_task_time <= (cycle_time - counter[station_cut_off-1])) & (count_station <= 2):
                allocation_table.iloc[i,6] = 'Yes'
                allocation_table.iloc[i,5] = station_cut_off
                counter[station_cut_off-1]+=current_task_time
                count_station += 1

            elif (previous_task_allocated == 'Yes') & (current_task_allocated == 'No') & (current_task_time <= (cycle_time - counter[current_station-1])) & (count_station <= 2):
                allocation_table.iloc[i,6] = 'Yes'
                allocation_table.iloc[i,5] = current_station
                counter[current_station-1]+=current_task_time  
                count_station += 1  
                
            elif (previous_task_allocated == 'Yes') & (current_task_allocated == 'No'):
                allocation_table.iloc[i,6] = 'Yes'
                allocation_table.iloc[i,5] = current_station + 1
                current_station+=1
                counter[current_station-1]+=current_task_time 
                count_station = 0
                current_station+=1

            else:
                allocation_table.iloc[i,6] = 'Yes'
                allocated_station = allocation_table[allocation_table['Task Number'] == current_task]['Workstation'].tolist()[0]
                allocation_table.iloc[i,5] = allocated_station
                allocation_table.iloc[i,4] = 0
            
    
    #reassign the starting workstations from 1 to respective workstations
    # Phân bổ các máy trạm từ 1 đến máy trạm tương ứng
    reassign = allocation_table[allocation_table['Task Description'] == 'START']['Task Number'].tolist()

    for start in reassign:
        next_task = allocation_table[allocation_table['Task Number'] == start]['Next Task'].tolist()[0]
        next_task_station = allocation_table[allocation_table['Task Number'] == next_task]['Workstation'].tolist()[0]
        allocation_table.loc[allocation_table['Task Number'] == start,'Workstation'] = next_task_station
        
    return allocation_table            

data = import_data(file_path)
graph = precedence_graph(data)
Line_Balance = create_LB_Table(data,graph)
solution = find_feasable_allocation(data,Line_Balance,cycle_time,workstations)
print(solution)