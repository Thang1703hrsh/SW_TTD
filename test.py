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

input_data = pd.read_excel(file_path, sheet_name='cap_ultra_merge_2',skiprows=3,usecols='B:G' , nrows = 7)
cycle_time = max(input_data['ST (Minutes)'])
workstations = (input_data['ST (Minutes)']).count()
print(input_data)