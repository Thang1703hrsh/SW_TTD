import pandas as pd
import matplotlib.pyplot as plt

file_path = r'Line Balancing.xlsm'
df = pd.read_excel(file_path,sheet_name='Example_1',skiprows=3,usecols='B:F')
print(df['ST (Minutes)'])