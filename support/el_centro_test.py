import pandas as pd
import os
cwd = os.getcwd()

print(cwd)

file_path = 'el_centro.txt'

# Using Pandas with a column specification
col_specification = [(0, 10), (20, 30), (30, 40), (30, 40), (50, 60), (70, 80)]

data = pd.read_fwf(file_path, colspecs='infer', header=None)

times = pd.concat([data[0], data[2], data[4]], sort=True)

# times.sortvalues(ascending=True)

accels = pd.concat([data[1], data[3], data[5]], sort=True)


new_data = pd.concat([times, accels], axis=1)

print('test')