import pandas as pd

file_path = r'\support\el_centro.txt'

# Using Pandas with a column specification
col_specification = [(0, 10), (20, 30), (30, 40), (30, 40), (50, 60), (70, 80)]

data = pd.read_fwf(file_path, colspecs=col_specification)
