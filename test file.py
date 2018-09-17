import pandas as pd

mdl_dof_df = pd.DataFrame([{1: True, 2: False, 3: True, 4: False, 5: True, 6: False}])

a = mdl_dof_df.iloc[[0]]

mdl_dof_df.iloc[[0]] = pd.DataFrame([{1: True, 2: True, 3: True, 4: True, 5: True, 6: True}])

b = mdl_dof_df.iloc[[0]]
