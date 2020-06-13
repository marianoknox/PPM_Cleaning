# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np

df = pd.read_excel("./PPM2.xlsx")
row = len(df)
col = len(df.columns)
print("No. of Rows:" + str(row))
print("No. of Columns:" + str(col))
df.head()