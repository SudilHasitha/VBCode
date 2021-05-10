# -*- coding: utf-8 -*-
"""
Created on Mon Apr  6 09:19:14 2020

@author: sudil
"""

import pandas as pd


data = pd.read_csv('products.csv')
data.head()
data.columns

result_status_success = data['result_status']=="SUCCESS"
result = data[result_status_success]

file_name="D:\Dialog\LakalFernando\Helakuru_Charging_Success"    
file_name = result.to_csv(file_name+str(".csv"), index = False, header=True)

