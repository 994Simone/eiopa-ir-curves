# -*- coding: utf-8 -*-
"""
Created on Mon Feb 28 10:12:41 2022

@author: HP557LA
"""
import pandas as pd
import numpy as np

#Function to clean the dataframe imported from EIOPA excel file
def rfrPrepareInput(df):
    len_df = len(df.columns)
    
    for ii in range(len_df):
        df = df.rename(columns={df.columns[ii]:"column"+str(ii)})
    
    df = df.drop(["column0","column1"], axis = 1)
    df = df.drop([1,2,3,4,5,6,7,8]).reset_index(drop = True)
    
    years = np.arange(151)
    years = pd.DataFrame(years)
    years = years.replace(0,"Maturity")
    df["Maturity"] = years
    years = df.pop("Maturity")
    df.insert(0, 'Maturity', years)
    
    headers = df.iloc[[0]]
    headers = headers.values.tolist()
    df.columns = headers
    
    df = df.drop([0]).reset_index(drop = True)
    return df