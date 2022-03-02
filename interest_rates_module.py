# -*- coding: utf-8 -*-
"""
Created on Sun Feb 27 09:55:59 2022

@author: HP557LA
"""
#Function to calculate interest rates forward 1y
def calculate_1y_forwardRates(df,curve):
    df[curve+"_shift"] = df[curve].shift(1, fill_value=0)
    df[curve+"_forward"] = (((1 + df[curve]) ** df["Maturity"]) / 
                        ((1 + df[curve+"_shift"]) ** (df["Maturity"] - 1)) - 1)
    return df