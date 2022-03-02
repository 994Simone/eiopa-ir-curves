# -*- coding: utf-8 -*-
"""
Created on Fri Feb 25 23:08:52 2022

@author: HP557LA
"""
import pandas as pd
from rfrPrepareInput_module import rfrPrepareInput
from interest_rates_module import calculate_1y_forwardRates

#Non-fixed input (user can change them)
inputPath = r"C:\Users\HP557LA\Downloads\December 2021 - EN"
outputPath = r"C:\OutputPython"
curve_tmp = "\curve_tmp.xlsx"
curve_final = "\curve_final.xlsx"

#Fixed input (unless EIOPA changes them)
fileName = "\EIOPA_RFR_20211231_Term_Structures.xlsx"
sheetsName = ["RFR_spot_with_VA","Spot_WITH_VA_shock_UP",
              "Spot_WITH_VA_shock_DOWN","RFR_spot_no_VA","Spot_NO_VA_shock_UP",
              "Spot_NO_VA_shock_DOWN"]
curveTypes = ["Central","Central_ShockUp","Central_ShockDown","Central_noVA",
              "Central_noVA_ShockUp","Central_noVA_ShockDown"]
countryToRemove = ["Austria","Belgium","Cyprus","Estonia","Finland","France",
                   "Germany","Greece","Ireland","Italy","Latvia","Lithuania",
                   "Luxembourg","Malta","Netherlands","Portugal","Slovakia",
                   "Slovenia","Spain", "Liechtenstein"]
economy = ["EUR","BGN","HRK","CZK","DKK","HUF","ISK","NOK","PLN","RON","RUB",
           "SEK","CHF","GBP","AUD","BRL","CAD","CLP","CNY","COP","HKD","INR",
           "JPY","MYR","MXN","NZD","SGD","ZAR","KRW","TWD","THB","TRY","USD"]

#Code starts here
dfList = [pd.DataFrame()]*len(sheetsName)

#Import input as dataframe
for ii in range(len(sheetsName)):
    dfList[ii] = pd.read_excel(inputPath + fileName,sheet_name=sheetsName[ii])

#Delete useless information from dataframes (1° step)
for ii in range(len(dfList)):
    dfList[ii] = rfrPrepareInput(dfList[ii])

#Save dataframes
with pd.ExcelWriter(outputPath + curve_tmp) as writer:
    for ii in range(len(dfList)):
        dfList[ii].to_excel(writer, sheet_name="Sheet"+str(ii))  

#Load dataframes to clean them
for ii in range(len(dfList)):    
    dfList[ii] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet"+str(ii))

#Delete useless information from dataframes (2° step)
for ii in range(len(dfList)):
    dfList[ii] = dfList[ii].drop(["Unnamed: 0"], axis=1)
    dfList[ii] = dfList[ii].drop([0]).reset_index(drop = True)
    dfList[ii] = dfList[ii].drop(countryToRemove, axis = 1)

headers = dfList[0].columns.values.tolist()

#Create empy list of dataframes, to fill it later
dfListSplit = [pd.DataFrame()]* len(headers[1:])

#Add column "Maturity" to each dataframe in list
df = dfList[0]["Maturity"].rename("Maturity")
for ii in range(len(dfListSplit)):
    dfListSplit[ii] = pd.concat((dfListSplit[ii],df),axis=1)

#Split dataframe by country
headers.remove("Maturity")
for ii in range(len(headers)):
    for jj in range(len(dfList)):
        df = dfList[jj][headers[ii]].rename(headers[ii]+"_"+curveTypes[jj])
        dfListSplit[ii] = pd.concat((dfListSplit[ii],df),axis=1)

#Calculate 1y forward rates for each country and curve
for ii in range(len(economy)):
    for jj in range(len(curveTypes)):
        countryCurve = headers[ii]+"_"+curveTypes[jj]
        calculate_1y_forwardRates(dfListSplit[ii],countryCurve)
        dfListSplit[ii] = dfListSplit[ii].drop(columns=[countryCurve+"_shift"])
    dfListSplit[ii].insert(0,"Economy",economy[ii])

#Save the final output
with pd.ExcelWriter(outputPath + curve_final) as writer:
    for ii in range(len(economy)):
        dfListSplit[ii].to_excel(writer, sheet_name=economy[ii],index=False)