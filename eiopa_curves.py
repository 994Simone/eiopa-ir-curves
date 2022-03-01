# -*- coding: utf-8 -*-
"""
Created on Fri Feb 25 23:08:52 2022

@author: HP557LA
"""

import pandas as pd
import gc
gc.collect()
from rfrPrepareInput_module import rfrPrepareInput

#Input
inputPath = r"C:\Users\HP557LA\Downloads\December 2021 - EN"
outputPath = r"C:\OutputPython"
curve_tmp = "\curve_tmp.xlsx"
curve_final = "\curve_final.xlsx"

#Code starts here
df1 = pd.read_excel(inputPath + "\EIOPA_RFR_20211231_Term_Structures.xlsx",
                                 sheet_name="RFR_spot_with_VA")

df2 = pd.read_excel(inputPath + "\EIOPA_RFR_20211231_Term_Structures.xlsx",
                                 sheet_name="Spot_WITH_VA_shock_UP")

df3 = pd.read_excel(inputPath + "\EIOPA_RFR_20211231_Term_Structures.xlsx",
                                 sheet_name="Spot_WITH_VA_shock_DOWN")

df4 = pd.read_excel(inputPath + "\EIOPA_RFR_20211231_Term_Structures.xlsx",
                                 sheet_name="RFR_spot_no_VA")

df5 = pd.read_excel(inputPath + "\EIOPA_RFR_20211231_Term_Structures.xlsx",
                                 sheet_name="Spot_NO_VA_shock_UP")

df6 = pd.read_excel(inputPath + "\EIOPA_RFR_20211231_Term_Structures.xlsx",
                                 sheet_name="Spot_NO_VA_shock_DOWN")

del inputPath

dfList = [df1,df2,df3,df4,df5,df6]

countryToRemove = ["Austria","Belgium","Cyprus","Estonia","Finland","France",
                   "Germany","Greece","Ireland","Italy","Latvia","Lithuania",
                   "Luxembourg","Malta","Netherlands","Portugal","Slovakia",
                   "Slovenia","Spain", "Liechtenstein"]

count = 0
for jj in dfList:
    dfList[count] = rfrPrepareInput(jj)
    count = count + 1

del jj, count, df1,df2,df3,df4,df5,df6

with pd.ExcelWriter(outputPath + curve_tmp) as writer:
    dfList[0].to_excel(writer, sheet_name="Sheet1")  
    dfList[1].to_excel(writer, sheet_name="Sheet2")
    dfList[2].to_excel(writer, sheet_name="Sheet3")
    dfList[3].to_excel(writer, sheet_name="Sheet4")
    dfList[4].to_excel(writer, sheet_name="Sheet5")
    dfList[5].to_excel(writer, sheet_name="Sheet6")

#Ricarico gli sheet così sono df lavorabili
dfList[0] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet1")
dfList[1] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet2")
dfList[2] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet3")
dfList[3] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet4")
dfList[4] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet5")
dfList[5] = pd.read_excel(outputPath + curve_tmp, sheet_name="Sheet6")

count = 0
for ii in dfList:
    ii = ii.drop(["Unnamed: 0"], axis=1)
    ii = ii.drop([0]).reset_index(drop = True)
    ii = ii.drop(countryToRemove, axis = 1)
    dfList[count] = ii
    count = count + 1

del ii, count, countryToRemove

headers = dfList[0].columns.values.tolist()

#creo df vuoti da inserire nella lista di df
df1 = pd.DataFrame()
df2 = pd.DataFrame()
df3 = pd.DataFrame()
df4 = pd.DataFrame()
df5 = pd.DataFrame()
df6 = pd.DataFrame()
df7 = pd.DataFrame()
df8 = pd.DataFrame()
df9 = pd.DataFrame()
df10 = pd.DataFrame()
df11 = pd.DataFrame()
df12 = pd.DataFrame()
df13 = pd.DataFrame()
df14 = pd.DataFrame()
df15 = pd.DataFrame()
df16 = pd.DataFrame()
df17 = pd.DataFrame()
df18 = pd.DataFrame()
df19 = pd.DataFrame()
df20 = pd.DataFrame()
df21 = pd.DataFrame()
df22 = pd.DataFrame()
df23 = pd.DataFrame()
df24 = pd.DataFrame()
df25 = pd.DataFrame()
df26 = pd.DataFrame()
df27 = pd.DataFrame()
df28 = pd.DataFrame()
df29 = pd.DataFrame()
df30 = pd.DataFrame()
df31 = pd.DataFrame()
df32 = pd.DataFrame()
df33 = pd.DataFrame()

dfListSplit = [df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,df14,df15,
               df16,df17,df18,df19,df20,df21,df22,df23,df24,df25,df26,df27,df28,
               df29,df30,df31,df32,df33]

del df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,df14,df15,df16,df17,df18,df19,df20,df21,df22,df23,df24,df25,df26,df27,df28,df29,df30,df31,df32,df33

curveTypes = ["Central","Central_ShockUp","Central_ShockDown","Central_noVA",
              "Central_noVA_ShockUp","Central_noVA_ShockDown"]

count = 0
df = dfList[0]["Maturity"]
for ii in dfListSplit:
    dfListSplit[count] = pd.concat((dfListSplit[count],df.rename("Maturity")),
                                   axis=1)
    count = count + 1  
  
count = 0
for jj in headers[1:]:
    countCurves = 0
    for ii in dfList:
        df = ii[jj]
        dfListSplit[count] = pd.concat((dfListSplit[count],
                                        df.rename(jj+"_"+curveTypes[countCurves])),
                                        axis=1)
        countCurves = countCurves + 1
    count = count + 1

del count, countCurves, df, ii, jj, dfList

economy = ["EUR","BGN","HRK","CZK","DKK","HUF","ISK","NOK","PLN","RON",
           "RUB","SEK","CHF","GBP",	"AUD","BRL","CAD","CLP","CNY","COP","HKD",
           "INR","JPY","MYR","MXN",	"NZD","SGD","ZAR","KRW","TWD","THB","TRY","USD"]

# calcolo dei forward rates 1y e separo le curve per currency
count = 0
countCrncy = 1
for df in dfListSplit:
    for ii in curveTypes:
        crncy = headers[countCrncy]+"_"
        df[crncy+ii+"_shift"] = df[crncy+ii].shift(1, fill_value=0)
        df[crncy+ii+"_forward"] = (((1 + df[crncy+ii]) ** df["Maturity"]) / 
                            ((1 + df[crncy+ii+"_shift"]) ** (df["Maturity"] - 1)) - 1)
        df = df.drop(columns=[crncy+ii+"_shift"])
    df.insert(0,"Economy",economy[count])
    dfListSplit[count] = df
    countCrncy = countCrncy +1
    count = count + 1

del count, countCrncy,df,crncy,ii

count = 0
with pd.ExcelWriter(outputPath + curve_final) as writer:
    for ii in economy:
        dfListSplit[count].to_excel(writer, sheet_name=ii,index=False)
        count = count + 1

del count, ii