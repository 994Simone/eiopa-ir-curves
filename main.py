import pandas as pd
from module_interest_rates import calculate_1y_forwardRates

# Non-fixed input (user can change them)
inputPath = r"C:\Users\HP557LA\Downloads\December 2021 - EN"
outputPath = r"C:\OutputPython"
curve_tmp = "\curve_tmp.xlsx"
outputFile = "\curve_final.xlsx"

# Fixed input (unless EIOPA changes them)
fileName = "\EIOPA_RFR_20211231_Term_Structures.xlsx"
sheetsName = ["RFR_spot_with_VA","Spot_WITH_VA_shock_UP",
              "Spot_WITH_VA_shock_DOWN","RFR_spot_no_VA","Spot_NO_VA_shock_UP",
              "Spot_NO_VA_shock_DOWN"]
curveTypes = ["Central","Central_ShockUp","Central_ShockDown","Central_noVA",
              "Central_noVA_ShockUp","Central_noVA_ShockDown"]
economy = ["EUR","BGN","HRK","CZK","DKK","HUF","ISK","NOK","PLN","RON","RUB",
            "SEK","CHF","GBP","AUD","BRL","CAD","CLP","CNY","COP","HKD","INR",
            "JPY","MYR","MXN","NZD","SGD","ZAR","KRW","TWD","THB","TRY","USD"]
rowsToRemove = [0,2,3,4,5,6,7,8,9]
colsToUse = [1,2,5,6,8,9,15,16,25,26,28,29,33,34,35,36,37,38,39,40,41,42,43,
              44,45,46,47,48,49,50,51,52,53,54]

# Import external input as dataframe
dfList = [pd.DataFrame()]*len(sheetsName)
for sheetPosition in range(len(sheetsName)):
    dfList[sheetPosition] = pd.read_excel(inputPath + fileName,
                            sheet_name=sheetsName[sheetPosition],skiprows=rowsToRemove,
                            usecols=colsToUse)
    dfList[sheetPosition] = dfList[sheetPosition].rename(columns={"Main menu":"Maturity"})

# Split dataframe by country
headers = dfList[0].columns.values.tolist()
dfListSplit = [pd.DataFrame()]* len(headers[1:])
headers.remove("Maturity")
dfMaturity = dfList[0]["Maturity"]
for headerPosition in range(len(headers)):
    dfListSplit[headerPosition] = pd.concat((dfListSplit[headerPosition],dfMaturity),axis=1)
    for dfPosition in range(len(dfList)):
        df = dfList[dfPosition][headers[headerPosition]].rename(headers[headerPosition]+"_"+curveTypes[dfPosition])
        dfListSplit[headerPosition] = pd.concat((dfListSplit[headerPosition],df),axis=1)

# Calculate 1y forward rates for each country and curve
for economyPosition in range(len(economy)):
    for curvePosition in range(len(curveTypes)):
        countryCurve = headers[economyPosition]+"_"+curveTypes[curvePosition]
        calculate_1y_forwardRates(dfListSplit[economyPosition],countryCurve)
        dfListSplit[economyPosition] = dfListSplit[economyPosition].drop(columns=[countryCurve+"_shift"])
    dfListSplit[economyPosition].insert(0,"Economy",economy[economyPosition])

# Save the final output
with pd.ExcelWriter(outputPath + outputFile) as writer:
    for economyPosition in range(len(economy)):
        dfListSplit[economyPosition].to_excel(writer, sheet_name=economy[economyPosition],index=False)
