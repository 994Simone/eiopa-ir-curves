import os
import pandas as pd

# Change if there's a new file
input_name = r"\EIOPA_RFR_20211231_Term_Structures.xlsx"

# Fixed input (unless EIOPA changes them)
curveTypes = ["Central","Central_ShockUp","Central_ShockDown","Central_noVA",
              "Central_noVA_ShockUp","Central_noVA_ShockDown"]
sheetsName = ["RFR_spot_with_VA","Spot_WITH_VA_shock_UP",
                  "Spot_WITH_VA_shock_DOWN","RFR_spot_no_VA","Spot_NO_VA_shock_UP",
                  "Spot_NO_VA_shock_DOWN"]
economy = ["EUR","BGN","HRK","CZK","DKK","HUF","ISK","NOK","PLN","RON","RUB",
            "SEK","CHF","GBP","AUD","BRL","CAD","CLP","CNY","COP","HKD","INR",
            "JPY","MYR","MXN","NZD","SGD","ZAR","KRW","TWD","THB","TRY","USD"]

def main():
    # Import input and save into list of dataframes
    dfList = import_input(input_name)
    
    # Calculate 1y forward rates for each country and curve
    dfList = calc_frwd_rates(dfList)
    
    # Save final output in excel and json
    save_result(dfList)


# Functions--------------------------------------------------------------------

# Import external input as dataframe
def import_input(input_name):
    rowsToRemove = [0,2,3,4,5,6,7,8,9]
    colsToUse = [1,2,5,6,8,9,15,16,25,26,28,29,33,34,35,36,37,38,39,40,41,42,43,
                  44,45,46,47,48,49,50,51,52,53,54]
    
    dfList = [pd.DataFrame()]*len(sheetsName)
    
    path = os.path.abspath(r"input"+input_name)
    
    for sheetPosition in range(len(sheetsName)):
        dfList[sheetPosition] = pd.read_excel(path, sheet_name=sheetsName[sheetPosition],
                                              skiprows=rowsToRemove,
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
            df = dfList[dfPosition][headers[headerPosition]].rename(curveTypes[dfPosition])
            dfListSplit[headerPosition] = pd.concat((dfListSplit[headerPosition],df),axis=1)

    return dfListSplit

# Calculate 1y forward rates for each country and curve
def calc_frwd_rates(dfList):
    for economyPosition in range(len(economy)):
        for curvePosition in range(len(curveTypes)):
            calc_1y_forward_rates(dfList[economyPosition],curveTypes[curvePosition])
            dfList[economyPosition] = dfList[economyPosition].drop(columns=[curveTypes[curvePosition]+"_shift"])
        dfList[economyPosition].insert(0,"Economy",economy[economyPosition])
    return dfList

# Function to calculate interest rates forward 1y
def calc_1y_forward_rates(df,curve):
    df[curve+"_shift"] = df[curve].shift(1, fill_value=0)
    df[curve+"_forward"] = (((1 + df[curve]) ** df["Maturity"]) / 
                        ((1 + df[curve+"_shift"]) ** (df["Maturity"] - 1)) - 1)
    return df

def save_result(dfList):
    
    # Create path if not exists
    if not os.path.exists(os.path.abspath("output")):
        os.mkdir("output")
    
    # Save the final output in Excel
    with pd.ExcelWriter(os.path.abspath(r"output\curve_final.xlsx")) as writer:
        for economyPosition in range(len(economy)):
            dfList[economyPosition].to_excel(writer, sheet_name=economy[economyPosition],index=False)
            
    # Save output in .json for unittesting
    dfOutput = pd.DataFrame()
    for df in dfList:
        dfOutput = pd.concat((dfOutput,df)).reset_index(drop=True)
    dfOutput.to_json(os.path.abspath(r"output\curve_final.json"))

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()