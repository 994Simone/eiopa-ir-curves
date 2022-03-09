# Function to calculate interest rates forward 1y

def calc_1y_forward_rates(df,curve):
    df[curve+"_shift"] = df[curve].shift(1, fill_value=0)
    df[curve+"_forward"] = (((1 + df[curve]) ** df["Maturity"]) / 
                        ((1 + df[curve+"_shift"]) ** (df["Maturity"] - 1)) - 1)
    return df
