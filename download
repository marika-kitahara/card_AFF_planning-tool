import pandas as pd

def optimize_budget(df, mode="CV最大"):

    if mode == "CV最大":
        result = df.loc[df.groupby(["date","media"])["cv"].idxmax()]

    elif mode == "CPA最小":
        result = df.loc[df.groupby(["date","media"])["cpa"].idxmin()]

    return result.reset_index(drop=True)