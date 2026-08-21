import pandas as pd
from config.constants import BUDGET_STEP, UP_RATE, DOWN_RATE

def simulate_plan(df):

    results = []

    for _, row in df.iterrows():

        base_cv = row["forecast_cv"]
        base_cost = row["cost"]

        for label, delta in {
            "梅": -BUDGET_STEP,
            "竹": 0,
            "松": BUDGET_STEP,
        }.items():

            if delta >= 0:
                multiplier = UP_RATE ** (delta / BUDGET_STEP)
            else:
                multiplier = DOWN_RATE ** abs(delta / BUDGET_STEP)

            new_cv = base_cv * multiplier
            new_cost = base_cost + delta

            cpa = new_cost / new_cv if new_cv != 0 else 0

            # ✅ ←ここが超重要！！
            result = {
                "date": row["date"],
                "media": row["media"],
                "plan": label,
                "cv": new_cv,
                "cost": new_cost,
                "cpa": cpa,
            }
            # 複合施策時に、どの施策由来の予測かを後工程まで保持する。
            if "CPN名" in row.index:
                result["CPN名"] = row["CPN名"]
            if "planning_segment" in row.index:
                result["planning_segment"] = row["planning_segment"]
            results.append(result)

    return pd.DataFrame(results)
