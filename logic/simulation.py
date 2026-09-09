import pandas as pd
from config.constants import BUDGET_STEP, UP_RATE, DOWN_RATE


def simulate_plan(df):
    """
    松竹梅を「単価(CPA)」基準でシミュレーションする。

    基礎単価 = 基礎COST / 基礎予測CV
    梅 = 基礎単価 - BUDGET_STEP
    竹 = 基礎単価
    松 = 基礎単価 + BUDGET_STEP

    CVは従来どおり単価変更に応じて UP_RATE / DOWN_RATE で変化させ、
    COSTは「予測CV × 採用単価」で算出する。
    単価が0円以下になる候補は営業向けプランとして成立しないため生成しない。
    """
    results = []

    for _, row in df.iterrows():
        base_cv = float(pd.to_numeric(row.get("forecast_cv", 0), errors="coerce") or 0)
        base_cost = float(pd.to_numeric(row.get("cost", 0), errors="coerce") or 0)

        # CVがない場合は意味のある単価を作れないため候補外。
        if base_cv <= 0:
            continue

        base_unit_price = base_cost / base_cv

        for label, unit_delta in {
            "梅": -BUDGET_STEP,
            "竹": 0,
            "松": BUDGET_STEP,
        }.items():
            unit_price = base_unit_price + unit_delta

            # マイナス/0円単価は最適プラン候補にしない。
            if unit_price <= 0:
                continue

            if unit_delta >= 0:
                multiplier = UP_RATE ** (unit_delta / BUDGET_STEP)
            else:
                multiplier = DOWN_RATE ** abs(unit_delta / BUDGET_STEP)

            new_cv = base_cv * multiplier
            new_cost = new_cv * unit_price

            result = {
                "date": row["date"],
                "media": row["media"],
                "plan": label,
                "cv": new_cv,
                "cost": new_cost,
                # CPA = 採用単価。最適プラン/手動設定の単価にもこの値が連動する。
                "cpa": unit_price,
            }
            if "CPN名" in row.index:
                result["CPN名"] = row["CPN名"]
            if "planning_segment" in row.index:
                result["planning_segment"] = row["planning_segment"]
            results.append(result)

    return pd.DataFrame(results)
