import pandas as pd

def optimize_budget(df, mode="CV最大"):

    # 最適プランではCPA<=0（マイナス/0円）の候補を採用しない。
    # 返金・取消・調整等でcostが負になるローデータは残すが、
    # 営業向けの「最適CPA」候補としては表示対象外にする。
    work = df.copy()
    if "cpa" in work.columns:
        cpa = pd.to_numeric(work["cpa"], errors="coerce")
        work = work.loc[cpa.gt(0)].copy()

    if work.empty:
        return work.reset_index(drop=True)

    if mode == "CV最大":
        result = work.loc[work.groupby(["date","media"])["cv"].idxmax()]

    elif mode in {"CPA最小", "単価最小"}:
        result = work.loc[work.groupby(["date","media"])["cpa"].idxmin()]

    else:
        result = work

    return result.reset_index(drop=True)
