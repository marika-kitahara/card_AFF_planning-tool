# 固定値は、業務上明示されたものだけに限定する。
MAGITOKU_AFTER_FACTOR = 0.90

BUDGET_STEP = 1000
UP_RATE = 1.15
DOWN_RATE = 0.85

# 係数算出時の参照期間
RECENT_NORMAL_DAYS = 28
TRANSITION_NORMAL_DAYS = 28

# 大規模媒体判定
PREMIUM_MEDIA_KEYWORDS = ("ハピタス", "モッピー", "LINE")

# 定常予測の学習設定
RECENCY_MONTH_DECAY = 0.80  # 1か月古い実績は直近月の80%の重み
OUTLIER_MIN_DAYS = 14      # これ未満の日数では異常値を自動除外しない
OUTLIER_MAD_Z_THRESHOLD = 3.5  # log1p(CV)のMADベースrobust z-score閾値
FACTOR_PRIOR_DAYS = 14       # 係数の少数実績を1.0側へ縮める事前日数
