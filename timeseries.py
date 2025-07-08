#valley = np.argmin(underwater) #previous
import numpy as np
import pandas as pd

# 假設你有一個報酬率累積值
returns = pd.Series([0.0, 0.1, -0.05, 0.07, -0.10, 0.04, 0.06])
cum_returns = (1 + returns).cumprod()

# 計算累積最高點（峰值）
rolling_max = cum_returns.cummax()

# 計算 underwater（從最高點下跌的比例）
underwater = cum_returns / rolling_max - 1

valley = underwater.index[np.argmin(underwater)]

print("underwater:",underwater)

print("\nvalley index:", valley)