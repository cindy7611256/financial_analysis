import pandas as pd
x=pd.DataFrame()
x['test']=[100,110,210]
x['pct']=x['test'].pct_change()
print(x)