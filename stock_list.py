import pandas as pd
import requests

url = "https://isin.twse.com.tw/isin/C_public.jsp?strMode=2"

headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}

r = requests.get(url, headers=headers, timeout=10)

r.encoding = "big5"

dfs = pd.read_html(r.text)

df = dfs[0]

# 整理表格
df.columns = df.iloc[0]   # 把第 0 列設成欄位名稱
df = df.drop([0, 1])      # 刪除前兩列（標題＋篩選行）
df = df.reset_index(drop=True)

# 拆分「有價證券代號及名稱」
df[['證券代號', '公司名稱']] = df['有價證券代號及名稱'].str.split('　', expand=True)
df = df.drop(columns=['有價證券代號及名稱'])

print(df.head())


"""
print("Status code:", r.status_code)
print("HTML 前 500 字:\n", r.text[:500])

try:
    dfs = pd.read_html(r.text)
    print(f"\n✅ 找到 {len(dfs)} 個表格\n")
    for i, df in enumerate(dfs):
        print(f"\n--- Table {i} ---\n")
        print(df.head())
        print(type(df))
        print(len(df))
    dfs_1=dfs.iloc[0,:]
    dfs_2=dfs.iloc[1:,:]
    print(dfs_1)
    print(dfs_2)
except ValueError:
    print("⚠️ 沒有找到任何 <table> 結構，無法使用 pd.read_html()")
"""