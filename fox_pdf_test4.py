import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import re
import camelot # 引入 Camelot 庫

# 定義財務術語映射字典 (英文到中文)
# 確保這裡的中文翻譯與下方類別映射中的鍵一致
financial_terms_map = {
    # 淨銷售額與收入
    "Net sales": "銷貨收入",
    "Net Sales": "淨銷售額",
    "Revenue": "營收",
    "Cost of sales": "銷貨成本",
    "Gross profit": "毛利",

    # 營業費用
    "Operating expenses:": "營業費用：",
    "Goodwill impairment": "商譽減損",
    "General and administrative": "一般及管理費用",
    "Sales and marketing": "銷售及行銷費用",
    "Research and development": "研發費用",
    "Amortization of purchased intangibles": "已購無形資產攤銷",
    "Total operating expenses": "營業費用合計",

    # 營業(虧損)收入及其他
    "(Loss) income from operations": "營業（虧損）／收入",
    "Interest expense": "利息費用",
    "Other (income) expense, net": "其他（收入）／費用, 淨額",
    "Loss before income taxes": "稅前淨損",
    "Benefit from income taxes": "所得稅利益",
    "Net loss": "本期淨損",

    # 歸屬於母公司之損失
    "Net Loss Attributable to FOX Stockholders": "歸屬於FOX股東之本期淨損",
    "Less: net loss attributable to non-controlling interest": "減：非控制權益之淨損",
    "Net loss attributable to non-controlling interest": "非控制權益之淨損",
    "Non-controlling interest": "非控制權益（少數股東權益）",

    # 綜合損失表
    "Other comprehensive (loss) income": "其他綜合（損失）收益",
    "Interest rate swap Change in net unrealized gains": "利率交換：未實現淨收益變動",
    "Reclassification of net gains on interest rate swap to net earnings": "將利率交換的淨收益重分類至（本期）損益",
    "Tax effects": "稅務影響",
    "Net change, net of tax effects": "稅後淨變動",
    "Foreign currency translation adjustments": "外幣換算差額調整",
    "Other comprehensive loss": "其他綜合損失",
    "Comprehensive loss": "綜合損失",
    "Less: comprehensive loss attributable to non-controlling interest": "減：歸屬於非控制性權益之綜合損失（少數股東）",
    "Comprehensive loss attributable to FOX stockholders": "歸屬於FOX股東之綜合損失",
    "The accompanying notes are an integral part of these condensed consolidated financial statements.": "附註為本簡明合併財務報表不可分割的一部分",


    # 資產
    "Assets": "資產",
    "Current assets:": "流動資產：",
    "Cash and cash equivalents": "現金及約當現金",
    "Accounts receivable (net of allowances of...)": "應收帳款（扣除備抵呆帳後）",
    "Inventory": "存貨",
    "Prepaids and other current assets": "預付費用及其他流動資產",
    "Total current assets": "流動資產合計",
    "Property, plant and equipment, net": "不動產、廠房及設備（淨額）",
    "Lease right-of-use assets": "租賃使用權資產",
    "Deferred tax assets, net": "遞延所得稅資產（淨額）",
    "Goodwill": "商譽",
    "Trademarks and brands, net": "商標及品牌權（淨額）",
    "Customer and distributor relationships, net": "客戶及經銷合作關係（淨額）",
    "Core technologies, net": "核心技術（淨額）",
    "Other assets": "其他資產",
    "Total assets": "資產總額",

    # 負債及股東權益
    "Liabilities and stockholders' equity": "負債及股東權益",
    "Current liabilities:": "流動負債：",
    "Accounts payable": "應付帳款",
    "Accrued expenses": "應計費用",
    "Current portion of long-term debt": "長期負債—一年內到期部分",
    "Total current liabilities": "流動負債合計",
    "Revolver": "週轉信貸",
    "Term Loans, less current portion": "分期貸款（扣除一年內到期部分）",
    "Other liabilities": "其他負債",
    "Total liabilities": "負債總額",
    "Commitments and contingencies": "承諾及或有事項（參見附註8）",

    # 股東權益
    "Stockholders’ equity": "股東權益",
    "Preferred stock, $0.001 par value — 10,000 authorized...": "優先股，每股面值0.001美元，核准發行10,000股，目前無發行",
    "Common stock, $0.001 par value — 90,000 authorized...": "普通股，每股面值0.001美元，核准發行90,000股，已發行/流通股數如表",
    "Additional paid-in capital": "資本公積",
    "Treasury stock, at cost; ...": "庫藏股（成本法計算；持有股數如表）",
    "Accumulated other comprehensive (loss) income": "累積其他綜合（損失）收益",
    "Retained earnings": "保留盈餘",
    "Total stockholders’ equity": "股東權益合計",
    "Total liabilities and stockholders’ equity": "負債及股東權益總額",

    # 現金流量表 - 營業活動
    "OPERATING ACTIVITIES:": "營業活動：",
    "Adjustments to reconcile net loss to net cash (used in) provided by operating activities:": "調整項目，使淨損與營業活動產生之現金(流入)流出數相符",
    "Provision for inventory reserve": "存貨跌價（損失）提列",
    "Stock-based compensation": "股票基礎報酬費用",
    "Amortization of acquired inventory step-up": "取得存貨增值攤銷",
    "Amortization of loan fees": "借款費用攤銷",
    "Amortization of deferred gains on prior swap settlements": "前期交換結算遞延收益攤銷",
    "Proceeds from interest rate swap settlements": "利率交換結算所得",
    "Loss on disposal of property and equipment": "處分不動產廠房設備損失",
    "Deferred taxes": "遞延所得稅",
    "Changes in operating assets and liabilities, net of effects of acquisitions:": "營業資產及負債變動（扣除併購影響）",
    "Accounts receivable": "應收帳款變動", # 現金流量表中的應收帳款變動
    "Inventory": "存貨變動", # 現金流量表中的存貨變動
    "Income taxes": "所得稅變動", # 現金流量表中的所得稅變動
    "Prepaids and other assets": "預付費用及其他資產變動",
    "Accounts payable": "應付帳款變動",
    "Accrued expenses and other liabilities": "應計費用及其他負債變動",
    "Net cash (used in) provided by operating activities": "營業活動現金流入（流出）淨額",

    # 現金流量表 - 投資活動
    "INVESTING ACTIVITIES:": "投資活動：",
    "Purchases of property and equipment": "購置不動產、廠房及設備",
    "Acquisitions of businesses, net of cash acquired": "企業合併（扣除取得現金）",
    "Acquisition of other assets, net of cash acquired": "取得其他資產（扣除取得現金）",
    "Net cash used in investing activities": "投資活動現金流出淨額",

    # 現金流量表 - 籌資活動
    "FINANCING ACTIVITIES:": "籌資活動：",
    "Proceeds from revolver": "取得週轉信貸款",
    "Payments on revolver": "償還週轉信貸款",
    "Repayment of term debt": "償還分期貸款",
    "Purchase and retirement of common stock": "購買並註銷普通股",
    "Repurchases from stock compensation program, net": "員工股票酬勞計畫購回（淨額）",
    "Net cash provided by (used in) financing activities": "籌資活動現金流入（流出）淨額",

    # 現金流量表 - 其他
    "EFFECT OF EXCHANGE RATE CHANGES ON CASH AND CASH EQUIVALENTS": "匯率變動對現金及約當現金之影響",
    "CHANGE IN CASH AND CASH EQUIVALENTS": "現金及約當現金增加（減少）",
    "CASH AND CASH EQUIVALENTS—Beginning of period": "期初現金及約當現金",
    "CASH AND CASH EQUIVALENTS—End of period": "期末現金及約當現金",
    "Other expense, net": "其他費用，淨額",
    "Income before income taxes": "稅前利潤",
    "(Benefit) provision for income taxes": "所得稅（利益）費用",
    "Net income": "淨利",

Net income attributable to Fox stockholders: 歸屬於福克斯股東的淨利

Earnings per share: 每股盈餘

Basic: 基本

Diluted: 稀釋

Weighted-average shares used to compute earnings per share: 計算每股盈餘所用的加權平均股數

    # 其他通用術語（來自原始數據）
    "Powered Vehicles Group": "動力車輛事業部",
    "Aftermarket Applications Group": "售後應用事業部",
    "Specialty Sports Group": "專業運動事業部",
    "Adjusted EBITDA": "調整後EBITDA",
    "Net income margin": "淨利率",
    "Adjusted EBITDA margin": "調整後EBITDA利潤率",
    "Unallocated corporate expenses": "未分配公司費用",
    "Acquisition related costs and expenses": "併購相關成本及費用",
    "Purchase accounting inventory fair value adjustment amortization": "購買會計存貨公允價值調整攤銷",

    # 表格標題關鍵字 (用於識別 Sheet 名稱)
    "Condensed Consolidated Balance ": "簡明合併資產負債表",
    "Condensed Consolidated Statemen": "簡明合併損益表",
    "Notes to Condensed Consolidated": "簡明合併財務報表附註",
    "Income from operations": "營業收入",
    "Consolidated Statements of Cash Flows": "合併現金流量表",
    "Consolidated Statements of Comprehensive Loss": "合併綜合損失表",
    "Consolidated Statements of Loss": "合併損益表",
    "Balance Sheet": "資產負債表",
    "Income Statement": "損益表",
    "Cash Flow Statement": "現金流量表",
    "Comprehensive Income": "綜合收益",
    "Comprehensive Loss": "綜合損失",
    "Consolidated Balance Sheets": "合併資產負債表", # 確保大小寫和複數形式匹配
    "Consolidated Statements of Income": "合併損益表", # 確保大小寫和複數形式匹配
    "Consolidated Statements of Comprehensive Income": "合併綜合損益表", # 確保大小寫和複數形式匹配
}

# 定義財務項目類別映射字典 (中文項目到類別)
# 這裡的鍵必須是 financial_terms_map 翻譯後的中文名稱
financial_category_map = {
    # 損益表相關
    "銷貨收入": "收入",
    "銷貨成本": "成本",
    "商譽減損": "營業費用",
    "一般及管理費用": "營業費用",
    "銷售及行銷費用": "營業費用",
    "研發費用": "營業費用",
    "已購無形資產攤銷": "營業費用",
    "利息費用": "非營業項目",
    "其他（收入）／費用, 淨額": "非營業項目",
    "所得稅利益": "稅前/稅後",
    "減：非控制權益之淨損": "稅前/稅後",
    "本期淨損": "稅前/稅後", # 淨損在損益表通常歸類在稅前/稅後結果

    # 資產負債表相關
    "現金及約當現金": "流動資產",
    "應收帳款（扣除備抵呆帳後）": "流動資產",
    "存貨": "流動資產",
    "預付費用及其他流動資產": "流動資產",
    "不動產、廠房及設備（淨額）": "非流動資產",
    "租賃使用權資產": "非流動資產",
    "遞延所得稅資產（淨額）": "非流動資產",
    "商譽": "無形資產",
    "商標及品牌權（淨額）": "無形資產",
    "客戶及經銷合作關係（淨額）": "無形資產",
    "核心技術（淨額）": "無形資產",
    "其他資產": "其他資產",

    "應付帳款": "流動負債",
    "應計費用": "流動負債",
    "長期負債—一年內到期部分": "流動負債",
    "週轉信貸": "流動負債",
    "分期貸款（扣除一年內到期部分）": "非流動負債",
    "其他負債": "非流動負債",
    "非控制權益（少數股東權益）": "權益",
    "股東權益": "權益",
    "優先股，每股面值0.001美元，核准發行10,000股，目前無發行": "權益",
    "普通股，每股面值0.001美元，核准發行90,000股，已發行/流通股數如表": "權益",
    "資本公積": "權益",
    "庫藏股": "權益",
    "累積其他綜合（損失）收益": "權益",
    "保留盈餘": "權益",

    # 現金流量表相關
    "淨損": "營業活動", # 淨損在現金流量表通常是營業活動調整項
    "折舊及攤銷": "營業活動",
    "存貨跌價（損失）提列": "營業活動",
    "股票基礎報酬費用": "營業活動",
    "取得存貨增值攤銷": "營業活動",
    "借款費用攤銷": "營業活動",
    "前期交換結算遞延收益攤銷": "營業活動",
    "利率交換結算所得": "營業活動",
    "處分不動產廠房設備損失": "營業活動",
    "遞延所得稅": "營業活動",
    "應收帳款變動": "營業活動",
    "存貨變動": "營業活動",
    "所得稅變動": "營業活動",
    "預付費用及其他資產變動": "營業活動",
    "應付帳款變動": "營業活動",
    "應計費用及其他負債變動": "營業活動",
    "營業活動現金流入（流出）淨額": "營業活動",
    "購置不動產、廠房及設備": "投資活動",
    "企業合併（扣除取得現金）": "投資活動",
    "取得其他資產（扣除取得現金）": "投資活動",
    "投資活動現金流出淨額": "投資活動",
    "取得週轉信貸款": "籌資活動",
    "償還週轉信貸款": "籌資活動",
    "償還分期貸款": "籌資活動",
    "購買並註銷普通股": "籌資活動",
    "員工股票酬勞計畫購回（淨額）": "籌資活動",
    "籌資活動現金流入（流出）淨額": "籌資活動",
    "匯率變動對現金及約當現金之影響": "其他",
    "現金及約當現金增加（減少）": "其他",
    "期初現金及約當現金": "其他",
    "期末現金及約當現金": "其他",

    # 其他來自原始數據的項目
    "動力車輛事業部": "事業部",
    "售後應用事業部": "事業部",
    "專業運動事業部": "事業部",
    "調整後EBITDA": "調整後指標",
    "淨利率": "利潤率",
    "調整後EBITDA利潤率": "利潤率",
    "未分配公司費用": "公司費用",
    "所得稅（利益）費用": "稅務",
    "併購相關成本及費用": "併購相關",
    "購買會計存貨公允價值調整攤銷": "併購相關",
    "營業費用：": "營業費用", # 確保標題行也能被歸類
    "營業（虧損）／收入": "營業結果",
    "資產": "總計",
    "負債及股東權益": "總計",
    "流動資產：": "流動資產",
    "流動負債：": "流動負債",
    "營業活動：": "營業活動",
    "投資活動：": "投資活動",
    "籌資活動：": "籌資活動",
}


# ➤ 抽取段落文字
def extract_paragraphs(pdf_path):
    doc = fitz.open(pdf_path)
    paragraphs = []
    for page in doc:
        text = page.get_text()
        for p in text.split('\n\n'):
            p = p.strip()
            # 排除純數字、純空白或只包含標點符號的行
            if p and not re.match(r"^[\d,.\s\-\(\)]+$", p): 
                paragraphs.append(p)
    doc.close() # 關閉文檔
    return paragraphs

# ➤ 清理表格欄位標題 & 雜訊列，並進行翻譯和組合
def clean_table(df_input):
    if df_input.empty:
        return None

    df_cleaned = df_input.copy()

    # 濾掉所有儲存格都是空字串或只包含空白的行
    all_empty_rows = df_cleaned.apply(lambda row: all(str(cell).strip() == '' for cell in row), axis=1)
    df_cleaned = df_cleaned[~all_empty_rows]

    if df_cleaned.empty: # 如果清理後沒有有效的行，則返回 None
        return None

    # --- 新增翻譯、組合和類別邏輯 ---
    # 假設財報項目名稱總是在第一列 (索引為 0)
    if len(df_cleaned.columns) > 0:
        # 1. 儲存原始英文項目名稱
        original_english_items = df_cleaned.iloc[:, 0].astype(str).copy()

        # 2. 將第一列的內容轉換為字串，然後使用映射字典進行替換 (翻譯成中文)
        def translate_item(item_text, mapping_dict):
            # 嘗試精確匹配
            if item_text in mapping_dict:
                return mapping_dict[item_text]
            # 嘗試部分匹配（從長到短，避免短詞覆蓋長詞）
            for eng_term in sorted(mapping_dict.keys(), key=len, reverse=True):
                if eng_term in item_text:
                    # 如果找到部分匹配，替換並返回
                    return item_text.replace(eng_term, mapping_dict[eng_term])
            return item_text # 如果都沒有匹配，返回原始文本
        
        df_cleaned.iloc[:, 0] = original_english_items.apply(lambda x: translate_item(x, financial_terms_map))
        
        # 3. 創建一個新的欄位，包含英文和中文的組合
        combined_items = []
        for i, eng_item in enumerate(original_english_items):
            # 獲取已經翻譯好的中文項目（可能因為沒有匹配而仍然是英文）
            translated_chinese_item = df_cleaned.iloc[i, 0] 
            
            # 檢查原始英文項目是否在我們的映射字典中（精確匹配）
            if eng_item in financial_terms_map:
                combined_items.append(f"{eng_item}: {financial_terms_map[eng_item]}")
            else:
                # 如果不在字典中，則只顯示翻譯後的項目（因為翻譯函數可能已經處理了部分匹配）
                combined_items.append(translated_chinese_item) 
                
        # 4. 根據翻譯後的中文項目，獲取其類別
        categories = []
        for translated_item in df_cleaned.iloc[:, 0]:
            # 嘗試精確匹配
            if translated_item in financial_category_map:
                categories.append(financial_category_map[translated_item])
            else:
                # 如果沒有精確匹配，嘗試部分匹配（從長到短）
                found_category = "未分類" # 默認為未分類
                for chi_term in sorted(financial_category_map.keys(), key=len, reverse=True):
                    if chi_term in translated_item:
                        found_category = financial_category_map[chi_term]
                        break
                categories.append(found_category)
        
        # 將新的組合欄位插入到 DataFrame 的最前面
        df_cleaned.insert(0, '項目類別', categories) # 先插入類別
        df_cleaned.insert(1, '項目名稱 (中英文)', combined_items) # 再插入中英文組合，使其在類別之後
        
    # --- 翻譯、組合和類別邏輯結束 ---

    return df_cleaned

# ➤ 抽取表格 (優先使用 Camelot，失敗則回退到 pdfplumber)
def extract_tables(pdf_path):
    tables = []
    doc = fitz.open(pdf_path) # 用 fitz 打開一次，方便獲取頁面文本用於標題
    
    print("📊 嘗試使用 Camelot 抽取表格...")
    try:
        # 'stream' 模式適用於沒有明確線條的表格，'lattice' 模式適用於有線條的表格
        # 您可能需要根據實際 PDF 的表格類型調整 flavor
        # 也可以嘗試指定 table_areas 或 row_tol 等參數來優化抽取
        camelot_tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        print(f"✅ Camelot 成功抽取到 {len(camelot_tables)} 個表格。")

        for table_obj in camelot_tables:
            df_raw = table_obj.df # Camelot 直接返回 DataFrame
            df = clean_table(df_raw) # 呼叫 clean_table 處理 DataFrame
            
            if df is not None:
                # 🔍 嘗試從頁面文字中抓出有意義的表格名稱
                page_text = doc[table_obj.page - 1].get_text() # 使用 fitz 獲取頁面文本，注意頁碼是從 0 開始
                title = "表格_Page" + str(table_obj.page) # 使用 Camelot 的頁碼 (1-based)
                if page_text:
                    for line in page_text.split('\n'):
                        # 檢查行中是否包含關鍵字來判斷表格標題
                        # 這裡也使用 InStr 進行部分匹配，以提高識別率
                        # 優先匹配長的標題關鍵字
                        for keyword in sorted(financial_terms_map.keys(), key=len, reverse=True):
                            # 限制只匹配表格標題相關的關鍵字，避免將項目名稱誤認為表格標題
                            if keyword in line and any(k in keyword for k in ["Balance", "Statemen", "Notes", "Income", "Cash Flow", "Comprehensive", "Consolidated"]): 
                                title = financial_terms_map.get(keyword, line.strip())[:31] # 嘗試翻譯標題，如果沒有則使用原始行
                                break
                        if title != "表格_Page" + str(table_obj.page): # 如果找到標題了就跳出
                            break
                tables.append((title, df))

    except Exception as e:
        print(f"⚠️ 使用 Camelot 抽取表格時發生錯誤: {e}")
        print("請確保您已安裝 Ghostscript 並正確配置環境變數。")
        print("將回退到使用 pdfplumber 抽取表格作為備用方案...")
        
        # 回退到 pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                raw_table = page.extract_table()
                if raw_table and len(raw_table) > 1:
                    # 將 pdfplumber 的原始列表轉換為 DataFrame 再傳給 clean_table
                    df = clean_table(pd.DataFrame(raw_table[1:], columns=raw_table[0]))
                else:
                    df = None # 如果 raw_table 無效，則為 None
                
                if df is not None:
                    text = page.extract_text()
                    title = "表格_Page" + str(page_num + 1)
                    if text:
                        for line in text.split('\n'):
                            for keyword in sorted(financial_terms_map.keys(), key=len, reverse=True):
                                if keyword in line and any(k in keyword for k in ["Balance", "Statemen", "Notes", "Income", "Cash Flow", "Comprehensive", "Consolidated"]):
                                    title = financial_terms_map.get(keyword, line.strip())[:31]
                                    break
                            if title != "表格_Page" + str(page_num + 1):
                                break
                    tables.append((title, df))
    
    doc.close() # 關閉文檔
    return tables

# ➤ 統整 & 輸出 Excel
def export_pdf_data_to_excel(pdf_path, output_excel):
    print("📖 正在抽取段落文字...")
    paragraphs = extract_paragraphs(pdf_path)
    df_paragraphs = pd.DataFrame(paragraphs, columns=["段落內容 Paragraph_Text"])

    print("📊 正在抽取與清理表格...")
    tables = extract_tables(pdf_path)

    print("📁 匯出到 Excel...")
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df_paragraphs.to_excel(writer, sheet_name="段落_Paragraphs", index=False)
        for sheet_name, df_table in tables:
            # 限制 Sheet 名稱不超過 31 字元 (Excel 的限制)
            df_table.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    print(f"✅ 完成！已輸出：{output_excel}")

# ✅ 程式執行入口
if __name__ == "__main__":
    # 請將 "fox_factory_2025_Q1.pdf" 替換為您的 PDF 文件路徑
    # 請將 "Fox財報資料_整理版1.xlsx" 替換為您希望輸出的 Excel 文件路徑
    export_pdf_data_to_excel("fox_factory_2024_Q4.pdf", "Fox財報資料_2024_Q4.xlsx")
