import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import re

# ➤ 抽取段落文字
def extract_paragraphs(pdf_path):
    doc = fitz.open(pdf_path)
    paragraphs = []
    for page in doc:
        text = page.get_text()
        for p in text.split('\n\n'):
            p = p.strip()
            if p and not re.match(r"^[\d,.\s]+$", p):  # 排除純數字或空白
                paragraphs.append(p)
    return paragraphs

# ➤ 清理表格欄位標題 & 雜訊列
def clean_table(raw_table):
    if not raw_table or len(raw_table) < 2:
        return None

    header = raw_table[0]
    rows = raw_table[1:]

    # 濾掉空列或全為 None 的列
    cleaned_rows = [
        row for row in rows if any(cell is not None and str(cell).strip() != '' for cell in row)
    ]

    # 修正 header 長度不足
    while len(header) < len(cleaned_rows[0]):
        header.append(f"Unnamed_{len(header)}")

    df = pd.DataFrame(cleaned_rows, columns=header)
    return df

# ➤ 抽取表格
def extract_tables(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            raw_table = page.extract_table()
            df = clean_table(raw_table)
            if df is not None:
                tables.append((f"表格_Page{page_num + 1}", df))
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
            # 限制 Sheet 名稱不超過 31 字元
            df_table.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    print(f"✅ 完成！已輸出：{output_excel}")

def extract_tables(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            raw_table = page.extract_table()
            df = clean_table(raw_table)
            if df is not None:
                # 🔍 嘗試從頁面文字中抓出有意義的表格名稱
                text = page.extract_text()
                title = "表格_Page" + str(page_num + 1)
                if text:
                    for line in text.split('\n'):
                        if any(keyword in line for keyword in ["Assets","Liabilities","stockholders","net sales","cost of sales","Consolidated", "Revenue", "Income", "Balance Sheet"]):
                            title = line.strip()[:31]  # 限制長度
                            break
                tables.append((title, df))
    return tables


# ✅ 程式執行入口
if __name__ == "__main__":
    export_pdf_data_to_excel("fox_factory_2025_Q1.pdf", "Fox財報資料_整理版1.xlsx")
