import pdfplumber
import fitz  # PyMuPDF
import pandas as pd

# ➤ 抽取 PDF 中的段落文字
def extract_paragraphs(pdf_path):
    doc = fitz.open(pdf_path)
    paragraphs = []
    for page in doc:
        text = page.get_text()
        for p in text.split('\n\n'):
            p = p.strip()
            if p and not p.replace('.', '').isdigit():  # 排除純數字
                paragraphs.append(p)
    return paragraphs

# ➤ 抽取 PDF 中的表格（每頁一個表格）
def extract_tables(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                tables.append((f"Table_{page_num + 1}", df))
    return tables

# ➤ 統整流程：抽段落與表格 ➜ 輸出 Excel
def export_pdf_data_to_excel(pdf_path, output_excel):
    print("📖 正在抽取段落文字...")
    paragraphs = extract_paragraphs(pdf_path)
    df_paragraphs = pd.DataFrame(paragraphs, columns=["Paragraph_Text"])

    print("📊 正在抽取表格...")
    tables = extract_tables(pdf_path)

    print("📁 匯出到 Excel...")
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df_paragraphs.to_excel(writer, sheet_name="Paragraphs", index=False)
        for sheet_name, df_table in tables:
            df_table.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Excel 限 31 字元

    print(f"✅ 已完成！Excel 檔案：{output_excel}")

# ✅ 執行
export_pdf_data_to_excel("fox_factory_2025_Q1.pdf", "Fox財報資料_整理版.xlsx")
