import os
import win32com.client as win32
import fitz
from tqdm import tqdm

# 엑셀 파일 경로
excel_file = "Calc.xlsx"

# 셀에 입력할 값들
cell_values = [f"SEC{i}" for i in range(1, 101)]

# 병합된 PDF 파일 경로
merged_pdf_file = "Appendix_A.pdf"

# PDF 파일 경로
pdf_files = []

# 엑셀 애플리케이션 객체 생성
excel_app = win32.Dispatch("Excel.Application")
excel_app.Visible = False
excel_app.DisplayAlerts = False

# 엑셀 파일 열기
workbook = excel_app.Workbooks.Open(os.path.abspath(excel_file))

# ULS_check 시트를 PDF로 출력하는 함수
def export_to_pdf(workbook, sheet_name, output_path):
    sheet = workbook.Sheets(sheet_name)
    sheet.ExportAsFixedFormat(0, output_path)

# 셀에 값 입력 및 저장 및 PDF로 출력
sheet_name = "ULS_check"
range_c2 = workbook.Sheets(sheet_name).Range("C2")

for value in tqdm(cell_values):
    # C2 셀에 값 입력
    range_c2.Value = value
    
    # 엑셀 파일 저장
    workbook.Save()

    # PDF 파일 경로
    pdf_file = f"ULS_check_{value}.pdf"
    pdf_files.append(pdf_file)
    
    # PDF로 출력
    export_to_pdf(workbook, sheet_name, os.path.abspath(pdf_file))

# 엑셀 파일 닫기
workbook.Close()

# 엑셀 애플리케이션 종료
excel_app.Quit()

# PDF 파일 병합
merger = fitz.open()

for pdf_file in pdf_files:
    pdf = fitz.open(pdf_file)
    merger.insert_pdf(pdf)
    pdf.close()

merger.save(merged_pdf_file)
merger.close()