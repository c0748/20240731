# ソースコード
import openpyxl
import os
from datetime import datetime

wb = openpyxl.Workbook()
current_date = datetime.now()
output_folder = f"請求書_{current_date.strftime('%Y年%m月')}"
# フォルダを作成
os.makedirs(output_folder, exist_ok=True)
# この1行を追記
output_file = f"{output_folder}/請求書_{current_date.strftime('%Y年%m月')}.xlsx"
# この1行を追記
wb.save(output_file)
