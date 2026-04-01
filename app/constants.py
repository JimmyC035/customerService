import os
import sys

# 處理 PyInstaller 打包後的檔案路徑
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = os.path.join(BASE_DIR, 'master_data.xlsx')
PRODUCT_DB_FILE = os.path.join(BASE_DIR, 'product_db.xlsx')

# 欄位定義
COLS = [
    "日期", "訂購人", "電話", "地址", "贈品", "品項", "備註", "尺寸",
    "數量", "單價", "總價", "折扣", "特殊折扣", "實拿",
    "成本(品項+贈品)", "利潤", "付款方式", "是否已付"
]

# 色彩主題
COLORS = {
    "bg": "#eef2f7",
    "card": "#ffffff",
    "primary": "#4a6fa5",
    "primary_hover": "#3b5d8c",
    "success": "#2e7d32",
    "text": "#2c3e50",
    "text_light": "#7f8c8d",
    "border": "#dce3eb",
    "input_bg": "#f8fafc",
    "table_header": "#4a6fa5",
    "table_row_alt": "#f5f8fb",
    "table_selected": "#4a6fa5",
}

# 表格欄位寬度
COL_WIDTHS = {
    "日期": 100, "訂購人": 80, "電話": 110, "地址": 200,
    "贈品": 80, "品項": 100, "備註": 120, "尺寸": 70,
    "數量": 55, "單價": 80, "總價": 80, "折扣": 60,
    "特殊折扣": 75, "實拿": 80, "成本(品項+贈品)": 120,
    "利潤": 75, "付款方式": 80, "是否已付": 75,
}

# 付款方式選項
PAYMENT_METHODS = ["現金", "刷卡", "匯款", "第三方支付"]
PAYMENT_STATUS = ["是", "否"]
