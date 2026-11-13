"""
StockMatch_Miya 定数定義モジュール
アプリケーション全体で使用する定数、列設定、初期値、カラーマップを集約
"""

APP_VERSION = "v1.5.1"
APP_NAME = f"StockMatch_Miya-{APP_VERSION}"
APP_TITLE = f"{APP_NAME} - 注文品納期・現品/請求照合システム"
APP_USER_MODEL_ID = f"StockMatch.Miya.{APP_VERSION}"

# 指示書列(J列)のフィルターオプション
INSTRUCTION_FILTER_OPTIONS = [
    ("HAS_ONLY", "「有り」のみ"),
    ("HAS_OR_BLANK", "有り または 空白 のみ"),
    ("BLANK_ONLY", "「空白」のみ"),
    ("ALL", "すべて表示 (フィルタなし)"),
]
INSTRUCTION_CODE_TO_LABEL = dict(INSTRUCTION_FILTER_OPTIONS)
INSTRUCTION_LABEL_TO_CODE = {lbl: code for code, lbl in INSTRUCTION_FILTER_OPTIONS}

# タブ1（納期超過・照合結果一覧）の全18列定義
# (col_id, 表示タイトル, 推奨列幅, 配置)
COL_CONFIGS_MAIN = [
    ("status_badge", "総合判定", 155, "center"),
    ("mismatch_reason", "不一致理由・照合診断", 260, "w"),
    ("due_date", "納期", 175, "center"),
    ("project_name", "プロジェクト名", 150, "w"),
    ("order_no", "管理No", 115, "center"),
    ("ship_to", "出荷先", 120, "w"),
    ("line_name", "品名、ライン名等", 165, "w"),
    ("product_detail", "製品詳細", 190, "w"),
    ("prod_qty", "製作数", 95, "center"),
    ("drawing_no", "図面番号", 150, "w"),
    ("part_no", "品番", 135, "w"),
    ("item_name", "品名", 175, "w"),
    ("maker", "メーカー", 135, "w"),
    ("unit_qty", "単数", 90, "center"),
    ("spec", "仕様", 180, "w"),
    ("supplier", "仕入先", 145, "w"),
    ("stock_qty", "在庫数", 95, "center"),
    ("order_date", "注文日", 120, "center"),
]
COL_TITLES_MAIN = {col_id: title for col_id, title, _, _ in COL_CONFIGS_MAIN}

# タブ1のデフォルト表示列（計18項目）
DEFAULT_VISIBLE_COLUMNS_MAIN = [col_id for col_id, _, _, _ in COL_CONFIGS_MAIN]

# タブ2（注文書未検出一覧）の全18列定義
COL_CONFIGS_MISSING = [
    ("status_badge", "状態判定", 155, "center"),
    ("mismatch_reason", "未検出診断内容", 260, "w"),
    ("due_date", "納期", 175, "center"),
    ("project_name", "プロジェクト名", 150, "w"),
    ("order_no", "管理No", 115, "center"),
    ("ship_to", "出荷先", 120, "w"),
    ("line_name", "品名、ライン名等", 165, "w"),
    ("product_detail", "製品詳細", 190, "w"),
    ("prod_qty", "製作数", 95, "center"),
    ("drawing_no", "図面番号", 150, "w"),
    ("part_no", "品番", 135, "w"),
    ("item_name", "品名", 175, "w"),
    ("maker", "メーカー", 135, "w"),
    ("unit_qty", "単数", 90, "center"),
    ("spec", "仕様", 180, "w"),
    ("supplier", "仕入先", 145, "w"),
    ("stock_qty", "在庫数", 95, "center"),
    ("order_date", "注文日", 120, "center"),
]
COL_TITLES_MISSING = {col_id: title for col_id, title, _, _ in COL_CONFIGS_MISSING}

# Excel直接展開用のカラー定義 (BGR 16進数文字列)
EXCEL_COLOR_HEX = {
    "HEADER_BG": "1E293B",
    "HEADER_FG": "FFFFFF",
    "BORDER": "CBD5E1",
    "UNMATCHED": {"bg": "EDE9FE", "fg": "4C1D95"},
    "BOTH_DELAY": {"bg": "FEE2E2", "fg": "7F1D1D"},
    "ACTUAL_DELAY": {"bg": "FFEDD5", "fg": "7C2D12"},
    "INVOICE_DELAY": {"bg": "FEF9C3", "fg": "713F12"},
    "OK": {"bg": "FFFFFF", "fg": "334155"},
    "MISSING_ORDER": {"bg": "FEF3C7", "fg": "92400E"},
}
