"""
StockMatch_Miya 定数定義モジュール
アプリケーション全体で使用する定数、列設定、初期値、カラーマップを集約
"""

APP_VERSION = "v1.5.0"
APP_NAME = "StockMatch_Miya-v1.5.0"
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

# タブ1（納期超過・照合結果一覧）の全21列定義
# (col_id, 表示タイトル, 推奨列幅, 配置)
COL_CONFIGS_MAIN = [
    ("status_badge", "総合判定", 165, "center"),
    ("mismatch_reason", "不一致理由・照合診断", 320, "w"),
    ("req_type", "依頼区分", 80, "center"),
    ("order_no", "管理No(注文番号)", 115, "center"),
    ("instruction_disp", "指示書(J列)", 80, "center"),
    ("ship_to", "出荷先", 110, "w"),
    ("line_name", "品名・ライン名等", 145, "w"),
    ("product_detail", "製品詳細", 145, "w"),
    ("supplier", "発注先", 130, "w"),
    ("maker", "メーカー", 100, "w"),
    ("item_name", "品名", 145, "w"),
    ("model_spec", "型番・規格(図番)", 165, "w"),
    ("qty", "数量", 60, "center"),
    ("due_date", "注文書納期", 100, "center"),
    ("days_overdue", "基準日超過", 90, "center"),
    ("actual_date", "現品確認日(物)", 110, "center"),
    ("actual_status", "現品状況", 135, "center"),
    ("invoice_date", "請求日(伝票)", 110, "center"),
    ("invoice_status", "請求状況", 135, "center"),
    ("person", "担当", 70, "center"),
    ("source_info", "照合元情報", 190, "w"),
]
COL_TITLES_MAIN = {col_id: title for col_id, title, _, _ in COL_CONFIGS_MAIN}

# タブ1のデフォルト表示列（計15項目 / 品名・現品確認日・現品状況・請求日・請求状況・担当はデフォルト非表示）
DEFAULT_VISIBLE_COLUMNS_MAIN = [
    "status_badge",       # 総合判定
    "mismatch_reason",    # 不一致理由・照合診断
    "req_type",           # 依頼区分
    "order_no",           # 管理No
    "instruction_disp",   # 指示書
    "ship_to",            # 出荷先
    "line_name",          # 品名・ライン名等
    "product_detail",     # 製品詳細
    "supplier",           # 発注先
    "maker",              # メーカー
    "model_spec",         # 型番・規格(図番)
    "qty",                # 数量
    "due_date",           # 注文書納期
    "days_overdue",       # 基準日超過
    "source_info",        # 照合元情報
]

# タブ2（注文書未検出一覧）の全18列定義
COL_CONFIGS_MISSING = [
    ("status_badge", "状態判定", 150, "center"),
    ("mgmt_sheet", "管理表シート", 100, "center"),
    ("mgmt_row", "行番号", 65, "center"),
    ("order_no", "管理No(8桁)", 105, "center"),
    ("instruction_disp", "指示書(J列)", 80, "center"),
    ("ship_to", "出荷先", 110, "w"),
    ("line_name", "品名・ライン名等", 145, "w"),
    ("product_detail", "製品詳細", 145, "w"),
    ("supplier", "業者 / 発注先", 140, "w"),
    ("maker", "メーカー", 100, "w"),
    ("item_name", "品名", 160, "w"),
    ("model_spec", "型番・規格・図番", 180, "w"),
    ("qty", "購入数", 60, "center"),
    ("due_date", "管理表納期", 95, "center"),
    ("actual_date", "現品確認日", 95, "center"),
    ("invoice_date", "請求日", 95, "center"),
    ("person", "担当", 75, "center"),
    ("mismatch_reason", "未検出診断内容", 320, "w"),
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
