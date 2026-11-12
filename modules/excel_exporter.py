"""
StockMatch_Miya Excel連携・出力モジュール
COMオートメーションによる一時ファイルなし直接展開、および注文書の誤保存・誤編集を5重ガードする安全読取専用オープンを集約
"""

import os
import sys
import datetime

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    HAS_WIN32COM = True
except ImportError:
    win32com = None
    pythoncom = None
    HAS_WIN32COM = False

from .constants import EXCEL_COLOR_HEX, COL_CONFIGS_MAIN, COL_CONFIGS_MISSING
from .date_utils import (
    rgb_to_bgr_int,
    parse_flexible_date,
    is_order_date_expired,
    is_order_no_excluded,
    is_instruction_matched,
)

# Excel展開用の各列の最小幅・最大幅の定義 (ColumnWidth: 半角文字数換算)
# ・文字数が少ない列（製作数・単数・在庫数・管理No等）はオートフィルター▼でヘッダーが隠れない最小幅を確保
# ・文字数が長くなりやすい列（製品詳細・仕様・不一致理由等）は広がりすぎて画面を圧迫しない最大幅に抑制
EXCEL_COL_WIDTH_LIMITS = {
    "status_badge": (20.0, 26.0),     # 総合判定 / 状態判定
    "mismatch_reason": (28.0, 38.0),  # 不一致理由・照合診断 / 未検出診断内容
    "due_date": (20.0, 26.0),         # 納期
    "project_name": (14.0, 22.0),     # プロジェクト名
    "order_no": (13.0, 16.0),         # 管理No
    "ship_to": (14.0, 20.0),          # 出荷先
    "line_name": (16.0, 24.0),        # 品名、ライン名等
    "product_detail": (18.0, 32.0),   # 製品詳細
    "prod_qty": (11.0, 13.0),         # 製作数
    "drawing_no": (13.0, 18.0),       # 図面番号
    "part_no": (12.0, 16.0),          # 品番
    "item_name": (18.0, 28.0),        # 品名
    "maker": (13.0, 20.0),            # メーカー
    "unit_qty": (11.0, 13.0),         # 単数
    "spec": (18.0, 28.0),             # 仕様
    "supplier": (14.0, 20.0),         # 仕入先
    "stock_qty": (11.0, 13.0),        # 在庫数
    "order_date": (13.0, 16.0),       # 注文日
}

# Excel展開用のデータセル配置 (xlCenter: -4108, xlLeft: -4131)
EXCEL_COL_ALIGNMENTS = {
    "status_badge": -4108,    # 中央
    "mismatch_reason": -4131, # 左
    "due_date": -4108,        # 中央
    "project_name": -4131,    # 左
    "order_no": -4108,        # 中央
    "ship_to": -4131,         # 左
    "line_name": -4131,       # 左
    "product_detail": -4131,  # 左
    "prod_qty": -4108,        # 中央
    "drawing_no": -4131,      # 左
    "part_no": -4131,         # 左
    "item_name": -4131,       # 左
    "maker": -4131,           # 左
    "unit_qty": -4108,        # 中央
    "spec": -4131,            # 左
    "supplier": -4131,        # 左
    "stock_qty": -4108,       # 中央
    "order_date": -4108,      # 中央
}


def format_excel_sheet_columns(ws, col_configs: list[tuple]):
    """
    Excelシートの各列に対して、AutoFit後の幅をベースに
    ・文字やオートフィルター▼アイコンが隠れない最小幅を保証
    ・広すぎる列（製品詳細・仕様・不一致理由等）が画面を破壊しない最大幅を制限
    ・適切な配置（中央揃え・左揃え）を適用
    """
    try:
        ws.Columns.AutoFit()
        for idx, col_info in enumerate(col_configs, start=1):
            col_id = col_info[0]
            min_w, max_w = EXCEL_COL_WIDTH_LIMITS.get(col_id, (12.0, 30.0))
            col_range = ws.Columns(idx)
            try:
                cur_w = float(col_range.ColumnWidth)
            except Exception:
                cur_w = min_w
            # AutoFit後の幅に適度なマージン(+2.0)を加えつつ、[min_w, max_w] の範囲に収める
            optimal_w = max(min_w, min(max_w, round(cur_w + 2.0, 1)))
            col_range.ColumnWidth = optimal_w

            # 2行目以降のデータ領域に適切な配置（中央／左）を適用
            align = EXCEL_COL_ALIGNMENTS.get(col_id, -4131)
            try:
                row_count = ws.UsedRange.Rows.Count
                if row_count > 1:
                    data_col_rng = ws.Range(ws.Cells(2, idx), ws.Cells(row_count, idx))
                    data_col_rng.HorizontalAlignment = align
            except Exception:
                pass
    except Exception:
        pass


def configure_excel_page_setup(ws, title: str):
    """Excelシートに業務帳票としての最適な印刷設定（A4横・1ページ収まり・見出し行リピート）を適用"""
    try:
        ps = ws.PageSetup
        ps.Orientation = 2  # xlLandscape (横向き)
        ps.PaperSize = 9    # xlPaperA4 (A4サイズ)
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        ps.PrintTitleRows = "$1:$1"
        ps.LeftMargin = 20
        ps.RightMargin = 20
        ps.TopMargin = 30
        ps.BottomMargin = 25
        ps.HeaderMargin = 15
        ps.FooterMargin = 15
        today_str = datetime.date.today().strftime("%Y/%m/%d")
        ps.LeftHeader = f"&10&\"Meiryo UI\"【StockMatch_Miya】 {title}"
        ps.RightHeader = f"&9&\"Meiryo UI\"作成日: {today_str}"
        ps.CenterFooter = "&10&\"Meiryo UI\"- &P / &N ページ -"
        ps.PrintGridlines = True
    except Exception:
        pass


def open_excel_read_only_at_location(
    filepath: str,
    sheet_name: str | None = None,
    row_index: int | None = None,
    is_mgmt: bool = False,
) -> tuple[bool, str]:
    """
    指定されたExcelファイルを起動し、該当シート・該当行をハイライト選択した状態で開く。
    注文書原本を誤編集・上書き保存してしまう事故を防止するため、注文書の場合は5重の保護ガードを適用する。
    """
    if not os.path.isfile(filepath):
        return False, f"ファイルが見つかりません:\n{filepath}"

    abs_path = os.path.abspath(filepath)

    if not HAS_WIN32COM:
        # win32com がない場合の安全フォールバック
        if not is_mgmt:
            return False, "win32com未検出のため、注文書原本の安全保護起動ができません。\npip install pywin32 を実行してください。"
        try:
            os.startfile(abs_path)
            return True, "通常モードで起動しました。"
        except Exception as e:
            return False, str(e)

    try:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = True
        try:
            excel.WindowState = -4137  # xlMaximized
        except Exception:
            pass

        norm_target = os.path.normcase(abs_path)
        target_wb = None

        # すでにExcelで開かれているか確認
        try:
            for wb in excel.Workbooks:
                if os.path.normcase(wb.FullName) == norm_target:
                    target_wb = wb
                    break
        except Exception:
            pass

        if target_wb:
            # 注文書ですでに通常モードで開かれていた場合は読み取り専用へ安全に切り替え
            if not is_mgmt and not target_wb.ReadOnly:
                try:
                    target_wb.ChangeFileAccess(Mode=3)  # xlReadOnly
                except Exception:
                    pass
        else:
            # 新規オープン
            if is_mgmt:
                target_wb = excel.Workbooks.Open(abs_path, ReadOnly=False)
            else:
                target_wb = excel.Workbooks.Open(
                    abs_path,
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                )

        target_wb.Activate()

        # 注文書専用のメモリ上ロック保護（セル選択・スクロール・コピーは許可し、セルの書き換えや削除を禁止）
        if not is_mgmt:
            try:
                for ws in target_wb.Worksheets:
                    try:
                        ws.Protect(
                            Password="StockMatch_ReadOnly_Guard",
                            DrawingObjects=False,
                            Contents=True,
                            Scenarios=False,
                            UserInterfaceOnly=True,
                            AllowFormattingCells=True,
                            AllowFormattingColumns=True,
                            AllowFormattingRows=True,
                            AllowInsertingColumns=False,
                            AllowInsertingRows=False,
                            AllowInsertingHyperlinks=False,
                            AllowDeletingColumns=False,
                            AllowDeletingRows=False,
                            AllowSorting=False,
                            AllowFiltering=True,
                            AllowUsingPivotTables=False,
                        )
                    except Exception:
                        pass
                try:
                    target_wb.Protect(
                        Password="StockMatch_ReadOnly_Guard",
                        Structure=True,
                        Windows=False,
                    )
                except Exception:
                    pass
            except Exception:
                pass

        # 該当シートおよび該当行のジャンプ・選択
        target_ws = None
        if sheet_name:
            try:
                target_ws = target_wb.Worksheets(sheet_name)
            except Exception:
                try:
                    s_clean = sheet_name.replace(" ", "")
                    for ws in target_wb.Worksheets:
                        if ws.Name.replace(" ", "") == s_clean:
                            target_ws = ws
                            break
                except Exception:
                    pass

        if target_ws is None:
            target_ws = target_wb.ActiveSheet

        target_ws.Activate()

        if row_index and row_index >= 1:
            try:
                scroll_row = max(1, row_index - 3)
                excel.ActiveWindow.ScrollRow = scroll_row
                target_ws.Rows(row_index).Select()
            except Exception:
                pass

        # 閉じる際の「変更を保存しますか？」確認ポップアップを抑止
        if not is_mgmt:
            try:
                target_wb.Saved = True
            except Exception:
                pass

        return True, "正常にExcelで開きました。"
    except Exception as e:
        return False, f"Excel起動中にエラーが発生しました:\n{e}"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def export_to_excel_direct(
    analyzed_records: list[dict],
    missing_order_records: list[dict],
    excluded_nos: set[str],
    base_date: datetime.date,
    exclude_months: int,
    inst_mode: str,
    cancel_event,
    progress_queue,
) -> tuple[bool, str]:
    """
    一時ファイルを一切作成せず、Excel COMオートメーション経由で
    新規ブック（未保存）を直接起動し、書式・フィルター・判定色・A4横印刷設定を再現して全データを転送する
    """
    if not HAS_WIN32COM:
        return False, "win32com モジュールが未検出です。"

    excel = None
    wb = None
    try:
        def _check_cancelled():
            if cancel_event.is_set():
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                raise Exception("ユーザーによってExcel直接展開が中止されました。")

        pythoncom.CoInitialize()
        progress_queue.put(("PROGRESS", (15.0, "Excelアプリケーションを起動中...")))
        _check_cancelled()

        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        progress_queue.put(("PROGRESS", (25.0, "新規Excelブックを作成中...")))
        _check_cancelled()

        wb = excel.Workbooks.Add()

        while wb.Sheets.Count < 2:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))

        ws1 = wb.Sheets(1)
        ws1.Name = "納期超過・照合結果"

        ws2 = wb.Sheets(2)
        ws2.Name = "注文書未検出(管理表のみ)"

        header_bg = rgb_to_bgr_int(EXCEL_COLOR_HEX["HEADER_BG"])
        header_fg = rgb_to_bgr_int(EXCEL_COLOR_HEX["HEADER_FG"])
        border_color = rgb_to_bgr_int(EXCEL_COLOR_HEX["BORDER"])

        color_map = {
            k: {
                "bg": rgb_to_bgr_int(v["bg"]),
                "fg": rgb_to_bgr_int(v["fg"]),
            }
            for k, v in EXCEL_COLOR_HEX.items()
            if isinstance(v, dict)
        }

        # -------------------------------------------------------------
        # 1. シート1: 納期超過・照合結果
        # -------------------------------------------------------------
        progress_queue.put(("PROGRESS", (45.0, "シート1 (納期超過・照合結果) にデータを転送中...")))
        _check_cancelled()

        headers_main = [title for _, title, _, _ in COL_CONFIGS_MAIN]

        rows_main_data = []
        tags_main = []

        def _is_record_excluded(r):
            o_no = r.get("order_no", "")
            if is_order_no_excluded(o_no, excluded_nos):
                return True
            o_dt = r.get("order_date")
            if is_order_date_expired(o_dt, base_date, exclude_months):
                return True
            inst_raw = r.get("instruction_raw", "")
            if not is_instruction_matched(inst_raw, inst_mode):
                return True
            return False

        for r in analyzed_records:
            if _is_record_excluded(r):
                continue
            rows_main_data.append(
                [
                    r["status_badge"],
                    r["mismatch_reason"],
                    r["due_date_str"],
                    r.get("project_name", ""),
                    r["order_no"],
                    r.get("ship_to", ""),
                    r.get("line_name", ""),
                    r.get("product_detail", ""),
                    r.get("prod_qty", ""),
                    r.get("drawing_no", ""),
                    r.get("part_no", ""),
                    r["item_name"],
                    r["maker"],
                    r.get("unit_qty", ""),
                    r.get("spec", ""),
                    r["supplier"],
                    r.get("stock_qty", ""),
                    r.get("order_date_str", "-"),
                ]
            )
            tags_main.append(r["row_tag"])

        num_main_cols = len(headers_main)
        last_col_letter_main = chr(64 + num_main_cols)

        # ヘッダー書き込み
        header_range_main = ws1.Range(f"A1:{last_col_letter_main}1")
        header_range_main.Value = [headers_main]
        header_range_main.Interior.Color = header_bg
        header_range_main.Font.Color = header_fg
        header_range_main.Font.Bold = True
        header_range_main.Font.Name = "Meiryo UI"
        header_range_main.Font.Size = 10
        header_range_main.RowHeight = 24
        header_range_main.VerticalAlignment = -4108  # xlCenter

        if rows_main_data:
            num_rows = len(rows_main_data)
            data_range = ws1.Range(f"A2:{last_col_letter_main}{num_rows + 1}")
            data_range.Value = rows_main_data
            data_range.Font.Name = "Meiryo UI"
            data_range.Font.Size = 9
            data_range.VerticalAlignment = -4108
            data_range.RowHeight = 22

            # 色分け適用
            for r_i, tag in enumerate(tags_main, start=2):
                if tag in color_map:
                    row_rng = ws1.Range(f"A{r_i}:{last_col_letter_main}{r_i}")
                    row_rng.Interior.Color = color_map[tag]["bg"]
                    row_rng.Font.Color = color_map[tag]["fg"]

            total_range = ws1.Range(f"A1:{last_col_letter_main}{num_rows + 1}")
            for border_id in (7, 8, 9, 10, 11, 12):
                total_range.Borders(border_id).LineStyle = 1
                total_range.Borders(border_id).Color = border_color

        format_excel_sheet_columns(ws1, COL_CONFIGS_MAIN)
        ws1.Range("A2").Select()
        excel.ActiveWindow.FreezePanes = True
        header_range_main.AutoFilter()
        configure_excel_page_setup(ws1, "納期超過・照合結果 一覧")

        # -------------------------------------------------------------
        # 2. シート2: 注文書未検出(管理表のみ)
        # -------------------------------------------------------------
        progress_queue.put(("PROGRESS", (75.0, "シート2 (注文書未検出) にデータを転送中...")))
        _check_cancelled()

        headers_missing = [title for _, title, _, _ in COL_CONFIGS_MISSING]

        rows_missing_data = []
        for r in missing_order_records:
            if _is_record_excluded(r):
                continue
            rows_missing_data.append(
                [
                    r["status_badge"],
                    r["mismatch_reason"],
                    r.get("due_date_str", r.get("due_date", "-")),
                    r.get("project_name", ""),
                    r["order_no"],
                    r.get("ship_to", ""),
                    r.get("line_name", ""),
                    r.get("product_detail", ""),
                    r.get("prod_qty", ""),
                    r.get("drawing_no", ""),
                    r.get("part_no", ""),
                    r["item_name"],
                    r["maker"],
                    r.get("unit_qty", ""),
                    r.get("spec", ""),
                    r["supplier"],
                    r.get("stock_qty", ""),
                    r.get("order_date_str", "-"),
                ]
            )

        num_missing_cols = len(headers_missing)
        last_col_missing = chr(64 + num_missing_cols)

        ws2.Activate()
        header_range_missing = ws2.Range(f"A1:{last_col_missing}1")
        header_range_missing.Value = [headers_missing]
        header_range_missing.Interior.Color = header_bg
        header_range_missing.Font.Color = header_fg
        header_range_missing.Font.Bold = True
        header_range_missing.Font.Name = "Meiryo UI"
        header_range_missing.Font.Size = 10
        header_range_missing.RowHeight = 24
        header_range_missing.VerticalAlignment = -4108

        if rows_missing_data:
            num_rows2 = len(rows_missing_data)
            data_range2 = ws2.Range(f"A2:{last_col_missing}{num_rows2 + 1}")
            data_range2.Value = rows_missing_data
            data_range2.Font.Name = "Meiryo UI"
            data_range2.Font.Size = 9
            data_range2.VerticalAlignment = -4108
            data_range2.RowHeight = 22

            row_rng2 = ws2.Range(f"A2:{last_col_missing}{num_rows2 + 1}")
            row_rng2.Interior.Color = color_map["MISSING_ORDER"]["bg"]
            row_rng2.Font.Color = color_map["MISSING_ORDER"]["fg"]

            total_range2 = ws2.Range(f"A1:{last_col_missing}{num_rows2 + 1}")
            for border_id in (7, 8, 9, 10, 11, 12):
                total_range2.Borders(border_id).LineStyle = 1
                total_range2.Borders(border_id).Color = border_color

        format_excel_sheet_columns(ws2, COL_CONFIGS_MISSING)
        ws2.Range("A2").Select()
        excel.ActiveWindow.FreezePanes = True
        header_range_missing.AutoFilter()
        configure_excel_page_setup(ws2, "注文書未検出(管理表のみ) 一覧")

        # シート1を選択状態にして表示
        ws1.Activate()
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.Visible = True

        try:
            excel.WindowState = -4137  # xlMaximized
        except Exception:
            pass

        return True, "正常にExcelへ直接展開しました。"
    except Exception as e:
        if cancel_event.is_set():
            return False, "処理が中止されました。"
        return False, str(e)
    finally:
        if excel is not None:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
