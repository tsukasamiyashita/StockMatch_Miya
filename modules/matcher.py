"""
StockMatch_Miya 照合エンジンモジュール
注文品管理表および注文書Excelの読込・パース、ファイル探索、類似度スコアリング、双方向突合・遅延診断を集約
"""

import os
import re
import datetime
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl

try:
    from scipy.optimize import linear_sum_assignment  # type: ignore
    HAS_SCIPY = True
except ImportError:
    linear_sum_assignment = None
    HAS_SCIPY = False

from .date_utils import (
    AnalysisCancelledException,
    parse_flexible_date,
    is_filled_confirmation,
    normalize_key,
    clean_company_name,
    make_bigrams,
    fast_similarity,
    extract_leading_number,
    is_zero_quantity,
    clean_cell_str,
    extract_8digit_no,
    is_order_no_excluded,
    classify_instruction_col,
)


def collect_valid_order_files(
    order_dir: str,
    mgmt_file_path: str,
    valid_mgmt_order_nos: set[str],
    progress_cb=None,
    cancel_event=None,
) -> tuple[list[dict], int, int]:
    """
    注文書フォルダ内（サブフォルダ含む全階層）から有効な注文書Excelファイルを高速探索・読み込み
    枝刈りスキャン（管理表にない管理Noフォルダを丸ごとスキップ）と並列読み込みを実施
    """
    if not os.path.isdir(order_dir):
        return [], 0, 0

    norm_mgmt = os.path.normcase(os.path.abspath(mgmt_file_path)) if mgmt_file_path else ""
    target_files = []
    ignored_count = 0

    # 1. 枝刈り探索
    for root, dirs, files in os.walk(order_dir):
        if cancel_event and cancel_event.is_set():
            raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

        # 管理表にない管理Noフォルダを配下ごと丸ごとスキップ
        if valid_mgmt_order_nos and root != order_dir:
            dir_name = os.path.basename(root)
            d_no8 = extract_8digit_no(dir_name)
            if d_no8 and d_no8 not in valid_mgmt_order_nos:
                dirs.clear()
                ignored_count += len(files)
                continue

        # サブフォルダ名も同様に検査
        if valid_mgmt_order_nos:
            keep_dirs = []
            for d in dirs:
                d_no = extract_8digit_no(d)
                if d_no and d_no not in valid_mgmt_order_nos:
                    ignored_count += 1
                else:
                    keep_dirs.append(d)
            dirs[:] = keep_dirs

        for f in files:
            f_low = f.lower()
            if not f_low.endswith((".xlsx", ".xlsm")) or f.startswith("~$"):
                continue

            full_path = os.path.join(root, f)
            if norm_mgmt and os.path.normcase(os.path.abspath(full_path)) == norm_mgmt:
                continue

            # 見積専用ファイルを事前除外
            f_clean = f.replace(" ", "")
            if "見積" in f_clean and not any(k in f_clean for k in ("注文", "発注")):
                ignored_count += 1
                continue

            f_no8 = extract_8digit_no(f)
            if not f_no8:
                p_dir_name = os.path.basename(root)
                f_no8 = extract_8digit_no(p_dir_name)

            if valid_mgmt_order_nos and f_no8 and f_no8 not in valid_mgmt_order_nos:
                ignored_count += 1
                continue

            target_files.append((full_path, f, f_no8))

    total_target = len(target_files)
    if total_target == 0:
        return [], 0, ignored_count

    all_order_items = []
    valid_file_count = 0

    # 2. マルチスレッド並列パース
    max_workers = min(12, max(2, (os.cpu_count() or 4) * 2))

    def _worker(item):
        if cancel_event and cancel_event.is_set():
            return None, 0
        fpath, fname, fno8 = item
        items = parse_order_book(fpath, fname, fno8, cancel_event=cancel_event)
        return items, len(items)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_worker, it): it for it in target_files}
        done_cnt = 0
        for fut in as_completed(futures):
            if cancel_event and cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
            done_cnt += 1
            if progress_cb and (done_cnt == 1 or done_cnt % 5 == 0 or done_cnt == total_target):
                pct = 35.0 + (38.0 * done_cnt / total_target)
                progress_cb(pct, f"[2/3] 注文書ファイルを高速読込中 ({done_cnt}/{total_target} 件)...")
            res, count = fut.result()
            if res:
                all_order_items.extend(res)
                valid_file_count += 1

    return all_order_items, valid_file_count, ignored_count


def parse_management_book(
    filepath: str,
    progress_cb=None,
    cancel_event=None,
) -> list[dict]:
    """注文品管理表Excelを読み込み、明細行リストを生成して返す"""
    if not os.path.isfile(filepath):
        return []

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    mgmt_rows = []
    sheet_names = wb.sheetnames
    total_sheets = len(sheet_names)

    for s_idx, sheet_name in enumerate(sheet_names, start=1):
        if cancel_event and cancel_event.is_set():
            wb.close()
            raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
        if progress_cb:
            pct = 5.0 + (28.0 * s_idx / max(1, total_sheets))
            progress_cb(
                pct,
                f"[1/3] 注文品管理表を読み込み中 ({s_idx}/{total_sheets} シート: {sheet_name})...",
            )

        s_clean = sheet_name.replace(" ", "")
        if any(k in s_clean for k in ("購入品", "購入", "購入依頼")):
            sheet_cat = "PURCHASE"
            sheet_cat_label = "購入依頼"
        elif any(k in s_clean for k in ("外注加工", "外注", "外依")):
            sheet_cat = "OUTSOURCE"
            sheet_cat_label = "外注依頼"
        else:
            sheet_cat = "OTHER"
            sheet_cat_label = "管理表品目"

        sheet_default_order_no = extract_8digit_no(sheet_name)

        ws = wb[sheet_name]
        raw_rows = []
        empty_streak = 0
        for r_idx, row_tuple in enumerate(ws.iter_rows(max_col=42, values_only=True), start=1):
            if not row_tuple or not any(v is not None and str(v).strip() != "" for v in row_tuple):
                empty_streak += 1
                if r_idx > 20 and empty_streak >= 60:
                    break
                continue
            empty_streak = 0
            raw_rows.append((r_idx, row_tuple))

        if not raw_rows:
            continue

        header_pos_in_list = None
        col_map = {}

        # ヘッダー探索（先頭18行以内）
        for list_i, (r_idx, row_tuple) in enumerate(raw_rows[:18]):
            if r_idx > 25:
                break
            prev_tuple = raw_rows[list_i - 1][1] if list_i > 0 and raw_rows[list_i - 1][0] == r_idx - 1 else ()
            row_vals = []
            for c_idx in range(1, len(row_tuple) + 1):
                v_curr = unicodedata.normalize("NFKC", str(row_tuple[c_idx - 1] or "").strip())
                v_above = (
                    unicodedata.normalize("NFKC", str(prev_tuple[c_idx - 1] or "").strip())
                    if c_idx - 1 < len(prev_tuple)
                    else ""
                )
                combined = re.sub(r"\s+", "", f"{v_above}{v_curr}")
                row_vals.append((c_idx, combined, re.sub(r"\s+", "", v_curr)))

            joined = " ".join(v[1] for v in row_vals)
            if ("品名" in joined or "購入数" in joined) and (
                "図面" in joined or "品番" in joined or "仕様" in joined or "現品" in joined or "メーカー" in joined or "管理" in joined
            ):
                header_pos_in_list = list_i
                for c_idx, v_comb, v_curr in row_vals:
                    if not v_comb:
                        continue
                    v_up = v_comb.upper()

                    if "現品確認" in v_comb or (c_idx == 1 and "現品" in v_comb):
                        col_map.setdefault("actual_date", c_idx)
                    elif "請求日" in v_comb or (c_idx >= 30 and "請求" in v_comb):
                        col_map.setdefault("invoice_date", c_idx)
                    elif any(
                        k in v_up
                        for k in ("管理NO", "管理番号", "注文番号", "発注番号", "手配番号", "工事番号", "工事NO", "製番")
                    ):
                        col_map.setdefault("order_no", c_idx)
                    elif "プロジェクト" in v_comb or "プロジェクト名" in v_comb or "PJ名" in v_comb:
                        col_map.setdefault("project_name", c_idx)
                    elif "発注日" in v_comb or "注文日" in v_comb:
                        col_map.setdefault("order_date", c_idx)
                    elif "図面番号" in v_comb or "図番" in v_comb:
                        col_map.setdefault("drawing_no", c_idx)
                    elif "品番" in v_comb or "型番" in v_comb or "部品番号" in v_comb:
                        col_map.setdefault("part_no", c_idx)
                    elif "出荷先" in v_comb or "納入先" in v_comb or "客先" in v_comb or "得意先" in v_comb:
                        col_map.setdefault("ship_to", c_idx)
                    elif "製品詳細" in v_comb or "詳細" in v_comb:
                        col_map.setdefault("product_detail", c_idx)
                    elif "ライン名" in v_comb or "品名、ライン名" in v_comb or "品名・ライン名" in v_comb:
                        col_map["line_name"] = c_idx
                    elif ("品名" in v_comb or "名称" in v_comb or "商品名" in v_comb) and "品名、ライン名" not in v_comb:
                        if c_idx >= 15:
                            col_map["item_name"] = c_idx
                        elif "item_name" not in col_map:
                            col_map["item_name"] = c_idx
                    elif "仕入先" in v_comb:
                        col_map.setdefault("supplier", c_idx)
                    elif "メーカー" in v_comb or "発注先" in v_comb or "業者" in v_comb:
                        col_map.setdefault("maker", c_idx)
                    elif "製作数" in v_comb or "作数" in v_comb:
                        col_map.setdefault("prod_qty", c_idx)
                    elif "単数" in v_comb:
                        col_map.setdefault("unit_qty", c_idx)
                    elif "仕様" in v_comb or "規格" in v_comb or "材質" in v_comb or "寸法" in v_comb:
                        col_map.setdefault("spec", c_idx)
                    elif "購入数" in v_comb or v_curr == "購入数":
                        col_map["qty"] = c_idx
                    elif "必要数" in v_comb or v_curr == "必要数":
                        col_map.setdefault("req_qty", c_idx)
                    elif "在庫数" in v_comb or v_curr == "在庫数":
                        col_map.setdefault("stock_qty", c_idx)
                    elif ("数量" in v_comb or v_curr in ("数", "発注数")) and "購入数" not in col_map:
                        col_map.setdefault("qty", c_idx)
                    elif "単価" in v_curr and "金額" not in v_curr and "必要数" not in v_curr:
                        col_map.setdefault("unit_price", c_idx)
                    elif "金額" in v_comb or "単価X必要数" in v_comb or "単価×必要数" in v_comb:
                        col_map.setdefault("amount", c_idx)
                    elif "納期" in v_comb or "納入日" in v_comb:
                        col_map.setdefault("due_date", c_idx)
                    elif "担当" in v_comb:
                        col_map.setdefault("person", c_idx)
                    elif "指示書" in v_comb:
                        col_map.setdefault("instruction_doc", c_idx)
                break

        if header_pos_in_list is None:
            continue

        col_map.setdefault("actual_date", 1)       # A列: 現品確認日
        col_map.setdefault("instruction_doc", 10)  # J列: 指示書
        col_map.setdefault("project_name", 11)     # K列: プロジェクト名
        col_map.setdefault("order_no", 13)         # M列: 管理No
        col_map.setdefault("ship_to", 14)          # N列: 出荷先
        col_map.setdefault("line_name", 15)        # O列: 品名、ライン名等
        col_map.setdefault("product_detail", 16)   # P列: 製品詳細
        col_map.setdefault("prod_qty", 17)         # Q列: 製作数
        col_map.setdefault("drawing_no", 18)       # R列: 図面番号
        col_map.setdefault("part_no", 19)          # S列: 品番
        col_map.setdefault("item_name", 20)        # T列: 品名
        col_map.setdefault("maker", 21)            # U列: メーカー
        col_map.setdefault("unit_qty", 22)         # V列: 単数
        col_map.setdefault("spec", 23)             # W列: 仕様
        col_map.setdefault("supplier", 24)         # X列: 仕入先
        col_map.setdefault("stock_qty", 27)        # AA列: 在庫数
        col_map["qty"] = 28                        # AB列: 購入数
        col_map.setdefault("order_date", 33)       # AG列: 注文日
        col_map.setdefault("due_date", 34)         # AH列: 納期
        col_map.setdefault("invoice_date", 35)     # AI列: 請求日

        data_slice = raw_rows[header_pos_in_list + 1 :]

        def _get_col(rtuple: tuple, col_1based: int | None):
            if not col_1based or col_1based < 1 or col_1based > len(rtuple):
                return None
            return rtuple[col_1based - 1]

        last_seen_order_no_8 = sheet_default_order_no

        for r_idx, row_tuple in data_slice:
            project_name_val = _get_col(row_tuple, col_map.get("project_name"))
            drawing_val = _get_col(row_tuple, col_map.get("drawing_no"))
            part_val = _get_col(row_tuple, col_map.get("part_no"))
            item_val = _get_col(row_tuple, col_map.get("item_name"))
            maker_val = _get_col(row_tuple, col_map.get("maker"))
            spec_val = _get_col(row_tuple, col_map.get("spec"))
            ship_to_val = _get_col(row_tuple, col_map.get("ship_to"))
            line_val = _get_col(row_tuple, col_map.get("line_name"))
            prod_detail_val = _get_col(row_tuple, col_map.get("product_detail"))
            order_no_val = _get_col(row_tuple, col_map.get("order_no"))
            prod_qty_val = _get_col(row_tuple, col_map.get("prod_qty"))
            unit_qty_val = _get_col(row_tuple, col_map.get("unit_qty"))
            supplier_val = _get_col(row_tuple, col_map.get("supplier"))
            stock_qty_val = _get_col(row_tuple, col_map.get("stock_qty"))

            qty_val = _get_col(row_tuple, 28)
            if qty_val is None:
                qty_val = _get_col(row_tuple, col_map.get("qty"))

            if not any([project_name_val, drawing_val, part_val, item_val, spec_val, maker_val, line_val, prod_detail_val]):
                continue

            order_date_raw = _get_col(row_tuple, col_map.get("order_date"))
            order_date = parse_flexible_date(order_date_raw, allow_period=False)
            inferred_year = order_date.year if order_date else datetime.date.today().year

            a_col_val = _get_col(row_tuple, 1)
            mapped_actual_val = _get_col(row_tuple, col_map.get("actual_date"))
            actual_raw = a_col_val if is_filled_confirmation(a_col_val) else mapped_actual_val

            ai_col_val = _get_col(row_tuple, 35)
            mapped_inv_val = _get_col(row_tuple, col_map.get("invoice_date"))
            invoice_raw = ai_col_val if is_filled_confirmation(ai_col_val) else mapped_inv_val

            # AH列 (34列目) の納期セル値を優先取得
            ah_col_val = _get_col(row_tuple, 34)
            mapped_due_val = _get_col(row_tuple, col_map.get("due_date"))
            due_raw = ah_col_val if (ah_col_val is not None and str(ah_col_val).strip() != "") else mapped_due_val

            actual_date = parse_flexible_date(actual_raw, default_year=inferred_year, allow_period=False)
            actual_filled = actual_date is not None or is_filled_confirmation(actual_raw)

            invoice_date = parse_flexible_date(invoice_raw, default_year=inferred_year, allow_period=False)
            invoice_filled = invoice_date is not None or is_filled_confirmation(invoice_raw)

            # 注文品管理表の納期(AH列)の日付型判定:
            # セル値が datetime.date または datetime.datetime 型である場合のみ有効な「日付型」と判定
            # 文字列型(メモ・未定等)や非日付型の場合は無視(False)とする
            mgmt_due_is_date_type = False
            mgmt_due_date = None
            if isinstance(due_raw, datetime.datetime):
                mgmt_due_date = due_raw.date()
                mgmt_due_is_date_type = True
            elif isinstance(due_raw, datetime.date):
                mgmt_due_date = due_raw
                mgmt_due_is_date_type = True

            project_name_str = clean_cell_str(project_name_val)
            drawing_str = clean_cell_str(drawing_val)
            part_str = clean_cell_str(part_val)
            item_str = clean_cell_str(item_val)
            maker_str = clean_cell_str(maker_val)
            spec_str = clean_cell_str(spec_val)
            ship_to_str = clean_cell_str(ship_to_val)
            line_str = clean_cell_str(line_val)
            prod_detail_str = clean_cell_str(prod_detail_val)
            prod_qty_str = clean_cell_str(prod_qty_val)
            unit_qty_str = clean_cell_str(unit_qty_val)
            supplier_str = clean_cell_str(supplier_val) or maker_str
            stock_qty_str = clean_cell_str(stock_qty_val)
            qty_str = clean_cell_str(qty_val)
            unit_price_str = clean_cell_str(_get_col(row_tuple, col_map.get("unit_price")))
            amount_str = clean_cell_str(_get_col(row_tuple, col_map.get("amount")))

            raw_order_no_str = clean_cell_str(order_no_val)
            exact_order_no_8 = extract_8digit_no(raw_order_no_str)

            if not exact_order_no_8:
                for c_scan in (13, 11, 2, 3, 4):
                    cand_8 = extract_8digit_no(_get_col(row_tuple, c_scan))
                    if cand_8 and not cand_8.startswith("202"):
                        exact_order_no_8 = cand_8
                        if not raw_order_no_str:
                            raw_order_no_str = cand_8
                        break

            if exact_order_no_8:
                last_seen_order_no_8 = exact_order_no_8
            elif last_seen_order_no_8:
                exact_order_no_8 = last_seen_order_no_8

            order_no_str = raw_order_no_str or exact_order_no_8 or sheet_default_order_no
            person_str = clean_cell_str(_get_col(row_tuple, col_map.get("person")))

            # J列 (10列目) の指示書列を読み取り
            j_col_val = _get_col(row_tuple, 10)
            mapped_inst_val = _get_col(row_tuple, col_map.get("instruction_doc"))
            inst_raw_val = j_col_val if (j_col_val is not None and str(j_col_val).strip() != "") else mapped_inst_val
            instruction_raw = clean_cell_str(inst_raw_val)
            instruction_state = classify_instruction_col(instruction_raw)
            instruction_disp = instruction_raw if instruction_raw else "(空白)"

            model_display_parts = [p for p in (drawing_str, part_str, spec_str) if p and p != "-"]
            model_spec_display = " / ".join(dict.fromkeys(model_display_parts))

            norm_drawing = normalize_key(drawing_str)
            norm_part = normalize_key(part_str)
            norm_item = normalize_key(item_str)
            norm_spec = normalize_key(spec_str)
            norm_line = normalize_key(line_str)
            norm_maker = clean_company_name(maker_str)
            norm_combined = normalize_key(f"{drawing_str} {part_str} {item_str} {spec_str} {line_str}")

            mgmt_rows.append(
                {
                    "sheet": sheet_name,
                    "sheet_cat": sheet_cat,
                    "sheet_cat_label": sheet_cat_label,
                    "row": r_idx,
                    "project_name": project_name_str,
                    "order_no": order_no_str,
                    "order_no_8": exact_order_no_8,
                    "norm_order_no": normalize_key(order_no_str),
                    "order_date": order_date,
                    "instruction_raw": instruction_raw,
                    "instruction_disp": instruction_disp,
                    "instruction_state": instruction_state,
                    "ship_to": ship_to_str,
                    "product_detail": prod_detail_str,
                    "prod_qty": prod_qty_str,
                    "drawing_no": drawing_str,
                    "part_no": part_str,
                    "item_name": item_str,
                    "line_name": line_str,
                    "maker": maker_str,
                    "unit_qty": unit_qty_str,
                    "supplier": supplier_str,
                    "stock_qty": stock_qty_str,
                    "spec": spec_str,
                    "model_spec": model_spec_display,
                    "qty": qty_str,
                    "qty_num": extract_leading_number(qty_str),
                    "unit_price": unit_price_str,
                    "amount": amount_str,
                    "actual_date": actual_date,
                    "actual_raw": clean_cell_str(actual_raw),
                    "actual_filled": actual_filled,
                    "invoice_date": invoice_date,
                    "invoice_raw": clean_cell_str(invoice_raw),
                    "invoice_filled": invoice_filled,
                    "mgmt_due_date": mgmt_due_date,
                    "mgmt_due_is_date_type": mgmt_due_is_date_type,
                    "mgmt_due_raw": clean_cell_str(due_raw),
                    "person": person_str,
                    "norm_drawing": norm_drawing,
                    "bg_drawing": make_bigrams(norm_drawing),
                    "norm_part": norm_part,
                    "bg_part": make_bigrams(norm_part),
                    "norm_item": norm_item,
                    "bg_item": make_bigrams(norm_item),
                    "norm_line": norm_line,
                    "bg_line": make_bigrams(norm_line),
                    "norm_spec": norm_spec,
                    "bg_spec": make_bigrams(norm_spec),
                    "norm_maker": norm_maker,
                    "bg_maker": make_bigrams(norm_maker),
                    "norm_combined": norm_combined,
                    "bg_combined": make_bigrams(norm_combined),
                    "matched": False,
                    "matched_by": "",
                }
            )

    wb.close()
    return mgmt_rows


def parse_order_book(
    filepath: str,
    filename: str,
    file_order_no_8: str | None = None,
    cancel_event=None,
) -> list[dict]:
    """1つの注文書Excelファイルを読み込み、明細行のリストを生成して返す"""
    extracted_items = []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    except Exception:
        return []

    file_supplier = ""
    f_clean = filename.replace(" ", "")
    m_sup = re.search(r"[【\[\(]([^】\]\)]+)[】\]\)]", filename)
    if m_sup:
        file_supplier = m_sup.group(1).strip()
    elif "_" in filename:
        parts = filename.split("_")
        if len(parts) >= 2 and parts[1].strip():
            file_supplier = parts[1].strip()

    try:
        for sheet_name in wb.sheetnames:
            if cancel_event and cancel_event.is_set():
                break

            s_clean = sheet_name.replace(" ", "")
            if "見積" in s_clean and not any(k in s_clean for k in ("注文", "発注")):
                continue

            ws = wb[sheet_name]
            raw_rows = []
            empty_streak = 0
            row2_text = ""
            skip_as_estimate = False

            for r_idx, row_tuple in enumerate(ws.iter_rows(max_row=220, max_col=28, values_only=True), start=1):
                if not row_tuple or not any(v is not None and str(v).strip() != "" for v in row_tuple):
                    empty_streak += 1
                    if r_idx > 20 and empty_streak >= 18:
                        break
                    continue
                empty_streak = 0

                # 2行目のタイトルを検査し、見積依頼書なら即座に打ち切る
                if r_idx == 2:
                    row2_text = " ".join(
                        unicodedata.normalize("NFKC", str(v or "")).strip() for v in row_tuple if v is not None
                    )
                    if "見積" in row2_text:
                        skip_as_estimate = True
                        break
                    if row2_text and not any(k in row2_text for k in ("注文", "発注")):
                        if "見積" in filename or "見積" in sheet_name:
                            skip_as_estimate = True
                            break

                raw_rows.append((r_idx, row_tuple))

            if skip_as_estimate or not raw_rows:
                continue

            if not row2_text and ("見積" in filename or "見積" in sheet_name) and not any(
                k in filename or k in sheet_name for k in ("注文", "発注")
            ):
                continue

            sheet_order_no = file_order_no_8
            sheet_supplier = file_supplier
            sheet_order_date = None
            default_year = datetime.date.today().year

            # 4行目にある注文日付の明示的探索
            for r_idx, row_tuple in raw_rows:
                if r_idx == 4:
                    for cell_v in row_tuple:
                        if cell_v is not None and str(cell_v).strip() != "":
                            d_cand4 = parse_flexible_date(cell_v, allow_period=False)
                            if d_cand4:
                                sheet_order_date = d_cand4
                                default_year = d_cand4.year
                                break
                    break

            for r_idx, row_tuple in raw_rows:
                if r_idx > 16:
                    break
                for c_idx in range(1, min(21, len(row_tuple) + 1)):
                    cell_val = str(row_tuple[c_idx - 1] or "").strip()
                    if not cell_val:
                        continue

                    if not sheet_order_no and (
                        "注文番号" in cell_val or "発注番号" in cell_val or "管理No" in cell_val or "No." in cell_val
                    ):
                        m_in = extract_8digit_no(cell_val)
                        if m_in:
                            sheet_order_no = m_in
                        else:
                            for offset in range(1, 9):
                                if c_idx - 1 + offset >= len(row_tuple):
                                    break
                                right_val = str(row_tuple[c_idx - 1 + offset] or "").strip()
                                m_r8 = extract_8digit_no(right_val)
                                if m_r8:
                                    sheet_order_no = m_r8
                                    break

                    if "注文日" in cell_val or "発注日" in cell_val:
                        d_in_cell = parse_flexible_date(cell_val, allow_period=False)
                        if d_in_cell:
                            sheet_order_date = d_in_cell
                            default_year = d_in_cell.year
                        else:
                            for offset in range(1, 9):
                                if c_idx - 1 + offset >= len(row_tuple):
                                    break
                                d_candidate = parse_flexible_date(row_tuple[c_idx - 1 + offset], allow_period=False)
                                if d_candidate:
                                    sheet_order_date = d_candidate
                                    default_year = d_candidate.year
                                    break

                    if "御中" in cell_val:
                        cleaned_sup = re.sub(r"(御中|様|殿)$", "", cell_val).strip()
                        if cleaned_sup:
                            for offset in range(1, 6):
                                if c_idx - 1 - offset >= 0:
                                    left_val = str(row_tuple[c_idx - 1 - offset] or "").strip()
                                    if left_val and not any(k in left_val for k in ("注文", "発注", "No")):
                                        cleaned_sup = f"{left_val} {cleaned_sup}".strip()
                                        break
                            sheet_supplier = cleaned_sup

            # 明細ヘッダー探索（6〜25行目）
            header_idx = None
            col_map = {}
            for list_i, (r_idx, row_tuple) in enumerate(raw_rows):
                if r_idx < 6:
                    continue
                if r_idx > 25:
                    break
                row_str = " ".join(unicodedata.normalize("NFKC", str(v or "")).strip() for v in row_tuple)
                if ("品名" in row_str or "数量" in row_str) and (
                    "図面" in row_str or "品番" in row_str or "型番" in row_str or "仕様" in row_str or "単価" in row_str or "納期" in row_str
                ):
                    header_idx = list_i
                    for c_idx in range(1, len(row_tuple) + 1):
                        v = unicodedata.normalize("NFKC", str(row_tuple[c_idx - 1] or "")).strip()
                        if not v:
                            continue
                        if "図面番号" in v or "図番" in v:
                            col_map.setdefault("drawing_no", c_idx)
                        elif "品番" in v or "型番" in v or "部品番号" in v:
                            col_map.setdefault("part_no", c_idx)
                        elif "品名" in v or "名称" in v:
                            col_map.setdefault("item_name", c_idx)
                        elif "メーカー" in v or "製造元" in v:
                            col_map.setdefault("maker", c_idx)
                        elif "仕様" in v or "規格" in v or "寸法" in v:
                            col_map.setdefault("spec", c_idx)
                        elif "数量" in v or "数" == v:
                            col_map.setdefault("qty", c_idx)
                        elif "単価" in v:
                            col_map.setdefault("unit_price", c_idx)
                        elif "金額" in v:
                            col_map.setdefault("amount", c_idx)
                        elif "納期" in v or "希望納期" in v or "納入期日" in v:
                            col_map.setdefault("due_date", c_idx)
                    break

            if header_idx is None:
                continue

            def _get_val(rtuple: tuple, col_1b: int | None):
                if not col_1b or col_1b < 1 or col_1b > len(rtuple):
                    return None
                return rtuple[col_1b - 1]

            data_rows = raw_rows[header_idx + 1 :]
            for r_idx, row_tuple in data_rows:
                drawing_val = _get_val(row_tuple, col_map.get("drawing_no"))
                part_val = _get_val(row_tuple, col_map.get("part_no"))
                item_val = _get_val(row_tuple, col_map.get("item_name"))
                maker_val = _get_val(row_tuple, col_map.get("maker"))
                spec_val = _get_val(row_tuple, col_map.get("spec"))
                qty_val = _get_val(row_tuple, col_map.get("qty"))
                due_val = _get_val(row_tuple, col_map.get("due_date"))

                if not any([drawing_val, part_val, item_val, spec_val, maker_val, qty_val]):
                    continue

                d_str = clean_cell_str(drawing_val)
                p_str = clean_cell_str(part_val)
                i_str = clean_cell_str(item_val)
                m_str = clean_cell_str(maker_val)
                s_str = clean_cell_str(spec_val)
                q_str = clean_cell_str(qty_val)

                if is_zero_quantity(q_str):
                    continue

                due_date = parse_flexible_date(due_val, default_year=default_year, allow_period=True)

                req_type_label = "購入依頼"
                if any(k in s_clean for k in ("外注", "加工", "外依")) or any(
                    k in filename for k in ("外注", "加工", "外依")
                ):
                    req_type_label = "外注依頼"

                model_spec_parts = [p for p in (d_str, p_str, s_str) if p and p != "-"]
                model_spec_disp = " / ".join(dict.fromkeys(model_spec_parts))

                norm_drawing = normalize_key(d_str)
                norm_part = normalize_key(p_str)
                norm_item = normalize_key(i_str)
                norm_spec = normalize_key(s_str)
                norm_maker = clean_company_name(m_str)
                norm_sup = clean_company_name(sheet_supplier)
                norm_combined = normalize_key(f"{d_str} {p_str} {i_str} {s_str}")

                extracted_items.append(
                    {
                        "file": filename,
                        "full_path": filepath,
                        "sheet": sheet_name,
                        "row": r_idx,
                        "req_type_label": req_type_label,
                        "order_no": sheet_order_no or "",
                        "order_no_8": sheet_order_no or "",
                        "norm_order_no": normalize_key(sheet_order_no or ""),
                        "order_date": sheet_order_date,
                        "supplier": sheet_supplier,
                        "norm_supplier": norm_sup,
                        "maker": m_str,
                        "norm_maker": norm_maker,
                        "drawing_no": d_str,
                        "norm_drawing": norm_drawing,
                        "bg_drawing": make_bigrams(norm_drawing),
                        "part_no": p_str,
                        "norm_part": norm_part,
                        "bg_part": make_bigrams(norm_part),
                        "item_name": i_str,
                        "norm_item": norm_item,
                        "bg_item": make_bigrams(norm_item),
                        "spec": s_str,
                        "norm_spec": norm_spec,
                        "bg_spec": make_bigrams(norm_spec),
                        "model_spec": model_spec_disp,
                        "qty": q_str,
                        "qty_num": extract_leading_number(q_str),
                        "due_date": due_date,
                        "due_raw": clean_cell_str(due_val),
                        "unit_price": clean_cell_str(_get_val(row_tuple, col_map.get("unit_price"))),
                        "amount": clean_cell_str(_get_val(row_tuple, col_map.get("amount"))),
                        "norm_combined": norm_combined,
                        "bg_combined": make_bigrams(norm_combined),
                    }
                )
    finally:
        wb.close()

    return extracted_items


def score_candidate_pair(o_item: dict, m_row: dict) -> float:
    """注文書明細と管理表明細の類似度スコアを算出（単価・金額は除外）"""
    o_sup = o_item.get("norm_supplier", "")
    o_mkr = o_item.get("norm_maker", "")
    m_mkr = m_row.get("norm_maker", "")

    sup_matched = False
    if o_sup and m_mkr:
        if o_sup == m_mkr or o_sup in m_mkr or m_mkr in o_sup:
            sup_matched = True
    if o_mkr and m_mkr:
        if o_mkr == m_mkr or o_mkr in m_mkr or m_mkr in o_mkr:
            sup_matched = True

    if o_sup and m_mkr and len(o_sup) >= 2 and len(m_mkr) >= 2 and not sup_matched:
        if not any(k in m_mkr for k in ("市販", "購入", "外注", "製作")) and not any(k in o_sup for k in ("市販", "購入")):
            return -800.0

    o_d = o_item["norm_drawing"]
    o_p = o_item["norm_part"]
    o_s = o_item["norm_spec"]
    o_i = o_item["norm_item"]

    m_d = m_row["norm_drawing"]
    m_p = m_row["norm_part"]
    m_s = m_row["norm_spec"]
    m_i = m_row["norm_item"]
    m_line = m_row.get("norm_line", "")

    score = 0.0
    strong_evidence = False

    if sup_matched:
        score += 250.0

    if o_d and m_d:
        if o_d == m_d:
            score += 500.0
            strong_evidence = True
        else:
            score -= 300.0

    if o_p and m_p:
        if o_p == m_p:
            score += 500.0
            strong_evidence = True
        elif len(o_p) >= 3 and len(m_p) >= 3 and (o_p in m_p or m_p in o_p):
            score += 250.0
            strong_evidence = True
        else:
            if len(o_p) >= 2 and len(m_p) >= 2:
                score -= 150.0

    o_long_codes = [c for c in (o_d, o_p, o_s) if len(c) >= 4]
    m_long_codes = [c for c in (m_d, m_p, m_s) if len(c) >= 4]

    for oc in o_long_codes:
        for mc in m_long_codes:
            if oc == mc:
                score += 650.0
                strong_evidence = True
            elif len(oc) >= 5 and len(mc) >= 5 and (oc in mc or mc in oc):
                score += 350.0
                strong_evidence = True

    if o_s and m_p and len(m_p) >= 3 and (o_s == m_p or m_p in o_s or o_s in m_p):
        score += 450.0
        strong_evidence = True
    if o_p and m_s and len(o_p) >= 3 and (o_p == m_s or o_p in m_s or m_s in o_p):
        score += 450.0
        strong_evidence = True

    name_sim = 0.0
    if o_i and m_i:
        if o_i == m_i:
            name_sim = 1.0
            score += 400.0
            strong_evidence = True
        elif o_i in m_i or m_i in o_i:
            name_sim = 0.92
            score += 280.0
            strong_evidence = True
        else:
            name_sim = fast_similarity(o_i, m_i, o_item["bg_item"], m_row["bg_item"])
            if name_sim >= 0.7:
                score += 240.0 * name_sim
                strong_evidence = True
            elif name_sim >= 0.4:
                score += 100.0 * name_sim

    if not strong_evidence and o_i and m_line:
        if o_i == m_line or o_i in m_line or m_line in o_i:
            score += 250.0
            strong_evidence = True

    if not strong_evidence and not (o_d or o_p):
        comb_sim = fast_similarity(o_item["norm_combined"], m_row["norm_combined"], o_item["bg_combined"], m_row["bg_combined"])
        if comb_sim >= 0.5:
            score += 300.0 * comb_sim
            strong_evidence = True

    o_q = o_item.get("qty_num")
    m_q = m_row.get("qty_num")
    if o_q and m_q:
        if o_q == m_q:
            score += 120.0
        else:
            score -= 100.0

    if (o_d and m_d and o_d == m_d) or (o_p and m_p and o_p == m_p):
        strong_evidence = True

    if not strong_evidence:
        return -1.0

    return score


def build_field_comparison_data(
    o_item: dict | None, m_row: dict | None, reason_summary: str, extra_notes: str = ""
) -> dict:
    """照合詳細・プレビュー用の項目比較表データを構築"""
    rows = []

    def _cmp_field(label: str, o_raw: str, m_raw: str, o_norm: str, m_norm: str):
        o_disp = o_raw if o_raw else "(空欄)"
        m_disp = m_raw if m_raw else "(空欄)"
        if not m_row:
            return (label, o_disp, "— (該当行なし)", "管理表未記載", "DIFF")
        if not o_item:
            return (label, "— (注文書なし)", m_disp, "注文書未検出", "DIFF")
        if not o_norm and not m_norm:
            return (label, o_disp, m_disp, "両方空欄 (-)", "INFO")
        if o_norm == m_norm:
            return (label, o_disp, m_disp, "◎ 完全一致", "MATCH")
        if o_norm and m_norm and (o_norm in m_norm or m_norm in o_norm):
            return (label, o_disp, m_disp, "△ 部分一致", "PARTIAL")
        return (label, o_disp, m_disp, "× 不一致", "DIFF")

    o_no = (o_item.get("order_no_8") or o_item.get("order_no", "")) if o_item else ""
    m_no = (m_row.get("order_no_8") or m_row.get("order_no", "")) if m_row else ""
    rows.append(_cmp_field("① 管理No (8桁)", o_no, m_no, o_no, m_no))

    rows.append(
        _cmp_field(
            "② 図面番号",
            o_item.get("drawing_no", "") if o_item else "",
            m_row.get("drawing_no", "") if m_row else "",
            o_item.get("norm_drawing", "") if o_item else "",
            m_row.get("norm_drawing", "") if m_row else "",
        )
    )
    rows.append(
        _cmp_field(
            "③ 品番",
            o_item.get("part_no", "") if o_item else "",
            m_row.get("part_no", "") if m_row else "",
            o_item.get("norm_part", "") if o_item else "",
            m_row.get("norm_part", "") if m_row else "",
        )
    )
    rows.append(
        _cmp_field(
            "④ 品名",
            o_item.get("item_name", "") if o_item else "",
            m_row.get("item_name", "") if m_row else "",
            o_item.get("norm_item", "") if o_item else "",
            m_row.get("norm_item", "") if m_row else "",
        )
    )
    rows.append(
        _cmp_field(
            "⑤ 仕様・規格",
            o_item.get("spec", "") if o_item else "",
            m_row.get("spec", "") if m_row else "",
            o_item.get("norm_spec", "") if o_item else "",
            m_row.get("norm_spec", "") if m_row else "",
        )
    )
    rows.append(
        _cmp_field(
            "⑥ メーカー",
            o_item.get("maker", "") if o_item else "",
            m_row.get("maker", "") if m_row else "",
            o_item.get("norm_maker", "") if o_item else "",
            m_row.get("norm_maker", "") if m_row else "",
        )
    )
    rows.append(
        _cmp_field(
            "⑦ 購入数 (数量)",
            o_item.get("qty", "") if o_item else "",
            m_row.get("qty", "") if m_row else "",
            o_item.get("qty_num", "") if o_item else "",
            m_row.get("qty_num", "") if m_row else "",
        )
    )

    o_p_disp = (o_item.get("unit_price", "") if o_item else "") or "(空欄)"
    m_p_disp = (m_row.get("unit_price", "") if m_row else "") or "(空欄)"
    rows.append(("⑧ 単価", o_p_disp, m_p_disp, "— (照合対象外)", "IGNORED"))

    o_a_disp = (o_item.get("amount", "") if o_item else "") or "(空欄)"
    m_a_disp = (m_row.get("amount", "") if m_row else "") or "(空欄)"
    rows.append(("⑨ 金額", o_a_disp, m_a_disp, "— (照合対象外)", "IGNORED"))

    m_inst_disp = (m_row.get("instruction_disp", "") if m_row else "") or "— (管理表なし)"
    rows.append(("⑩ 指示書 (管理表J列)", "— (注文書側なし)", m_inst_disp, f"状態: {m_inst_disp}", "INFO"))

    m_ship_disp = (m_row.get("ship_to", "") if m_row else "") or "(空欄)"
    rows.append(("⑪ 出荷先 (管理表N列)", "— (注文書側なし)", m_ship_disp, "管理表情報", "INFO"))

    m_line_disp = (m_row.get("line_name", "") if m_row else "") or "(空欄)"
    rows.append(("⑫ 品名、ライン名等 (管理表O列)", "— (注文書側なし)", m_line_disp, "管理表情報", "INFO"))

    m_prod_disp = (m_row.get("product_detail", "") if m_row else "") or "(空欄)"
    rows.append(("⑬ 製品詳細 (管理表P列)", "— (注文書側なし)", m_prod_disp, "管理表情報", "INFO"))

    m_pj_disp = (m_row.get("project_name", "") if m_row else "") or "(空欄)"
    rows.append(("⑭ プロジェクト名 (管理表K列)", "— (注文書側なし)", m_pj_disp, "管理表情報", "INFO"))

    m_prod_qty_disp = (m_row.get("prod_qty", "") if m_row else "") or "(空欄)"
    rows.append(("⑮ 製作数 (管理表Q列)", "— (注文書側なし)", m_prod_qty_disp, "管理表情報", "INFO"))

    m_unit_qty_disp = (m_row.get("unit_qty", "") if m_row else "") or "(空欄)"
    rows.append(("⑯ 単数 (管理表V列)", "— (注文書側なし)", m_unit_qty_disp, "管理表情報", "INFO"))

    o_sup_disp = (o_item.get("supplier", "") if o_item else "") or "(空欄)"
    m_sup_disp = (m_row.get("supplier", "") if m_row else "") or "(空欄)"
    rows.append(_cmp_field("⑰ 仕入先 (管理表X列)", o_sup_disp, m_sup_disp, o_sup_disp, m_sup_disp))

    m_stock_disp = (m_row.get("stock_qty", "") if m_row else "") or "(空欄)"
    rows.append(("⑱ 在庫数 (管理表AA列)", "— (注文書側なし)", m_stock_disp, "管理表情報", "INFO"))

    cross_notes = []
    if o_item and m_row:
        if o_item["norm_spec"] and m_row["norm_part"] and o_item["norm_spec"] == m_row["norm_part"]:
            cross_notes.append("・注文書の『仕様』と管理表の『品番』が一致しています。")
        if o_item["norm_part"] and m_row["norm_spec"] and o_item["norm_part"] == m_row["norm_spec"]:
            cross_notes.append("・注文書の『品番』と管理表の『仕様』が一致しています。")
        if m_row.get("ship_to"):
            cross_notes.append(f"・管理表の『出荷先』: {m_row['ship_to']}")
        if m_row.get("line_name"):
            cross_notes.append(f"・管理表の『品名・ライン名等』: {m_row['line_name']}")
        if m_row.get("product_detail"):
            cross_notes.append(f"・管理表の『製品詳細』: {m_row['product_detail']}")

    detail_lines = [f"【診断結果サマリー】\n  {reason_summary}", ""]
    if o_item:
        o_dt_str = o_item["order_date"].strftime("%Y-%m-%d") if o_item.get("order_date") else "未記載"
        detail_lines.append(
            f"【注文書データ元】\n  ファイル: {o_item['file']} (シート: {o_item['sheet']} / 行: {o_item['row']}) / 依頼区分: {o_item['req_type_label']} / 注文日: {o_dt_str}"
        )
    else:
        detail_lines.append("【注文書データ元】\n  （注文書フォルダ内に該当する注文書が見当たりません）")

    if m_row:
        m_dt_str = m_row["order_date"].strftime("%Y-%m-%d") if m_row.get("order_date") else "未記載"
        m_inst_str = m_row.get("instruction_disp") or "(空白)"
        m_ship_str = m_row.get("ship_to") or "-"
        m_line_str = m_row.get("line_name") or "-"
        m_prod_str = m_row.get("product_detail") or "-"
        detail_lines.append(
            f"【管理表の登録データ】\n  シート: {m_row['sheet']} (行: {m_row['row']}) / 管理No: {m_row.get('order_no_8') or m_row.get('order_no')} / 指示書(J列): {m_inst_str} / 注文日: {m_dt_str}\n  出荷先: {m_ship_str} / 品名・ライン名等: {m_line_str} / 製品詳細: {m_prod_str} / AB列購入数: {m_row.get('qty') or '未記載'} / 納期: {m_row.get('mgmt_due_raw') or '未記載'}"
        )
    else:
        detail_lines.append("【管理表の登録データ】\n  （管理表内に該当する明細が登録されていません）")

    if cross_notes:
        detail_lines.append("\n【追加情報・列構造】\n" + "\n".join(cross_notes))
    if extra_notes:
        detail_lines.append("\n【詳細調査ログ】\n" + extra_notes)

    return {
        "comparison_rows": rows,
        "detail_text": "\n".join(detail_lines),
        "order_full_path": o_item.get("full_path", "") if o_item else "",
        "order_file": o_item.get("file", "") if o_item else "",
        "order_sheet": o_item.get("sheet", "") if o_item else "",
        "order_row": o_item.get("row") if o_item else None,
        "mgmt_sheet": m_row.get("sheet", "") if m_row else "",
        "mgmt_row": m_row.get("row") if m_row else None,
        "order_no": (o_item.get("order_no_8") if o_item else None) or (m_row.get("order_no_8") if m_row else ""),
    }


def diagnose_unmatched_item(
    o_item: dict,
    mgmt_rows: list[dict],
) -> tuple[str, dict]:
    """管理表と一致しなかった注文書明細の理由を詳細診断"""
    o_no = o_item.get("order_no_8", "")

    same_no_rows = [
        m
        for m in mgmt_rows
        if m.get("order_no_8") == o_no or m.get("norm_order_no") == o_no
    ]

    best_cand = None
    best_raw_sim = -999.0
    for m in same_no_rows:
        raw_sc = score_candidate_pair(o_item, m)
        if raw_sc > best_raw_sim:
            best_raw_sim = raw_sc
            best_cand = m

    if best_cand:
        matched_parts = []
        diff_parts = []

        def _check_diff(label: str, o_v: str, m_v: str, o_n: str, m_n: str):
            if not o_n and not m_n:
                return
            if o_n == m_n or (o_n and m_n and (o_n in m_n or m_n in o_n)):
                matched_parts.append(label)
            else:
                diff_parts.append(f"{label}[注文:{o_v or '空'}≠管理:{m_v or '空'}]")

        _check_diff("図番", o_item["drawing_no"], best_cand["drawing_no"], o_item["norm_drawing"], best_cand["norm_drawing"])
        _check_diff("品番", o_item["part_no"], best_cand["part_no"], o_item["norm_part"], best_cand["norm_part"])
        _check_diff("品名", o_item["item_name"], best_cand["item_name"], o_item["norm_item"], best_cand["norm_item"])
        _check_diff("仕様", o_item["spec"], best_cand["spec"], o_item["norm_spec"], best_cand["norm_spec"])
        _check_diff("購入数", o_item["qty"], best_cand["qty"], o_item["qty_num"], best_cand["qty_num"])

        m_str = f"一致:{','.join(matched_parts)} / " if matched_parts else ""
        d_str = "不一致:" + ", ".join(diff_parts) if diff_parts else "全項目不一致"
        reason = (
            f"【項目不一致】管理No[{o_no}] 最有力:{best_cand['sheet']}(行{best_cand['row']}) "
            f"⇒ {m_str}{d_str}"
        )

        cand_list_lines = [
            f"・管理表内の管理No [{o_no}] の該当行数: 計 {len(same_no_rows)} 行",
            "・管理Noが同じ管理表データの抜粋:",
        ]
        for idx_c, r_c in enumerate(same_no_rows[:15], start=1):
            used_mark = " (照合済)" if r_c.get("matched") else " (未照合)"
            cand_list_lines.append(
                f"   {idx_c}. {r_c['sheet']}(行{r_c['row']}){used_mark} | 業者:{r_c.get('maker')} | 図番:{r_c['drawing_no']} | 品番:{r_c['part_no']} | 品名:{r_c['item_name']} | 仕様:{r_c['spec']} | 購入数:{r_c['qty']}"
            )

        diag = build_field_comparison_data(
            o_item, best_cand, reason, "\n".join(cand_list_lines)
        )
        diag["best_cand_instruction_raw"] = best_cand.get("instruction_raw", "")
        diag["best_cand_instruction_disp"] = best_cand.get("instruction_disp", "(空白)")
        diag["best_cand_instruction_state"] = best_cand.get("instruction_state", "BLANK")
        diag["best_cand_ship_to"] = best_cand.get("ship_to", "")
        diag["best_cand_line_name"] = best_cand.get("line_name", "")
        diag["best_cand_product_detail"] = best_cand.get("product_detail", "")
        return reason, diag

    supp_cand = same_no_rows[0] if same_no_rows else None
    reason = f"【明細未記載】管理No[{o_no}]は管理表にありますが一致する明細行がありません"
    diag = build_field_comparison_data(o_item, None, reason)
    diag["best_cand_instruction_raw"] = supp_cand.get("instruction_raw", "") if supp_cand else ""
    diag["best_cand_instruction_disp"] = supp_cand.get("instruction_disp", "(未記載)") if supp_cand else "(未記載)"
    diag["best_cand_instruction_state"] = supp_cand.get("instruction_state", "BLANK") if supp_cand else "BLANK"
    diag["best_cand_ship_to"] = supp_cand.get("ship_to", "") if supp_cand else ""
    diag["best_cand_line_name"] = supp_cand.get("line_name", "") if supp_cand else ""
    diag["best_cand_product_detail"] = supp_cand.get("product_detail", "") if supp_cand else ""
    return reason, diag


def evaluate_single_record(
    req_type: str,
    order_no: str,
    instruction_raw: str,
    instruction_disp: str,
    instruction_state: str,
    ship_to: str,
    line_name: str,
    product_detail: str,
    supplier: str,
    maker: str,
    item_name: str,
    model_spec: str,
    qty: str,
    due_date: datetime.date | None,
    actual_date: datetime.date | None,
    actual_raw: str,
    actual_filled: bool,
    invoice_date: datetime.date | None,
    invoice_raw: str,
    invoice_filled: bool,
    person: str,
    source_info: str,
    base_date: datetime.date,
    is_fallback_order_date: bool = False,
    order_date: datetime.date | None = None,
    grace_days: int = 7,
    is_unmatched: bool = False,
    mismatch_reason: str = "",
    diagnostic_data: dict | None = None,
    project_name: str = "",
    prod_qty: str = "",
    drawing_no: str = "",
    part_no: str = "",
    unit_qty: str = "",
    spec: str = "",
    stock_qty: str = "",
) -> dict:
    """1つの照合明細について超過状況を判定し、表示用レコード辞書を構築"""
    actual_overdue = False
    invoice_overdue = False
    overdue_days_num = None
    days_overdue_str = "-"

    if due_date:
        delta_from_today = (base_date - due_date).days
        if delta_from_today > 0:
            overdue_days_num = delta_from_today
            if is_fallback_order_date and order_date:
                days_since_order = (base_date - order_date).days
                days_overdue_str = f"{delta_from_today}日超過 (注文後{days_since_order}日)"
            elif is_fallback_order_date:
                days_overdue_str = f"{delta_from_today}日超過"
            else:
                days_overdue_str = f"{delta_from_today}日超過"
        elif delta_from_today == 0:
            overdue_days_num = 0
            days_overdue_str = "本日満了" if is_fallback_order_date else "本日納期"
        else:
            overdue_days_num = delta_from_today
            days_overdue_str = f"残り{-delta_from_today}日"

    if is_unmatched:
        status_badge = "【不一致】管理表未記載"
        row_tag = "UNMATCHED"
        actual_display_str = "不一致"
        actual_status = "不一致 (管理表未記載)"
        invoice_display_str = "不一致"
        invoice_status = "不一致 (管理表未記載)"
        actual_overdue = True
        invoice_overdue = True
    else:
        if not actual_filled:
            if due_date and base_date > due_date:
                actual_overdue = True
                actual_status = f"未確認 (超過{(base_date - due_date).days}日)"
            elif due_date is None:
                actual_overdue = True
                actual_status = "未確認 (納期不明)"
            else:
                actual_status = "未確認 (納期前)"
            actual_display_str = "未入力"
        else:
            actual_overdue = False
            if actual_date:
                actual_display_str = actual_date.strftime("%Y-%m-%d")
                if due_date and actual_date > due_date and not is_fallback_order_date:
                    delay_d = (actual_date - due_date).days
                    actual_status = f"確認済 ({delay_d}日遅れ入荷)"
                else:
                    actual_status = "確認済 (正常)"
            else:
                actual_display_str = actual_raw
                actual_status = "確認済 (正常)"

        if not invoice_filled:
            if due_date and base_date > due_date:
                invoice_overdue = True
                invoice_status = f"未確認 (超過{(base_date - due_date).days}日)"
            elif actual_date is not None and base_date > actual_date:
                invoice_overdue = True
                invoice_status = f"伝票未着 (現品確認後{(base_date - actual_date).days}日経過)"
            elif due_date is None:
                invoice_overdue = True
                invoice_status = "未確認 (納期不明)"
            else:
                invoice_status = "未確認 (納期前)"
            invoice_display_str = "未入力"
        else:
            invoice_overdue = False
            if invoice_date:
                invoice_display_str = invoice_date.strftime("%Y-%m-%d")
                if due_date and invoice_date > due_date and not is_fallback_order_date:
                    delay_inv = (invoice_date - due_date).days
                    invoice_status = f"確認済 (納期+{delay_inv}日)"
                else:
                    invoice_status = "確認済 (正常)"
            else:
                invoice_display_str = invoice_raw
                invoice_status = "確認済 (正常)"

        if actual_overdue and invoice_overdue:
            status_badge = "【警告】現品・請求 超過"
            row_tag = "BOTH_DELAY"
        elif actual_overdue:
            status_badge = "【現品超過】現品未確認"
            row_tag = "ACTUAL_DELAY"
        elif invoice_overdue:
            status_badge = "【請求超過】伝票未確認"
            row_tag = "INVOICE_DELAY"
        else:
            status_badge = "正常 / 確認済"
            row_tag = "OK"

    due_str = due_date.strftime("%Y-%m-%d") if due_date else "未記載"
    if is_fallback_order_date and due_date:
        if order_date:
            due_str = f"未記載 ({order_date.strftime('%Y-%m-%d')}注文+{grace_days}日)"
        else:
            due_str = f"未記載 (注文+{grace_days}日)"

    return {
        "status_badge": status_badge,
        "mismatch_reason": mismatch_reason,
        "project_name": project_name,
        "req_type": req_type,
        "order_no": order_no,
        "instruction_raw": instruction_raw,
        "instruction_disp": instruction_disp,
        "instruction_state": instruction_state,
        "ship_to": ship_to,
        "line_name": line_name,
        "product_detail": product_detail,
        "prod_qty": prod_qty,
        "drawing_no": drawing_no,
        "part_no": part_no,
        "order_date": order_date,
        "order_date_str": order_date.strftime("%Y-%m-%d") if order_date else "-",
        "supplier": supplier,
        "maker": maker,
        "item_name": item_name,
        "unit_qty": unit_qty,
        "spec": spec,
        "stock_qty": stock_qty,
        "model_spec": model_spec,
        "qty": qty,
        "due_date": due_str,
        "due_date_str": due_str,
        "days_overdue": days_overdue_str,
        "days_overdue_str": days_overdue_str,
        "overdue_days_num": overdue_days_num,
        "actual_date": actual_display_str,
        "actual_date_str": actual_display_str,
        "actual_status": actual_status,
        "invoice_date": invoice_display_str,
        "invoice_date_str": invoice_display_str,
        "invoice_status": invoice_status,
        "person": person,
        "source_info": source_info,
        "actual_overdue": actual_overdue,
        "invoice_overdue": invoice_overdue,
        "is_unmatched": is_unmatched,
        "is_any_delay": actual_overdue or invoice_overdue or is_unmatched,
        "row_tag": row_tag,
        "diagnostic_data": diagnostic_data,
    }


def match_and_evaluate(
    mgmt_rows: list[dict],
    order_items: list[dict],
    base_date: datetime.date,
    grace_days: int = 7,
    progress_cb=None,
    cancel_event=None,
) -> tuple[list[dict], list[dict]]:
    """
    注文書明細と管理表行を照合し、
    1. main_results: 注文書由来の照合・遅延一覧（タブ1用）
    2. missing_order_results: 注文書未検出・未手配の管理表明細一覧（タブ2用 / ※AB列購入数0は完全除外）
    の2つに完全に分離して返す。
    """
    for m in mgmt_rows:
        m["matched"] = False
        m["matched_by"] = ""

    main_results = []
    missing_order_results = []

    # 管理No（8桁）でグルーピング
    mgmt_by_no = {}
    for m in mgmt_rows:
        no = m.get("order_no_8") or m.get("norm_order_no")
        if no:
            mgmt_by_no.setdefault(no, []).append(m)

    orders_by_no = {}
    for o in order_items:
        no = o.get("order_no_8") or o.get("norm_order_no")
        orders_by_no.setdefault(no, []).append(o)

    total_orders = len(order_items)
    processed_orders = 0

    all_order_nos = set(mgmt_by_no.keys()) | set(orders_by_no.keys())

    for o_no in all_order_nos:
        if cancel_event and cancel_event.is_set():
            raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

        o_list = orders_by_no.get(o_no, [])
        m_list = mgmt_by_no.get(o_no, [])

        if o_list and not m_list:
            for o_item in o_list:
                if cancel_event and cancel_event.is_set():
                    raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
                processed_orders += 1
                if progress_cb and (processed_orders == 1 or processed_orders % 5 == 0 or processed_orders == total_orders):
                    pct = 75.0 + (18.0 * processed_orders / max(1, total_orders))
                    progress_cb(pct, f"[3/3] 注文書明細と管理表を照合中 ({processed_orders}/{total_orders} 件目)...")

                reason_txt, diag_data = diagnose_unmatched_item(o_item, mgmt_rows)
                final_due = o_item.get("due_date")
                order_dt = o_item.get("order_date")
                is_fallback = False
                if not final_due and order_dt:
                    final_due = order_dt + datetime.timedelta(days=grace_days)
                    is_fallback = True

                rec = evaluate_single_record(
                    req_type=o_item.get("req_type_label", "注文"),
                    order_no=o_no,
                    instruction_raw="",
                    instruction_disp="(未記載)",
                    instruction_state="BLANK",
                    ship_to="",
                    line_name="",
                    product_detail="",
                    supplier=o_item.get("supplier", ""),
                    maker=o_item.get("maker", ""),
                    item_name=o_item.get("item_name", ""),
                    model_spec=o_item.get("model_spec", ""),
                    qty=o_item.get("qty", ""),
                    due_date=final_due,
                    actual_date=None,
                    actual_raw="",
                    actual_filled=False,
                    invoice_date=None,
                    invoice_raw="",
                    invoice_filled=False,
                    person="【不一致】",
                    source_info=f"【不一致】注文書:{o_item['file']}(行{o_item['row']})",
                    base_date=base_date,
                    is_fallback_order_date=is_fallback,
                    order_date=order_dt,
                    grace_days=grace_days,
                    is_unmatched=True,
                    mismatch_reason=reason_txt,
                    diagnostic_data=diag_data,
                )
                main_results.append(rec)
            continue

        if o_list and m_list:
            cost_matrix = []
            score_matrix = []
            for o_item in o_list:
                row_costs = []
                row_scores = []
                for m_row in m_list:
                    sc = score_candidate_pair(o_item, m_row)
                    row_scores.append(sc)
                    cost = 2000.0 - sc if sc > 0 else 50000.0
                    row_costs.append(cost)
                cost_matrix.append(row_costs)
                score_matrix.append(row_scores)

            matched_pairs = {}
            if o_list and m_list:
                if HAS_SCIPY:
                    row_ind, col_ind = linear_sum_assignment(cost_matrix)
                    for r_idx, c_idx in zip(row_ind, col_ind):
                        if score_matrix[r_idx][c_idx] >= 140.0:
                            matched_pairs[r_idx] = (c_idx, score_matrix[r_idx][c_idx])
                else:
                    used_cols = set()
                    candidates = []
                    for r_idx in range(len(o_list)):
                        for c_idx in range(len(m_list)):
                            candidates.append((score_matrix[r_idx][c_idx], r_idx, c_idx))
                    candidates.sort(reverse=True, key=lambda x: x[0])
                    for sc, r_idx, c_idx in candidates:
                        if sc < 140.0:
                            break
                        if r_idx not in matched_pairs and c_idx not in used_cols:
                            matched_pairs[r_idx] = (c_idx, sc)
                            used_cols.add(c_idx)

            for r_idx, o_item in enumerate(o_list):
                if cancel_event and cancel_event.is_set():
                    raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
                processed_orders += 1
                if progress_cb and (processed_orders == 1 or processed_orders % 5 == 0 or processed_orders == total_orders):
                    pct = 75.0 + (18.0 * processed_orders / max(1, total_orders))
                    progress_cb(pct, f"[3/3] 注文書明細と管理表を照合・診断中 ({processed_orders}/{total_orders} 件目)...")

                order_due = o_item.get("due_date")
                order_dt = o_item.get("order_date")
                matched_mgmt = m_list[matched_pairs[r_idx][0]] if r_idx in matched_pairs else None
                if not order_dt and matched_mgmt:
                    order_dt = matched_mgmt.get("order_date")

                is_fallback = False
                # 注文品管理表の納期(AH列)に納期が記入してあり、かつ型が日付型である場合は、
                # 注文書に書かれている納期ではなく注文品管理表の納期を参照する。
                # 納期の型が日付型ではない場合(文字列・メモ・空欄等)は無視して、注文書の納期を参照する。
                if matched_mgmt and matched_mgmt.get("mgmt_due_is_date_type") and matched_mgmt.get("mgmt_due_date"):
                    final_due = matched_mgmt.get("mgmt_due_date")
                elif order_due:
                    final_due = order_due
                elif order_dt:
                    final_due = order_dt + datetime.timedelta(days=grace_days)
                    is_fallback = True
                else:
                    final_due = None

                if r_idx in matched_pairs:
                    c_idx, sc = matched_pairs[r_idx]
                    best_mgmt = m_list[c_idx]
                    best_mgmt["matched"] = True
                    best_mgmt["matched_by"] = f"{o_item['file']}(行{o_item['row']})"

                    final_supplier = o_item.get("supplier") or best_mgmt.get("supplier", "")
                    final_maker = o_item.get("maker") or best_mgmt.get("maker", "")
                    match_summary = f"◎ 一致 ({best_mgmt['sheet']} 行{best_mgmt['row']} / スコア:{int(sc)})"
                    diag_data = build_field_comparison_data(o_item, best_mgmt, match_summary)

                    rec = evaluate_single_record(
                        project_name=best_mgmt.get("project_name", ""),
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=best_mgmt.get("order_no_8") or best_mgmt.get("order_no") or o_no,
                        instruction_raw=best_mgmt.get("instruction_raw", ""),
                        instruction_disp=best_mgmt.get("instruction_disp", "(空白)"),
                        instruction_state=best_mgmt.get("instruction_state", "BLANK"),
                        ship_to=best_mgmt.get("ship_to", ""),
                        line_name=best_mgmt.get("line_name", ""),
                        product_detail=best_mgmt.get("product_detail", ""),
                        prod_qty=best_mgmt.get("prod_qty", ""),
                        drawing_no=best_mgmt.get("drawing_no") or o_item.get("drawing_no", ""),
                        part_no=best_mgmt.get("part_no") or o_item.get("part_no", ""),
                        item_name=best_mgmt.get("item_name") or o_item.get("item_name", ""),
                        maker=final_maker,
                        unit_qty=best_mgmt.get("unit_qty", ""),
                        spec=best_mgmt.get("spec") or o_item.get("spec", ""),
                        supplier=final_supplier,
                        stock_qty=best_mgmt.get("stock_qty", ""),
                        model_spec=o_item.get("model_spec") or best_mgmt.get("model_spec", ""),
                        qty=o_item.get("qty") or best_mgmt.get("qty", ""),
                        due_date=final_due,
                        actual_date=best_mgmt.get("actual_date"),
                        actual_raw=best_mgmt.get("actual_raw", ""),
                        actual_filled=best_mgmt.get("actual_filled", False),
                        invoice_date=best_mgmt.get("invoice_date"),
                        invoice_raw=best_mgmt.get("invoice_raw", ""),
                        invoice_filled=best_mgmt.get("invoice_filled", False),
                        person=best_mgmt.get("person", ""),
                        source_info=f"[{o_item.get('req_type_label', '')}] 注文書:{o_item['file']} ⇔ 管理表:{best_mgmt['sheet']}(行{best_mgmt['row']})",
                        base_date=base_date,
                        is_fallback_order_date=is_fallback,
                        order_date=order_dt,
                        grace_days=grace_days,
                        is_unmatched=False,
                        mismatch_reason=match_summary,
                        diagnostic_data=diag_data,
                    )
                else:
                    reason_txt, diag_data = diagnose_unmatched_item(o_item, mgmt_rows)
                    rec = evaluate_single_record(
                        project_name="",
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=o_no,
                        instruction_raw=diag_data.get("best_cand_instruction_raw", ""),
                        instruction_disp=diag_data.get("best_cand_instruction_disp", "(未記載)"),
                        instruction_state=diag_data.get("best_cand_instruction_state", "BLANK"),
                        ship_to=diag_data.get("best_cand_ship_to", ""),
                        line_name=diag_data.get("best_cand_line_name", ""),
                        product_detail=diag_data.get("best_cand_product_detail", ""),
                        prod_qty="",
                        drawing_no=o_item.get("drawing_no", ""),
                        part_no=o_item.get("part_no", ""),
                        item_name=o_item.get("item_name", ""),
                        maker=o_item.get("maker", ""),
                        unit_qty="",
                        spec=o_item.get("spec", ""),
                        supplier=o_item.get("supplier", ""),
                        stock_qty="",
                        model_spec=o_item.get("model_spec", ""),
                        qty=o_item.get("qty", ""),
                        due_date=final_due,
                        actual_date=None,
                        actual_raw="",
                        actual_filled=False,
                        invoice_date=None,
                        invoice_raw="",
                        invoice_filled=False,
                        person="【不一致】",
                        source_info=f"【不一致】注文書:{o_item['file']}(行{o_item['row']})",
                        base_date=base_date,
                        is_fallback_order_date=is_fallback,
                        order_date=order_dt,
                        grace_days=grace_days,
                        is_unmatched=True,
                        mismatch_reason=reason_txt,
                        diagnostic_data=diag_data,
                    )
                main_results.append(rec)

    # 2. 注文書未検出（管理表にあるのに注文書が見当たらない）明細の抽出（タブ2用）
    for m in mgmt_rows:
        if cancel_event and cancel_event.is_set():
            raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
        if m.get("matched"):
            continue

        raw_qty = m.get("qty", "")
        qty_num_str = m.get("qty_num", "")
        if is_zero_quantity(raw_qty) or is_zero_quantity(qty_num_str):
            continue

        m_no = m.get("order_no_8") or m.get("order_no", "")
        reason_txt = f"管理表『{m['sheet']}』(行{m['row']})に記載ありますが、対応する注文書ファイルが作成・手配されていません"

        diag_data = build_field_comparison_data(None, m, reason_txt)

        rec_missing = {
            "status_badge": "未発注 (注文書なし)",
            "mismatch_reason": reason_txt,
            "project_name": m.get("project_name", ""),
            "mgmt_sheet": m["sheet"],
            "mgmt_row": str(m["row"]),
            "order_no": m_no,
            "instruction_raw": m.get("instruction_raw", ""),
            "instruction_disp": m.get("instruction_disp", "(空白)"),
            "instruction_state": m.get("instruction_state", "BLANK"),
            "ship_to": m.get("ship_to", ""),
            "line_name": m.get("line_name", ""),
            "product_detail": m.get("product_detail", ""),
            "prod_qty": m.get("prod_qty", ""),
            "drawing_no": m.get("drawing_no", ""),
            "part_no": m.get("part_no", ""),
            "item_name": m.get("item_name", ""),
            "maker": m.get("maker", ""),
            "unit_qty": m.get("unit_qty", ""),
            "spec": m.get("spec", ""),
            "supplier": m.get("supplier") or m.get("maker", ""),
            "stock_qty": m.get("stock_qty", ""),
            "order_date": m.get("order_date"),
            "order_date_str": m["order_date"].strftime("%Y-%m-%d") if m.get("order_date") else "-",
            "model_spec": m.get("model_spec", ""),
            "qty": m.get("qty", ""),
            "due_date": m.get("mgmt_due_raw") or (m["mgmt_due_date"].strftime("%Y-%m-%d") if m.get("mgmt_due_date") else "-"),
            "due_date_str": m.get("mgmt_due_raw") or (m["mgmt_due_date"].strftime("%Y-%m-%d") if m.get("mgmt_due_date") else "-"),
            "actual_date": m["actual_date"].strftime("%Y-%m-%d") if m.get("actual_date") else (m.get("actual_raw") or "未入力"),
            "invoice_date": m["invoice_date"].strftime("%Y-%m-%d") if m.get("invoice_date") else (m.get("invoice_raw") or "未入力"),
            "person": m.get("person", ""),
            "mismatch_reason": reason_txt,
            "source_info": f"管理表:{m['sheet']}(行{m['row']})",
            "row_tag": "MISSING_ORDER",
            "diagnostic_data": diag_data,
        }
        missing_order_results.append(rec_missing)

    main_results.sort(
        key=lambda r: (
            not (r["is_unmatched"] or r["is_any_delay"]),
            not r["is_unmatched"],
            -(r["overdue_days_num"] or -9999),
        )
    )

    return main_results, missing_order_results
