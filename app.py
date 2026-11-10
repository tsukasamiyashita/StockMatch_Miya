import os
import sys
import re
import json
import ctypes
import calendar
import unicodedata
import datetime
import difflib
import threading
import queue
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd  # type: ignore
import openpyxl  # type: ignore

try:
    from scipy.optimize import linear_sum_assignment  # type: ignore
    HAS_SCIPY = True
except ImportError:
    linear_sum_assignment = None
    HAS_SCIPY = False

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    HAS_WIN32COM = True
except ImportError:
    win32com = None
    pythoncom = None
    HAS_WIN32COM = False


class AnalysisCancelledException(Exception):
    """ユーザーによって照合処理が中断された場合のカスタム例外"""
    pass


class StockMatchApp:
    """
    StockMatch_Miya-v1.5.0
    注文品管理表（1つのファイル）と注文書フォルダ内（サブフォルダ含む全階層）の全注文書Excelを双方向照合し、
    ・納期超過（現品確認日・請求日）および管理表未記載の注文書
    ・注文書未検出（管理表にあるのに注文書が見当たらない未発注明細 / ※AB列購入数0は自動除外）
    を2つの独立したタブ一覧で明確に分離して確認・可視化するデスクトップアプリ
    """

    if getattr(sys, "frozen", False):
        APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
        RESOURCE_DIR = getattr(sys, "_MEIPASS", APP_DIR)
    else:
        APP_DIR = os.path.dirname(os.path.abspath(__file__))
        RESOURCE_DIR = APP_DIR

    DEFAULT_CONFIG_FILE_PATH = os.path.join(APP_DIR, "config.json")
    CONFIG_POINTER_FILE_PATH = os.path.join(APP_DIR, "config_location.txt")
    CONFIG_FILE_PATH = DEFAULT_CONFIG_FILE_PATH
    ICON_FILE_PATH = os.path.join(APP_DIR, "icon.ico")

    INSTRUCTION_FILTER_OPTIONS = [
        ("HAS_ONLY", "「有り」のみ"),
        ("HAS_OR_BLANK", "有り または 空白 のみ"),
        ("BLANK_ONLY", "「空白」のみ"),
        ("ALL", "すべて表示 (フィルタなし)"),
    ]
    INSTRUCTION_CODE_TO_LABEL = dict(INSTRUCTION_FILTER_OPTIONS)
    INSTRUCTION_LABEL_TO_CODE = {lbl: code for code, lbl in INSTRUCTION_FILTER_OPTIONS}

    # 照合結果一覧（タブ1）のデフォルト表示列（計15項目）
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

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("StockMatch_Miya-v1.5.0 - 注文品納期・現品/請求照合システム")

        # メインウィンドウおよび全Toplevelウィンドウに icon.ico を適用
        self._apply_window_icon(self.root, is_main=True)

        # 起動時にウィンドウを最大化表示
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")

        self.root.minsize(1080, 720)

        # 設定ファイルの保存先パス管理（ポインタファイルまたはデフォルト）
        self.config_file_path = tk.StringVar(value=self._resolve_initial_config_path())

        # 状態保持用変数
        self.mgmt_file_path = tk.StringVar()
        self.order_dir_path = tk.StringVar()
        self.base_date_str = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.blank_due_grace_days = tk.StringVar(value="7")
        self.exclude_after_months = tk.StringVar(value="12")
        self.excluded_order_nos_str = tk.StringVar(value="")
        self.instruction_filter_mode = tk.StringVar(value="HAS_ONLY")
        self.instruction_filter_label_var = tk.StringVar(value=self.INSTRUCTION_CODE_TO_LABEL["HAS_ONLY"])
        self.filter_mode = tk.StringVar(value="BOTH_DELAY")
        self.search_keyword = tk.StringVar()
        self.missing_search_keyword = tk.StringVar()
        self.visible_columns_main = list(self.DEFAULT_VISIBLE_COLUMNS_MAIN)

        # 進捗表示・非同期スレッド管理用
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_text_var = tk.StringVar(value="待機中")
        self.is_running = False
        self.is_excel_running = False
        self.cancel_event = threading.Event()
        self.excel_cancel_event = threading.Event()
        self.ui_queue = queue.Queue()

        # 解析結果データ保持（2系統に分離）
        self.analyzed_records = []       # タブ1: 注文書由来の照合・遅延レコード
        self.missing_order_records = []   # タブ2: 注文書未検出（管理表のみ）レコード
        self.cached_mgmt_rows = []       # 注文品管理表の全行キャッシュ（プレビュー参照用）

        # 高速化用インメモリキャッシュ（更新日時・ファイルサイズベースで未変更ファイルの再読込を省略）
        self._mgmt_file_cache = None     # (norm_path, mtime, size, mgmt_rows)
        self._order_file_cache = {}      # norm_path -> (mtime, size, extracted_items)
        self._cache_lock = threading.Lock()

        # 各タブの列フィルター保持用辞書
        self.column_filters_main = {}
        self.column_filters_missing = {}

        self._setup_styles()
        self._build_menu()
        self._build_ui()
        self._load_config_and_autoload()

    def _get_grace_days(self) -> int:
        """納期未記載時に注文日を基準とする猶予日数（デフォルト7日）を整数で取得"""
        val = unicodedata.normalize("NFKC", self.blank_due_grace_days.get().strip())
        try:
            days = int(val)
            return max(0, days)
        except ValueError:
            return 7

    def _get_exclude_months(self) -> int:
        """注文日から何ヶ月経過したら除外するか（デフォルト12ヶ月、0は除外しない）を整数で取得"""
        val = unicodedata.normalize("NFKC", self.exclude_after_months.get().strip())
        if not val:
            return 12
        try:
            months = int(val)
            return max(0, months)
        except ValueError:
            return 12

    @staticmethod
    def _add_months(d: datetime.date, months: int) -> datetime.date:
        """指定された日付 d に months ヶ月を加算した日付を返す（月末日補正あり）"""
        total_months = (d.year * 12 + (d.month - 1)) + months
        target_year = total_months // 12
        target_month = (total_months % 12) + 1
        max_day = calendar.monthrange(target_year, target_month)[1]
        return datetime.date(target_year, target_month, min(d.day, max_day))

    def _is_order_date_expired(
        self,
        order_date: datetime.date | None,
        base_date: datetime.date | None = None,
        exclude_months: int | None = None,
    ) -> bool:
        """注文日(order_date)から指定ヶ月数(exclude_months)以上が経過しているかを判定"""
        if order_date is None:
            return False
        if exclude_months is None:
            exclude_months = self._get_exclude_months()
        if exclude_months <= 0:
            return False
        if base_date is None:
            base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
        cutoff_date = self._add_months(order_date, exclude_months)
        return base_date >= cutoff_date

    @staticmethod
    def _classify_instruction_col(raw_val: str) -> str:
        """
        注文品管理表の指示書列(J列)の値を判定する:
          - "BLANK": 空白（未記入）
          - "HAS"  : 有り（「有」「有り」「あり」「○」「済」等）
          - "NONE" : 無し・その他（「無」「無し」「なし」「不要」「-」「×」等）
        """
        s = unicodedata.normalize("NFKC", str(raw_val or "")).strip()
        if not s:
            return "BLANK"
        s_clean = re.sub(r"\s+", "", s).upper()
        if not s_clean:
            return "BLANK"
        if any(k in s_clean for k in ("無", "なし", "ナシ", "不要", "対象外", "否", "×", "✕", "✗")):
            return "NONE"
        if s_clean in ("-", "ー", "―", "‐", "NO", "NONE", "N/A", "0", "FALSE"):
            return "NONE"
        if any(k in s_clean for k in ("有", "あり", "アリ", "○", "◯", "◎", "済", "完", "YES", "OK")):
            return "HAS"
        return "NONE"

    def _get_instruction_filter_mode(self) -> str:
        """現在の指示書列(J列)フィルターモードコードを取得"""
        lbl = self.instruction_filter_label_var.get().strip()
        if lbl in self.INSTRUCTION_LABEL_TO_CODE:
            code = self.INSTRUCTION_LABEL_TO_CODE[lbl]
            self.instruction_filter_mode.set(code)
            return code
        code = self.instruction_filter_mode.get().strip().upper()
        if code in self.INSTRUCTION_CODE_TO_LABEL:
            return code
        return "HAS_ONLY"

    def _set_instruction_filter_mode(self, mode_code: str):
        """指示書列(J列)フィルターモードコードをセットし、UI表示ラベルと同期する"""
        code = str(mode_code or "").strip().upper()
        if code not in self.INSTRUCTION_CODE_TO_LABEL:
            code = "HAS_ONLY"
        self.instruction_filter_mode.set(code)
        self.instruction_filter_label_var.set(self.INSTRUCTION_CODE_TO_LABEL[code])

    @staticmethod
    def _is_instruction_matched(inst_state: str, filter_mode: str) -> bool:
        """レコードの指示書列状態(inst_state)がフィルター条件(filter_mode)に一致するか判定"""
        state = (inst_state or "BLANK").upper()
        if filter_mode == "ALL":
            return True
        if filter_mode == "HAS_OR_BLANK":
            return state in ("HAS", "BLANK")
        if filter_mode == "HAS_ONLY":
            return state == "HAS"
        if filter_mode == "BLANK_ONLY":
            return state == "BLANK"
        return state == "HAS"

    def _parse_lines_to_order_nos(self, text: str) -> list[str]:
        """行区切りテキスト（Excel貼り付け・カンマ区切り含む）から管理Noリストを抽出・重複除去して返す"""
        if not text:
            return []
        norm = unicodedata.normalize("NFKC", str(text))
        lines = norm.splitlines()
        seen = set()
        result = []
        for line in lines:
            line_clean = line.strip()
            if not line_clean:
                continue
            # カンマやタブ、セミコロンで分割
            tokens = re.split(r"[\t,;\s、，:：]+", line_clean)
            for token in tokens:
                t = token.strip().upper()
                if not t:
                    continue
                # 8桁数字があればそれを優先抽出
                m8 = re.search(r"\b\d{8}\b", t)
                val = m8.group(0) if m8 else t
                if val and val not in seen:
                    seen.add(val)
                    result.append(val)
        return result

    def _get_excluded_order_nos_list(self) -> list[str]:
        """現在設定されている除外管理Noの重複なしリストを取得"""
        raw = self.excluded_order_nos_str.get().strip()
        return self._parse_lines_to_order_nos(raw)

    def _get_excluded_order_nos(self) -> set[str]:
        """一覧から除外する管理Noの集合（set）を大文字・正規化済みで取得"""
        items = self._get_excluded_order_nos_list()
        res = set()
        for it in items:
            clean = it.strip().upper()
            if clean:
                res.add(clean)
                m8 = re.search(r"\b\d{8}\b", clean)
                if m8:
                    res.add(m8.group(0))
        return res

    def _is_order_no_excluded(self, order_no_val: str, excluded_set: set[str] | None = None) -> bool:
        """指定された管理Noが除外対象に含まれるかを判定"""
        if excluded_set is None:
            excluded_set = self._get_excluded_order_nos()
        if not excluded_set or not order_no_val:
            return False
        val_norm = unicodedata.normalize("NFKC", str(order_no_val)).strip().upper()
        if not val_norm:
            return False
        if val_norm in excluded_set:
            return True
        m8 = re.search(r"\b\d{8}\b", val_norm)
        if m8 and m8.group(0) in excluded_set:
            return True
        return False

    def _is_record_excluded(
        self,
        rec: dict,
        excluded_nos: set[str] | None = None,
        base_date: datetime.date | None = None,
        exclude_months: int | None = None,
        inst_filter_mode: str | None = None,
    ) -> bool:
        """レコードが除外管理No・注文日経過月数除外・指示書列(J列)フィルターにより除外されるかを判定"""
        if self._is_order_no_excluded(rec.get("order_no", ""), excluded_nos):
            return True
        if self._is_order_date_expired(rec.get("order_date"), base_date=base_date, exclude_months=exclude_months):
            return True
        if inst_filter_mode is None:
            inst_filter_mode = self._get_instruction_filter_mode()
        if not self._is_instruction_matched(rec.get("instruction_state", "BLANK"), inst_filter_mode):
            return True
        return False

    def _get_visible_main_records(self) -> list[dict]:
        """除外設定（除外管理No・注文日経過月数・指示書列フィルター）を反映したタブ1の有効レコード一覧を取得"""
        excluded_nos = self._get_excluded_order_nos()
        base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
        exclude_months = self._get_exclude_months()
        inst_mode = self._get_instruction_filter_mode()
        return [
            r for r in self.analyzed_records
            if not self._is_record_excluded(r, excluded_nos, base_date, exclude_months, inst_mode)
        ]

    def _get_visible_missing_records(self) -> list[dict]:
        """除外設定（除外管理No・注文日経過月数・指示書列フィルター）を反映したタブ2の有効レコード一覧を取得"""
        excluded_nos = self._get_excluded_order_nos()
        base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
        exclude_months = self._get_exclude_months()
        inst_mode = self._get_instruction_filter_mode()
        return [
            r for r in self.missing_order_records
            if not self._is_record_excluded(r, excluded_nos, base_date, exclude_months, inst_mode)
        ]

    def _on_date_or_month_filter_changed(self, _event=None):
        """判定基準日・経過除外月数・指示書列(J列)フィルターが変更された際に一覧と件数サマリーを即時更新する"""
        self._get_instruction_filter_mode()
        if self.analyzed_records or self.missing_order_records:
            self._update_kpis()
            self._refresh_main_table()
            self._refresh_missing_table()

    def _update_excluded_summary_badge(self):
        """メイン画面の除外設定ボタンおよびサマリーバッジの表示を更新"""
        items = self._get_excluded_order_nos_list()
        cnt = len(items)
        if hasattr(self, "btn_open_exclude"):
            self.btn_open_exclude.configure(text=f"🚫 除外する管理Noを設定・編集... (現在 {cnt}件)")
        if hasattr(self, "lbl_excluded_summary"):
            if cnt == 0:
                self.lbl_excluded_summary.configure(
                    text="  除外設定なし (全件を照合)  ",
                    bg="#f1f5f9",
                    fg="#475569",
                )
            else:
                sample_str = ", ".join(items[:3]) + ("..." if cnt > 3 else "")
                self.lbl_excluded_summary.configure(
                    text=f"  ✔ 除外指定中: {cnt}件 ({sample_str})  ",
                    bg="#fee2e2",
                    fg="#991b1b",
                )

    def _clear_excluded_order_nos(self):
        """除外管理Noの指定を全解除する"""
        if not self._get_excluded_order_nos_list():
            messagebox.showinfo("情報", "現在、除外設定されている管理Noはありません。")
            return
        if messagebox.askyesno("全解除の確認", "設定されている除外管理Noをすべて解除しますか？"):
            self.excluded_order_nos_str.set("")
            self._update_excluded_summary_badge()
            self._update_kpis()
            self._refresh_main_table()
            self._refresh_missing_table()
            self._save_config(show_msg=False)
            self.status_var.set("除外管理Noの設定をすべて解除しました。")

    def _open_exclude_dialog(self):
        from modules.dialogs import open_exclude_dialog

        def on_apply(text_val):
            self.excluded_order_nos_str.set(text_val)
            self._update_excluded_summary_badge()
            self._on_date_or_month_filter_changed()
            self._save_config()

        open_exclude_dialog(
            parent=self.root,
            current_excluded_text=self.excluded_order_nos_str.get(),
            icon_path=self._resolve_icon_path(),
            on_apply=on_apply,
        )
    def _resolve_icon_path(self) -> str | None:
        """icon.ico の実在パスを探索して返す"""
        candidates = [
            self.ICON_FILE_PATH,
            os.path.join(self.RESOURCE_DIR, "icon.ico"),
            os.path.abspath("icon.ico"),
        ]
        for path in candidates:
            if path and os.path.isfile(path):
                return path
        return None

    def _apply_window_icon(self, window: tk.Wm, is_main: bool = False):
        """指定ウィンドウに icon.ico を設定する"""
        icon_path = self._resolve_icon_path()
        if not icon_path:
            return
        try:
            if is_main:
                window.iconbitmap(default=icon_path)
            window.iconbitmap(icon_path)
        except Exception:
            pass

    def _setup_styles(self):
        """UIテーマ・スタイルの設定（タブ選択状態の視認性を劇的に向上）"""
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("TFrame", background="#f8fafc")
        style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabel", background="#f8fafc", font=("Meiryo UI", 9))
        style.configure("Card.TLabel", background="#ffffff", font=("Meiryo UI", 9))
        style.configure("Header.TLabel", background="#1e293b", foreground="#ffffff", font=("Meiryo UI", 13, "bold"))
        style.configure("SubHeader.TLabel", background="#1e293b", foreground="#cbd5e1", font=("Meiryo UI", 9))

        style.configure(
            "Primary.TButton",
            font=("Meiryo UI", 10, "bold"),
            background="#1d4ed8",
            foreground="#ffffff",
            padding=(16, 7),
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1e40af"), ("disabled", "#94a3b8")],
            foreground=[("disabled", "#e2e8f0")],
        )

        style.configure(
            "Cancel.TButton",
            font=("Meiryo UI", 10, "bold"),
            background="#dc2626",
            foreground="#ffffff",
            padding=(14, 7),
        )
        style.map(
            "Cancel.TButton",
            background=[("active", "#b91c1c"), ("disabled", "#e2e8f0")],
            foreground=[("disabled", "#94a3b8")],
        )

        style.configure(
            "Excel.TButton",
            font=("Meiryo UI", 9, "bold"),
            background="#166534",
            foreground="#ffffff",
            padding=(12, 6),
        )
        style.map(
            "Excel.TButton",
            background=[("active", "#15803d"), ("disabled", "#94a3b8")],
            foreground=[("disabled", "#e2e8f0")],
        )

        style.configure("Action.TButton", font=("Meiryo UI", 9), padding=(10, 5))
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background="#2563eb",
            thickness=14,
        )

        style.configure(
            "Treeview",
            font=("Meiryo UI", 9),
            rowheight=28,
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#0f172a",
        )
        style.configure(
            "Treeview.Heading",
            font=("Meiryo UI", 9, "bold"),
            background="#e2e8f0",
            foreground="#1e293b",
            padding=(4, 6),
        )

        # タブウィジェット (Notebook) スタイル
        style.configure(
            "TNotebook",
            background="#f8fafc",
            borderwidth=0,
            tabmargins=[2, 4, 2, 0]
        )
        style.configure(
            "TNotebook.Tab",
            font=("Meiryo UI", 10),
            padding=(20, 8),
            background="#e2e8f0",
            foreground="#64748b",
            borderwidth=1,
            lightcolor="#cbd5e1",
            darkcolor="#94a3b8",
        )
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", "#1e293b"),
                ("active", "#cbd5e1"),
            ],
            foreground=[
                ("selected", "#ffffff"),
                ("active", "#0f172a"),
            ],
            font=[
                ("selected", ("Meiryo UI", 10, "bold")),
            ],
            expand=[
                ("selected", [1, 2, 1, 0]),
            ],
        )

    def _build_menu(self):
        """上部メニューバー"""
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="注文品管理表ファイルを選択...", command=self._select_mgmt_file, accelerator="Ctrl+O")
        file_menu.add_command(label="注文書フォルダを選択...", command=self._select_order_dir, accelerator="Ctrl+D")
        file_menu.add_separator()
        file_menu.add_command(label="照合・遅延チェックを実行", command=self.run_analysis, accelerator="F5")
        file_menu.add_command(label="Excelで直接開く (一時ファイルなし)", command=self._open_in_excel_direct, accelerator="Ctrl+S")
        file_menu.add_command(label="実行中の処理を中止", command=self.cancel_active_task, accelerator="Esc")
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit, accelerator="Alt+F4")
        menubar.add_cascade(label="ファイル(F)", menu=file_menu)

        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(
            label="現在の設定を保存 (パス・超過日数・経過除外月数・指示書フィルタ・除外管理No・表示列)",
            command=self._save_config_manual,
        )
        config_menu.add_command(
            label="照合結果一覧の表示列設定...",
            command=self._open_column_visibility_dialog,
        )
        config_menu.add_separator()
        config_menu.add_command(label="設定ファイル(config.json)の保存先を指定・変更...", command=self._open_config_location_dialog)
        config_menu.add_command(label="設定ファイルの保存先をアプリ配置フォルダ(初期値)に戻す", command=lambda: self._reset_config_path_to_default(confirm=True))
        config_menu.add_separator()
        config_menu.add_command(label="設定ファイル保存先フォルダを開く", command=self._open_current_config_folder)
        config_menu.add_command(label="アプリ配置フォルダを開く", command=self._open_app_folder)
        menubar.add_cascade(label="設定(S)", menu=config_menu)

        ref_menu = tk.Menu(menubar, tearoff=0)
        ref_menu.add_command(
            label="選択行の「注文書」をExcelで開く 【読み取り専用・保存ブロック】",
            command=self._open_selected_row_order,
        )
        ref_menu.add_command(
            label="選択行の「注文品管理表」をExcelで開く (該当行へジャンプ)",
            command=self._open_active_tab_selected_mgmt,
        )
        ref_menu.add_command(
            label="選択行の照合詳細・シートプレビューを表示",
            command=self._open_active_tab_selected_detail,
        )
        ref_menu.add_separator()
        ref_menu.add_command(label="注文品管理表ファイルをExcelで開く", command=self._open_mgmt_file_direct)
        ref_menu.add_command(label="注文書フォルダをエクスプローラーで開く", command=self._open_order_dir_direct)
        menubar.add_cascade(label="参照(R)", menu=ref_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使い方・操作ガイド(H)...", command=lambda: self._show_help_guide(tab_index=0), accelerator="F1")
        help_menu.add_command(label="照合・判定ロジックの詳細...", command=lambda: self._show_help_guide(tab_index=1))
        help_menu.add_command(label="ショートカット・便利機能一覧...", command=lambda: self._show_help_guide(tab_index=2))
        help_menu.add_separator()
        help_menu.add_command(label="バージョン情報...", command=self._show_about)
        menubar.add_cascade(label="ヘルプ(H)", menu=help_menu)

        self.root.config(menu=menubar)

        self.root.bind_all("<F1>", lambda _e: self._show_help_guide(tab_index=0))
        self.root.bind_all("<F5>", lambda _e: self.run_analysis())
        self.root.bind_all("<Escape>", lambda _e: self.cancel_active_task())
        self.root.bind_all("<Control-o>", lambda _e: self._select_mgmt_file())
        self.root.bind_all("<Control-d>", lambda _e: self._select_order_dir())
        self.root.bind_all("<Control-s>", lambda _e: self._open_in_excel_direct())

    def _build_ui(self):
        """メイン画面レイアウトの構築"""
        self.root.configure(bg="#f8fafc")

        # 1. トップヘッダーバー
        header_frame = tk.Frame(self.root, bg="#1e293b", height=56, padx=16, pady=10)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="StockMatch_Miya-v1.5.0", style="Header.TLabel")
        title_lbl.pack(side=tk.LEFT)

        sub_lbl = ttk.Label(
            header_frame,
            text="  |  注文書 ⇔ 管理表（個別タブ分離：納期超過照合 × 注文書未検出）",
            style="SubHeader.TLabel",
        )
        sub_lbl.pack(side=tk.LEFT, padx=(8, 0))

        help_btn_header = tk.Button(
            header_frame,
            text="？ ヘルプ / 操作ガイド (F1)",
            bg="#334155",
            fg="#f8fafc",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief="flat",
            font=("Meiryo UI", 9, "bold"),
            padx=12,
            pady=4,
            cursor="hand2",
            command=lambda: self._show_help_guide(tab_index=0),
        )
        help_btn_header.pack(side=tk.RIGHT)

        # 2. コントロールパネル
        ctrl_container = ttk.Frame(self.root, padding=(16, 10))
        ctrl_container.pack(fill=tk.X)

        file_card = ttk.Frame(ctrl_container, style="Card.TFrame", padding=12)
        file_card.pack(fill=tk.X)

        # ① 注文品管理表ファイル行
        row1 = ttk.Frame(file_card, style="Card.TFrame")
        row1.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(
            row1,
            text="① 注文品管理表ファイル:",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        mgmt_entry = ttk.Entry(row1, textvariable=self.mgmt_file_path, width=65)
        mgmt_entry.pack(side=tk.LEFT, padx=(4, 8), fill=tk.X, expand=True)
        ttk.Button(row1, text="ファイルを参照...", style="Action.TButton", command=self._select_mgmt_file).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(row1, text="📋 管理表を開く", style="Action.TButton", command=self._open_mgmt_file_direct).pack(
            side=tk.LEFT
        )

        # ② 注文書フォルダ行
        row2 = ttk.Frame(file_card, style="Card.TFrame")
        row2.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(
            row2,
            text="② 注文書フォルダパス:",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        order_entry = ttk.Entry(row2, textvariable=self.order_dir_path, width=65)
        order_entry.pack(side=tk.LEFT, padx=(4, 8), fill=tk.X, expand=True)
        ttk.Button(row2, text="フォルダを参照...", style="Action.TButton", command=self._select_order_dir).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(row2, text="📂 フォルダを開く", style="Action.TButton", command=self._open_order_dir_direct).pack(
            side=tk.LEFT
        )

        # ③ 基準日・超過設定・経過月数除外行（設定入力項目と「設定を保存」ボタンを集約）
        row3 = ttk.Frame(file_card, style="Card.TFrame")
        row3.pack(fill=tk.X, pady=(2, 6))

        ttk.Label(
            row3,
            text="③ 判定基準日 (今日):",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        date_entry = ttk.Entry(row3, textvariable=self.base_date_str, width=12, font=("Meiryo UI", 9))
        date_entry.pack(side=tk.LEFT, padx=(4, 6))
        date_entry.bind("<Return>", self._on_date_or_month_filter_changed)
        date_entry.bind("<FocusOut>", self._on_date_or_month_filter_changed)

        ttk.Button(
            row3,
            text="今日に戻す",
            style="Action.TButton",
            command=lambda: [
                self.base_date_str.set(datetime.date.today().strftime("%Y-%m-%d")),
                self._on_date_or_month_filter_changed(),
            ],
        ).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(
            row3,
            text="納期未記載時:",
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Label(
            row3,
            text="注文日(4行目)から",
            style="Card.TLabel",
            font=("Meiryo UI", 9),
        ).pack(side=tk.LEFT, padx=(4, 2))
        grace_entry = ttk.Entry(
            row3, textvariable=self.blank_due_grace_days, width=4, justify="center", font=("Meiryo UI", 9)
        )
        grace_entry.pack(side=tk.LEFT, padx=(2, 2))
        ttk.Label(
            row3,
            text="日経過で超過",
            style="Card.TLabel",
            font=("Meiryo UI", 9),
        ).pack(side=tk.LEFT, padx=(2, 14))

        ttk.Label(
            row3,
            text="経過除外:",
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Label(
            row3,
            text="注文日から",
            style="Card.TLabel",
            font=("Meiryo UI", 9),
        ).pack(side=tk.LEFT, padx=(4, 2))
        exclude_months_entry = ttk.Entry(
            row3, textvariable=self.exclude_after_months, width=4, justify="center", font=("Meiryo UI", 9)
        )
        exclude_months_entry.pack(side=tk.LEFT, padx=(2, 2))
        exclude_months_entry.bind("<Return>", self._on_date_or_month_filter_changed)
        exclude_months_entry.bind("<FocusOut>", self._on_date_or_month_filter_changed)
        ttk.Label(
            row3,
            text="ヶ月経過で除外",
            style="Card.TLabel",
            font=("Meiryo UI", 9),
        ).pack(side=tk.LEFT, padx=(2, 10))

        self.save_cfg_btn = ttk.Button(
            row3,
            text="💾 設定を保存",
            style="Action.TButton",
            command=self._save_config_manual,
        )
        self.save_cfg_btn.pack(side=tk.RIGHT)

        self.cfg_loc_btn = ttk.Button(
            row3,
            text="⚙ 保存先設定...",
            style="Action.TButton",
            command=self._open_config_location_dialog,
        )
        self.cfg_loc_btn.pack(side=tk.RIGHT, padx=(0, 6))

        # ④ 除外管理No設定 ＆ 注文品管理表の指示書列(J列)フィルター設定行
        row_ex = ttk.Frame(file_card, style="Card.TFrame")
        row_ex.pack(fill=tk.X, pady=(2, 6))

        ttk.Label(
            row_ex,
            text="④ 一覧から除外する管理No:",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)

        self.btn_open_exclude = ttk.Button(
            row_ex,
            text="🚫 除外する管理Noを設定・編集... (現在 0件)",
            style="Action.TButton",
            command=self._open_exclude_dialog,
        )
        self.btn_open_exclude.pack(side=tk.LEFT, padx=(4, 8))

        self.lbl_excluded_summary = tk.Label(
            row_ex,
            text="  除外設定なし (全件を照合)  ",
            bg="#f1f5f9",
            fg="#475569",
            font=("Meiryo UI", 9),
            padx=8,
            pady=2,
            relief="solid",
            borderwidth=1,
        )
        self.lbl_excluded_summary.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(
            row_ex,
            text="全解除",
            style="Action.TButton",
            command=self._clear_excluded_order_nos,
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(
            row_ex,
            text="📋 管理表 指示書列(J列)フィルタ:",
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.cmb_instruction_filter = ttk.Combobox(
            row_ex,
            textvariable=self.instruction_filter_label_var,
            values=[lbl for _, lbl in self.INSTRUCTION_FILTER_OPTIONS],
            state="readonly",
            width=22,
            font=("Meiryo UI", 9),
        )
        self.cmb_instruction_filter.pack(side=tk.LEFT, padx=(0, 8))
        self.cmb_instruction_filter.bind("<<ComboboxSelected>>", self._on_date_or_month_filter_changed)

        ttk.Label(
            row_ex,
            text="※J列が「有り」または「空白」の案件等にワンクリックで絞り込めます",
            style="Card.TLabel",
            foreground="#64748b",
            font=("Meiryo UI", 8),
        ).pack(side=tk.LEFT)

        # ⑤ 進捗状況 ＆ 照合実行ボタン行（メインの実行ボタンを進捗行に大きく配置）
        row4 = ttk.Frame(file_card, style="Card.TFrame")
        row4.pack(fill=tk.X, pady=(6, 0))

        ttk.Label(
            row4,
            text="進捗状況:",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)

        self.progress_bar = ttk.Progressbar(
            row4,
            variable=self.progress_var,
            maximum=100.0,
            mode="determinate",
            style="Horizontal.TProgressbar",
            length=260,
        )
        self.progress_bar.pack(side=tk.LEFT, padx=(4, 10))

        self.progress_lbl = ttk.Label(
            row4,
            textvariable=self.progress_text_var,
            style="Card.TLabel",
            foreground="#1d4ed8",
            font=("Meiryo UI", 9, "bold"),
        )
        self.progress_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.run_btn = ttk.Button(
            row4,
            text="▶ 照合・遅延チェック実行 (F5)",
            style="Primary.TButton",
            command=self.run_analysis,
        )
        self.run_btn.pack(side=tk.RIGHT)

        self.cancel_btn = ttk.Button(
            row4,
            text="⏹ 中止 (Esc)",
            style="Cancel.TButton",
            command=self.cancel_active_task,
            state=tk.DISABLED,
        )
        self.cancel_btn.pack(side=tk.RIGHT, padx=(0, 8))

        # 3. サマリーKPIカード 兼 ワンクリック表示フィルタエリア
        kpi_section = ttk.Frame(self.root, padding=(16, 0, 16, 8))
        kpi_section.pack(fill=tk.X)

        kpi_header_row = ttk.Frame(kpi_section)
        kpi_header_row.pack(fill=tk.X, pady=(0, 3))
        ttk.Label(
            kpi_header_row,
            text="■ 照合件数サマリー ＆ 表示フィルタ切替 （各カードのボタンを押すと一覧を即座に絞り込みます）",
            font=("Meiryo UI", 8, "bold"),
            foreground="#334155",
        ).pack(side=tk.LEFT)

        kpi_frame = ttk.Frame(kpi_section)
        kpi_frame.pack(fill=tk.X)

        self.kpi_vars = {
            "order_total": tk.StringVar(value="0 件"),
            "all_delay": tk.StringVar(value="0 件"),
            "both_delay": tk.StringVar(value="0 件"),
            "actual_delay": tk.StringVar(value="0 件"),
            "invoice_delay": tk.StringVar(value="0 件"),
            "unmatched": tk.StringVar(value="0 件"),
            "ok": tk.StringVar(value="0 件"),
            "missing_order": tk.StringVar(value="0 件"),
        }

        self.kpi_cards = {}

        kpi_specs = [
            (
                "ALL",
                "注文書明細 総数",
                self.kpi_vars["order_total"],
                "#1e293b",
                "#f1f5f9",
                "#2563eb",
                "#e2e8f0",
                "▼ 全件を表示",
                "✔ 全件を表示中",
            ),
            (
                "ALL_DELAY",
                "すべての異常(超過・不一致)",
                self.kpi_vars["all_delay"],
                "#881337",
                "#ffe4e6",
                "#e11d48",
                "#fecdd3",
                "▼ 異常をすべて表示",
                "✔ フィルター表示中",
            ),
            (
                "BOTH_DELAY",
                "【警告】現品・請求の両方超過",
                self.kpi_vars["both_delay"],
                "#991b1b",
                "#fee2e2",
                "#dc2626",
                "#fecaca",
                "▼ この条件で表示",
                "✔ フィルター表示中",
            ),
            (
                "ACTUAL_DELAY",
                "【現品超過】現品確認日の超過",
                self.kpi_vars["actual_delay"],
                "#9a3412",
                "#ffedd5",
                "#ea580c",
                "#fed7aa",
                "▼ この条件で表示",
                "✔ フィルター表示中",
            ),
            (
                "INVOICE_DELAY",
                "【請求超過】請求日(伝票)超過",
                self.kpi_vars["invoice_delay"],
                "#854d0e",
                "#fef9c3",
                "#ca8a04",
                "#fef08a",
                "▼ この条件で表示",
                "✔ フィルター表示中",
            ),
            (
                "UNMATCHED_ONLY",
                "【不一致】管理表との不一致",
                self.kpi_vars["unmatched"],
                "#4c1d95",
                "#ede9fe",
                "#7c3aed",
                "#ddd6fe",
                "▼ この条件で表示",
                "✔ フィルター表示中",
            ),
            (
                "OK_ONLY",
                "正常・納期前・確認済",
                self.kpi_vars["ok"],
                "#166534",
                "#dcfce7",
                "#16a34a",
                "#bbf7d0",
                "▼ この条件で表示",
                "✔ フィルター表示中",
            ),
            (
                "MISSING_TAB",
                "【別タブ】注文書未検出",
                self.kpi_vars["missing_order"],
                "#92400e",
                "#fef3c7",
                "#d97706",
                "#fde68a",
                "▷ 未発注タブを表示",
                "✔ 未発注タブ表示中",
            ),
        ]

        for idx_k, spec in enumerate(kpi_specs):
            kpi_frame.grid_columnconfigure(idx_k, weight=1, uniform="kpi_col")
            card_w = self._create_kpi_card(kpi_frame, *spec)
            pad_r = 0 if idx_k == len(kpi_specs) - 1 else 5
            card_w.grid(row=0, column=idx_k, sticky="nsew", padx=(0, pad_r))

        # 4. タブ切り替えコンテナ（Notebook）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        # --- タブ1: 納期超過・照合結果 一覧 ---
        tab1 = ttk.Frame(self.notebook)
        self.notebook.add(tab1, text="  【選択中】▶ ① 納期超過・照合結果 一覧 (0件)  ")
        self._build_main_table_tab(tab1)

        # --- タブ2: 注文書未検出（管理表のみ）一覧 ---
        tab2 = ttk.Frame(self.notebook)
        self.notebook.add(tab2, text="  ▷ ② 【未発注・注文書なし】管理表のみ記載 一覧 (0件)  ")
        self._build_missing_order_tab(tab2)

        # タブ切り替えイベント
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # 初期選択状態のスタイル反映
        self._update_kpi_card_styles()

        # 5. ステータスバー
        self.status_var = tk.StringVar(
            value="準備完了: ①注文品管理表ファイル と ②注文書フォルダパス を指定して「照合・遅延チェック実行」を押してください。"
        )
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bg="#e2e8f0",
            fg="#334155",
            anchor="w",
            padx=12,
            pady=4,
            font=("Meiryo UI", 9),
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _on_tab_changed(self, _event=None):
        """タブが切り替わったときにタイトル文言とKPIカードのアクティブ表示を更新する"""
        try:
            curr_tab_idx = self.notebook.index("current")
            total_main = len(self._get_visible_main_records())
            total_missing = len(self._get_visible_missing_records())

            if curr_tab_idx == 0:
                self.notebook.tab(0, text=f"  【選択中】▶ ① 納期超過・照合結果 一覧 ({total_main}件)  ")
                self.notebook.tab(1, text=f"  ▷ ② 【未発注・注文書なし】管理表のみ記載 一覧 ({total_missing}件)  ")
            else:
                self.notebook.tab(0, text=f"  ▷ ① 納期超過・照合結果 一覧 ({total_main}件)  ")
                self.notebook.tab(1, text=f"  【選択中】▶ ② 【未発注・注文書なし】管理表のみ記載 一覧 ({total_missing}件)  ")
            self._update_kpi_card_styles()
        except Exception:
            pass

    def _build_main_table_tab(self, parent):
        """タブ1: 納期超過・照合結果テーブルの構築"""
        filter_frame = ttk.Frame(parent, padding=(4, 6, 4, 6))
        filter_frame.pack(fill=tk.X)

        ttk.Label(filter_frame, text="現在の表示フィルタ:", font=("Meiryo UI", 9, "bold")).pack(
            side=tk.LEFT, padx=(0, 6)
        )

        self.active_filter_badge = tk.Label(
            filter_frame,
            text="  ✔ 【警告】現品・請求の両方超過 (0 件)  ",
            bg="#fee2e2",
            fg="#991b1b",
            font=("Meiryo UI", 9, "bold"),
            padx=8,
            pady=3,
            relief="solid",
            borderwidth=1,
        )
        self.active_filter_badge.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(
            filter_frame,
            text="全件表示に切替",
            style="Action.TButton",
            command=lambda: self._apply_kpi_filter("ALL"),
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(filter_frame, text="キーワード検索:", font=("Meiryo UI", 9, "bold")).pack(
            side=tk.LEFT, padx=(4, 4)
        )
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_keyword, width=18)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind("<KeyRelease>", lambda _e: self._refresh_main_table())

        ttk.Button(
            filter_frame,
            text="クリア",
            style="Action.TButton",
            command=lambda: [self.search_keyword.set(""), self._refresh_main_table()],
        ).pack(side=tk.LEFT, padx=(4, 0))

        # 右側アクションボタン群（右端に「Excelで直接開く」ボタン、その左に「中止」ボタンを配置）
        self.excel_open_btn = ttk.Button(
            filter_frame,
            text="📊 Excelで直接開く",
            style="Excel.TButton",
            command=self._open_in_excel_direct,
        )
        self.excel_open_btn.pack(side=tk.RIGHT, padx=(6, 0))

        self.excel_cancel_btn = ttk.Button(
            filter_frame,
            text="⏹ 中止",
            style="Cancel.TButton",
            command=self.cancel_excel_open,
            state=tk.DISABLED,
        )
        self.excel_cancel_btn.pack(side=tk.RIGHT, padx=(4, 0))

        clear_col_btn = ttk.Button(
            filter_frame,
            text="列フィルター解除",
            style="Action.TButton",
            command=self._clear_all_main_column_filters,
        )
        clear_col_btn.pack(side=tk.RIGHT, padx=(4, 0))

        col_visibility_btn = ttk.Button(
            filter_frame,
            text="👁 表示列の設定...",
            style="Action.TButton",
            command=self._open_column_visibility_dialog,
        )
        col_visibility_btn.pack(side=tk.RIGHT, padx=(4, 0))

        open_mgmt_btn = ttk.Button(
            filter_frame,
            text="📋 管理表を参照",
            style="Action.TButton",
            command=lambda: self._open_selected_row_mgmt(is_missing_tab=False),
        )
        open_mgmt_btn.pack(side=tk.RIGHT, padx=(4, 0))

        open_order_btn = ttk.Button(
            filter_frame,
            text="📄 注文書を参照 (読取専用)",
            style="Action.TButton",
            command=self._open_selected_row_order,
        )
        open_order_btn.pack(side=tk.RIGHT, padx=(4, 0))

        inspect_btn = ttk.Button(
            filter_frame,
            text="🔍 照合詳細・プレビュー",
            style="Action.TButton",
            command=lambda: self._open_selected_row_detail(is_missing_tab=False),
        )
        inspect_btn.pack(side=tk.RIGHT, padx=(4, 0))

        table_container = ttk.Frame(parent)
        table_container.pack(fill=tk.BOTH, expand=True)

        columns = (
            "status_badge",
            "mismatch_reason",
            "req_type",
            "order_no",
            "instruction_disp",
            "ship_to",
            "line_name",
            "product_detail",
            "supplier",
            "maker",
            "item_name",
            "model_spec",
            "qty",
            "due_date",
            "days_overdue",
            "actual_date",
            "actual_status",
            "invoice_date",
            "invoice_status",
            "person",
            "source_info",
        )

        self.tree_main = ttk.Treeview(table_container, columns=columns, show="headings", selectmode="browse")

        self.col_configs_main = [
            ("status_badge", "総合判定", 130, "center"),
            ("mismatch_reason", "不一致理由・照合診断", 210, "w"),
            ("req_type", "依頼区分", 65, "center"),
            ("order_no", "管理No(注文番号)", 100, "center"),
            ("instruction_disp", "指示書(J列)", 65, "center"),
            ("ship_to", "出荷先", 95, "w"),
            ("line_name", "品名・ライン名等", 115, "w"),
            ("product_detail", "製品詳細", 115, "w"),
            ("supplier", "発注先", 105, "w"),
            ("maker", "メーカー", 80, "w"),
            ("item_name", "品名", 110, "w"),
            ("model_spec", "型番・規格(図番)", 135, "w"),
            ("qty", "数量", 45, "center"),
            ("due_date", "注文書納期", 85, "center"),
            ("days_overdue", "基準日超過", 75, "center"),
            ("actual_date", "現品確認日(物)", 85, "center"),
            ("actual_status", "現品状況", 95, "center"),
            ("invoice_date", "請求日(伝票)", 85, "center"),
            ("invoice_status", "請求状況", 95, "center"),
            ("person", "担当", 55, "center"),
            ("source_info", "照合元情報", 135, "w"),
        ]

        self.col_titles_main = {col_id: title for col_id, title, _, _ in self.col_configs_main}

        for col_id, _, width, anchor in self.col_configs_main:
            self._update_heading_display_main(col_id)
            self.tree_main.column(col_id, width=width, anchor=anchor)

        self._apply_column_visibility_main()

        self.tree_main.bind("<Button-1>", self._on_tree_main_click)
        self.tree_main.bind("<Double-1>", lambda e: self._on_tree_double_click(e, is_missing_tab=False))
        self.tree_main.bind("<Button-3>", lambda e: self._on_tree_right_click(e, is_missing_tab=False))

        self.tree_main.tag_configure("UNMATCHED", background="#ede9fe", foreground="#4c1d95")
        self.tree_main.tag_configure("BOTH_DELAY", background="#fee2e2", foreground="#7f1d1d")
        self.tree_main.tag_configure("ACTUAL_DELAY", background="#ffedd5", foreground="#7c2d12")
        self.tree_main.tag_configure("INVOICE_DELAY", background="#fef9c3", foreground="#713f12")
        self.tree_main.tag_configure("OK", background="#ffffff", foreground="#334155")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree_main.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree_main.xview)
        self.tree_main.configure(yscrollcommand=vsb.set)

        self.tree_main.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

    def _build_missing_order_tab(self, parent):
        """タブ2: 注文書未検出（管理表のみ）テーブルの構築"""
        filter_frame = ttk.Frame(parent, padding=(4, 6, 4, 6))
        filter_frame.pack(fill=tk.X)

        ttk.Label(
            filter_frame,
            text="※注文品管理表に記載があるものの、注文書フォルダ内に対応する注文書Excelが存在しない未発注案件です（※AB列購入数0は自動除外）。",
            font=("Meiryo UI", 9, "bold"),
            foreground="#92400e",
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(filter_frame, text="キーワード:").pack(side=tk.LEFT, padx=(10, 4))
        missing_search_entry = ttk.Entry(filter_frame, textvariable=self.missing_search_keyword, width=20)
        missing_search_entry.pack(side=tk.LEFT)
        missing_search_entry.bind("<KeyRelease>", lambda _e: self._refresh_missing_table())

        clear_missing_col_btn = ttk.Button(
            filter_frame,
            text="列フィルター解除 (Excel風)",
            style="Action.TButton",
            command=self._clear_all_missing_column_filters,
        )
        clear_missing_col_btn.pack(side=tk.RIGHT, padx=(6, 0))

        open_missing_mgmt_btn = ttk.Button(
            filter_frame,
            text="📋 管理表を参照",
            style="Action.TButton",
            command=lambda: self._open_selected_row_mgmt(is_missing_tab=True),
        )
        open_missing_mgmt_btn.pack(side=tk.RIGHT, padx=(6, 0))

        inspect_missing_btn = ttk.Button(
            filter_frame,
            text="🔍 詳細・プレビュー",
            style="Action.TButton",
            command=lambda: self._open_selected_row_detail(is_missing_tab=True),
        )
        inspect_missing_btn.pack(side=tk.RIGHT, padx=(6, 0))

        table_container = ttk.Frame(parent)
        table_container.pack(fill=tk.BOTH, expand=True)

        columns = (
            "status_badge",
            "mgmt_sheet",
            "mgmt_row",
            "order_no",
            "instruction_disp",
            "ship_to",
            "line_name",
            "product_detail",
            "supplier",
            "maker",
            "item_name",
            "model_spec",
            "qty",
            "due_date",
            "actual_date",
            "invoice_date",
            "person",
            "mismatch_reason",
        )

        self.tree_missing = ttk.Treeview(table_container, columns=columns, show="headings", selectmode="browse")

        self.col_configs_missing = [
            ("status_badge", "状態判定", 130, "center"),
            ("mgmt_sheet", "管理表シート", 85, "center"),
            ("mgmt_row", "行番号", 55, "center"),
            ("order_no", "管理No(8桁)", 95, "center"),
            ("instruction_disp", "指示書(J列)", 65, "center"),
            ("ship_to", "出荷先", 95, "w"),
            ("line_name", "品名・ライン名等", 115, "w"),
            ("product_detail", "製品詳細", 115, "w"),
            ("supplier", "業者 / 発注先", 105, "w"),
            ("maker", "メーカー", 80, "w"),
            ("item_name", "品名", 115, "w"),
            ("model_spec", "型番・規格・図番", 135, "w"),
            ("qty", "購入数", 45, "center"),
            ("due_date", "管理表納期", 85, "center"),
            ("actual_date", "現品確認日", 85, "center"),
            ("invoice_date", "請求日", 85, "center"),
            ("person", "担当", 55, "center"),
            ("mismatch_reason", "未検出診断内容", 210, "w"),
        ]

        self.col_titles_missing = {col_id: title for col_id, title, _, _ in self.col_configs_missing}

        for col_id, _, width, anchor in self.col_configs_missing:
            self._update_heading_display_missing(col_id)
            self.tree_missing.column(col_id, width=width, anchor=anchor)

        self.tree_missing.bind("<Button-1>", self._on_tree_missing_click)
        self.tree_missing.bind("<Double-1>", lambda e: self._on_tree_double_click(e, is_missing_tab=True))
        self.tree_missing.bind("<Button-3>", lambda e: self._on_tree_right_click(e, is_missing_tab=True))

        self.tree_missing.tag_configure("MISSING_ORDER", background="#fef3c7", foreground="#92400e")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree_missing.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree_missing.xview)
        self.tree_missing.configure(yscrollcommand=vsb.set)

        self.tree_missing.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

    def _create_kpi_card(
        self,
        parent,
        filter_key: str,
        title: str,
        value_var: tk.StringVar,
        fg_color: str,
        bg_color: str,
        accent_color: str,
        hover_bg: str,
        inactive_btn_text: str,
        active_btn_text: str,
    ) -> tk.Frame:
        """KPIサマリー表示 兼 ワンクリックフィルター切替カードの生成"""
        card_outer = tk.Frame(
            parent,
            bg=bg_color,
            highlightbackground="#cbd5e1",
            highlightthickness=2,
            cursor="hand2",
        )

        top_strip = tk.Frame(card_outer, bg=hover_bg, height=4, cursor="hand2")
        top_strip.pack(fill=tk.X, side=tk.TOP)

        body = tk.Frame(card_outer, bg=bg_color, padx=7, pady=5, cursor="hand2")
        body.pack(fill=tk.BOTH, expand=True)

        lbl_title = tk.Label(
            body,
            text=title,
            bg=bg_color,
            fg=fg_color,
            font=("Meiryo UI", 8, "bold"),
            anchor="w",
            cursor="hand2",
        )
        lbl_title.pack(fill=tk.X)

        lbl_val = tk.Label(
            body,
            textvariable=value_var,
            bg=bg_color,
            fg=fg_color,
            font=("Meiryo UI", 14, "bold"),
            anchor="e",
            cursor="hand2",
        )
        lbl_val.pack(fill=tk.X, pady=(1, 4))

        btn_filter = tk.Button(
            body,
            text=inactive_btn_text,
            bg="#ffffff",
            fg=fg_color,
            activebackground=accent_color,
            activeforeground="#ffffff",
            font=("Meiryo UI", 8, "bold"),
            relief="solid",
            bd=1,
            pady=2,
            cursor="hand2",
            command=lambda k=filter_key: self._apply_kpi_filter(k),
        )
        btn_filter.pack(fill=tk.X)

        self.kpi_cards[filter_key] = {
            "card_outer": card_outer,
            "top_strip": top_strip,
            "body": body,
            "lbl_title": lbl_title,
            "lbl_val": lbl_val,
            "btn_filter": btn_filter,
            "title": title,
            "value_var": value_var,
            "fg_color": fg_color,
            "bg_color": bg_color,
            "accent_color": accent_color,
            "hover_bg": hover_bg,
            "inactive_btn_text": inactive_btn_text,
            "active_btn_text": active_btn_text,
        }

        def _on_card_click(_e=None):
            self._apply_kpi_filter(filter_key)

        def _on_enter(_e=None):
            card_outer.configure(highlightbackground=accent_color)
            if not self._is_kpi_card_active(filter_key):
                btn_filter.configure(bg=hover_bg)

        def _on_leave(_e=None):
            self._update_kpi_card_styles()

        for w in (card_outer, top_strip, body, lbl_title, lbl_val):
            w.bind("<Button-1>", _on_card_click)
            w.bind("<Enter>", _on_enter)
            w.bind("<Leave>", _on_leave)

        btn_filter.bind("<Enter>", _on_enter)
        btn_filter.bind("<Leave>", _on_leave)

        return card_outer

    def _is_kpi_card_active(self, filter_key: str) -> bool:
        """指定されたKPIカードが現在アクティブなフィルター（または表示中タブ）かどうかを判定"""
        try:
            curr_tab = self.notebook.index("current")
        except Exception:
            curr_tab = 0
        if filter_key == "MISSING_TAB":
            return curr_tab == 1
        return curr_tab == 0 and self.filter_mode.get() == filter_key

    def _apply_kpi_filter(self, filter_key: str):
        """KPIカードのボタン（またはカード本体）が押されたときにフィルターとタブを切り替える"""
        if filter_key == "MISSING_TAB":
            try:
                self.notebook.select(1)
            except Exception:
                pass
            self._update_kpi_card_styles()
            return

        self.filter_mode.set(filter_key)
        try:
            if self.notebook.index("current") != 0:
                self.notebook.select(0)
        except Exception:
            pass
        self._refresh_main_table()
        self._update_kpi_card_styles()

    def _update_kpi_card_styles(self):
        """現在の選択フィルターおよび選択タブに応じて、全KPIカードとフィルターバッジの見た目を更新する"""
        if not hasattr(self, "kpi_cards"):
            return

        for f_key, info in self.kpi_cards.items():
            is_act = self._is_kpi_card_active(f_key)
            card_outer = info["card_outer"]
            top_strip = info["top_strip"]
            btn_filter = info["btn_filter"]
            fg_color = info["fg_color"]
            bg_color = info["bg_color"]
            accent_color = info["accent_color"]
            hover_bg = info["hover_bg"]

            if is_act:
                card_outer.configure(highlightbackground=accent_color, highlightthickness=2, bg=bg_color)
                top_strip.configure(bg=accent_color)
                btn_filter.configure(
                    text=info["active_btn_text"],
                    bg=accent_color,
                    fg="#ffffff",
                    activebackground=accent_color,
                    activeforeground="#ffffff",
                    relief="flat",
                    bd=0,
                )
            else:
                card_outer.configure(highlightbackground="#cbd5e1", highlightthickness=1, bg=bg_color)
                top_strip.configure(bg=hover_bg)
                btn_filter.configure(
                    text=info["inactive_btn_text"],
                    bg="#ffffff",
                    fg=fg_color,
                    activebackground=hover_bg,
                    activeforeground=fg_color,
                    relief="solid",
                    bd=1,
                )

        if hasattr(self, "active_filter_badge"):
            curr_mode = self.filter_mode.get()
            act_info = self.kpi_cards.get(curr_mode)
            if act_info:
                shown_cnt = len(self.tree_main.get_children()) if hasattr(self, "tree_main") else 0
                total_cnt = len(self._get_visible_main_records())
                extra_tags = []
                inst_mode = self._get_instruction_filter_mode()
                if inst_mode == "HAS_OR_BLANK":
                    extra_tags.append("J列:有り/空白のみ")
                elif inst_mode == "HAS_ONLY":
                    extra_tags.append("J列:有りのみ")
                elif inst_mode == "BLANK_ONLY":
                    extra_tags.append("J列:空白のみ")
                if self.search_keyword.get().strip():
                    extra_tags.append("キーワード絞込中")
                if self.column_filters_main:
                    extra_tags.append(f"列フィルタ{len(self.column_filters_main)}列")
                extra_str = f" / {'・'.join(extra_tags)}" if extra_tags else ""
                badge_text = (
                    f"  ✔ {act_info['title']}  （表示: {shown_cnt} 件 / 全 {total_cnt} 件{extra_str}）  "
                )
                self.active_filter_badge.configure(
                    text=badge_text,
                    bg=act_info["bg_color"],
                    fg=act_info["fg_color"],
                )

    # =========================================================================
    # 注文品管理表・注文書の直接参照＆プレビュー機能
    # =========================================================================
    def _open_excel_at_location(self, filepath: str, sheet_name: str = "", row_index: int | None = None, is_mgmt: bool = False):
        from modules.excel_exporter import open_excel_read_only_at_location
        open_excel_read_only_at_location(filepath, sheet_name, row_index, is_mgmt)
    def _open_mgmt_file_direct(self):
        """上部の設定パスにある注文品管理表ファイルをExcelで開く"""
        mgmt_path = self.mgmt_file_path.get().strip().strip('"')
        self._open_excel_at_location(mgmt_path, label_for_status="注文品管理表", read_only_strict=False)

    def _open_order_dir_direct(self):
        """上部の設定パスにある注文書フォルダをエクスプローラーで開く"""
        order_dir = self.order_dir_path.get().strip().strip('"')
        if not order_dir or not os.path.isdir(order_dir):
            messagebox.showwarning("フォルダ未検出", f"有効な注文書フォルダが指定されていません:\n{order_dir or '(未指定)'}")
            return
        try:
            os.startfile(os.path.abspath(order_dir))
        except Exception as e:
            messagebox.showerror("フォルダ起動エラー", f"フォルダを開けませんでした:\n{e}")

    def _get_selected_record(self, is_missing_tab: bool) -> dict | None:
        tree = self.tree_missing if is_missing_tab else self.tree_main
        records = self.missing_order_records if is_missing_tab else self.analyzed_records
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("行未選択", "一覧テーブルから参照したい行をクリックして選択してください。")
            return None
        item_id = sel[0]
        prefix = "mis_" if is_missing_tab else "rec_"
        rec_idx_str = item_id.replace(prefix, "")
        if not rec_idx_str.isdigit():
            return None
        rec_idx = int(rec_idx_str)
        if not (0 <= rec_idx < len(records)):
            return None
        return records[rec_idx]

    def _open_selected_row_order(self):
        """タブ1の選択行に対応する注文書Excelを【読み取り専用・保存ブロック】で開き、該当シート・該当行を選択表示する"""
        rec = self._get_selected_record(is_missing_tab=False)
        if not rec:
            return
        diag = rec.get("diagnostic_data", {})
        order_path = diag.get("order_full_path", "")
        order_sheet = diag.get("order_sheet", "")
        order_row = diag.get("order_row")
        if not order_path:
            messagebox.showinfo("注文書なし", "この明細には対応する注文書ファイルパスが紐付いていません。")
            return
        self._open_excel_at_location(
            order_path,
            sheet_name=order_sheet,
            row_num=order_row,
            label_for_status="注文書",
            read_only_strict=True,
        )

    def _open_selected_row_mgmt(self, is_missing_tab: bool = False):
        """選択行に対応する注文品管理表Excelを開き、該当シート・該当行を選択表示する"""
        rec = self._get_selected_record(is_missing_tab=is_missing_tab)
        if not rec:
            return
        mgmt_path = self.mgmt_file_path.get().strip().strip('"')
        diag = rec.get("diagnostic_data", {})
        mgmt_sheet = diag.get("mgmt_sheet") or rec.get("mgmt_sheet", "")
        mgmt_row = diag.get("mgmt_row") or (int(rec["mgmt_row"]) if str(rec.get("mgmt_row", "")).isdigit() else None)

        # 不一致行などで mgmt_row が未確定の場合、同じ管理Noの先頭行を探す
        if not mgmt_row and self.cached_mgmt_rows:
            target_no = rec.get("order_no", "")
            for m in self.cached_mgmt_rows:
                if m.get("order_no_8") == target_no or m.get("order_no") == target_no:
                    mgmt_sheet = m.get("sheet", "")
                    mgmt_row = m.get("row")
                    break

        self._open_excel_at_location(
            mgmt_path,
            sheet_name=mgmt_sheet,
            row_num=mgmt_row,
            label_for_status="注文品管理表",
        )

    def _open_active_tab_selected_mgmt(self):
        try:
            is_missing = self.notebook.index("current") == 1
        except Exception:
            is_missing = False
        self._open_selected_row_mgmt(is_missing_tab=is_missing)

    def _open_active_tab_selected_detail(self):
        try:
            is_missing = self.notebook.index("current") == 1
        except Exception:
            is_missing = False
        self._open_selected_row_detail(is_missing_tab=is_missing)

    def _open_order_file_parent_folder(self, order_full_path: str):
        if order_full_path and os.path.isfile(order_full_path):
            folder = os.path.dirname(os.path.abspath(order_full_path))
            try:
                os.startfile(folder)
                return
            except Exception:
                pass
        self._open_order_dir_direct()

    def _on_tree_right_click(self, event, is_missing_tab: bool):
        """一覧行の右クリック時にコンテキストメニューを表示し、注文書・管理表を即座に参照可能にする"""
        tree = self.tree_missing if is_missing_tab else self.tree_main
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            return
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        tree.selection_set(row_id)
        tree.focus(row_id)

        rec = self._get_selected_record(is_missing_tab=is_missing_tab)
        if not rec:
            return
        diag = rec.get("diagnostic_data", {})
        order_path = diag.get("order_full_path", "")
        order_sheet = diag.get("order_sheet", "")
        order_row = diag.get("order_row")
        mgmt_sheet = diag.get("mgmt_sheet") or rec.get("mgmt_sheet", "")
        mgmt_row = diag.get("mgmt_row") or rec.get("mgmt_row", "")

        menu = tk.Menu(self.root, tearoff=0, font=("Meiryo UI", 9))
        menu.add_command(
            label="🔍 照合詳細・シートプレビューを表示 (ダブルクリック)",
            command=lambda: self._open_selected_row_detail(is_missing_tab=is_missing_tab),
        )
        menu.add_separator()
        if not is_missing_tab and order_path:
            o_loc = f" ({os.path.basename(order_path)} / 行{order_row})" if order_row else ""
            menu.add_command(
                label=f"📄 この明細の『注文書』をExcelで開く 【読取専用・保存ブロック】{o_loc}",
                command=self._open_selected_row_order,
            )
        m_loc = f" ({mgmt_sheet} / 行{mgmt_row})" if mgmt_sheet and mgmt_row else ""
        menu.add_command(
            label=f"📋 この明細の『注文品管理表』をExcelで開く{m_loc}",
            command=lambda: self._open_selected_row_mgmt(is_missing_tab=is_missing_tab),
        )
        if not is_missing_tab and order_path:
            menu.add_separator()
            menu.add_command(
                label="📂 この注文書の保存先フォルダを開く",
                command=lambda: self._open_order_file_parent_folder(order_path),
            )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _on_tree_double_click(self, event, is_missing_tab: bool):
        tree = self.tree_missing if is_missing_tab else self.tree_main
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            return
        self._open_selected_row_detail(is_missing_tab=is_missing_tab)

    def _open_selected_row_detail(self, is_missing_tab: bool = False):
        rec, is_missing = self._get_selected_record()
        if not rec:
            messagebox.showinfo("行選択", "詳細・照合診断を表示する行を一覧から選択してください。")
            return
        from modules.dialogs import open_detail_preview_dialog

        open_detail_preview_dialog(
            parent=self.root,
            rec=rec,
            cached_mgmt_rows=self.cached_mgmt_rows,
            is_missing_tab=is_missing,
            icon_path=self._resolve_icon_path(),
            on_open_order=lambda r: self._open_excel_at_location(
                r.get("filepath", ""), r.get("sheet_name", ""), r.get("row_index"), is_mgmt=False
            ),
            on_open_mgmt=lambda r: self._open_excel_at_location(
                self.mgmt_file_path.get().strip(),
                r.get("mgmt_sheet", "") or r.get("sheet_name", ""),
                r.get("mgmt_row_index"),
                is_mgmt=True,
            ),
        )
    def _show_help_guide(self, tab_index: int = 0):
        from modules.dialogs import show_help_guide
        show_help_guide(self.root, initial_tab=tab_index, icon_path=self._resolve_icon_path())
    def _show_about(self):
        from modules.dialogs import show_about
        show_about(self.root, icon_path=self._resolve_icon_path())
    def _resolve_initial_config_path(cls) -> str:
        """起動時に config.json の保存先をポインタファイル等から決定して返す"""
        ptr_file = cls.CONFIG_POINTER_FILE_PATH
        if os.path.isfile(ptr_file):
            try:
                with open(ptr_file, "r", encoding="utf-8") as f:
                    custom_target = f.read().strip()
                if custom_target:
                    if os.path.isdir(custom_target):
                        cand = os.path.join(custom_target, "config.json")
                    elif custom_target.lower().endswith(".json"):
                        cand = custom_target
                    else:
                        cand = os.path.join(custom_target, "config.json")
                    cand_dir = os.path.dirname(cand)
                    if os.path.isdir(cand_dir) or not os.path.exists(cand_dir):
                        return os.path.abspath(cand)
            except Exception:
                pass
        return os.path.abspath(cls.DEFAULT_CONFIG_FILE_PATH)

    def get_config_file_path(self) -> str:
        """現在アクティブな config.json の絶対パスを取得"""
        if hasattr(self, "config_file_path"):
            val = self.config_file_path.get().strip().strip('"')
            if val:
                return os.path.abspath(val)
        return os.path.abspath(self.DEFAULT_CONFIG_FILE_PATH)

    def _open_config_folder(self):
        """互換用: 現在の設定ファイル保存先フォルダを開く"""
        self._open_current_config_folder()

    def _open_current_config_folder(self):
        """現在設定されている config.json の保存先フォルダを開く"""
        try:
            cfg_p = self.get_config_file_path()
            folder = os.path.dirname(cfg_p)
            if os.path.isdir(folder):
                os.startfile(folder)
            else:
                os.startfile(self.APP_DIR)
        except Exception as e:
            messagebox.showinfo("フォルダパス", f"設定ファイル保存先フォルダ:\n{self.get_config_file_path()}\n\n({e})")

    def _open_app_folder(self):
        """アプリ本体の配置フォルダを開く"""
        try:
            os.startfile(self.APP_DIR)
        except Exception as e:
            messagebox.showinfo("フォルダパス", f"アプリ配置フォルダ:\n{self.APP_DIR}\n\n({e})")

    # =========================================================================
    # 設定ファイル処理（注文書・管理表・Excelファイルへの書き込み完全禁止ガード）
    # =========================================================================
    def _is_strictly_allowed_path(self, target_filepath: str) -> bool:
        """
        書き込み可能なパスを設定ファイル (config.json) およびポインタファイルに厳格制限する。
        Excelファイル（.xlsx/.xlsm/.xls等）や注文書フォルダ配下、注文品管理表への書き込みは絶対拒否する。
        """
        try:
            if not target_filepath:
                return False
            norm_target = os.path.normcase(os.path.realpath(os.path.abspath(target_filepath)))
            # Excel関連拡張子への書き込みはいかなる場合も絶対拒否
            if norm_target.endswith((".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".csv")):
                return False
            # 注文品管理表ファイル自体への書き込みを絶対拒否
            mgmt_p = self.mgmt_file_path.get().strip().strip('"') if hasattr(self, "mgmt_file_path") else ""
            if mgmt_p:
                norm_mgmt = os.path.normcase(os.path.realpath(os.path.abspath(mgmt_p)))
                if norm_target == norm_mgmt:
                    return False

            allowed_norms = [
                os.path.normcase(os.path.realpath(os.path.abspath(self.get_config_file_path()))),
                os.path.normcase(os.path.realpath(os.path.abspath(self.DEFAULT_CONFIG_FILE_PATH))),
                os.path.normcase(os.path.realpath(os.path.abspath(self.CONFIG_POINTER_FILE_PATH))),
            ]
            if norm_target in allowed_norms:
                return True
            # 新規保存先として指定された .json ファイルまたはポインタファイル
            if norm_target.endswith(".json") or norm_target.endswith(".txt"):
                return True
            return False
        except Exception:
            return False

    def _set_config_file_path(self, new_target: str, copy_existing: bool = True) -> bool:
        """config.json の保存先を指定されたパスに変更し、ポインタファイルに記録する"""
        if not new_target:
            return False
        clean = new_target.strip().strip('"')
        if os.path.isdir(clean):
            target_path = os.path.join(clean, "config.json")
        elif clean.lower().endswith(".json"):
            target_path = clean
        else:
            target_path = os.path.join(clean, "config.json")

        target_abs = os.path.abspath(target_path)
        old_path = self.get_config_file_path()

        if target_abs.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            messagebox.showerror("エラー", "Excelファイルを設定ファイルの保存先として指定することはできません。")
            return False

        target_dir = os.path.dirname(target_abs)
        try:
            os.makedirs(target_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("エラー", f"保存先フォルダを作成・確認できませんでした:\n{target_dir}\n\n({e})")
            return False

        # 既存設定のコピー
        if copy_existing and old_path and os.path.isfile(old_path) and os.path.normcase(old_path) != os.path.normcase(target_abs):
            if not os.path.isfile(target_abs):
                try:
                    import shutil
                    shutil.copy2(old_path, target_abs)
                except Exception:
                    pass

        # デフォルト保存先と同じならポインタファイルを削除、異なるなら記録
        norm_target = os.path.normcase(target_abs)
        norm_default = os.path.normcase(os.path.abspath(self.DEFAULT_CONFIG_FILE_PATH))

        if norm_target == norm_default:
            try:
                if os.path.isfile(self.CONFIG_POINTER_FILE_PATH):
                    os.remove(self.CONFIG_POINTER_FILE_PATH)
            except Exception:
                pass
        else:
            try:
                with open(self.CONFIG_POINTER_FILE_PATH, "w", encoding="utf-8") as f:
                    f.write(target_abs)
            except Exception as e:
                messagebox.showwarning(
                    "保存先警告",
                    f"保存先ポインタファイルへの記録に失敗しました:\n{e}\n次回起動時に反映されない可能性があります。"
                )

        self.config_file_path.set(target_abs)
        return True

    def _reset_config_path_to_default(self, confirm: bool = True) -> bool:
        """ボタン一つで config.json の保存先をデフォルト（アプリと同じ場所）に戻す"""
        default_path = os.path.abspath(self.DEFAULT_CONFIG_FILE_PATH)
        curr_path = self.get_config_file_path()
        if os.path.normcase(curr_path) == os.path.normcase(default_path):
            if confirm:
                messagebox.showinfo("情報", f"設定ファイル(config.json)の保存先は既にデフォルト（アプリと同じ場所）に設定されています:\n{default_path}")
            return True

        if confirm:
            if not messagebox.askyesno(
                "デフォルトに戻す確認",
                "設定ファイル(config.json)の保存先を、初期値「アプリと同じ場所」に戻しますか？\n\n"
                f"現在の保存先:\n{curr_path}\n\n"
                f"戻し先（アプリ配置フォルダ）:\n{default_path}"
            ):
                return False

        success = self._set_config_file_path(default_path, copy_existing=True)
        if success:
            self._load_config_and_autoload()
            if confirm:
                messagebox.showinfo(
                    "設定完了",
                    f"設定ファイル(config.json)の保存先をアプリと同じ場所に戻しました:\n{default_path}"
                )
        return success

    def _open_config_location_dialog(self):
        from modules.dialogs import open_config_location_dialog
        open_config_location_dialog(
            parent=self.root,
            config_mgr=self,
            icon_path=self._resolve_icon_path(),
            on_success=lambda p: (self.config_file_path.set(p), self._load_config_and_autoload()),
        )
    def _save_config(self, show_msg: bool = False):
        mgmt = self.mgmt_file_path.get().strip().strip('"')
        order = self.order_dir_path.get().strip().strip('"')
        grace_days = self._get_grace_days()
        exclude_months = self._get_exclude_months()
        inst_filter_mode = self._get_instruction_filter_mode()
        excluded_list = self._get_excluded_order_nos_list()

        self.blank_due_grace_days.set(str(grace_days))
        self.exclude_after_months.set(str(exclude_months))
        self._set_instruction_filter_mode(inst_filter_mode)

        data = {
            "mgmt_file_path": mgmt,
            "order_dir_path": order,
            "blank_due_grace_days": grace_days,
            "exclude_after_months": exclude_months,
            "instruction_filter_mode": inst_filter_mode,
            "excluded_order_nos": "\n".join(excluded_list),
            "visible_columns_main": self.visible_columns_main,
            "last_updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        target_file = self.get_config_file_path()

        if not self._is_strictly_allowed_path(target_file):
            err_msg = f"保存エラー: 許可されていない保存先への書き込み要求です。\n対象: {target_file}"
            if show_msg:
                messagebox.showerror("保存エラー", err_msg)
            return

        try:
            target_dir = os.path.dirname(target_file)
            if target_dir:
                os.makedirs(target_dir, exist_ok=True)

            with open(target_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            # 設定保存と同時に一覧と件数サマリーの除外状態を同期
            self._update_excluded_summary_badge()
            self._update_kpis()
            self._refresh_main_table()
            self._refresh_missing_table()

            if show_msg:
                ex_info = (
                    f"\n・除外管理No: {len(excluded_list)}件設定中"
                    if excluded_list
                    else "\n・除外管理No: 未設定"
                )
                months_info = (
                    f"\n・経過月数による自動除外: 注文日から {exclude_months} ヶ月経過で除外"
                    if exclude_months > 0
                    else "\n・経過月数による自動除外: なし (0ヶ月)"
                )
                inst_label = self.INSTRUCTION_CODE_TO_LABEL.get(inst_filter_mode, inst_filter_mode)
                inst_info = f"\n・管理表 指示書列(J列)フィルタ: {inst_label}"
                messagebox.showinfo(
                    "設定保存完了",
                    f"設定を保存しました:\n{target_file}\n\n"
                    f"・納期未記載時の超過基準: 注文日から {grace_days} 日経過"
                    f"{months_info}{inst_info}{ex_info}",
                )
        except Exception as e:
            if show_msg:
                messagebox.showerror("保存エラー", f"設定ファイルへの保存に失敗しました:\n{e}")

    def _save_config_manual(self):
        self._save_config(show_msg=True)

    def _load_config_and_autoload(self):
        cfg_path = self.get_config_file_path()
        if os.path.exists(cfg_path) and self._is_strictly_allowed_path(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                mgmt = data.get("mgmt_file_path", "")
                order = data.get("order_dir_path", "")
                grace_days = data.get("blank_due_grace_days", 7)
                exclude_months = data.get("exclude_after_months", 12)
                inst_filter_mode = data.get("instruction_filter_mode", "HAS_ONLY")
                excluded_nos = data.get("excluded_order_nos", "")
                if mgmt and os.path.isfile(mgmt):
                    self.mgmt_file_path.set(mgmt)
                if order and os.path.isdir(order):
                    self.order_dir_path.set(order)
                if grace_days is not None:
                    self.blank_due_grace_days.set(str(grace_days))
                if exclude_months is not None:
                    self.exclude_after_months.set(str(exclude_months))
                if inst_filter_mode:
                    self._set_instruction_filter_mode(str(inst_filter_mode))
                if isinstance(excluded_nos, list):
                    self.excluded_order_nos_str.set("\n".join(str(x) for x in excluded_nos))
                elif excluded_nos:
                    parsed = self._parse_lines_to_order_nos(str(excluded_nos))
                    self.excluded_order_nos_str.set("\n".join(parsed))

                visible_cols = data.get("visible_columns_main")
                if isinstance(visible_cols, list) and visible_cols:
                    all_cids = [cfg[0] for cfg in self.col_configs_main]
                    valid_cols = [c for c in visible_cols if c in all_cids]
                    if valid_cols:
                        self.visible_columns_main = valid_cols
                        self._apply_column_visibility_main()
            except Exception:
                pass

        self._update_excluded_summary_badge()

        if not self.mgmt_file_path.get() or not self.order_dir_path.get():
            self._autoload_default_files()

    def _autoload_default_files(self):
        cwd = self.APP_DIR
        xlsx_files = [
            f for f in os.listdir(cwd) if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
        ]

        if not self.mgmt_file_path.get():
            mgmt_candidates = [f for f in xlsx_files if "管理表" in f]
            if mgmt_candidates:
                self.mgmt_file_path.set(os.path.join(cwd, sorted(mgmt_candidates)[0]))

        if not self.order_dir_path.get():
            order_candidates = [f for f in xlsx_files if "管理表" not in f and "納期超過チェック結果" not in f]
            if order_candidates:
                self.order_dir_path.set(cwd)

    def _select_mgmt_file(self):
        initial_dir = os.path.dirname(self.mgmt_file_path.get().strip().strip('"'))
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = self.APP_DIR

        path = filedialog.askopenfilename(
            title="注文品管理表のExcelファイルを選択",
            initialdir=initial_dir,
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls"), ("All Files", "*.*")],
        )
        if path:
            self.mgmt_file_path.set(path)

    def _select_order_dir(self):
        initial_dir = self.order_dir_path.get().strip().strip('"')
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = self.APP_DIR
        folder = filedialog.askdirectory(title="注文書が入っているフォルダを選択", initialdir=initial_dir)
        if folder:
            self.order_dir_path.set(folder)

    def _report_progress(self, percent: float, message: str):
        if self.cancel_event.is_set():
            raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
        self.ui_queue.put(("PROGRESS", (percent, message)))

    def _collect_valid_order_files(self, order_dir: str, mgmt_file_path: str, valid_mgmt_order_nos: set, progress_cb=None):
        from modules.matcher import collect_valid_order_files
        return collect_valid_order_files(order_dir, mgmt_file_path, valid_mgmt_order_nos, progress_cb=progress_cb, cancel_event=self.cancel_event)
    @staticmethod
    def _parse_flexible_date(date_val) -> datetime.date | None:
        from modules.date_utils import parse_flexible_date
        return parse_flexible_date(date_val)
    def _parse_management_book(self, filepath: str, progress_cb=None):
        from modules.matcher import parse_management_book
        return parse_management_book(filepath, progress_cb=progress_cb, cancel_event=self.cancel_event)
    def _parse_order_book(self, filepath: str, filename: str, file_order_no_8: str = ""):
        from modules.matcher import parse_order_book
        return parse_order_book(filepath, filename, file_order_no_8, cancel_event=self.cancel_event)
    def run_analysis(self):
        if self.is_running:
            return

        mgmt_path = self.mgmt_file_path.get().strip().strip('"')
        if not mgmt_path or not os.path.isfile(mgmt_path):
            messagebox.showerror("ファイルエラー", "有効な『注文品管理表』のファイルパスを指定してください。")
            return

        order_dir = self.order_dir_path.get().strip().strip('"')
        if not order_dir or not os.path.isdir(order_dir):
            messagebox.showerror("フォルダエラー", "有効な『注文書フォルダ』のパスを指定してください。")
            return

        base_date = self._parse_flexible_date(self.base_date_str.get())
        if not base_date:
            messagebox.showerror("日付エラー", "判定基準日は YYYY-MM-DD 形式で入力してください。")
            return

        self.is_running = True
        self.cancel_event.clear()
        self.run_btn.configure(state=tk.DISABLED, text="照合処理中...")
        self.cancel_btn.configure(state=tk.NORMAL, text="⏹ 中止 (Esc)")
        self.excel_open_btn.configure(state=tk.DISABLED)
        self.excel_cancel_btn.configure(state=tk.DISABLED)
        self.progress_var.set(2.0)
        self.progress_text_var.set("解析を開始しています...")
        self.status_var.set("実行中: 注文品管理表と注文書フォルダを照合しています...")

        grace_days = self._get_grace_days()
        worker = threading.Thread(
            target=self._analysis_worker,
            args=(mgmt_path, order_dir, base_date, grace_days),
            daemon=True,
        )
        worker.start()
        self.root.after(40, self._poll_ui_queue)

    def cancel_active_task(self):
        """実行中のタスク（照合またはExcel直接展開）を安全に中断する"""
        if self.is_running:
            self.cancel_analysis()
        elif self.is_excel_running:
            self.cancel_excel_open()

    def cancel_analysis(self):
        """実行中の照合・遅延チェック処理を安全に中断する"""
        if not self.is_running:
            return
        self.cancel_event.set()
        self.cancel_btn.configure(state=tk.DISABLED, text="中止処理中...")
        self.status_var.set("照合処理の中止を要求しました。安全に中断しています...")
        self.progress_text_var.set("処理を中止しています...")

    def cancel_excel_open(self):
        """実行中のExcel直接展開処理を安全に中断する"""
        if not self.is_excel_running:
            return
        self.excel_cancel_event.set()
        self.excel_cancel_btn.configure(state=tk.DISABLED, text="中止中...")
        self.cancel_btn.configure(state=tk.DISABLED, text="中止中...")
        self.status_var.set("Excel展開処理の中止を要求しました。安全に中断しています...")
        self.progress_text_var.set("Excel処理を中止しています...")

    def _analysis_worker(self, mgmt_path: str, order_dir: str, base_date: datetime.date, grace_days: int = 7):
        try:
            self._report_progress(5.0, "[1/3] 注文品管理表を開いています...")
            mgmt_rows = self._parse_management_book(mgmt_path, progress_cb=self._report_progress)
            if self.cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

            excluded_nos = self._get_excluded_order_nos()
            if excluded_nos:
                # 除外指定された管理Noの行を管理表データから除外（タブ2の注文書未検出一覧からも除外される）
                mgmt_rows = [
                    m
                    for m in mgmt_rows
                    if not self._is_order_no_excluded(m.get("order_no_8") or m.get("order_no", ""), excluded_nos)
                ]

            valid_mgmt_order_nos = set()
            for m in mgmt_rows:
                no8 = m.get("order_no_8") or m.get("norm_order_no")
                if no8 and not self._is_order_no_excluded(no8, excluded_nos):
                    valid_mgmt_order_nos.add(no8)

            ex_msg = f" (除外指定: {len(excluded_nos)}件)" if excluded_nos else ""
            self._report_progress(
                35.0,
                f"[2/3] 管理表にある管理No ({len(valid_mgmt_order_nos)}種類{ex_msg}) の注文書を探索中...",
            )
            order_items, valid_order_files, ignored_files = self._collect_valid_order_files(
                order_dir, mgmt_path, valid_mgmt_order_nos, progress_cb=self._report_progress
            )
            if self.cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

            if excluded_nos:
                # 抽出された注文書明細からも除外管理Noをフィルタリング（タブ1の照合一覧からも除外される）
                order_items = [
                    o
                    for o in order_items
                    if not self._is_order_no_excluded(o.get("order_no_8") or o.get("order_no", ""), excluded_nos)
                ]

            self._report_progress(75.0, f"[3/3] 明細の双方向照合を実行中 ({len(order_items)} 件)...")
            main_records, missing_records = self._match_and_evaluate(
                mgmt_rows, order_items, base_date, grace_days=grace_days, progress_cb=self._report_progress
            )
            if self.cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

            self.ui_queue.put(
                (
                    "DONE",
                    {
                        "main_records": main_records,
                        "missing_records": missing_records,
                        "mgmt_rows": mgmt_rows,
                        "base_date": base_date,
                        "valid_order_files": valid_order_files,
                        "ignored_files": ignored_files,
                        "order_items_count": len(order_items),
                        "mgmt_rows_count": len(mgmt_rows),
                        "mgmt_no_count": len(valid_mgmt_order_nos),
                    },
                )
            )
        except AnalysisCancelledException:
            self.ui_queue.put(("CANCELLED", None))
        except Exception as e:
            self.ui_queue.put(("ERROR", str(e)))

    def _poll_ui_queue(self):
        try:
            while True:
                msg_type, payload = self.ui_queue.get_nowait()
                if msg_type == "PROGRESS":
                    pct, text = payload
                    self.progress_var.set(pct)
                    self.progress_text_var.set(text)
                    self.status_var.set(f"実行中: {text}")
                elif msg_type == "DONE":
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL, text="▶ 照合・遅延チェック実行 (F5)")
                    self.cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止 (Esc)")
                    self.excel_cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止")
                    self.excel_open_btn.configure(state=tk.NORMAL)

                    self.analyzed_records = payload["main_records"]
                    self.missing_order_records = payload["missing_records"]
                    self.cached_mgmt_rows = payload.get("mgmt_rows", [])
                    base_date = payload["base_date"]
                    valid_order_files = payload["valid_order_files"]
                    mgmt_rows_count = payload["mgmt_rows_count"]
                    mgmt_no_count = payload.get("mgmt_no_count", 0)

                    self.column_filters_main.clear()
                    self.column_filters_missing.clear()
                    self._update_all_headings_main()
                    self._update_all_headings_missing()

                    self._update_kpis()
                    self._refresh_main_table()
                    self._refresh_missing_table()

                    vis_main = self._get_visible_main_records()
                    vis_missing = self._get_visible_missing_records()
                    delay_cnt = sum(1 for r in vis_main if r["is_any_delay"])
                    unmatched_cnt = sum(1 for r in vis_main if r.get("is_unmatched"))
                    missing_cnt = len(vis_missing)
                    excluded_main_cnt = len(self.analyzed_records) - len(vis_main)
                    ex_note = f" (フィルタ除外:{excluded_main_cnt}件)" if excluded_main_cnt > 0 else ""

                    self.progress_var.set(100.0)
                    self.progress_text_var.set(
                        f"完了！ (注文書 {valid_order_files}ファイル / 管理表 {mgmt_rows_count}行・管理No {mgmt_no_count}種)"
                    )
                    self.status_var.set(
                        f"照合完了 (基準日: {base_date.strftime('%Y-%m-%d')}) — "
                        f"注文書明細: {len(vis_main)}件{ex_note} (超過:{delay_cnt}件 / 管理表未記載:{unmatched_cnt}件) ｜ "
                        f"【別タブ】注文書未検出(管理表のみ): {missing_cnt}件"
                    )
                    return
                elif msg_type == "CANCELLED":
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL, text="▶ 照合・遅延チェック実行 (F5)")
                    self.cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止 (Esc)")
                    self.excel_cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止")
                    if self.analyzed_records or self.missing_order_records:
                        self.excel_open_btn.configure(state=tk.NORMAL)
                    else:
                        self.excel_open_btn.configure(state=tk.DISABLED)
                    self.progress_var.set(0.0)
                    self.progress_text_var.set("照合が中止されました")
                    self.status_var.set("ユーザーによって照合・遅延チェック処理が中止されました。")
                    messagebox.showinfo("処理中止", "照合・遅延チェック処理を中止しました。")
                    return
                elif msg_type == "ERROR":
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL, text="▶ 照合・遅延チェック実行 (F5)")
                    self.cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止 (Esc)")
                    self.excel_cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止")
                    if self.analyzed_records or self.missing_order_records:
                        self.excel_open_btn.configure(state=tk.NORMAL)
                    else:
                        self.excel_open_btn.configure(state=tk.DISABLED)
                    self.progress_var.set(0.0)
                    self.progress_text_var.set("エラーが発生しました")
                    self.status_var.set(f"エラー: {payload}")
                    messagebox.showerror("解析エラー", f"処理中にエラーが発生しました:\n{payload}")
                    return
        except queue.Empty:
            pass

        if self.is_running:
            self.root.after(40, self._poll_ui_queue)

    def _score_candidate_pair(self, o_item: dict, m_row: dict) -> float:
        from modules.matcher import score_candidate_pair
        return score_candidate_pair(o_item, m_row)
    def _build_field_comparison_data(self, o_item: dict, m_row: dict, reason_summary: str = "", extra_notes: list | None = None) -> list:
        from modules.matcher import build_field_comparison_data
        return build_field_comparison_data(o_item, m_row, reason_summary=reason_summary, extra_notes=extra_notes)
    def _diagnose_unmatched_item(self, o_item: dict, mgmt_rows: list[dict]) -> dict:
        from modules.matcher import diagnose_unmatched_item
        return diagnose_unmatched_item(o_item, mgmt_rows)
    def _match_and_evaluate(self, mgmt_rows: list, order_items: list, base_date: datetime.date, grace_days: int = 7, progress_cb=None):
        from modules.matcher import match_and_evaluate
        return match_and_evaluate(mgmt_rows, order_items, base_date, grace_days=grace_days, progress_cb=progress_cb, cancel_event=self.cancel_event)
    def _evaluate_single_record(self, **kwargs) -> dict:
        from modules.matcher import evaluate_single_record
        return evaluate_single_record(**kwargs)
    def _update_kpis(self):
        visible_main = self._get_visible_main_records()
        visible_missing = self._get_visible_missing_records()

        total = len(visible_main)
        unmatched_cnt = sum(1 for r in visible_main if r.get("is_unmatched"))
        both_cnt = sum(
            1
            for r in visible_main
            if not r.get("is_unmatched") and r["actual_overdue"] and r["invoice_overdue"]
        )
        actual_cnt = sum(
            1
            for r in visible_main
            if not r.get("is_unmatched") and r["actual_overdue"] and not r["invoice_overdue"]
        )
        invoice_cnt = sum(
            1
            for r in visible_main
            if not r.get("is_unmatched") and not r["actual_overdue"] and r["invoice_overdue"]
        )
        all_delay_cnt = unmatched_cnt + both_cnt + actual_cnt + invoice_cnt
        ok_cnt = sum(
            1
            for r in visible_main
            if not r.get("is_unmatched") and not r["actual_overdue"] and not r["invoice_overdue"]
        )
        missing_cnt = len(visible_missing)

        self.kpi_vars["order_total"].set(f"{total} 件")
        self.kpi_vars["all_delay"].set(f"{all_delay_cnt} 件")
        self.kpi_vars["unmatched"].set(f"{unmatched_cnt} 件")
        self.kpi_vars["both_delay"].set(f"{both_cnt} 件")
        self.kpi_vars["actual_delay"].set(f"{actual_cnt} 件")
        self.kpi_vars["invoice_delay"].set(f"{invoice_cnt} 件")
        self.kpi_vars["ok"].set(f"{ok_cnt} 件")
        self.kpi_vars["missing_order"].set(f"{missing_cnt} 件")

        self._on_tab_changed()

    # =========================================================================
    # テーブル描画＆フィルター処理（タブ1: メイン照合一覧）
    # =========================================================================
    def _refresh_main_table(self):
        for item in self.tree_main.get_children():
            self.tree_main.delete(item)

        mode = self.filter_mode.get()
        kw = self.search_keyword.get().strip().lower()
        excluded_nos = self._get_excluded_order_nos()
        base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
        exclude_months = self._get_exclude_months()
        inst_mode = self._get_instruction_filter_mode()

        for idx, rec in enumerate(self.analyzed_records):
            if self._is_record_excluded(rec, excluded_nos, base_date, exclude_months, inst_mode):
                continue

            if mode == "ALL_DELAY":
                if not (rec.get("is_unmatched") or rec.get("is_any_delay")):
                    continue
            elif mode == "UNMATCHED_ONLY":
                if not rec.get("is_unmatched"):
                    continue
            elif mode == "BOTH_DELAY":
                if rec.get("is_unmatched") or not (rec["actual_overdue"] and rec["invoice_overdue"]):
                    continue
            elif mode == "ACTUAL_DELAY":
                if rec.get("is_unmatched") or not (rec["actual_overdue"] and not rec["invoice_overdue"]):
                    continue
            elif mode == "INVOICE_DELAY":
                if rec.get("is_unmatched") or not (not rec["actual_overdue"] and rec["invoice_overdue"]):
                    continue
            elif mode == "OK_ONLY":
                if rec.get("is_unmatched") or rec["actual_overdue"] or rec["invoice_overdue"]:
                    continue

            skip_by_col = False
            for col_id, allowed_values in self.column_filters_main.items():
                val = str(rec.get(col_id, "") or "").strip()
                if not val:
                    val = "(空白)"
                if val not in allowed_values:
                    skip_by_col = True
                    break
            if skip_by_col:
                continue

            if kw:
                searchable = " ".join(
                    [
                        rec["status_badge"],
                        rec["mismatch_reason"],
                        rec["req_type"],
                        rec["order_no"],
                        rec.get("instruction_disp", ""),
                        rec.get("ship_to", ""),
                        rec.get("line_name", ""),
                        rec.get("product_detail", ""),
                        rec["supplier"],
                        rec["maker"],
                        rec["item_name"],
                        rec["model_spec"],
                        rec["person"],
                        rec["actual_status"],
                        rec["invoice_status"],
                    ]
                ).lower()
                if kw not in searchable:
                    continue

            self.tree_main.insert(
                "",
                tk.END,
                iid=f"rec_{idx}",
                values=(
                    rec["status_badge"],
                    rec["mismatch_reason"],
                    rec["req_type"],
                    rec["order_no"],
                    rec.get("instruction_disp", ""),
                    rec.get("ship_to", ""),
                    rec.get("line_name", ""),
                    rec.get("product_detail", ""),
                    rec["supplier"],
                    rec["maker"],
                    rec["item_name"],
                    rec["model_spec"],
                    rec["qty"],
                    rec["due_date_str"],
                    rec["days_overdue_str"],
                    rec["actual_date_str"],
                    rec["actual_status"],
                    rec["invoice_date_str"],
                    rec["invoice_status"],
                    rec["person"],
                    rec["source_info"],
                ),
                tags=(rec["row_tag"],),
            )

        self._update_kpi_card_styles()

    def _update_heading_display_main(self, col_id: str):
        orig_title = self.col_titles_main.get(col_id, col_id)
        if col_id in self.column_filters_main:
            display_title = f"{orig_title} [▼]"
        else:
            display_title = f"{orig_title} ▾"
        self.tree_main.heading(col_id, text=display_title)

    def _update_all_headings_main(self):
        for col_id, _, _, _ in self.col_configs_main:
            self._update_heading_display_main(col_id)

    def _on_tree_main_click(self, event):
        region = self.tree_main.identify_region(event.x, event.y)
        if region == "heading":
            col = self.tree_main.identify_column(event.x)
            col_index = int(col.replace("#", "")) - 1
            disp_cols = self.tree_main.cget("displaycolumns")
            if not disp_cols or disp_cols == "#all" or disp_cols == ("#all",):
                actual_cols = [c[0] for c in self.col_configs_main]
            else:
                actual_cols = list(disp_cols)

            if 0 <= col_index < len(actual_cols):
                col_id = actual_cols[col_index]
                self._open_column_filter_dialog(
                    col_id=col_id,
                    title=self.col_titles_main.get(col_id, col_id),
                    records=self._get_visible_main_records(),
                    filter_dict=self.column_filters_main,
                    refresh_func=self._refresh_main_table,
                    update_heading_func=self._update_heading_display_main,
                    sort_func=lambda c, r: self._sort_tree(self.tree_main, c, r),
                    x=event.x_root,
                    y=event.y_root,
                )


    # =========================================================================
    # 表示列設定（タブ1: 照合結果一覧）
    # =========================================================================
    def _apply_column_visibility_main(self):
        """self.visible_columns_main に基づいて tree_main の表示列（displaycolumns）を設定する"""
        all_col_ids = [c[0] for c in self.col_configs_main]
        valid_visible = [cid for cid in all_col_ids if cid in self.visible_columns_main]
        if not valid_visible:
            valid_visible = list(self.DEFAULT_VISIBLE_COLUMNS_MAIN)
        self.visible_columns_main = valid_visible
        try:
            self.tree_main.configure(displaycolumns=tuple(valid_visible))
        except Exception:
            pass

    def _open_column_visibility_dialog(self):
        from modules.dialogs import open_column_visibility_dialog

        def on_apply(new_cols, save):
            self._apply_column_visibility_main(new_cols)
            if save:
                self._save_config()

        open_column_visibility_dialog(
            parent=self.root,
            col_configs_main=self.col_configs_main,
            default_visible_cols=self.DEFAULT_VISIBLE_COLUMNS_MAIN,
            current_visible_cols=self.visible_columns_main,
            icon_path=self._resolve_icon_path(),
            on_apply=on_apply,
        )
    def _clear_all_main_column_filters(self):
        self.column_filters_main.clear()
        self._update_all_headings_main()
        self._refresh_main_table()

    # =========================================================================
    # テーブル描画＆フィルター処理（タブ2: 注文書未検出一覧）
    # =========================================================================
    def _refresh_missing_table(self):
        for item in self.tree_missing.get_children():
            self.tree_missing.delete(item)

        kw = self.missing_search_keyword.get().strip().lower()
        excluded_nos = self._get_excluded_order_nos()
        base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
        exclude_months = self._get_exclude_months()
        inst_mode = self._get_instruction_filter_mode()

        for idx, rec in enumerate(self.missing_order_records):
            if self._is_record_excluded(rec, excluded_nos, base_date, exclude_months, inst_mode):
                continue

            skip_by_col = False
            for col_id, allowed_values in self.column_filters_missing.items():
                val = str(rec.get(col_id, "") or "").strip()
                if not val:
                    val = "(空白)"
                if val not in allowed_values:
                    skip_by_col = True
                    break
            if skip_by_col:
                continue

            if kw:
                searchable = " ".join(
                    [
                        rec["status_badge"],
                        rec["mgmt_sheet"],
                        rec["mgmt_row"],
                        rec["order_no"],
                        rec.get("instruction_disp", ""),
                        rec.get("ship_to", ""),
                        rec.get("line_name", ""),
                        rec.get("product_detail", ""),
                        rec["supplier"],
                        rec["maker"],
                        rec["item_name"],
                        rec["model_spec"],
                        rec["person"],
                        rec["mismatch_reason"],
                    ]
                ).lower()
                if kw not in searchable:
                    continue

            self.tree_missing.insert(
                "",
                tk.END,
                iid=f"mis_{idx}",
                values=(
                    rec["status_badge"],
                    rec["mgmt_sheet"],
                    rec["mgmt_row"],
                    rec["order_no"],
                    rec.get("instruction_disp", ""),
                    rec.get("ship_to", ""),
                    rec.get("line_name", ""),
                    rec.get("product_detail", ""),
                    rec["supplier"],
                    rec["maker"],
                    rec["item_name"],
                    rec["model_spec"],
                    rec["qty"],
                    rec["due_date"],
                    rec["actual_date"],
                    rec["invoice_date"],
                    rec["person"],
                    rec["mismatch_reason"],
                ),
                tags=(rec["row_tag"],),
            )

    def _update_heading_display_missing(self, col_id: str):
        orig_title = self.col_titles_missing.get(col_id, col_id)
        if col_id in self.column_filters_missing:
            display_title = f"{orig_title} [▼]"
        else:
            display_title = f"{orig_title} ▾"
        self.tree_missing.heading(col_id, text=display_title)

    def _update_all_headings_missing(self):
        for col_id, _, _, _ in self.col_configs_missing:
            self._update_heading_display_missing(col_id)

    def _on_tree_missing_click(self, event):
        region = self.tree_missing.identify_region(event.x, event.y)
        if region == "heading":
            col = self.tree_missing.identify_column(event.x)
            col_index = int(col.replace("#", "")) - 1
            if 0 <= col_index < len(self.col_configs_missing):
                col_id = self.col_configs_missing[col_index][0]
                self._open_column_filter_dialog(
                    col_id=col_id,
                    title=self.col_titles_missing.get(col_id, col_id),
                    records=self._get_visible_missing_records(),
                    filter_dict=self.column_filters_missing,
                    refresh_func=self._refresh_missing_table,
                    update_heading_func=self._update_heading_display_missing,
                    sort_func=lambda c, r: self._sort_tree(self.tree_missing, c, r),
                    x=event.x_root,
                    y=event.y_root,
                )

    def _clear_all_missing_column_filters(self):
        self.column_filters_missing.clear()
        self._update_all_headings_missing()
        self._refresh_missing_table()

    # =========================================================================
    # 共通Excel風列フィルターダイアログ
    # =========================================================================
    def _open_column_filter_dialog(self, event, col_id: str, is_missing: bool = False):
        from modules.dialogs import open_column_filter_dialog
        records = self._get_visible_missing_records() if is_missing else self._get_visible_main_records()
        col_title = self.col_titles_missing.get(col_id, col_id) if is_missing else self.col_titles_main.get(col_id, col_id)
        filter_dict = self.column_filters_missing if is_missing else self.column_filters_main
        refresh_cb = self._refresh_missing_table if is_missing else self._refresh_main_table
        heading_cb = self._update_heading_display_missing if is_missing else self._update_heading_display_main
        sort_cb = lambda cid, rev: self._sort_tree(self.tree_missing if is_missing else self.tree_main, cid, rev)

        open_column_filter_dialog(
            parent=self.root,
            col_id=col_id,
            title=col_title,
            records=records,
            filter_dict=filter_dict,
            on_refresh=refresh_cb,
            on_update_heading=heading_cb,
            on_sort=sort_cb,
            x=event.x_root,
            y=event.y_root,
        )
    def _sort_tree(self, tree: ttk.Treeview, col: str, reverse: bool):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, "", index)

    # =========================================================================
    # Excel直接展開機能（一時ファイル非作成・新規ブック展開）
    # =========================================================================
    @staticmethod
    def _open_in_excel_direct(self):
        """Excelがインストールされている場合、一時ファイルを作成せず未保存の新規ブックに直接結果を展開する"""
        if not self.analyzed_records and not self.missing_order_records:
            messagebox.showinfo("データなし", "Excelに展開する照合結果がありません。先に「照合・遅延チェック実行」を実行してください。")
            return
        if self.is_excel_running:
            messagebox.showwarning("処理中", "現在Excelへの展開処理を実行中です。完了または中止をお待ちください。")
            return

        from modules.excel_exporter import HAS_WIN32COM, export_to_excel_direct
        if not HAS_WIN32COM:
            messagebox.showerror(
                "pywin32未検出",
                "Excelへの直接展開には pywin32 パッケージが必要です。\n"
                "ExcelがインストールされているWindows環境で実行してください。",
            )
            return

        self.is_excel_running = True
        self.excel_cancel_event.clear()
        self.excel_open_btn.configure(state=tk.DISABLED, text="⏳ Excel展開中...")
        self.excel_cancel_btn.configure(state=tk.NORMAL, text="⏹ 中止")
        self.cancel_btn.configure(state=tk.NORMAL, text="⏹ 中止 (Esc)")
        self.run_btn.configure(state=tk.DISABLED)
        self.progress_var.set(0.0)
        self.progress_text_var.set("Excelアプリケーションを起動中...")

        excluded_nos = self._get_excluded_order_nos()
        base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
        exclude_months = self._get_exclude_months()
        inst_mode = self.instruction_filter_mode.get()

        def _worker():
            ok, msg = export_to_excel_direct(
                analyzed_records=self.analyzed_records,
                missing_order_records=self.missing_order_records,
                excluded_nos=excluded_nos,
                base_date=base_date,
                exclude_months=exclude_months,
                inst_mode=inst_mode,
                cancel_event=self.excel_cancel_event,
                progress_queue=self.ui_queue,
            )
            self.ui_queue.put(("EXCEL_DONE" if ok else "EXCEL_ERROR", msg))

        th = threading.Thread(target=_worker, daemon=True)
        th.start()
        self.root.after(50, self._poll_excel_queue)
    def _poll_excel_queue(self):
        try:
            while True:
                msg_type, payload = self.ui_queue.get_nowait()
                if msg_type == "PROGRESS":
                    pct, text = payload
                    self.progress_var.set(pct)
                    self.progress_text_var.set(text)
                    self.status_var.set(f"実行中: {text}")
                elif msg_type == "EXCEL_OPENED_SUCCESS":
                    self.is_excel_running = False
                    self.excel_open_btn.configure(state=tk.NORMAL, text="📊 Excelで直接開く")
                    self.excel_cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止")
                    self.cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止 (Esc)")
                    self.run_btn.configure(state=tk.NORMAL, text="▶ 照合・遅延チェック実行 (F5)")
                    self.progress_var.set(100.0)
                    self.progress_text_var.set("Excel展開完了！")
                    self.status_var.set("新規Excelブックに照合結果を展開しました（未保存の新規ブックとして開いています）。")
                    messagebox.showinfo(
                        "Excel起動完了",
                        "一時ファイルを作成せず、Excelの新規ブックとして直接展開しました。\n\n"
                        "・シート1: 納期超過・照合結果\n"
                        "・シート2: 注文書未検出(管理表のみ)\n"
                        "※アプリと同じ色分け・書式・オートフィルターが設定されています。\n"
                        "保存が必要な場合は、Excel側で「名前を付けて保存」を行ってください。",
                    )
                    return
                elif msg_type == "EXCEL_OPENED_CANCELLED":
                    self.is_excel_running = False
                    self.excel_open_btn.configure(state=tk.NORMAL, text="📊 Excelで直接開く")
                    self.excel_cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止")
                    self.cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止 (Esc)")
                    self.run_btn.configure(state=tk.NORMAL, text="▶ 照合・遅延チェック実行 (F5)")
                    self.progress_var.set(0.0)
                    self.progress_text_var.set("Excel展開が中止されました")
                    self.status_var.set("ユーザーによってExcel直接展開処理が中止されました。")
                    messagebox.showinfo("処理中止", "Excel直接展開処理を中止しました。")
                    return
                elif msg_type == "EXCEL_OPENED_ERROR":
                    self.is_excel_running = False
                    self.excel_open_btn.configure(state=tk.NORMAL, text="📊 Excelで直接開く")
                    self.excel_cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止")
                    self.cancel_btn.configure(state=tk.DISABLED, text="⏹ 中止 (Esc)")
                    self.run_btn.configure(state=tk.NORMAL, text="▶ 照合・遅延チェック実行 (F5)")
                    self.progress_var.set(0.0)
                    self.progress_text_var.set("Excel展開エラー")
                    self.status_var.set(f"Excel展開エラー: {payload}")
                    messagebox.showerror("Excel展開エラー", f"Excelの起動またはデータ転送中にエラーが発生しました:\n{payload}")
                    return
                else:
                    self.ui_queue.put((msg_type, payload))
                    break
        except queue.Empty:
            pass

        if self.is_excel_running:
            self.root.after(50, self._poll_excel_queue)


def main():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("StockMatch.Miya.v1.5.0")
    except Exception:
        pass

    root = tk.Tk()
    _app = StockMatchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
