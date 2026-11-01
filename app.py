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


class StockMatchApp:
    """
    StockMatch_Miya-v1.1.0
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

    CONFIG_FILE_PATH = os.path.join(APP_DIR, "config.json")
    ICON_FILE_PATH = os.path.join(APP_DIR, "icon.ico")

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("StockMatch_Miya-v1.1.0 - 注文品納期・現品/請求照合システム")

        # メインウィンドウおよび全Toplevelウィンドウに icon.ico を適用
        self._apply_window_icon(self.root, is_main=True)

        # 起動時にウィンドウを最大化表示
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")

        self.root.minsize(1080, 720)

        # 状態保持用変数
        self.mgmt_file_path = tk.StringVar()
        self.order_dir_path = tk.StringVar()
        self.base_date_str = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.blank_due_grace_days = tk.StringVar(value="7")
        self.filter_mode = tk.StringVar(value="BOTH_DELAY")
        self.search_keyword = tk.StringVar()
        self.missing_search_keyword = tk.StringVar()

        # 進捗表示・非同期スレッド管理用
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_text_var = tk.StringVar(value="待機中")
        self.is_running = False
        self.ui_queue = queue.Queue()

        # 解析結果データ保持（2系統に分離）
        self.analyzed_records = []       # タブ1: 注文書由来の照合・遅延レコード
        self.missing_order_records = []   # タブ2: 注文書未検出（管理表のみ）レコード

        # 各タブの列フィルター保持用辞書
        self.column_filters_main = {}
        self.column_filters_missing = {}

        self._setup_styles()
        self._build_menu()
        self._build_ui()
        self._load_config_and_autoload()

    def _get_grace_days(self) -> int:
        """納期未記載時に注文日を基準とする猶予日数（デフォルト7日）を整数で取得"""
        val = self.blank_due_grace_days.get().strip()
        try:
            days = int(val)
            return max(0, days)
        except ValueError:
            return 7

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
            font=("Meiryo UI", 9, "bold"),
            background="#2563eb",
            foreground="#ffffff",
            padding=(12, 6),
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1d4ed8"), ("disabled", "#94a3b8")],
            foreground=[("disabled", "#e2e8f0")],
        )

        style.configure(
            "Excel.TButton",
            font=("Meiryo UI", 9, "bold"),
            background="#166534",
            foreground="#ffffff",
            padding=(10, 6),
        )
        style.map(
            "Excel.TButton",
            background=[("active", "#15803d"), ("disabled", "#94a3b8")],
            foreground=[("disabled", "#e2e8f0")],
        )

        style.configure("Action.TButton", font=("Meiryo UI", 9), padding=(8, 4))
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
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit, accelerator="Alt+F4")
        menubar.add_cascade(label="ファイル(F)", menu=file_menu)

        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="現在の設定を保存 (パス・超過日数)", command=self._save_config_manual)
        config_menu.add_command(label="アプリ配置フォルダを開く", command=self._open_config_folder)
        menubar.add_cascade(label="設定(S)", menu=config_menu)

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
        self.root.bind_all("<Control-o>", lambda _e: self._select_mgmt_file())
        self.root.bind_all("<Control-d>", lambda _e: self._select_order_dir())
        self.root.bind_all("<Control-s>", lambda _e: self._open_in_excel_direct())

    def _build_ui(self):
        """メイン画面レイアウトの構築"""
        self.root.configure(bg="#f8fafc")

        # 1. トップヘッダーバー
        header_frame = tk.Frame(self.root, bg="#1e293b", height=56, padx=16, pady=10)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="StockMatch_Miya-v1.1.0", style="Header.TLabel")
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
        ttk.Button(row2, text="設定を保存", style="Action.TButton", command=self._save_config_manual).pack(
            side=tk.LEFT
        )

        # ③ 基準日・超過設定・実行ボタン行
        row3 = ttk.Frame(file_card, style="Card.TFrame")
        row3.pack(fill=tk.X, pady=(2, 6))

        ttk.Label(
            row3,
            text="③ 判定基準日 (今日):",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        date_entry = ttk.Entry(row3, textvariable=self.base_date_str, width=13, font=("Meiryo UI", 9))
        date_entry.pack(side=tk.LEFT, padx=(4, 6))

        ttk.Button(
            row3,
            text="今日に戻す",
            style="Action.TButton",
            command=lambda: self.base_date_str.set(datetime.date.today().strftime("%Y-%m-%d")),
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

        self.excel_open_btn = ttk.Button(
            row3,
            text="📊 Excelで直接開く",
            style="Excel.TButton",
            command=self._open_in_excel_direct,
        )
        self.excel_open_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self.run_btn = ttk.Button(
            row3, text="照合・遅延チェック実行", style="Primary.TButton", command=self.run_analysis
        )
        self.run_btn.pack(side=tk.RIGHT)

        # ④ 進捗バー
        row4 = ttk.Frame(file_card, style="Card.TFrame")
        row4.pack(fill=tk.X, pady=(4, 0))

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
        self.progress_bar.pack(side=tk.LEFT, padx=(4, 12))

        self.progress_lbl = ttk.Label(
            row4,
            textvariable=self.progress_text_var,
            style="Card.TLabel",
            foreground="#1d4ed8",
            font=("Meiryo UI", 9, "bold"),
        )
        self.progress_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 3. サマリーKPIカードエリア
        kpi_frame = ttk.Frame(self.root, padding=(16, 0, 16, 8))
        kpi_frame.pack(fill=tk.X)

        self.kpi_vars = {
            "order_total": tk.StringVar(value="0 件"),
            "both_delay": tk.StringVar(value="0 件"),
            "actual_delay": tk.StringVar(value="0 件"),
            "invoice_delay": tk.StringVar(value="0 件"),
            "unmatched": tk.StringVar(value="0 件"),
            "ok": tk.StringVar(value="0 件"),
            "missing_order": tk.StringVar(value="0 件"),
        }

        self._create_kpi_card(kpi_frame, "注文書明細 総数", self.kpi_vars["order_total"], "#334155", "#f1f5f9").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6)
        )
        self._create_kpi_card(
            kpi_frame, "【不一致】管理表未記載", self.kpi_vars["unmatched"], "#4c1d95", "#ede9fe"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(
            kpi_frame, "【警告】現品・請求 超過", self.kpi_vars["both_delay"], "#991b1b", "#fee2e2"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(
            kpi_frame, "【現品超過】現品確認 遅れ", self.kpi_vars["actual_delay"], "#9a3412", "#ffedd5"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(
            kpi_frame, "【請求超過】伝票(請求) 遅れ", self.kpi_vars["invoice_delay"], "#854d0e", "#fef9c3"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(kpi_frame, "正常・納期前・確認済", self.kpi_vars["ok"], "#166534", "#dcfce7").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6)
        )
        self._create_kpi_card(
            kpi_frame, "【別タブ】注文書未検出", self.kpi_vars["missing_order"], "#92400e", "#fef3c7"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

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
        """タブが切り替わったときにタイトル文言を更新し、選択中を強調表示する"""
        try:
            curr_tab_idx = self.notebook.index("current")
            total_main = len(self.analyzed_records)
            total_missing = len(self.missing_order_records)

            if curr_tab_idx == 0:
                self.notebook.tab(0, text=f"  【選択中】▶ ① 納期超過・照合結果 一覧 ({total_main}件)  ")
                self.notebook.tab(1, text=f"  ▷ ② 【未発注・注文書なし】管理表のみ記載 一覧 ({total_missing}件)  ")
            else:
                self.notebook.tab(0, text=f"  ▷ ① 納期超過・照合結果 一覧 ({total_main}件)  ")
                self.notebook.tab(1, text=f"  【選択中】▶ ② 【未発注・注文書なし】管理表のみ記載 一覧 ({total_missing}件)  ")
        except Exception:
            pass

    def _build_main_table_tab(self, parent):
        """タブ1: 納期超過・照合結果テーブルの構築"""
        filter_frame = ttk.Frame(parent, padding=(4, 6, 4, 6))
        filter_frame.pack(fill=tk.X)

        ttk.Label(filter_frame, text="表示フィルタ:", font=("Meiryo UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))

        modes = [
            ("すべての異常(超過・不一致)", "ALL_DELAY"),
            ("管理表との不一致", "UNMATCHED_ONLY"),
            ("現品確認日の超過", "ACTUAL_DELAY"),
            ("請求日(伝票)の超過", "INVOICE_DELAY"),
            ("現品・請求の両方超過", "BOTH_DELAY"),
            ("全件表示（正常含む）", "ALL"),
        ]
        for text, mode_val in modes:
            rb = ttk.Radiobutton(
                filter_frame,
                text=text,
                variable=self.filter_mode,
                value=mode_val,
                command=self._refresh_main_table,
            )
            rb.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(filter_frame, text="キーワード:").pack(side=tk.LEFT, padx=(10, 4))
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_keyword, width=18)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind("<KeyRelease>", lambda _e: self._refresh_main_table())

        clear_col_btn = ttk.Button(
            filter_frame,
            text="列フィルター解除 (Excel風)",
            style="Action.TButton",
            command=self._clear_all_main_column_filters,
        )
        clear_col_btn.pack(side=tk.RIGHT, padx=(6, 0))

        inspect_btn = ttk.Button(
            filter_frame,
            text="🔍 選択行の不一致・照合詳細を表示",
            style="Action.TButton",
            command=lambda: self._open_selected_row_detail(is_missing_tab=False),
        )
        inspect_btn.pack(side=tk.RIGHT, padx=(6, 0))

        table_container = ttk.Frame(parent)
        table_container.pack(fill=tk.BOTH, expand=True)

        columns = (
            "status_badge",
            "mismatch_reason",
            "req_type",
            "order_no",
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
            ("status_badge", "総合判定", 165, "center"),
            ("mismatch_reason", "不一致理由・照合診断", 340, "w"),
            ("req_type", "依頼区分", 80, "center"),
            ("order_no", "管理No(注文番号)", 115, "center"),
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

        self.col_titles_main = {col_id: title for col_id, title, _, _ in self.col_configs_main}

        for col_id, _, width, anchor in self.col_configs_main:
            self._update_heading_display_main(col_id)
            self.tree_main.column(col_id, width=width, anchor=anchor)

        self.tree_main.bind("<Button-1>", self._on_tree_main_click)
        self.tree_main.bind("<Double-1>", lambda e: self._on_tree_double_click(e, is_missing_tab=False))

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

        inspect_missing_btn = ttk.Button(
            filter_frame,
            text="🔍 選択行の詳細を確認",
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
            ("status_badge", "状態判定", 150, "center"),
            ("mgmt_sheet", "管理表シート", 100, "center"),
            ("mgmt_row", "行番号", 65, "center"),
            ("order_no", "管理No(8桁)", 105, "center"),
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

        self.col_titles_missing = {col_id: title for col_id, title, _, _ in self.col_configs_missing}

        for col_id, _, width, anchor in self.col_configs_missing:
            self._update_heading_display_missing(col_id)
            self.tree_missing.column(col_id, width=width, anchor=anchor)

        self.tree_missing.bind("<Button-1>", self._on_tree_missing_click)
        self.tree_missing.bind("<Double-1>", lambda e: self._on_tree_double_click(e, is_missing_tab=True))

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
        self, parent, title: str, value_var: tk.StringVar, fg_color: str, bg_color: str
    ) -> tk.Frame:
        """KPIサマリー表示用カードの生成"""
        card = tk.Frame(parent, bg=bg_color, highlightbackground="#cbd5e1", highlightthickness=1, padx=8, pady=8)
        lbl_title = tk.Label(card, text=title, bg=bg_color, fg=fg_color, font=("Meiryo UI", 8, "bold"), anchor="w")
        lbl_title.pack(fill=tk.X)
        lbl_val = tk.Label(
            card, textvariable=value_var, bg=bg_color, fg=fg_color, font=("Meiryo UI", 15, "bold"), anchor="e"
        )
        lbl_val.pack(fill=tk.X, pady=(4, 0))
        return card

    # =========================================================================
    # 不一致・照合詳細の比較インスペクタ
    # =========================================================================
    def _on_tree_double_click(self, event, is_missing_tab: bool):
        tree = self.tree_missing if is_missing_tab else self.tree_main
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            return
        self._open_selected_row_detail(is_missing_tab=is_missing_tab)

    def _open_selected_row_detail(self, is_missing_tab: bool):
        """選択されている行の不一致理由・項目比較ダイアログを開く"""
        tree = self.tree_missing if is_missing_tab else self.tree_main
        records = self.missing_order_records if is_missing_tab else self.analyzed_records

        sel = tree.selection()
        if not sel:
            messagebox.showinfo("行未選択", "一覧テーブルから確認したい行をクリックして選択してください。")
            return

        item_id = sel[0]
        prefix = "mis_" if is_missing_tab else "rec_"
        rec_idx_str = item_id.replace(prefix, "")
        if not rec_idx_str.isdigit():
            return
        rec_idx = int(rec_idx_str)
        if not (0 <= rec_idx < len(records)):
            return

        rec = records[rec_idx]
        diag = rec.get("diagnostic_data", {})

        win = tk.Toplevel(self.root)
        win.title(f"詳細情報 — 管理No: {rec['order_no']} / 品名: {rec['item_name']}")
        self._apply_window_icon(win)
        win.geometry("920x660")
        win.minsize(760, 520)
        win.transient(self.root)

        banner_bg = "#92400e" if is_missing_tab else ("#4c1d95" if rec.get("is_unmatched") else "#1e293b")

        top_banner = tk.Frame(win, bg=banner_bg, padx=14, pady=10)
        top_banner.pack(fill=tk.X)

        tk.Label(
            top_banner,
            text=f"判定: {rec['status_badge']}   |   {rec['mismatch_reason']}",
            bg=banner_bg,
            fg="#ffffff",
            font=("Meiryo UI", 10, "bold"),
            anchor="w",
        ).pack(fill=tk.X)

        body = ttk.Frame(win, padding=14)
        body.pack(fill=tk.BOTH, expand=True)

        lbl_cmp = ttk.Label(
            body,
            text="■ 注文書 と 注文品管理表 の項目別詳細（単価・金額は照合対象外）",
            font=("Meiryo UI", 9, "bold"),
        )
        lbl_cmp.pack(anchor="w", pady=(0, 6))

        cmp_cols = ("field", "order_val", "mgmt_val", "judge")
        cmp_tree = ttk.Treeview(body, columns=cmp_cols, show="headings", height=10)
        cmp_tree.heading("field", text="比較項目")
        cmp_tree.heading("order_val", text="注文書の記載内容")
        cmp_tree.heading("mgmt_val", text="注文品管理表の記載内容")
        cmp_tree.heading("judge", text="照合判定")

        cmp_tree.column("field", width=160, anchor="w")
        cmp_tree.column("order_val", width=285, anchor="w")
        cmp_tree.column("mgmt_val", width=285, anchor="w")
        cmp_tree.column("judge", width=140, anchor="center")

        cmp_tree.tag_configure("MATCH", background="#dcfce7", foreground="#166534")
        cmp_tree.tag_configure("PARTIAL", background="#fef9c3", foreground="#854d0e")
        cmp_tree.tag_configure("DIFF", background="#fee2e2", foreground="#991b1b")
        cmp_tree.tag_configure("INFO", background="#f1f5f9", foreground="#334155")
        cmp_tree.tag_configure("IGNORED", background="#f8fafc", foreground="#64748b")

        cmp_rows = diag.get("comparison_rows", [])
        for f_name, o_v, m_v, j_txt, tag in cmp_rows:
            cmp_tree.insert("", tk.END, values=(f_name, o_v, m_v, j_txt), tags=(tag,))

        cmp_tree.pack(fill=tk.X, pady=(0, 12))

        lbl_rep = ttk.Label(
            body,
            text="■ 詳細診断レポート",
            font=("Meiryo UI", 9, "bold"),
        )
        lbl_rep.pack(anchor="w", pady=(0, 6))

        txt_frame = ttk.Frame(body)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(txt_frame, wrap="word", font=("Meiryo UI", 9), padx=10, pady=8, relief="solid", borderwidth=1)
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        txt.insert("1.0", diag.get("detail_text", rec.get("mismatch_reason", "")))
        txt.configure(state="disabled")

        btn_bar = ttk.Frame(win, padding=(14, 8))
        btn_bar.pack(fill=tk.X)
        ttk.Button(btn_bar, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)

    # =========================================================================
    # ヘルプ・ガイドダイアログ
    # =========================================================================
    def _create_scrollable_text(self, parent) -> tk.Text:
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(
            container,
            wrap="word",
            font=("Meiryo UI", 9),
            padx=14,
            pady=12,
            relief="flat",
            bg="#ffffff",
            fg="#0f172a",
        )
        vsb = ttk.Scrollbar(container, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._configure_rich_text_tags(txt)
        return txt

    def _configure_rich_text_tags(self, txt: tk.Text):
        txt.tag_configure(
            "h1",
            font=("Meiryo UI", 12, "bold"),
            foreground="#1e293b",
            background="#e2e8f0",
            spacing1=8,
            spacing3=8,
            lmargin1=8,
            lmargin2=8,
        )
        txt.tag_configure(
            "h2",
            font=("Meiryo UI", 10, "bold"),
            foreground="#1d4ed8",
            spacing1=12,
            spacing3=4,
            lmargin1=6,
            lmargin2=6,
        )
        txt.tag_configure(
            "body",
            font=("Meiryo UI", 9),
            foreground="#1e293b",
            spacing1=2,
            spacing2=3,
            spacing3=2,
            lmargin1=18,
            lmargin2=32,
        )
        txt.tag_configure(
            "point",
            font=("Meiryo UI", 9, "bold"),
            foreground="#0f172a",
            background="#fef9c3",
            spacing1=3,
            spacing2=3,
            spacing3=3,
            lmargin1=18,
            lmargin2=32,
        )
        txt.tag_configure(
            "warn",
            font=("Meiryo UI", 9, "bold"),
            foreground="#991b1b",
            background="#fee2e2",
            spacing1=3,
            spacing2=3,
            spacing3=3,
            lmargin1=18,
            lmargin2=32,
        )
        txt.tag_configure(
            "badge_ok",
            font=("Meiryo UI", 9, "bold"),
            foreground="#166534",
            background="#dcfce7",
            lmargin1=18,
            lmargin2=32,
        )
        txt.tag_configure(
            "sub_note",
            font=("Meiryo UI", 9),
            foreground="#475569",
            spacing1=1,
            spacing2=2,
            spacing3=4,
            lmargin1=30,
            lmargin2=42,
        )

    def _show_help_guide(self, tab_index: int = 0):
        win = tk.Toplevel(self.root)
        win.title("StockMatch_Miya-v1.1.0 - ヘルプ＆操作ガイド")
        self._apply_window_icon(win)
        win.geometry("880x680")
        win.minsize(720, 520)
        win.transient(self.root)

        banner = tk.Frame(win, bg="#1e293b", padx=18, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text="📖 StockMatch_Miya-v1.1.0  取扱説明書・操作ガイド",
            bg="#1e293b",
            fg="#ffffff",
            font=("Meiryo UI", 12, "bold"),
            anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            banner,
            text="注文書と注文品管理表の双方向自動照合・タブ分離一覧表示・納期超過チェックツールの使い方",
            bg="#1e293b",
            fg="#cbd5e1",
            font=("Meiryo UI", 9),
            anchor="w",
        ).pack(fill=tk.X, pady=(2, 0))

        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=14, pady=(12, 6))

        tab1 = ttk.Frame(notebook, padding=8)
        notebook.add(tab1, text="  ① 基本的な使い方・設定  ")
        txt1 = self._create_scrollable_text(tab1)

        txt1.insert(tk.END, "  STEP 1 : ファイルとフォルダの指定  \n", "h1")
        txt1.insert(tk.END, "① 注文品管理表ファイルの選択\n", "h2")
        txt1.insert(tk.END, "・『ファイルを参照...』ボタンから管理表Excel（例: 注文品管理表.xlsx）を選択します。\n", "body")
        txt1.insert(tk.END, "・管理表に登録されているすべての「管理No（8桁）」を自動抽出し、一致する注文書を照合します。\n", "point")
        txt1.insert(tk.END, "・管理表のAB列（28列目）にある「購入数」列を正確に参照します。\n", "sub_note")

        txt1.insert(tk.END, "② 注文書フォルダパスの選択（サブフォルダ自動対応）\n", "h2")
        txt1.insert(tk.END, "・『フォルダを参照...』ボタンから注文書Excelが保管されている親フォルダを選択します。\n", "body")
        txt1.insert(tk.END, "★【サブフォルダ一括探索】指定したフォルダ直下だけでなく、配下の全サブフォルダ内も漏れなく自動スキャンします。\n", "point")
        txt1.insert(tk.END, "★【見積依頼書の自動除外】Excelシートの2行目を自動検査し、「見積」の記載があるシートは除外して「注文」シートのみを照合します。\n", "point")

        txt1.insert(tk.END, "\n  STEP 2 : 判定基準日・超過日数の設定と保存  \n", "h1")
        txt1.insert(tk.END, "■ 判定基準日（今日）と設定の保存\n", "h2")
        txt1.insert(tk.END, "・納期超過の判定基準日（初期値は今日）を YYYY-MM-DD 形式で指定します。\n", "body")
        txt1.insert(tk.END, "・『設定を保存』ボタンを押すと、管理表パス・注文書フォルダパスおよび納期未記載時の超過日数が config.json に保存されます。\n", "body")
        txt1.insert(tk.END, "■ 納期未記載時の注文日基準判定\n", "h2")
        txt1.insert(tk.END, "・注文書明細の納期が空欄の場合、注文書4行目に記載されている「注文日」から指定日数（初期値: 7日）を経過した時点で超過とみなします。\n", "point")

        txt1.insert(tk.END, "\n  STEP 3 : 2つの独立タブによる確認  \n", "h1")
        txt1.insert(tk.END, "■ 照合実行後の2系統タブ表示\n", "h2")
        txt1.insert(tk.END, "・『① 納期超過・照合結果 一覧』タブ: 注文書由来の納期遅延・確認状況、管理表未記載案件を確認できます（※初期表示フィルタは最も確認優先度の高い『現品・請求の両方超過』がデフォルト選択されます）。\n", "point")
        txt1.insert(tk.END, "・『② 【未発注・注文書なし】管理表のみ記載 一覧』タブ: 管理表に記載があるのに注文書が未手配の明細を独立して一覧確認できます。\n", "point")
        txt1.insert(tk.END, "★【購入数0の自動除外】管理表AB列の購入数が『0』の項目（在庫充当品等）は未発注一覧から自動除外されます。\n", "point")
        txt1.insert(tk.END, "・一覧の行を【ダブルクリック】すると項目別比較ウィンドウが開きます。\n", "body")
        txt1.configure(state="disabled")

        tab2 = ttk.Frame(notebook, padding=8)
        notebook.add(tab2, text="  ② 照合ルール・タブ分離の仕組み  ")
        txt2 = self._create_scrollable_text(tab2)

        txt2.insert(tk.END, "  1. 注文書未検出（管理表のみ）の独立表示  \n", "h1")
        txt2.insert(tk.END, "■ タブ②『【未発注・注文書なし】管理表のみ記載 一覧』\n", "h2")
        txt2.insert(tk.END, "・管理表に登録されているが、注文書Excelがフォルダ内に存在しない未発注案件を完全に別枠で集約します。\n", "point")
        txt2.insert(tk.END, "★【AB列購入数0の自動除外】管理表のAB列（28列目）「購入数」が『0』の項目は未発注一覧から自動的に除外されます。\n", "point")
        txt2.insert(tk.END, "・メインの納期照合一覧に混ざらないため、注文品管理の作業効率が大幅に向上します。\n", "body")

        txt2.insert(tk.END, "\n  2. 高精度マッチングの基本ルール  \n", "h1")
        txt2.insert(tk.END, "■ 単価・金額は照合対象外（誤判定の完全防止）\n", "h2")
        txt2.insert(tk.END, "★【重要】単価未入力・金額0円や消費税表記の違いによる誤判定を防ぐため、単価・金額は照合計算から完全に除外しています。\n", "point")
        txt2.insert(tk.END, "・照合は『①管理No(8桁)』『②図面番号』『③品番』『④品名』『⑤仕様・規格』『⑥発注先・メーカー』『⑦購入数(AB列)』で行われます。\n", "body")

        txt2.insert(tk.END, "\n  3. 現品確認日（A列）と請求日（AI列）の超過判定  \n", "h1")
        txt2.insert(tk.END, "■ 現品確認日（物）と 請求日（伝票）\n", "h2")
        txt2.insert(tk.END, "・日付または確認済の記載があれば『確認済 (正常)』と判定します。\n", "badge_ok")
        txt2.insert(tk.END, "・空欄のまま納期が判定基準日を過ぎている場合は『未確認 (超過X日)』として警告表示します。\n", "warn")

        txt2.insert(tk.END, "\n  4. 注文書納期未記載時・期間表記（上旬/中旬/下旬等）の自動判定  \n", "h1")
        txt2.insert(tk.END, "■ 期間表記（上旬・中旬・下旬・月末など）の期日自動変換\n", "h2")
        txt2.insert(tk.END, "・『11/上』『11/上旬』『11月上旬』『11/初』『11月頭』等 → その月の10日（例: 11/10）までとして判定します。\n", "point")
        txt2.insert(tk.END, "・『11/中』『11/中旬』『11月中旬』『11/中頃』『11月半ば』等 → その月の20日（例: 11/20）までとして判定します。\n", "point")
        txt2.insert(tk.END, "・『11/下』『11/下旬』『11月下旬』『11/末』『11月末』『11月いっぱい』『11月内』等 → その月の末日（例: 11/30）までとして判定します。\n", "point")
        txt2.insert(tk.END, "■ 納期未記載時の注文日（4行目）＋指定日数（初期値: 7日）による期日算出\n", "h2")
        txt2.insert(tk.END, "・注文書明細の納期が空欄の場合、注文書シート4行目の注文日を自動抽出し、設定日数経過後に納期超過として扱います。\n", "point")
        txt2.insert(tk.END, "・設定した日数は『設定を保存』により次回起動時にも自動的に引き継がれます。\n", "body")
        txt2.configure(state="disabled")

        tab3 = ttk.Frame(notebook, padding=8)
        notebook.add(tab3, text="  ③ 便利機能・ショートカットキー  ")
        txt3 = self._create_scrollable_text(tab3)

        txt3.insert(tk.END, "  Excel風 列フィルター＆ソート機能  \n", "h1")
        txt3.insert(tk.END, "■ 列ヘッダーをクリックして絞り込み・並べ替え\n", "h2")
        txt3.insert(tk.END, "・両タブそれぞれで独立して列ごとのExcel風フィルターメニューが開きます。\n", "body")
        txt3.insert(tk.END, "・値の検索やチェックボックスで特定の業者・管理Noだけに絞り込めます。\n", "body")

        txt3.insert(tk.END, "\n  キーボードショートカット一覧  \n", "h1")
        txt3.insert(tk.END, "・ F1 キー        :  このヘルプ＆操作ガイドを開く\n", "body")
        txt3.insert(tk.END, "・ F5 キー        :  照合・遅延チェックを実行する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + O       :  ① 注文品管理表ファイルを選択する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + D       :  ② 注文書フォルダを選択する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + S       :  一時ファイルを作成せずExcel新規ブックで開く\n", "body")
        txt3.insert(tk.END, "・ Alt + F4       :  アプリケーションを終了する\n", "body")
        txt3.configure(state="disabled")

        try:
            notebook.select(tab_index)
        except Exception:
            pass

        btn_box = ttk.Frame(win, padding=(14, 8))
        btn_box.pack(fill=tk.X)
        ttk.Button(btn_box, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)

    def _show_about(self):
        messagebox.showinfo(
            "バージョン情報",
            "StockMatch_Miya-v1.1.0\n"
            "注文書 ⇔ 管理表 双方向照合・現品/請求超過システム\n\n"
            "・Excel直接起動: 一時ファイルなしで未保存新規ブック展開（A4横・1ページ収まり印刷設定済）\n"
            "・タブ完全分離: 「①納期超過・照合一覧」と「②注文書未検出(管理表のみ)」を別一覧で確認可能\n"
            "・未発注一覧フィルタ: 注文品管理表のAB列「購入数」が0の明細は自動除外\n"
            "・納期未記載対応: 注文書納期が空欄時、注文日(4行目)から指定日数(初期値7日)経過で超過判定\n"
            "・照合方式: 全体最適1対1マッチング（※単価・金額は照合対象外）\n"
            "・自動分別: 2行目の記載により「見積依頼書」を除外し「注文書」のみ照合\n"
            "・探索範囲: 指定フォルダおよび配下の全サブフォルダを一括検査\n"
            f"・設定保存先: {self.CONFIG_FILE_PATH}\n\n"
            "© StockMatch_Miya System All Rights Reserved.",
        )

    def _open_config_folder(self):
        try:
            os.startfile(self.APP_DIR)
        except Exception as e:
            messagebox.showinfo("フォルダパス", f"アプリ配置フォルダ:\n{self.APP_DIR}\n\n({e})")

    # =========================================================================
    # 設定ファイル処理
    # =========================================================================
    def _is_strictly_allowed_path(self, target_filepath: str) -> bool:
        try:
            norm_target = os.path.normcase(os.path.realpath(os.path.abspath(target_filepath)))
            norm_allowed = os.path.normcase(os.path.realpath(os.path.abspath(self.CONFIG_FILE_PATH)))
            return norm_target == norm_allowed
        except Exception:
            return False

    def _save_config(self, show_msg: bool = False):
        mgmt = self.mgmt_file_path.get().strip().strip('"')
        order = self.order_dir_path.get().strip().strip('"')
        grace_days = self._get_grace_days()

        data = {
            "mgmt_file_path": mgmt,
            "order_dir_path": order,
            "blank_due_grace_days": grace_days,
            "last_updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        target_file = self.CONFIG_FILE_PATH

        if not self._is_strictly_allowed_path(target_file):
            err_msg = f"保存エラー: 許可されていない保存先への書き込み要求です。\n対象: {target_file}"
            if show_msg:
                messagebox.showerror("保存エラー", err_msg)
            return

        try:
            with open(target_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if show_msg:
                messagebox.showinfo("設定保存完了", f"パスおよび納期超過設定を保存しました:\n{target_file}\n(納期未記載時の超過基準: 注文日から{grace_days}日経過)")
        except Exception as e:
            if show_msg:
                messagebox.showerror("保存エラー", f"設定ファイルへの保存に失敗しました:\n{e}")

    def _save_config_manual(self):
        self._save_config(show_msg=True)

    def _load_config_and_autoload(self):
        cfg_path = self.CONFIG_FILE_PATH
        if os.path.exists(cfg_path) and self._is_strictly_allowed_path(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                mgmt = data.get("mgmt_file_path", "")
                order = data.get("order_dir_path", "")
                grace_days = data.get("blank_due_grace_days", 7)
                if mgmt and os.path.isfile(mgmt):
                    self.mgmt_file_path.set(mgmt)
                if order and os.path.isdir(order):
                    self.order_dir_path.set(order)
                if grace_days is not None:
                    self.blank_due_grace_days.set(str(grace_days))
            except Exception:
                pass

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
        self.ui_queue.put(("PROGRESS", (percent, message)))

    def _collect_valid_order_files(
        self, folder_path: str, mgmt_filepath: str, valid_mgmt_order_nos: set[str], progress_cb=None
    ) -> tuple[list[dict], int, int]:
        order_items = []
        valid_file_count = 0
        ignored_file_count = 0

        mgmt_abspath = os.path.normcase(os.path.realpath(os.path.abspath(mgmt_filepath))) if mgmt_filepath else ""

        file_list = []
        for root_dir, _dirs, files in os.walk(folder_path):
            for file_name in files:
                file_list.append((root_dir, file_name))

        file_list.sort(key=lambda item: os.path.join(item[0], item[1]))
        total_entries = len(file_list)

        for idx, (root_dir, entry_name) in enumerate(file_list, start=1):
            full_path = os.path.join(root_dir, entry_name)

            if progress_cb:
                pct = 35.0 + (40.0 * idx / max(1, total_entries))
                rel_display = os.path.relpath(full_path, folder_path)
                progress_cb(
                    pct,
                    f"[2/3] 注文書フォルダを確認中 ({idx}/{total_entries}): {rel_display}",
                )

            if entry_name.startswith("~$") or entry_name.startswith("."):
                ignored_file_count += 1
                continue

            if not entry_name.lower().endswith((".xlsx", ".xlsm")):
                ignored_file_count += 1
                continue

            full_abspath = os.path.normcase(os.path.realpath(os.path.abspath(full_path)))
            if (
                full_abspath == mgmt_abspath
                or "管理表" in entry_name
                or "納期超過チェック結果" in entry_name
            ):
                ignored_file_count += 1
                continue

            norm_entry = unicodedata.normalize("NFKC", entry_name)
            m_pre = re.match(r"^\s*(\d{8})", norm_entry)
            if m_pre:
                file_no_8 = m_pre.group(1)
                if file_no_8 not in valid_mgmt_order_nos:
                    ignored_file_count += 1
                    continue

            try:
                extracted = self._parse_order_book(full_path, valid_mgmt_order_nos)
                if extracted:
                    order_items.extend(extracted)
                    valid_file_count += 1
                else:
                    ignored_file_count += 1
            except Exception:
                ignored_file_count += 1

        return order_items, valid_file_count, ignored_file_count

    # =========================================================================
    # 日付・文字列正規化ユーティリティ
    # =========================================================================
    @staticmethod
    def _parse_flexible_date(
        val, default_year: int | None = None, allow_period: bool = True
    ) -> datetime.date | None:
        if val is None or pd.isna(val):
            return None

        if isinstance(val, datetime.datetime):
            return val.date()
        if isinstance(val, datetime.date):
            return val

        if isinstance(val, (int, float)):
            if 35000 <= val <= 60000:
                try:
                    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val))
                except Exception:
                    pass

        s = str(val).strip()
        if not s:
            return None

        s = unicodedata.normalize("NFKC", s).strip()
        if not s or s.lower() in (
            "-",
            "ー",
            "―",
            "‐",
            "未",
            "未定",
            "未確認",
            "未着",
            "別途",
            "相談",
            "要相談",
            "調整中",
            "none",
            "nan",
            "nat",
        ):
            return None

        def _make_valid_date(y: int, m: int, d: int) -> datetime.date | None:
            if not (1990 <= y <= 2099 and 1 <= m <= 12 and 1 <= d <= 31):
                return None
            max_d = calendar.monthrange(y, m)[1]
            return datetime.date(y, m, min(d, max_d))

        def _period_to_day(y: int, m: int, period_token: str) -> int | None:
            if not (1990 <= y <= 2099 and 1 <= m <= 12):
                return None
            p = period_token.strip()
            if any(k in p for k in ("下旬", "下", "末", "いっぱい", "一杯", "月内", "後半", "最終")):
                return calendar.monthrange(y, m)[1]
            if any(k in p for k in ("中旬", "中頃", "中ごろ", "中", "半ば", "なかば", "央")):
                return 20
            if "前半" in p:
                return 15
            if any(k in p for k in ("上旬", "上", "初旬", "初", "はじめ", "頭", "あたま")):
                return 10
            return None

        # 曜日表記 (月), (火) 等を除去し、残りの括弧は空白化（例: 11月(上旬) などを保持）
        s = re.sub(r"\([月火水木金土日祝・]+\)", " ", s)
        s = s.replace("(", " ").replace(")", " ")

        # 和暦（令和/R・平成/H）を西暦に変換
        s = re.sub(r"令和\s*元\s*年", "2019年", s)
        s = re.sub(
            r"(?:令和|(?<![A-Za-z])[Rr])\s*(\d{1,2})\s*(?:年|[./-])",
            lambda m_era: f"{2018 + int(m_era.group(1))}年",
            s,
        )
        s = re.sub(
            r"(?:平成|(?<![A-Za-z])[Hh])\s*(\d{1,2})\s*(?:年|[./-])",
            lambda m_era: f"{1988 + int(m_era.group(1))}年",
            s,
        )

        fallback_year = default_year or datetime.date.today().year
        m_yr4 = re.search(r"(?<!\d)(20\d{2}|19\d{2})\s*(?:年|[./-])", s)
        m_yr2 = re.search(r"(?:^|(?<!\d))(\d{2})\s*(?:年|(?=[./-]\s*\d{1,2}\s*[月./-]))", s)
        if m_yr4:
            inferred_y = int(m_yr4.group(1))
        elif m_yr2 and 15 <= int(m_yr2.group(1)) <= 99:
            inferred_y = 2000 + int(m_yr2.group(1))
        else:
            inferred_y = fallback_year

        # 1. 期間表現（11/上旬、11/上、11/中、11/下、11月末、11月いっぱい、11/上～中 など）の解析
        if allow_period:
            period_matches = list(
                re.finditer(
                    r"(?:(?<!\d)(?P<y>20\d{2}|19\d{2}|\d{2})\s*[年./-]\s*)?"
                    r"(?<!\d)(?P<m>1[0-2]|0?[1-9])\s*"
                    r"(?:"
                    r"(?:月|/|\.|-|ー|―)\s*"
                    r"(?:(?:上旬|中旬|下旬|初旬|前半|後半|上|中|下|初)\s*(?:[~〜～・/ー―-]|から)\s*)?"
                    r"(?P<p1>上旬|中旬|下旬|初旬|中頃|中ごろ|末日|末頃|末ごろ|月末|月内|いっぱい|一杯|前半|後半|半ば|なかば|初め|はじめ|あたま|最終|上|中(?!止)|下|初|末|頭|央)"
                    r"|"
                    r"\s*(?P<p2>上旬|中旬|下旬|初旬|中頃|中ごろ|末日|月末|月内|いっぱい|一杯)"
                    r")",
                    s,
                )
            )
            if period_matches:
                last_m = period_matches[-1]
                y_str = last_m.group("y")
                if y_str:
                    y_val = int(y_str)
                    if y_val < 100:
                        y_val += 2000
                else:
                    y_val = inferred_y
                m_val = int(last_m.group("m"))
                p_str = last_m.group("p1") or last_m.group("p2") or ""
                d_val = _period_to_day(y_val, m_val, p_str)
                if d_val is not None:
                    res = _make_valid_date(y_val, m_val, d_val)
                    if res:
                        return res

        # 2. 西暦4桁 + 月 + 日（日範囲 2026/11/10～15 等にも対応）
        m_full = re.search(
            r"(?<!\d)(20\d{2}|19\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})"
            r"(?:\s*日?\s*(?:[~〜～ー―-]|から)\s*(?:(?:\d{1,2})\s*[月./-]\s*)?(\d{1,2}))?(?!\d)",
            s,
        )
        if m_full:
            y_val = int(m_full.group(1))
            m_val = int(m_full.group(2))
            d_val = int(m_full.group(4)) if m_full.group(4) else int(m_full.group(3))
            res = _make_valid_date(y_val, m_val, d_val)
            if res:
                return res

        # 3. 西暦2桁 + 月 + 日（例: 26/11/15, 26.11.15頃）
        m_short_yr = re.search(
            r"(?:^|(?<!\d))(\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})(?!\d)",
            s,
        )
        if m_short_yr:
            yr2 = int(m_short_yr.group(1))
            if 15 <= yr2 <= 99:
                res = _make_valid_date(2000 + yr2, int(m_short_yr.group(2)), int(m_short_yr.group(3)))
                if res:
                    return res

        # 4. 月日範囲（例: 11/10～11/20, 11/25〜12/5）
        m_md_range2 = re.search(
            r"(?<!\d)(\d{1,2})\s*[月./-]\s*(\d{1,2})\s*日?\s*(?:[~〜～ー―]|から)\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})(?!\d)",
            s,
        )
        if m_md_range2:
            res = _make_valid_date(inferred_y, int(m_md_range2.group(3)), int(m_md_range2.group(4)))
            if res:
                return res

        # 5. 同月内の日範囲（例: 11/10～20, 11月10日〜15日, 11/10-15）
        m_md_range1 = re.search(
            r"(?<!\d)(\d{1,2})\s*(?:月|/|\.)\s*(\d{1,2})\s*日?\s*(?:[~〜～ー―-]|から)\s*(\d{1,2})(?!\d)",
            s,
        )
        if m_md_range1:
            res = _make_valid_date(inferred_y, int(m_md_range1.group(1)), int(m_md_range1.group(3)))
            if res:
                return res

        # 6. 月 + 日（例: 11/15, 11月15日, 11.15, 11-15, 納期11/15頃）
        m_md = re.search(
            r"(?<![0-9./-])(\d{1,2})\s*(?:月|/|\.|-)\s*(\d{1,2})(?:\s*日)?(?![0-9./-])",
            s,
        )
        if m_md:
            res = _make_valid_date(inferred_y, int(m_md.group(1)), int(m_md.group(2)))
            if res:
                return res

        # 7. 8桁数値日付（例: 20261115）
        m_8digit = re.match(r"^(20\d{2}|19\d{2})(\d{2})(\d{2})$", s)
        if m_8digit:
            res = _make_valid_date(int(m_8digit.group(1)), int(m_8digit.group(2)), int(m_8digit.group(3)))
            if res:
                return res

        # 8. 月のみ表記（例: 11月, 11月予定, 11月納品, 2026年11月）→ その月の末日
        if allow_period:
            m_month_only = re.search(
                r"(?<!\d)(?:(20\d{2}|19\d{2})\s*[年./-]\s*)?(1[0-2]|0?[1-9])\s*月"
                r"(?:\s*(?:予定|頃|ごろ|度|納品|納入|入荷|迄|まで|目標|希望))?\s*$",
                s,
            )
            if m_month_only:
                y_val = int(m_month_only.group(1)) if m_month_only.group(1) else inferred_y
                m_val = int(m_month_only.group(2))
                max_d = calendar.monthrange(y_val, m_val)[1]
                res = _make_valid_date(y_val, m_val, max_d)
                if res:
                    return res

        return None

    @staticmethod
    def _is_filled_confirmation(raw_val) -> bool:
        if raw_val is None or pd.isna(raw_val):
            return False
        s = str(raw_val).strip()
        if not s or s.lower() in ("none", "nan", "nat"):
            return False
        if s in ("-", "ー", "―", "‐", "未", "未定", "未確認", "未着", "×", "✕"):
            return False
        return True

    @staticmethod
    def _normalize_key(text: str) -> str:
        if not text or pd.isna(text):
            return ""
        s = str(text).strip()
        if s.lower() in ("none", "nan", "nat", "-", "ー", "―", "‐"):
            return ""
        s = unicodedata.normalize("NFKC", s)
        for corp in ("株式会社", "有限会社", "(株)", "（株）", "(有)", "（有）", "御中", "様", "殿"):
            s = s.replace(corp, "")
        s = s.replace("×", "X").replace("＊", "X").replace("*", "X")
        s = re.sub(r"[^0-9A-Za-z一-龥ぁ-んァ-ン]", "", s)
        return s.upper()

    @staticmethod
    def _clean_company_name(text: str) -> str:
        if not text:
            return ""
        s = unicodedata.normalize("NFKC", str(text)).strip()
        for corp in ("株式会社", "有限会社", "(株)", "（株）", "(有)", "（有）", "御中", "様", "殿", "本社", "営業所", "工場"):
            s = s.replace(corp, "")
        s = re.sub(r"[\s\-_・/]", "", s)
        return s.upper()

    @classmethod
    def _make_bigrams(cls, norm_str: str) -> frozenset:
        if len(norm_str) < 2:
            return frozenset()
        return frozenset(norm_str[i : i + 2] for i in range(len(norm_str) - 1))

    @classmethod
    def _fast_similarity(
        cls, na: str, bg_a: frozenset, nb: str, bg_b: frozenset
    ) -> float:
        if not na or not nb:
            return 0.0
        if na == nb:
            return 1.0
        if na in nb or nb in na:
            shorter = min(len(na), len(nb))
            longer = max(len(na), len(nb))
            return 0.85 + 0.15 * (shorter / longer)

        if bg_a and bg_b:
            inter = len(bg_a & bg_b)
            if inter == 0:
                return 0.0
            bigram_score = inter / min(len(bg_a), len(bg_b))
            if bigram_score >= 0.40:
                seq_ratio = difflib.SequenceMatcher(None, na, nb).quick_ratio()
                return max(bigram_score, seq_ratio)
            return bigram_score

        return difflib.SequenceMatcher(None, na, nb).quick_ratio()

    @staticmethod
    def _extract_Leading_number(text: str) -> str:
        if text is None or pd.isna(text):
            return ""
        s = str(text).strip()
        if s.lower() in ("none", "nan", "nat"):
            return ""
        s = unicodedata.normalize("NFKC", s).replace(",", "")
        m = re.search(r"(\d+(?:\.\d+)?)", s)
        if not m:
            return ""
        val_str = m.group(1)
        if val_str.endswith(".0"):
            val_str = val_str[:-2]
        return val_str

    @staticmethod
    def _is_zero_quantity(qty_str: str) -> bool:
        """購入数・数量文字列が 0 または 0.0、全角０等であるかを確実に判定する"""
        if qty_str is None or pd.isna(qty_str):
            return False
        clean = unicodedata.normalize("NFKC", str(qty_str)).strip()
        if clean in ("0", "0.0", "０", "0個", "0台", "0本", "0式", "0組", "0ヶ"):
            return True
        m = re.match(r"^0+(?:\.0+)?(?:\D.*)?$", clean)
        if m:
            return True
        try:
            num_match = re.search(r"^(\d+(?:\.\d+)?)", clean)
            if num_match and float(num_match.group(1)) == 0.0:
                return True
        except Exception:
            pass
        return False

    @staticmethod
    def _clean_cell_str(v) -> str:
        if v is None or pd.isna(v):
            return ""
        s = str(v).strip()
        if s.lower() in ("none", "nan", "nat"):
            return ""
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        return s

    @staticmethod
    def _extract_8digit_no(text: str) -> str:
        if not text:
            return ""
        s = unicodedata.normalize("NFKC", str(text)).strip()
        if s.endswith(".0"):
            s = s[:-2]
        m = re.search(r"(?<!\d)(\d{8})(?!\d)", s)
        if m:
            return m.group(1)
        digits_only = re.sub(r"\D", "", s)
        if len(digits_only) == 8 and not digits_only.startswith("202"):
            return digits_only
        m2 = re.search(r"(\d{8})", s)
        return m2.group(1) if m2 else ""

    # =========================================================================
    # Excel解析ロジック
    # =========================================================================
    def _parse_management_book(self, filepath: str, progress_cb=None) -> list[dict]:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        mgmt_rows = []
        sheet_names = wb.sheetnames
        total_sheets = len(sheet_names)

        for s_idx, sheet_name in enumerate(sheet_names, start=1):
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

            sheet_default_order_no = self._extract_8digit_no(sheet_name)

            ws = wb[sheet_name]
            raw_rows = []
            empty_streak = 0
            for r_idx, row_tuple in enumerate(ws.iter_rows(max_col=60, values_only=True), start=1):
                if not row_tuple or not any(v is not None and str(v).strip() != "" for v in row_tuple):
                    empty_streak += 1
                    if r_idx > 20 and empty_streak >= 100:
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
                        elif "発注日" in v_comb or "注文日" in v_comb:
                            col_map.setdefault("order_date", c_idx)
                        elif "図面番号" in v_comb or "図番" in v_comb:
                            col_map.setdefault("drawing_no", c_idx)
                        elif "品番" in v_comb or "型番" in v_comb or "部品番号" in v_comb:
                            col_map.setdefault("part_no", c_idx)
                        elif "ライン名" in v_comb or "品名、ライン名" in v_comb or "品名・ライン名" in v_comb:
                            col_map["line_name"] = c_idx
                        elif ("品名" in v_comb or "名称" in v_comb or "商品名" in v_comb) and "品名、ライン名" not in v_comb:
                            if c_idx >= 15:
                                col_map["item_name"] = c_idx
                            elif "item_name" not in col_map:
                                col_map["item_name"] = c_idx
                        elif "メーカー" in v_comb or "発注先" in v_comb or "仕入先" in v_comb or "業者" in v_comb:
                            col_map.setdefault("maker", c_idx)
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
                    break

            if header_pos_in_list is None:
                continue

            col_map.setdefault("actual_date", 1)    # A列: 現品確認日
            col_map.setdefault("order_no", 13)      # M列: 管理No
            col_map.setdefault("item_name", 20)     # T列: 品名
            col_map.setdefault("maker", 21)         # U列: メーカー
            col_map.setdefault("spec", 23)          # W列: 仕様
            col_map["qty"] = 28                     # AB列: 購入数
            col_map.setdefault("invoice_date", 35)  # AI列: 請求日

            data_slice = raw_rows[header_pos_in_list + 1 :]

            def _get_col(rtuple: tuple, col_1based: int | None):
                if not col_1based or col_1based < 1 or col_1based > len(rtuple):
                    return None
                return rtuple[col_1based - 1]

            last_seen_order_no_8 = sheet_default_order_no

            for r_idx, row_tuple in data_slice:
                drawing_val = _get_col(row_tuple, col_map.get("drawing_no"))
                part_val = _get_col(row_tuple, col_map.get("part_no"))
                item_val = _get_col(row_tuple, col_map.get("item_name"))
                maker_val = _get_col(row_tuple, col_map.get("maker"))
                spec_val = _get_col(row_tuple, col_map.get("spec"))
                line_val = _get_col(row_tuple, col_map.get("line_name"))
                order_no_val = _get_col(row_tuple, col_map.get("order_no"))

                qty_val = _get_col(row_tuple, 28)
                if qty_val is None:
                    qty_val = _get_col(row_tuple, col_map.get("qty"))

                if not any([drawing_val, part_val, item_val, spec_val, maker_val, line_val]):
                    continue

                order_date_raw = _get_col(row_tuple, col_map.get("order_date"))
                order_date = self._parse_flexible_date(order_date_raw, allow_period=False)
                inferred_year = order_date.year if order_date else datetime.date.today().year

                a_col_val = _get_col(row_tuple, 1)
                mapped_actual_val = _get_col(row_tuple, col_map.get("actual_date"))
                actual_raw = a_col_val if self._is_filled_confirmation(a_col_val) else mapped_actual_val

                ai_col_val = _get_col(row_tuple, 35)
                mapped_inv_val = _get_col(row_tuple, col_map.get("invoice_date"))
                invoice_raw = ai_col_val if self._is_filled_confirmation(ai_col_val) else mapped_inv_val

                due_raw = _get_col(row_tuple, col_map.get("due_date"))

                actual_date = self._parse_flexible_date(actual_raw, default_year=inferred_year, allow_period=False)
                actual_filled = actual_date is not None or self._is_filled_confirmation(actual_raw)

                invoice_date = self._parse_flexible_date(invoice_raw, default_year=inferred_year, allow_period=False)
                invoice_filled = invoice_date is not None or self._is_filled_confirmation(invoice_raw)

                mgmt_due_date = self._parse_flexible_date(due_raw, default_year=inferred_year, allow_period=True)

                drawing_str = self._clean_cell_str(drawing_val)
                part_str = self._clean_cell_str(part_val)
                item_str = self._clean_cell_str(item_val)
                maker_str = self._clean_cell_str(maker_val)
                spec_str = self._clean_cell_str(spec_val)
                line_str = self._clean_cell_str(line_val)
                qty_str = self._clean_cell_str(qty_val)
                unit_price_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("unit_price")))
                amount_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("amount")))

                raw_order_no_str = self._clean_cell_str(order_no_val)
                exact_order_no_8 = self._extract_8digit_no(raw_order_no_str)

                if not exact_order_no_8:
                    for c_scan in (13, 11, 2, 3, 4):
                        cand_8 = self._extract_8digit_no(_get_col(row_tuple, c_scan))
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
                person_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("person")))

                model_display_parts = [p for p in (drawing_str, part_str, spec_str) if p and p != "-"]
                model_spec_display = " / ".join(dict.fromkeys(model_display_parts))

                norm_drawing = self._normalize_key(drawing_str)
                norm_part = self._normalize_key(part_str)
                norm_item = self._normalize_key(item_str)
                norm_spec = self._normalize_key(spec_str)
                norm_line = self._normalize_key(line_str)
                norm_maker = self._clean_company_name(maker_str)
                norm_combined = self._normalize_key(f"{drawing_str} {part_str} {item_str} {spec_str} {line_str}")

                mgmt_rows.append(
                    {
                        "sheet": sheet_name,
                        "sheet_cat": sheet_cat,
                        "sheet_cat_label": sheet_cat_label,
                        "row": r_idx,
                        "order_no": order_no_str,
                        "order_no_8": exact_order_no_8,
                        "norm_order_no": self._normalize_key(order_no_str),
                        "order_date": order_date,
                        "drawing_no": drawing_str,
                        "part_no": part_str,
                        "item_name": item_str,
                        "line_name": line_str,
                        "maker": maker_str,
                        "supplier": maker_str,
                        "spec": spec_str,
                        "model_spec": model_spec_display,
                        "qty": qty_str,
                        "qty_num": self._extract_Leading_number(qty_str),
                        "unit_price": unit_price_str,
                        "amount": amount_str,
                        "actual_date": actual_date,
                        "actual_raw": self._clean_cell_str(actual_raw),
                        "actual_filled": actual_filled,
                        "invoice_date": invoice_date,
                        "invoice_raw": self._clean_cell_str(invoice_raw),
                        "invoice_filled": invoice_filled,
                        "mgmt_due_date": mgmt_due_date,
                        "mgmt_due_raw": self._clean_cell_str(due_raw),
                        "person": person_str,
                        "norm_drawing": norm_drawing,
                        "bg_drawing": self._make_bigrams(norm_drawing),
                        "norm_part": norm_part,
                        "bg_part": self._make_bigrams(norm_part),
                        "norm_item": norm_item,
                        "bg_item": self._make_bigrams(norm_item),
                        "norm_line": norm_line,
                        "bg_line": self._make_bigrams(norm_line),
                        "norm_spec": norm_spec,
                        "bg_spec": self._make_bigrams(norm_spec),
                        "norm_maker": norm_maker,
                        "bg_maker": self._make_bigrams(norm_maker),
                        "norm_combined": norm_combined,
                        "bg_combined": self._make_bigrams(norm_combined),
                        "matched": False,
                        "matched_by": "",
                    }
                )

        wb.close()
        return mgmt_rows

    def _parse_order_book(self, filepath: str, valid_mgmt_order_nos: set[str]) -> list[dict]:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        order_items = []
        filename = os.path.basename(filepath)
        norm_fname = unicodedata.normalize("NFKC", filename)

        prefix_match = re.match(r"^\s*(\d{8})", norm_fname)
        if prefix_match:
            file_order_no_8 = prefix_match.group(1)
        else:
            file_order_no_8 = self._extract_8digit_no(norm_fname)

        if file_order_no_8 and file_order_no_8 not in valid_mgmt_order_nos:
            wb.close()
            return []

        if "購依" in norm_fname or "購入" in norm_fname:
            req_type_code = "PURCHASE"
            req_type_label = "購入依頼"
        elif "外依" in norm_fname or "外注" in norm_fname:
            req_type_code = "OUTSOURCE"
            req_type_label = "外注依頼"
        else:
            req_type_code = "UNKNOWN"
            req_type_label = "一般注文"

        file_supplier = ""
        m_fname_sup = re.search(r"\d+(?:購依|外依)(.+?)注文", filename)
        if m_fname_sup:
            file_supplier = m_fname_sup.group(1).strip()

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            raw_rows = []
            empty_streak = 0
            for r_idx, row_tuple in enumerate(ws.iter_rows(max_row=300, max_col=35, values_only=True), start=1):
                if not row_tuple or not any(v is not None and str(v).strip() != "" for v in row_tuple):
                    empty_streak += 1
                    if r_idx > 25 and empty_streak >= 30:
                        break
                    continue
                empty_streak = 0
                raw_rows.append((r_idx, row_tuple))

            if not raw_rows:
                continue

            row2_text = ""
            for r_idx, row_tuple in raw_rows:
                if r_idx == 2:
                    row2_text = " ".join(
                        unicodedata.normalize("NFKC", str(v or "")).strip() for v in row_tuple
                    )
                    break

            if "見積" in row2_text:
                continue

            if row2_text and not any(k in row2_text for k in ("注文", "発注")):
                if "見積" in row2_text or "見積" in filename or "見積" in sheet_name:
                    continue

            sheet_order_no = file_order_no_8
            sheet_supplier = file_supplier
            sheet_order_date = None
            default_year = datetime.date.today().year

            # 注文書4行目にある注文日付の明示的探索
            for r_idx, row_tuple in raw_rows:
                if r_idx == 4:
                    for cell_v in row_tuple:
                        if cell_v is not None and str(cell_v).strip() != "":
                            d_cand4 = self._parse_flexible_date(cell_v, allow_period=False)
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
                        m_in = self._extract_8digit_no(cell_val)
                        if m_in:
                            sheet_order_no = m_in
                        else:
                            for offset in range(1, 9):
                                if c_idx - 1 + offset >= len(row_tuple):
                                    break
                                right_val = str(row_tuple[c_idx - 1 + offset] or "").strip()
                                m_r8 = self._extract_8digit_no(right_val)
                                if m_r8:
                                    sheet_order_no = m_r8
                                    break

                    if "注文日" in cell_val or "発注日" in cell_val:
                        for offset in range(1, 9):
                            if c_idx - 1 + offset >= len(row_tuple):
                                break
                            d_candidate = self._parse_flexible_date(row_tuple[c_idx - 1 + offset], allow_period=False)
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
                    elif (
                        (
                            "株式会社" in cell_val
                            or "有限会社" in cell_val
                            or "(株)" in cell_val
                            or "（株）" in cell_val
                            or "(有)" in cell_val
                            or "（有）" in cell_val
                        )
                        and r_idx <= 10
                        and "丸菱" not in cell_val
                    ):
                        cleaned_sup = re.sub(r"(御中|様|殿)$", "", cell_val).strip()
                        if cleaned_sup and not sheet_supplier:
                            sheet_supplier = cleaned_sup

            resolved_no_8 = file_order_no_8 or (
                sheet_order_no if len(sheet_order_no) == 8 and sheet_order_no.isdigit() else self._normalize_key(sheet_order_no)
            )
            if not resolved_no_8 or resolved_no_8 not in valid_mgmt_order_nos:
                continue

            header_pos_in_list = None
            col_map = {}
            for list_i, (r_idx, row_tuple) in enumerate(raw_rows):
                if r_idx > 25:
                    break
                row_vals = [str(v or "").strip() for v in row_tuple]
                joined = " ".join(row_vals)
                if "現品" in joined and "請求" in joined:
                    break
                if ("品名" in joined or "図面" in joined or "品番" in joined or "名称" in joined) and (
                    "数量" in joined or "購入数" in joined or "納期" in joined or "単価" in joined
                ):
                    header_pos_in_list = list_i
                    for c_idx, val in enumerate(row_vals, start=1):
                        v_clean = re.sub(r"\s+", "", unicodedata.normalize("NFKC", val))
                        if not v_clean:
                            continue
                        if "図面番号" in v_clean or "図番" in v_clean:
                            col_map.setdefault("drawing_no", c_idx)
                        elif "品番" in v_clean or "型番" in v_clean:
                            col_map.setdefault("part_no", c_idx)
                        elif "品名" in v_clean or "名称" in v_clean or "商品名" in v_clean:
                            col_map.setdefault("item_name", c_idx)
                        elif "メーカー" in v_clean:
                            col_map.setdefault("maker", c_idx)
                        elif "仕様" in v_clean or "規格" in v_clean or "材質" in v_clean or "寸法" in v_clean:
                            col_map.setdefault("spec", c_idx)
                        elif "購入数" in v_clean or "数量" in v_clean or v_clean in ("数", "必要数"):
                            col_map.setdefault("qty", c_idx)
                        elif "単位" in v_clean:
                            col_map.setdefault("unit", c_idx)
                        elif "単価" in v_clean and "金額" not in v_clean:
                            col_map.setdefault("unit_price", c_idx)
                        elif "金額" in v_clean:
                            col_map.setdefault("amount", c_idx)
                        elif "納期" in v_clean or "希望納期" in v_clean or "納入日" in v_clean:
                            col_map.setdefault("due_date", c_idx)
                        elif "備考" in v_clean or "摘要" in v_clean:
                            col_map.setdefault("note", c_idx)
                    break

            if header_pos_in_list is None:
                continue

            if "drawing_no" in col_map and "part_no" in col_map and col_map["part_no"] == col_map["drawing_no"] + 1:
                base_c = col_map["drawing_no"]
                col_map.setdefault("item_name", base_c + 2)
                col_map.setdefault("maker", base_c + 3)
                col_map.setdefault("spec", base_c + 4)
                col_map.setdefault("qty", base_c + 5)
                col_map.setdefault("unit_price", base_c + 6)
                col_map.setdefault("amount", base_c + 7)
                col_map.setdefault("due_date", base_c + 8)

            data_rows = raw_rows[header_pos_in_list + 1 :]

            if "item_name" in col_map and not any(k in col_map for k in ("drawing_no", "part_no", "spec")):
                item_col = col_map["item_name"]
                qty_col = col_map.get("qty", item_col + 9)
                active_sub_cols = []
                for c_scan in range(item_col + 1, qty_col):
                    for _, sample_tuple in data_rows[:18]:
                        if c_scan - 1 < len(sample_tuple):
                            val_scan = sample_tuple[c_scan - 1]
                            if val_scan is not None and str(val_scan).strip() != "":
                                active_sub_cols.append(c_scan)
                                break
                if len(active_sub_cols) == 1:
                    col_map["spec"] = active_sub_cols[0]
                elif len(active_sub_cols) >= 2:
                    col_map["drawing_no"] = active_sub_cols[0]
                    col_map["spec"] = active_sub_cols[1]

            sorted_cols = sorted(set(col_map.values()))

            def _get_non_empty_in_span(rtuple: tuple, start_col: int | None):
                if not start_col:
                    return None
                next_cols = [c for c in sorted_cols if c > start_col]
                end_col = next_cols[0] - 1 if next_cols else min(start_col + 4, len(rtuple))
                for c in range(start_col, max(start_col, end_col) + 1):
                    if 1 <= c <= len(rtuple):
                        v = rtuple[c - 1]
                        if v is not None and str(v).strip() != "":
                            return v
                return None

            for r_idx, row_tuple in data_rows:
                drawing_str = self._clean_cell_str(_get_non_empty_in_span(row_tuple, col_map.get("drawing_no")))
                part_str = self._clean_cell_str(_get_non_empty_in_span(row_tuple, col_map.get("part_no")))
                item_str = self._clean_cell_str(_get_non_empty_in_span(row_tuple, col_map.get("item_name")))
                maker_str = self._clean_cell_str(_get_non_empty_in_span(row_tuple, col_map.get("maker")))
                spec_str = self._clean_cell_str(_get_non_empty_in_span(row_tuple, col_map.get("spec")))

                if not any([drawing_str, part_str, item_str, spec_str]):
                    continue
                joined_check = f"{drawing_str}{part_str}{item_str}{spec_str}"
                if any(
                    stop_word in joined_check
                    for stop_word in ("合計", "小計", "消費税", "総合計", "特記事項", "以上", "振込", "備考")
                ):
                    continue

                qty_val = _get_non_empty_in_span(row_tuple, col_map.get("qty"))
                unit_val = _get_non_empty_in_span(row_tuple, col_map.get("unit"))
                unit_price_val = _get_non_empty_in_span(row_tuple, col_map.get("unit_price"))
                amount_val = _get_non_empty_in_span(row_tuple, col_map.get("amount"))
                due_raw = _get_non_empty_in_span(row_tuple, col_map.get("due_date"))

                due_date = self._parse_flexible_date(due_raw, default_year=default_year)

                qty_num_str = self._clean_cell_str(qty_val)
                unit_str = self._clean_cell_str(unit_val)
                qty_str = f"{qty_num_str}{unit_str}".strip()
                qty_num = self._extract_Leading_number(qty_str)

                unit_price_str = self._clean_cell_str(unit_price_val)
                amount_str = self._clean_cell_str(amount_val)

                if not qty_num and due_date is None:
                    continue

                model_parts = [p for p in (drawing_str, part_str, spec_str) if p and p != "-"]
                model_spec_display = " / ".join(dict.fromkeys(model_parts))

                norm_drawing = self._normalize_key(drawing_str)
                norm_part = self._normalize_key(part_str)
                norm_item = self._normalize_key(item_str)
                norm_spec = self._normalize_key(spec_str)
                norm_maker = self._clean_company_name(maker_str)
                norm_supplier = self._clean_company_name(sheet_supplier or file_supplier)
                norm_combined = self._normalize_key(f"{drawing_str} {part_str} {item_str} {spec_str}")

                order_items.append(
                    {
                        "file": filename,
                        "full_path": filepath,
                        "req_type_code": req_type_code,
                        "req_type_label": req_type_label,
                        "sheet": sheet_name,
                        "row": r_idx,
                        "order_no": sheet_order_no,
                        "order_no_8": resolved_no_8,
                        "norm_order_no": self._normalize_key(sheet_order_no),
                        "order_date": sheet_order_date,
                        "drawing_no": drawing_str,
                        "part_no": part_str,
                        "item_name": item_str,
                        "maker": maker_str,
                        "supplier": sheet_supplier or file_supplier,
                        "norm_supplier": norm_supplier,
                        "spec": spec_str,
                        "model_spec": model_spec_display,
                        "qty": qty_str,
                        "qty_num": qty_num,
                        "unit_price": unit_price_str,
                        "amount": amount_str,
                        "due_date": due_date,
                        "due_raw": self._clean_cell_str(due_raw),
                        "norm_drawing": norm_drawing,
                        "bg_drawing": self._make_bigrams(norm_drawing),
                        "norm_part": norm_part,
                        "bg_part": self._make_bigrams(norm_part),
                        "norm_item": norm_item,
                        "bg_item": self._make_bigrams(norm_item),
                        "norm_spec": norm_spec,
                        "bg_spec": self._make_bigrams(norm_spec),
                        "norm_maker": norm_maker,
                        "bg_maker": self._make_bigrams(norm_maker),
                        "norm_combined": norm_combined,
                        "bg_combined": self._make_bigrams(norm_combined),
                    }
                )

        wb.close()
        return order_items

    # =========================================================================
    # 照合処理
    # =========================================================================
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
        self.run_btn.configure(state=tk.DISABLED, text="照合処理中...")
        self.excel_open_btn.configure(state=tk.DISABLED)
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

    def _analysis_worker(self, mgmt_path: str, order_dir: str, base_date: datetime.date, grace_days: int = 7):
        try:
            self._report_progress(5.0, "[1/3] 注文品管理表を開いています...")
            mgmt_rows = self._parse_management_book(mgmt_path, progress_cb=self._report_progress)

            valid_mgmt_order_nos = set()
            for m in mgmt_rows:
                no8 = m.get("order_no_8") or m.get("norm_order_no")
                if no8:
                    valid_mgmt_order_nos.add(no8)

            self._report_progress(
                35.0,
                f"[2/3] 管理表にある管理No ({len(valid_mgmt_order_nos)}種類) の注文書を探索中...",
            )
            order_items, valid_order_files, ignored_files = self._collect_valid_order_files(
                order_dir, mgmt_path, valid_mgmt_order_nos, progress_cb=self._report_progress
            )

            self._report_progress(75.0, f"[3/3] 明細の双方向照合を実行中 ({len(order_items)} 件)...")
            main_records, missing_records = self._match_and_evaluate(
                mgmt_rows, order_items, base_date, grace_days=grace_days, progress_cb=self._report_progress
            )

            self.ui_queue.put(
                (
                    "DONE",
                    {
                        "main_records": main_records,
                        "missing_records": missing_records,
                        "base_date": base_date,
                        "valid_order_files": valid_order_files,
                        "ignored_files": ignored_files,
                        "order_items_count": len(order_items),
                        "mgmt_rows_count": len(mgmt_rows),
                        "mgmt_no_count": len(valid_mgmt_order_nos),
                    },
                )
            )
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
                    self.run_btn.configure(state=tk.NORMAL, text="照合・遅延チェック実行")
                    self.excel_open_btn.configure(state=tk.NORMAL)

                    self.analyzed_records = payload["main_records"]
                    self.missing_order_records = payload["missing_records"]
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

                    delay_cnt = sum(1 for r in self.analyzed_records if r["is_any_delay"])
                    unmatched_cnt = sum(1 for r in self.analyzed_records if r.get("is_unmatched"))
                    missing_cnt = len(self.missing_order_records)

                    self.progress_var.set(100.0)
                    self.progress_text_var.set(
                        f"完了！ (注文書 {valid_order_files}ファイル / 管理表 {mgmt_rows_count}行・管理No {mgmt_no_count}種)"
                    )
                    self.status_var.set(
                        f"照合完了 (基準日: {base_date.strftime('%Y-%m-%d')}) — "
                        f"注文書明細: {len(self.analyzed_records)}件 (超過:{delay_cnt}件 / 管理表未記載:{unmatched_cnt}件) ｜ "
                        f"【別タブ】注文書未検出(管理表のみ): {missing_cnt}件"
                    )
                    return
                elif msg_type == "ERROR":
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL, text="照合・遅延チェック実行")
                    self.excel_open_btn.configure(state=tk.NORMAL)
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
                name_sim = self._fast_similarity(o_i, o_item["bg_item"], m_i, m_row["bg_item"])
                if name_sim >= 0.50:
                    score += 220.0 * name_sim
        elif o_i and m_line:
            if o_i == m_line or o_i in m_line or m_line in o_i:
                name_sim = 0.85
                score += 180.0

        spec_sim = 0.0
        if o_s and m_s:
            if o_s == m_s:
                spec_sim = 1.0
                score += 350.0
                strong_evidence = True
            elif o_s in m_s or m_s in o_s:
                spec_sim = 0.92
                score += 260.0
                strong_evidence = True
            else:
                spec_sim = self._fast_similarity(o_s, o_item["bg_spec"], m_s, m_row["bg_spec"])
                if spec_sim >= 0.50:
                    score += 200.0 * spec_sim

        comb_sim = self._fast_similarity(
            o_item["norm_combined"],
            o_item["bg_combined"],
            m_row["norm_combined"],
            m_row["bg_combined"],
        )
        if comb_sim >= 0.60:
            score += 180.0 * comb_sim

        if name_sim >= 0.70 and spec_sim >= 0.65:
            strong_evidence = True
        elif name_sim >= 0.80 and (not o_s or not m_s):
            strong_evidence = True
        elif spec_sim >= 0.80 and (not o_i or not m_i):
            strong_evidence = True
        elif comb_sim >= 0.75:
            strong_evidence = True

        if o_item["qty_num"] and m_row["qty_num"]:
            if o_item["qty_num"] == m_row["qty_num"]:
                score += 150.0
                strong_evidence = True
            else:
                score -= 40.0

        if (o_d and m_d and o_d == m_d) or (o_p and m_p and o_p == m_p):
            strong_evidence = True

        if not strong_evidence:
            return -1.0

        return score

    def _build_field_comparison_data(
        self, o_item: dict | None, m_row: dict | None, reason_summary: str, extra_notes: str = ""
    ) -> dict:
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

        cross_notes = []
        if o_item and m_row:
            if o_item["norm_spec"] and m_row["norm_part"] and o_item["norm_spec"] == m_row["norm_part"]:
                cross_notes.append("・注文書の『仕様』と管理表の『品番』が一致しています。")
            if o_item["norm_part"] and m_row["norm_spec"] and o_item["norm_part"] == m_row["norm_spec"]:
                cross_notes.append("・注文書の『品番』と管理表の『仕様』が一致しています。")
            if m_row.get("line_name"):
                cross_notes.append(f"・管理表の『製品詳細 / ライン名等』: {m_row['line_name']}")

        detail_lines = [f"【診断結果サマリー】\n  {reason_summary}", ""]
        if o_item:
            detail_lines.append(
                f"【注文書データ元】\n  ファイル: {o_item['file']} (シート: {o_item['sheet']} / 行: {o_item['row']}) / 依頼区分: {o_item['req_type_label']}"
            )
        else:
            detail_lines.append("【注文書データ元】\n  （注文書フォルダ内に該当する注文書が見当たりません）")

        if m_row:
            detail_lines.append(
                f"【管理表の登録データ】\n  シート: {m_row['sheet']} (行: {m_row['row']}) / 管理No: {m_row.get('order_no_8') or m_row.get('order_no')} / AB列購入数: {m_row.get('qty') or '未記載'} / 納期: {m_row.get('mgmt_due_raw') or '未記載'}"
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
        }

    def _diagnose_unmatched_item(
        self,
        o_item: dict,
        mgmt_rows: list[dict],
    ) -> tuple[str, dict]:
        o_no = o_item.get("order_no_8", "")

        same_no_rows = [
            m
            for m in mgmt_rows
            if m.get("order_no_8") == o_no or m.get("norm_order_no") == o_no
        ]

        best_cand = None
        best_raw_sim = -999.0
        for m in same_no_rows:
            raw_sc = self._score_candidate_pair(o_item, m)
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

            diag = self._build_field_comparison_data(
                o_item, best_cand, reason, "\n".join(cand_list_lines)
            )
            return reason, diag

        reason = f"【明細未記載】管理No[{o_no}]は管理表にありますが一致する明細行がありません"
        diag = self._build_field_comparison_data(o_item, None, reason)
        return reason, diag

    def _match_and_evaluate(
        self,
        mgmt_rows: list[dict],
        order_items: list[dict],
        base_date: datetime.date,
        grace_days: int = 7,
        progress_cb=None,
    ) -> tuple[list[dict], list[dict]]:
        """
        注文書明細と管理表行を照合し、
        1. main_results: 注文書由来の照合・遅延一覧（タブ1用）
        2. missing_order_results: 注文書未検出・未手配の管理表明細一覧（タブ2用 / ※AB列購入数0は完全除外）
        の2つに完全に分離して返す。
        """
        main_results = []
        missing_order_results = []

        all_mgmt_order_nos_8 = set()
        mgmt_by_no = {}
        for m in mgmt_rows:
            m["matched"] = False
            m["matched_by"] = ""
            no8 = m.get("order_no_8") or m.get("norm_order_no")
            if no8:
                all_mgmt_order_nos_8.add(no8)
                mgmt_by_no.setdefault(no8, []).append(m)

        orders_by_no = {}
        for o_item in order_items:
            o_no = o_item.get("order_no_8", "")
            if o_no and o_no in all_mgmt_order_nos_8:
                orders_by_no.setdefault(o_no, []).append(o_item)

        total_orders = sum(len(order_list) for order_list in orders_by_no.values())
        processed_orders = 0

        # 1. 注文書明細ベースの照合（タブ1用）
        for o_no, o_list in orders_by_no.items():
            m_list = mgmt_by_no.get(o_no, [])

            cost_matrix = []
            score_matrix = []

            for o in o_list:
                row_costs = []
                row_scores = []
                req_code = o.get("req_type_code", "UNKNOWN")
                for m in m_list:
                    base_sc = self._score_candidate_pair(o, m)
                    if req_code == "PURCHASE" and m.get("sheet_cat") == "PURCHASE":
                        base_sc += 150.0
                    elif req_code == "OUTSOURCE" and m.get("sheet_cat") == "OUTSOURCE":
                        base_sc += 150.0

                    row_scores.append(base_sc)
                    cost = 10000.0 - max(-500.0, base_sc)
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
                processed_orders += 1
                if progress_cb and (processed_orders == 1 or processed_orders % 5 == 0 or processed_orders == total_orders):
                    pct = 75.0 + (18.0 * processed_orders / max(1, total_orders))
                    progress_cb(pct, f"[3/3] 注文書明細と管理表を照合・診断中 ({processed_orders}/{total_orders} 件目)...")

                order_due = o_item.get("due_date")
                order_dt = o_item.get("order_date")
                is_fallback = False
                if order_due:
                    final_due = order_due
                elif order_dt:
                    final_due = order_dt + datetime.timedelta(days=grace_days)
                    is_fallback = True
                else:
                    final_due = best_mgmt.get("mgmt_due_date") if r_idx in matched_pairs else None

                if r_idx in matched_pairs:
                    c_idx, sc = matched_pairs[r_idx]
                    best_mgmt = m_list[c_idx]
                    best_mgmt["matched"] = True
                    best_mgmt["matched_by"] = f"{o_item['file']}(行{o_item['row']})"

                    final_supplier = o_item.get("supplier") or best_mgmt.get("supplier", "")
                    final_maker = o_item.get("maker") or best_mgmt.get("maker", "")
                    match_summary = f"◎ 一致 ({best_mgmt['sheet']} 行{best_mgmt['row']} / スコア:{int(sc)})"
                    diag_data = self._build_field_comparison_data(o_item, best_mgmt, match_summary)

                    rec = self._evaluate_single_record(
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=best_mgmt.get("order_no_8") or best_mgmt.get("order_no") or o_no,
                        supplier=final_supplier,
                        maker=final_maker,
                        item_name=o_item.get("item_name") or best_mgmt.get("item_name", ""),
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
                    reason_txt, diag_data = self._diagnose_unmatched_item(o_item, mgmt_rows)
                    rec = self._evaluate_single_record(
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=o_no,
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

        # 2. 注文書未検出（管理表にあるのに注文書が見当たらない）明細の抽出（タブ2用）
        for m in mgmt_rows:
            if m.get("matched"):
                continue

            raw_qty = m.get("qty", "")
            qty_num_str = m.get("qty_num", "")
            if self._is_zero_quantity(raw_qty) or self._is_zero_quantity(qty_num_str):
                continue

            m_no = m.get("order_no_8") or m.get("order_no", "")
            reason_txt = f"管理表『{m['sheet']}』(行{m['row']})に記載ありますが、対応する注文書ファイルが作成・手配されていません"

            diag_data = self._build_field_comparison_data(None, m, reason_txt)

            rec_missing = {
                "status_badge": "未発注 (注文書なし)",
                "mgmt_sheet": m["sheet"],
                "mgmt_row": str(m["row"]),
                "order_no": m_no,
                "supplier": m.get("supplier") or m.get("maker", ""),
                "maker": m.get("maker", ""),
                "item_name": m.get("item_name", ""),
                "model_spec": m.get("model_spec", ""),
                "qty": m.get("qty", ""),
                "due_date": m.get("mgmt_due_raw") or (m["mgmt_due_date"].strftime("%Y-%m-%d") if m.get("mgmt_due_date") else "-"),
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
                r["order_no"],
            )
        )

        missing_order_results.sort(
            key=lambda r: (
                r["order_no"],
                r["mgmt_sheet"],
                int(r["mgmt_row"]) if r["mgmt_row"].isdigit() else 0,
            )
        )

        return main_results, missing_order_results

    def _evaluate_single_record(
        self,
        req_type: str,
        order_no: str,
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
    ) -> dict:
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
            "req_type": req_type,
            "order_no": order_no,
            "supplier": supplier,
            "maker": maker,
            "item_name": item_name,
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
            "is_any_delay": (actual_overdue or invoice_overdue) and not is_unmatched,
            "row_tag": row_tag,
            "diagnostic_data": diagnostic_data or {},
        }

    def _update_kpis(self):
        total = len(self.analyzed_records)
        unmatched_cnt = sum(1 for r in self.analyzed_records if r.get("is_unmatched"))
        both_cnt = sum(
            1
            for r in self.analyzed_records
            if not r.get("is_unmatched") and r["actual_overdue"] and r["invoice_overdue"]
        )
        actual_cnt = sum(
            1
            for r in self.analyzed_records
            if not r.get("is_unmatched") and r["actual_overdue"] and not r["invoice_overdue"]
        )
        invoice_cnt = sum(
            1
            for r in self.analyzed_records
            if not r.get("is_unmatched") and not r["actual_overdue"] and r["invoice_overdue"]
        )
        ok_cnt = sum(
            1
            for r in self.analyzed_records
            if not r.get("is_unmatched") and not r["actual_overdue"] and not r["invoice_overdue"]
        )
        missing_cnt = len(self.missing_order_records)

        self.kpi_vars["order_total"].set(f"{total} 件")
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

        for idx, rec in enumerate(self.analyzed_records):
            if mode == "ALL_DELAY":
                if not (rec.get("is_unmatched") or rec.get("is_any_delay")):
                    continue
            elif mode == "UNMATCHED_ONLY":
                if not rec.get("is_unmatched"):
                    continue
            elif mode == "ACTUAL_DELAY":
                if rec.get("is_unmatched") or not rec["actual_overdue"]:
                    continue
            elif mode == "INVOICE_DELAY":
                if rec.get("is_unmatched") or not rec["invoice_overdue"]:
                    continue
            elif mode == "BOTH_DELAY":
                if rec.get("is_unmatched") or not (rec["actual_overdue"] and rec["invoice_overdue"]):
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
            if 0 <= col_index < len(self.col_configs_main):
                col_id = self.col_configs_main[col_index][0]
                self._open_column_filter_dialog(
                    col_id=col_id,
                    title=self.col_titles_main.get(col_id, col_id),
                    records=self.analyzed_records,
                    filter_dict=self.column_filters_main,
                    refresh_func=self._refresh_main_table,
                    update_heading_func=self._update_heading_display_main,
                    sort_func=lambda c, r: self._sort_tree(self.tree_main, c, r),
                    x=event.x_root,
                    y=event.y_root,
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

        for idx, rec in enumerate(self.missing_order_records):
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
                    records=self.missing_order_records,
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
    def _open_column_filter_dialog(
        self,
        col_id: str,
        title: str,
        records: list[dict],
        filter_dict: dict,
        refresh_func,
        update_heading_func,
        sort_func,
        x: int,
        y: int,
    ):
        if not records:
            return

        all_unique_vals = set()
        for rec in records:
            val = str(rec.get(col_id, "") or "").strip()
            if not val:
                val = "(空白)"
            all_unique_vals.add(val)

        sorted_vals = sorted(all_unique_vals, key=lambda s: (s == "(空白)", s))
        currently_selected = set(filter_dict.get(col_id, all_unique_vals))

        win = tk.Toplevel(self.root)
        win.title(f"フィルター: {title}")
        self._apply_window_icon(win)

        win_w, win_h = 320, 440
        scr_w = win.winfo_screenwidth()
        scr_h = win.winfo_screenheight()
        pos_x = min(max(10, x), scr_w - win_w - 20)
        pos_y = min(max(10, y), scr_h - win_h - 40)
        win.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        sort_frame = tk.Frame(win, bg="#f8fafc", padx=10, pady=8)
        sort_frame.pack(fill=tk.X)

        btn_sort_asc = tk.Button(
            sort_frame,
            text="↑ 昇順で並べ替え (A → Z / 0 → 9)",
            bg="#ffffff",
            fg="#1e293b",
            relief="solid",
            borderwidth=1,
            font=("Meiryo UI", 9),
            anchor="w",
            padx=8,
            pady=3,
            command=lambda: [sort_func(col_id, False), win.destroy()],
        )
        btn_sort_asc.pack(fill=tk.X, pady=(0, 4))

        btn_sort_desc = tk.Button(
            sort_frame,
            text="↓ 降順で並べ替え (Z → A / 9 → 0)",
            bg="#ffffff",
            fg="#1e293b",
            relief="solid",
            borderwidth=1,
            font=("Meiryo UI", 9),
            anchor="w",
            padx=8,
            pady=3,
            command=lambda: [sort_func(col_id, True), win.destroy()],
        )
        btn_sort_desc.pack(fill=tk.X)

        tk.Frame(win, height=1, bg="#cbd5e1").pack(fill=tk.X, padx=10, pady=(4, 6))

        search_frame = tk.Frame(win, bg="#f8fafc", padx=10)
        search_frame.pack(fill=tk.X)
        tk.Label(search_frame, text="検索:", bg="#f8fafc", font=("Meiryo UI", 8, "bold")).pack(anchor="w")
        filter_search_var = tk.StringVar()
        f_entry = ttk.Entry(search_frame, textvariable=filter_search_var)
        f_entry.pack(fill=tk.X, pady=(2, 6))

        btn_frame = tk.Frame(win, bg="#f1f5f9", padx=10, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

        chk_outer = tk.Frame(win, bg="#ffffff", highlightbackground="#cbd5e1", highlightthickness=1)
        chk_outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        canvas = tk.Canvas(chk_outer, bg="#ffffff", highlightthickness=0)
        vsb = ttk.Scrollbar(chk_outer, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#ffffff")

        window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def _on_canvas_configure(e):
            canvas.itemconfig(window_id, width=e.width)

        def _on_frame_configure(_e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", _on_canvas_configure)
        scrollable_frame.bind("<Configure>", _on_frame_configure)
        canvas.configure(yscrollcommand=vsb.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        val_vars = {}
        chk_widgets = []
        updating_all_var = False

        all_var = tk.BooleanVar(value=(len(currently_selected) == len(sorted_vals)))

        def _toggle_select_all():
            nonlocal updating_all_var
            if updating_all_var:
                return
            target_val = all_var.get()
            k = filter_search_var.get().strip().lower()
            for v, var in val_vars.items():
                if not k or k in v.lower():
                    var.set(target_val)

        chk_all = tk.Checkbutton(
            scrollable_frame,
            text="(すべて選択)",
            variable=all_var,
            command=_toggle_select_all,
            bg="#ffffff",
            activebackground="#ffffff",
            font=("Meiryo UI", 9, "bold"),
            anchor="w",
            padx=4,
            pady=2,
        )
        chk_all.pack(fill=tk.X, padx=4, pady=(2, 2))

        def _on_single_check_changed():
            nonlocal updating_all_var
            updating_all_var = True
            all_checked = all(var.get() for var in val_vars.values())
            all_var.set(all_checked)
            updating_all_var = False

        for val in sorted_vals:
            var = tk.BooleanVar(value=(val in currently_selected))
            val_vars[val] = var
            chk = tk.Checkbutton(
                scrollable_frame,
                text=val,
                variable=var,
                command=_on_single_check_changed,
                bg="#ffffff",
                activebackground="#ffffff",
                font=("Meiryo UI", 9),
                anchor="w",
                padx=16,
                pady=1,
            )
            chk.pack(fill=tk.X, pady=1)
            chk_widgets.append((val, chk))

        def _on_filter_search(*_):
            k = filter_search_var.get().strip().lower()
            for val, w in chk_widgets:
                if not k or k in val.lower():
                    w.pack(fill=tk.X, pady=1)
                else:
                    w.pack_forget()
            canvas.yview_moveto(0)

        filter_search_var.trace_add("write", _on_filter_search)

        def _cleanup_and_destroy():
            canvas.unbind_all("<MouseWheel>")
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _cleanup_and_destroy)

        def _apply_filter():
            selected = {v for v, var in val_vars.items() if var.get()}
            if len(selected) == len(sorted_vals):
                filter_dict.pop(col_id, None)
            else:
                filter_dict[col_id] = selected
            update_heading_func(col_id)
            refresh_func()
            _cleanup_and_destroy()

        def _clear_filter():
            filter_dict.pop(col_id, None)
            update_heading_func(col_id)
            refresh_func()
            _cleanup_and_destroy()

        btn_ok = tk.Button(
            btn_frame,
            text="適用 (OK)",
            bg="#2563eb",
            fg="#ffffff",
            font=("Meiryo UI", 9, "bold"),
            relief="flat",
            padx=12,
            pady=4,
            command=_apply_filter,
        )
        btn_ok.pack(side=tk.RIGHT, padx=(6, 0))

        btn_cancel = tk.Button(
            btn_frame,
            text="キャンセル",
            bg="#ffffff",
            fg="#475569",
            font=("Meiryo UI", 9),
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=4,
            command=_cleanup_and_destroy,
        )
        btn_cancel.pack(side=tk.RIGHT)

        btn_clear = tk.Button(
            btn_frame,
            text="フィルター解除",
            bg="#ffffff",
            fg="#b91c1c",
            font=("Meiryo UI", 9),
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=4,
            command=_clear_filter,
        )
        btn_clear.pack(side=tk.LEFT)

    def _sort_tree(self, tree: ttk.Treeview, col: str, reverse: bool):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, "", index)

    # =========================================================================
    # Excel直接展開機能（一時ファイル非作成・新規ブック展開）
    # =========================================================================
    @staticmethod
    def _rgb_to_bgr_int(hex_color: str) -> int:
        """#RRGGBB 形式の文字列を Excel COM の BGR 整数に変換"""
        hex_color = hex_color.lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (b << 16) | (g << 8) | r

    @staticmethod
    def _configure_excel_page_setup(ws, excel, sheet_title: str):
        """
        業務帳票として最適なExcel印刷設定（A4横、横幅1ページ収まり、見出し行繰り返し等）を適用
        """
        try:
            ps = ws.PageSetup
            # 1. 用紙方向: 横向き (xlLandscape = 2)
            ps.Orientation = 2
            # 2. 用紙サイズ: A4 (xlPaperA4 = 9)
            ps.PaperSize = 9
            # 3. 横幅1ページに収める（縦は行数に合わせて自然に流れる）
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            # 4. タイトル行の繰り返し（2ページ目以降も見出し1行目を印刷）
            ps.PrintTitleRows = "$1:$1"
            # 5. グリッド線（枠線）の印刷
            ps.PrintGridlines = True
            # 6. 水平中央揃え
            ps.CenterHorizontally = True
            # 7. 余白設定（狭めにしてデータをゆったり印刷）
            try:
                ps.LeftMargin = excel.InchesToPoints(0.5)
                ps.RightMargin = excel.InchesToPoints(0.5)
                ps.TopMargin = excel.InchesToPoints(0.6)
                ps.BottomMargin = excel.InchesToPoints(0.6)
                ps.HeaderMargin = excel.InchesToPoints(0.3)
                ps.FooterMargin = excel.InchesToPoints(0.3)
            except Exception:
                pass
            # 8. ヘッダー・フッター
            ps.LeftHeader = f"&\"Meiryo UI,Bold\"StockMatch_Miya - {sheet_title}"
            ps.RightHeader = "&\"Meiryo UI\"作成日: &D"
            ps.LeftFooter = "&\"Meiryo UI\"&A"
            ps.CenterFooter = "&\"Meiryo UI\"&P / &N ページ"
            ps.RightFooter = ""
        except Exception:
            pass

    def _open_in_excel_direct(self):
        """
        一時ファイルを一切作成せず、ExcelのCOMオートメーション経由で
        新規ブック（未保存）を作成し、書式・フィルター・判定色を再現して直接展開する
        """
        if not self.analyzed_records and not self.missing_order_records:
            messagebox.showinfo("データなし", "先に「照合・遅延チェック実行」を行ってください。")
            return

        if not HAS_WIN32COM:
            messagebox.showerror(
                "モジュール未検出",
                "Excelを直接起動するには 'pywin32' パッケージが必要です。\n"
                "ターミナルまたはコマンドプロンプトで以下を実行してください:\n"
                "pip install pywin32",
            )
            return

        self.excel_open_btn.configure(state=tk.DISABLED, text="Excel生成中...")
        self.status_var.set("Excelを起動し、新規ブックにデータを転送しています（一時ファイルなし）...")

        def _excel_worker():
            try:
                pythoncom.CoInitialize()

                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                except Exception:
                    excel = win32com.client.Dispatch("Excel.Application")

                excel.Visible = True
                excel.ScreenUpdating = False
                excel.DisplayAlerts = False

                # 新規ブックの作成（一時ファイルはディスク上に作成されない）
                wb = excel.Workbooks.Add()

                # シートの準備（2シート構成）
                while wb.Sheets.Count < 2:
                    wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))

                ws1 = wb.Sheets(1)
                ws1.Name = "納期超過・照合結果"

                ws2 = wb.Sheets(2)
                ws2.Name = "注文書未検出(管理表のみ)"

                # カラー定義（BGR int）
                header_bg = self._rgb_to_bgr_int("1E293B")
                header_fg = self._rgb_to_bgr_int("FFFFFF")
                border_color = self._rgb_to_bgr_int("CBD5E1")

                color_map = {
                    "UNMATCHED": {
                        "bg": self._rgb_to_bgr_int("EDE9FE"),
                        "fg": self._rgb_to_bgr_int("4C1D95"),
                    },
                    "BOTH_DELAY": {
                        "bg": self._rgb_to_bgr_int("FEE2E2"),
                        "fg": self._rgb_to_bgr_int("7F1D1D"),
                    },
                    "ACTUAL_DELAY": {
                        "bg": self._rgb_to_bgr_int("FFEDD5"),
                        "fg": self._rgb_to_bgr_int("7C2D12"),
                    },
                    "INVOICE_DELAY": {
                        "bg": self._rgb_to_bgr_int("FEF9C3"),
                        "fg": self._rgb_to_bgr_int("713F12"),
                    },
                    "OK": {
                        "bg": self._rgb_to_bgr_int("FFFFFF"),
                        "fg": self._rgb_to_bgr_int("334155"),
                    },
                    "MISSING_ORDER": {
                        "bg": self._rgb_to_bgr_int("FEF3C7"),
                        "fg": self._rgb_to_bgr_int("92400E"),
                    },
                }

                # -------------------------------------------------------------
                # 1. シート1: 納期超過・照合結果
                # -------------------------------------------------------------
                headers_main = [
                    "総合判定",
                    "不一致理由・照合診断",
                    "依頼区分",
                    "管理No(注文番号)",
                    "発注先",
                    "メーカー",
                    "品名",
                    "型番・規格(図番)",
                    "数量",
                    "注文書納期",
                    "基準日超過",
                    "現品確認日(物)",
                    "現品状況",
                    "請求日(伝票)",
                    "請求状況",
                    "担当",
                    "照合元情報",
                ]

                rows_main_data = []
                tags_main = []
                for r in self.analyzed_records:
                    rows_main_data.append(
                        [
                            r["status_badge"],
                            r["mismatch_reason"],
                            r["req_type"],
                            r["order_no"],
                            r["supplier"],
                            r["maker"],
                            r["item_name"],
                            r["model_spec"],
                            r["qty"],
                            r["due_date_str"],
                            r["days_overdue_str"],
                            r["actual_date_str"],
                            r["actual_status"],
                            r["invoice_date_str"],
                            r["invoice_status"],
                            r["person"],
                            r["source_info"],
                        ]
                    )
                    tags_main.append(r["row_tag"])

                num_main_rows = len(rows_main_data)
                num_main_cols = len(headers_main)

                # ヘッダー書き込み
                ws1.Range(ws1.Cells(1, 1), ws1.Cells(1, num_main_cols)).Value = headers_main
                h_rng1 = ws1.Range(ws1.Cells(1, 1), ws1.Cells(1, num_main_cols))
                h_rng1.Font.Bold = True
                h_rng1.Font.Name = "Meiryo UI"
                h_rng1.Font.Size = 10
                h_rng1.Font.Color = header_fg
                h_rng1.Interior.Color = header_bg
                h_rng1.RowHeight = 26
                h_rng1.VerticalAlignment = -4108  # xlCenter

                if num_main_rows > 0:
                    data_rng1 = ws1.Range(ws1.Cells(2, 1), ws1.Cells(num_main_rows + 1, num_main_cols))
                    data_rng1.Value = rows_main_data
                    data_rng1.Font.Name = "Meiryo UI"
                    data_rng1.Font.Size = 9
                    data_rng1.RowHeight = 22
                    data_rng1.VerticalAlignment = -4108  # xlCenter

                    # 行ごとのハイライト適用
                    for idx, tag in enumerate(tags_main, start=2):
                        if tag in color_map:
                            row_rng = ws1.Range(ws1.Cells(idx, 1), ws1.Cells(idx, num_main_cols))
                            row_rng.Interior.Color = color_map[tag]["bg"]
                            row_rng.Font.Color = color_map[tag]["fg"]

                    # 罫線の設定
                    all_rng1 = ws1.Range(ws1.Cells(1, 1), ws1.Cells(num_main_rows + 1, num_main_cols))
                    for border_id in [7, 8, 9, 10, 11, 12]:  # xlEdgeLeft, Right, Top, Bottom, InsideV, InsideH
                        try:
                            all_rng1.Borders(border_id).LineStyle = 1  # xlContinuous
                            all_rng1.Borders(border_id).Color = border_color
                        except Exception:
                            pass

                    # オートフィルター設定
                    ws1.Range(ws1.Cells(1, 1), ws1.Cells(num_main_rows + 1, num_main_cols)).AutoFilter(Field=1)
                else:
                    ws1.Range(ws1.Cells(1, 1), ws1.Cells(1, num_main_cols)).AutoFilter(Field=1)

                ws1.Columns.AutoFit()

                # シート1の見出し（1行目）をスクロール時も隠れないようにウィンドウ枠固定
                ws1.Activate()
                excel.ActiveWindow.SplitRow = 1
                excel.ActiveWindow.SplitColumn = 0
                ws1.Range("A2").Select()
                excel.ActiveWindow.FreezePanes = True

                # シート1の印刷設定（A4横・横幅1ページ自動収まり・見出し行繰り返し等）
                self._configure_excel_page_setup(ws1, excel, "納期超過・照合結果")

                # -------------------------------------------------------------
                # 2. シート2: 注文書未検出(管理表のみ)
                # -------------------------------------------------------------
                headers_missing = [
                    "状態判定",
                    "管理表シート",
                    "行番号",
                    "管理No(8桁)",
                    "業者 / 発注先",
                    "メーカー",
                    "品名",
                    "型番・規格・図番",
                    "購入数",
                    "管理表納期",
                    "現品確認日",
                    "請求日",
                    "担当",
                    "未検出診断内容",
                ]

                rows_missing_data = []
                for r in self.missing_order_records:
                    rows_missing_data.append(
                        [
                            r["status_badge"],
                            r["mgmt_sheet"],
                            r["mgmt_row"],
                            r["order_no"],
                            r["supplier"],
                            r["maker"],
                            r["item_name"],
                            r["model_spec"],
                            r["qty"],
                            r["due_date"],
                            r["actual_date"],
                            r["invoice_date"],
                            r["person"],
                            r["mismatch_reason"],
                        ]
                    )

                num_missing_rows = len(rows_missing_data)
                num_missing_cols = len(headers_missing)

                ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, num_missing_cols)).Value = headers_missing
                h_rng2 = ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, num_missing_cols))
                h_rng2.Font.Bold = True
                h_rng2.Font.Name = "Meiryo UI"
                h_rng2.Font.Size = 10
                h_rng2.Font.Color = header_fg
                h_rng2.Interior.Color = header_bg
                h_rng2.RowHeight = 26
                h_rng2.VerticalAlignment = -4108  # xlCenter

                if num_missing_rows > 0:
                    data_rng2 = ws2.Range(ws2.Cells(2, 1), ws2.Cells(num_missing_rows + 1, num_missing_cols))
                    data_rng2.Value = rows_missing_data
                    data_rng2.Font.Name = "Meiryo UI"
                    data_rng2.Font.Size = 9
                    data_rng2.RowHeight = 22
                    data_rng2.VerticalAlignment = -4108  # xlCenter

                    # 未発注行の配色
                    m_bg = color_map["MISSING_ORDER"]["bg"]
                    m_fg = color_map["MISSING_ORDER"]["fg"]
                    for idx in range(2, num_missing_rows + 2):
                        row_rng = ws2.Range(ws2.Cells(idx, 1), ws2.Cells(idx, num_missing_cols))
                        row_rng.Interior.Color = m_bg
                        row_rng.Font.Color = m_fg

                    # 罫線の設定
                    all_rng2 = ws2.Range(ws2.Cells(1, 1), ws2.Cells(num_missing_rows + 1, num_missing_cols))
                    for border_id in [7, 8, 9, 10, 11, 12]:
                        try:
                            all_rng2.Borders(border_id).LineStyle = 1
                            all_rng2.Borders(border_id).Color = border_color
                        except Exception:
                            pass

                    ws2.Range(ws2.Cells(1, 1), ws2.Cells(num_missing_rows + 1, num_missing_cols)).AutoFilter(Field=1)
                else:
                    ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, num_missing_cols)).AutoFilter(Field=1)

                ws2.Columns.AutoFit()

                # シート2の見出し（1行目）をスクロール時も隠れないようにウィンドウ枠固定
                ws2.Activate()
                excel.ActiveWindow.SplitRow = 1
                excel.ActiveWindow.SplitColumn = 0
                ws2.Range("A2").Select()
                excel.ActiveWindow.FreezePanes = True

                # シート2の印刷設定（A4横・横幅1ページ自動収まり・見出し行繰り返し等）
                self._configure_excel_page_setup(ws2, excel, "注文書未検出(管理表のみ)")

                # 最終的に第1シートをアクティブにし、A1セルを選択
                ws1.Activate()
                ws1.Range("A1").Select()

                excel.ScreenUpdating = True
                excel.DisplayAlerts = True

                # Excelウィンドウを最前面に表示
                try:
                    excel.WindowState = -4137  # xlMaximized
                except Exception:
                    pass

                self.ui_queue.put(("EXCEL_OPENED_SUCCESS", None))

            except Exception as e:
                self.ui_queue.put(("EXCEL_OPENED_ERROR", str(e)))
            finally:
                pythoncom.CoUninitialize()

        worker = threading.Thread(target=_excel_worker, daemon=True)
        worker.start()
        self.root.after(50, self._poll_excel_queue)

    def _poll_excel_queue(self):
        try:
            while True:
                msg_type, payload = self.ui_queue.get_nowait()
                if msg_type == "EXCEL_OPENED_SUCCESS":
                    self.excel_open_btn.configure(state=tk.NORMAL, text="📊 Excelで直接開く")
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
                elif msg_type == "EXCEL_OPENED_ERROR":
                    self.excel_open_btn.configure(state=tk.NORMAL, text="📊 Excelで直接開く")
                    self.status_var.set(f"Excel展開エラー: {payload}")
                    messagebox.showerror("Excel展開エラー", f"Excelの起動またはデータ転送中にエラーが発生しました:\n{payload}")
                    return
                else:
                    self.ui_queue.put((msg_type, payload))
                    break
        except queue.Empty:
            pass

        self.root.after(50, self._poll_excel_queue)


def main():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("StockMatch.Miya.v1.1.0")
    except Exception:
        pass

    root = tk.Tk()
    _app = StockMatchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
