import os
import re
import json
import unicodedata
import datetime
import difflib
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl


class StockMatchApp:
    """
    StockMatch_Miya-v1.0.0
    注文品管理表（1つのファイル）と注文書フォルダ内（サブフォルダ含む全階層）の全注文書Excelを照合し、
    今日（判定基準日）時点で「現品確認日」および「請求日」の納期超過を検査・可視化するデスクトップアプリ
    """

    # 厳格な固定保存先定義（ここ以外への保存・上書きは一切遮断・不許可）
    ALLOWED_CONFIG_DIR = os.path.abspath(r"C:\StockMatch")
    CONFIG_FILE_PATH = os.path.abspath(r"C:\StockMatch\config.json")

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("StockMatch_Miya-v1.0.0 - 注文品納期・現品/請求照合システム")

        # 起動時にウィンドウを最大化表示
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")

        self.root.minsize(1060, 720)

        # 状態保持用変数
        self.mgmt_file_path = tk.StringVar()
        self.order_dir_path = tk.StringVar()
        self.base_date_str = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.filter_mode = tk.StringVar(value="ALL_DELAY")
        self.search_keyword = tk.StringVar()

        # 進捗表示・非同期スレッド管理用
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_text_var = tk.StringVar(value="待機中")
        self.is_running = False
        self.ui_queue = queue.Queue()

        # 解析結果データ保持リスト
        self.analyzed_records = []
        # 各列の個別フィルター保持用辞書 {col_id: set_of_selected_values}
        self.column_filters = {}

        self._setup_styles()
        self._build_menu()
        self._build_ui()
        self._load_config_and_autoload()

    def _setup_styles(self):
        """UIテーマ・スタイルの設定"""
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

    def _build_menu(self):
        """上部メニューバー（ファイル / 設定 / ヘルプ）の作成"""
        menubar = tk.Menu(self.root)

        # 1. ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="注文品管理表ファイルを選択...", command=self._select_mgmt_file, accelerator="Ctrl+O")
        file_menu.add_command(label="注文書フォルダを選択...", command=self._select_order_dir, accelerator="Ctrl+D")
        file_menu.add_separator()
        file_menu.add_command(label="照合・遅延チェックを実行", command=self.run_analysis, accelerator="F5")
        file_menu.add_command(label="CSV/Excel出力...", command=self._export_results, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit, accelerator="Alt+F4")
        menubar.add_cascade(label="ファイル(F)", menu=file_menu)

        # 2. 設定メニュー
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="現在のパス設定を保存 (C:\\StockMatch専用)", command=self._save_config_manual)
        config_menu.add_command(label="保存先フォルダを開く (C:\\StockMatch)", command=self._open_config_folder)
        menubar.add_cascade(label="設定(S)", menu=config_menu)

        # 3. ヘルプメニュー
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使い方・操作ガイド(H)...", command=self._show_help_guide, accelerator="F1")
        help_menu.add_command(label="照合判定ロジックについて...", command=self._show_logic_guide)
        help_menu.add_separator()
        help_menu.add_command(label="バージョン情報...", command=self._show_about)
        menubar.add_cascade(label="ヘルプ(H)", menu=help_menu)

        self.root.config(menu=menubar)

        # ショートカットキー割り当て
        self.root.bind_all("<F1>", lambda e: self._show_help_guide())
        self.root.bind_all("<F5>", lambda e: self.run_analysis())
        self.root.bind_all("<Control-o>", lambda e: self._select_mgmt_file())
        self.root.bind_all("<Control-d>", lambda e: self._select_order_dir())
        self.root.bind_all("<Control-s>", lambda e: self._export_results())

    def _build_ui(self):
        """メイン画面レイアウトの構築"""
        self.root.configure(bg="#f8fafc")

        # 1. トップヘッダーバー
        header_frame = tk.Frame(self.root, bg="#1e293b", height=56, padx=16, pady=10)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="StockMatch_Miya-v1.0.0", style="Header.TLabel")
        title_lbl.pack(side=tk.LEFT)

        sub_lbl = ttk.Label(
            header_frame,
            text="  |  注文書納期 × 管理表（現品確認日・請求日）超過照合システム",
            style="SubHeader.TLabel",
        )
        sub_lbl.pack(side=tk.LEFT, padx=(8, 0))

        help_btn_header = tk.Button(
            header_frame,
            text="？ ヘルプ / ガイド",
            bg="#334155",
            fg="#f8fafc",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief="flat",
            font=("Meiryo UI", 9),
            padx=10,
            pady=3,
            command=self._show_help_guide,
        )
        help_btn_header.pack(side=tk.RIGHT)

        # 2. コントロールパネル
        ctrl_container = ttk.Frame(self.root, padding=(16, 10))
        ctrl_container.pack(fill=tk.X)

        file_card = ttk.Frame(ctrl_container, style="Card.TFrame", padding=12)
        file_card.pack(fill=tk.X)

        # ① 注文品管理表ファイル選択行
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

        # ② 注文書フォルダ選択行（サブフォルダ含む全階層を走査）
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
        ttk.Button(row2, text="パス設定を保存", style="Action.TButton", command=self._save_config_manual).pack(
            side=tk.LEFT
        )

        # ③ 基準日・実行ボタン行
        row3 = ttk.Frame(file_card, style="Card.TFrame")
        row3.pack(fill=tk.X, pady=(2, 6))

        ttk.Label(
            row3,
            text="③ 判定基準日 (今日):",
            width=22,
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)
        date_entry = ttk.Entry(row3, textvariable=self.base_date_str, width=16, font=("Meiryo UI", 9))
        date_entry.pack(side=tk.LEFT, padx=(4, 8))

        ttk.Button(
            row3,
            text="今日に戻す",
            style="Action.TButton",
            command=lambda: self.base_date_str.set(datetime.date.today().strftime("%Y-%m-%d")),
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(
            row3,
            text="※ファイル名先頭8桁の管理Noが管理表と一致しない項目は【不一致】管理表未記載として明示されます。",
            style="Card.TLabel",
            foreground="#64748b",
        ).pack(side=tk.LEFT)

        self.export_btn = ttk.Button(
            row3, text="CSV/Excel出力", style="Action.TButton", command=self._export_results
        )
        self.export_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self.run_btn = ttk.Button(
            row3, text="照合・遅延チェック実行", style="Primary.TButton", command=self.run_analysis
        )
        self.run_btn.pack(side=tk.RIGHT)

        # ④ 進捗バー＆実行状況表示
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
        kpi_frame = ttk.Frame(self.root, padding=(16, 0, 16, 10))
        kpi_frame.pack(fill=tk.X)

        self.kpi_vars = {
            "total": tk.StringVar(value="0 件"),
            "unmatched": tk.StringVar(value="0 件"),
            "both_delay": tk.StringVar(value="0 件"),
            "actual_delay": tk.StringVar(value="0 件"),
            "invoice_delay": tk.StringVar(value="0 件"),
            "ok": tk.StringVar(value="0 件"),
        }

        self._create_kpi_card(kpi_frame, "検査対象明細 総数", self.kpi_vars["total"], "#334155", "#f1f5f9").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6)
        )
        self._create_kpi_card(
            kpi_frame, "【不一致】管理表未記載", self.kpi_vars["unmatched"], "#4c1d95", "#ede9fe"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(
            kpi_frame, "【重要】現品・請求 共に超過", self.kpi_vars["both_delay"], "#991b1b", "#fee2e2"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(
            kpi_frame, "【現品超過】現品確認 遅れ", self.kpi_vars["actual_delay"], "#9a3412", "#ffedd5"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(
            kpi_frame, "【請求超過】伝票(請求) 遅れ", self.kpi_vars["invoice_delay"], "#854d0e", "#fef9c3"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._create_kpi_card(kpi_frame, "正常・納期前・確認済", self.kpi_vars["ok"], "#166534", "#dcfce7").pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )

        # 4. フィルタ・検索バー
        filter_frame = ttk.Frame(self.root, padding=(16, 0, 16, 8))
        filter_frame.pack(fill=tk.X)

        ttk.Label(filter_frame, text="表示フィルタ:", font=("Meiryo UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))

        modes = [
            ("すべての異常(超過・不一致)を表示", "ALL_DELAY"),
            ("管理表との不一致のみ", "UNMATCHED_ONLY"),
            ("現品確認日の超過のみ", "ACTUAL_DELAY"),
            ("請求日(伝票)の超過のみ", "INVOICE_DELAY"),
            ("現品・請求の両方超過", "BOTH_DELAY"),
            ("全件表示（正常含む）", "ALL"),
        ]
        for text, mode_val in modes:
            rb = ttk.Radiobutton(
                filter_frame,
                text=text,
                variable=self.filter_mode,
                value=mode_val,
                command=self._refresh_table,
            )
            rb.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(filter_frame, text="キーワード検索:").pack(side=tk.LEFT, padx=(12, 4))
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_keyword, width=22)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind("<KeyRelease>", lambda e: self._refresh_table())

        self.clear_col_filters_btn = ttk.Button(
            filter_frame,
            text="列フィルター解除 (Excel風)",
            style="Action.TButton",
            command=self._clear_all_column_filters,
        )
        self.clear_col_filters_btn.pack(side=tk.RIGHT, padx=(8, 0))

        # 5. メイン結果テーブル（Treeview）
        table_container = ttk.Frame(self.root, padding=(16, 0, 16, 12))
        table_container.pack(fill=tk.BOTH, expand=True)

        columns = (
            "status_badge",
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

        self.tree = ttk.Treeview(table_container, columns=columns, show="headings", selectmode="browse")

        self.col_configs = [
            ("status_badge", "総合判定", 165, "center"),
            ("req_type", "依頼区分", 85, "center"),
            ("order_no", "管理No(注文番号)", 120, "center"),
            ("supplier", "発注先", 140, "w"),
            ("maker", "メーカー", 110, "w"),
            ("item_name", "品名", 150, "w"),
            ("model_spec", "型番・規格(図番)", 165, "w"),
            ("qty", "数量", 65, "center"),
            ("due_date", "注文書納期", 95, "center"),
            ("days_overdue", "基準日超過", 95, "center"),
            ("actual_date", "現品確認日(物)", 115, "center"),
            ("actual_status", "現品状況", 140, "center"),
            ("invoice_date", "請求日(伝票)", 115, "center"),
            ("invoice_status", "請求状況", 140, "center"),
            ("person", "担当", 75, "center"),
            ("source_info", "照合元情報", 180, "w"),
        ]

        self.col_titles = {col_id: title for col_id, title, _, _ in self.col_configs}

        for col_id, heading_text, width, anchor in self.col_configs:
            self._update_heading_display(col_id)
            self.tree.column(col_id, width=width, anchor=anchor)

        self.tree.bind("<Button-1>", self._on_tree_click)

        # 行の背景色タグ設定
        self.tree.tag_configure("UNMATCHED", background="#ede9fe", foreground="#4c1d95")  # 紫系（不一致）
        self.tree.tag_configure("BOTH_DELAY", background="#fee2e2", foreground="#7f1d1d")
        self.tree.tag_configure("ACTUAL_DELAY", background="#ffedd5", foreground="#7c2d12")
        self.tree.tag_configure("INVOICE_DELAY", background="#fef9c3", foreground="#713f12")
        self.tree.tag_configure("OK", background="#ffffff", foreground="#334155")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        # 6. ステータスバー
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

    def _create_kpi_card(
        self, parent, title: str, value_var: tk.StringVar, fg_color: str, bg_color: str
    ) -> tk.Frame:
        """KPIサマリー表示用カードの生成"""
        card = tk.Frame(parent, bg=bg_color, highlightbackground="#cbd5e1", highlightthickness=1, padx=10, pady=8)
        lbl_title = tk.Label(card, text=title, bg=bg_color, fg=fg_color, font=("Meiryo UI", 8, "bold"), anchor="w")
        lbl_title.pack(fill=tk.X)
        lbl_val = tk.Label(
            card, textvariable=value_var, bg=bg_color, fg=fg_color, font=("Meiryo UI", 15, "bold"), anchor="e"
        )
        lbl_val.pack(fill=tk.X, pady=(4, 0))
        return card

    # =========================================================================
    # ヘルプ・ガイドダイアログ
    # =========================================================================
    def _show_help_guide(self):
        """使い方・操作ガイドのモーダルウィンドウを表示"""
        win = tk.Toplevel(self.root)
        win.title("StockMatch_Miya-v1.0.0 - 使い方・操作ガイド")
        win.geometry("750x620")
        win.minsize(580, 450)
        win.transient(self.root)

        txt_frame = ttk.Frame(win, padding=12)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(txt_frame, wrap="word", font=("Meiryo UI", 9), padx=10, pady=10, relief="solid", borderwidth=1)
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        help_content = """【基本操作手順】
1. 『① 注文品管理表ファイル』の指定:
   - 「ファイルを参照...」から管理表Excel（例: 注文品管理表.xlsx）を選択します。

2. 『② 注文書フォルダパス』の指定:
   - 「フォルダを参照...」から注文書Excelが保管されている親フォルダを選択します。
   - 【サブフォルダ対応】指定フォルダ直下だけでなく、配下にあるすべてのサブフォルダ内も再帰的に走査し、すべての注文書Excelを自動検出・一括検査します。
   - 注文書以外の無関係なファイル（PDF、テキスト、一時ファイル、注文書フォーマットではないExcel等）は自動的にスキップされます。

3. 『管理No（8桁）による厳格な一致照合ルール』:
   - 注文書のファイル名先頭8桁が「管理No（注文番号）」となります（例: 22609030購依〇〇注文.xlsx ⇒ 管理No: 22609030）。
   - 注文品管理表の中に、この管理Noと一致する項目が存在しない場合は、無条件で「【不一致】管理表未記載」と判定されます。
   - 管理Noが一致する行の中から、図面番号・品番等によって該当明細を特定します。

4. 『パス設定を保存』ボタン:
   - ボタンを押すと、現在の管理表パスと注文書フォルダパスが C:\\StockMatch\\config.json に保存されます。
   - 【厳重ガード】本アプリの設定ファイルは「C:\\StockMatch\\config.json」専用に厳重に固定されており、これ以外のフォルダへ保存・上書きされることは一切ありません。
   - 次回起動時に自動復元されるため、毎回の選択が不要になります。

5. 『判定基準日』の確認:
   - デフォルトで本日の日付が入っています。過去や未来の納期超過を検証したい場合は日付を打ち替えて「照合・遅延チェック実行」を押してください。

6. 『照合・遅延チェック実行』ボタン (F5キー):
   - 照合計算を開始します。進捗バーで現在の状況がリアルタイムに表示されます。

--------------------------------------------------
【画面の見方と操作】
■ 不一致の表示（管理表未記載）:
  - 管理Noが管理表に存在しない場合や明細が見当たらない場合は「【不一致】管理表未記載」として紫色で強調表示されます。

■ 表示フィルタ / キーワード検索:
  - 「すべての異常(超過・不一致)を表示」や「管理表との不一致のみ」等でワンクリック切り替えが可能です。

■ Excel風列フィルター機能:
  - 各列のヘッダーをクリックすると、Excelと同じ「昇順・降順ソート」「値の検索」「個別チェックボックス絞り込み」が利用できます。

■ CSV/Excel出力 (Ctrl+Sキー):
  - 照合・超過判定結果の一覧を .xlsx または .csv 形式で出力・保存できます。
"""
        txt.insert("1.0", help_content)
        txt.configure(state="disabled")

        btn_box = ttk.Frame(win, padding=(12, 8))
        btn_box.pack(fill=tk.X)
        ttk.Button(btn_box, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)

    def _show_logic_guide(self):
        """照合判定ロジック詳細ガイドを表示"""
        win = tk.Toplevel(self.root)
        win.title("照合・遅延判定ロジックの詳細")
        win.geometry("700x560")
        win.minsize(560, 420)
        win.transient(self.root)

        txt_frame = ttk.Frame(win, padding=12)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(txt_frame, wrap="word", font=("Meiryo UI", 9), padx=10, pady=10, relief="solid", borderwidth=1)
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        logic_content = """【照合・遅延判定ロジック】

1. 管理No（ファイル名先頭8桁）の厳格判定:
   - 注文書ファイル名の先頭8桁から管理Noを抽出します（例: 22609030）。
   - 注文品管理表側の「管理No / 注文番号」列（またはシート名）にこの管理Noと一致するデータが存在しない場合、直ちに「【不一致】管理表未記載」と判定します。

2. 依頼区分の切り分け（購入依頼 vs 外注依頼）:
   - ファイル名に「購依」が含まれる場合は【購入依頼】として管理表の「購入品リスト」系シートを照合します。
   - ファイル名に「外依」が含まれる場合は【外注依頼】として管理表の「外注加工リスト」系シートを照合します。

3. 明細行の特定優先ルール:
   - 管理Noが一致する管理表データの中から、注文書の「図面番号」または「品番」（スペース・記号を除去した英数字）で1対1特定を行います。
   - 同一型番が複数ある場合は数量の一致度をタイブレークとして判定します。
   - 該当する行が存在しない場合は「【不一致】管理表未記載」と判定します。

4. 現品確認日（物：A列）の判定:
   - 管理表のA列（1列目）に日付や確認済み文字列が記入されている場合:
     ⇒ 「確認済 (正常)」と判定（現品超過フラグはOFF）。
   - 管理表のA列が未記入の場合:
     ⇒ 判定基準日（今日）が注文書の納期を過ぎていれば「未確認 (超過X日)」と判定。

5. 請求日（伝票：AI列）の判定:
   - 管理表のAI列（35列目）に日付や確認済み文字列が記入されている場合:
     ⇒ 「確認済 (正常)」と判定（請求超過フラグはOFF）。
   - 管理表のAI列が未記入の場合:
     ⇒ 判定基準日（今日）が注文書の納期を過ぎていれば「未確認 (超過X日)」と判定。
"""
        txt.insert("1.0", logic_content)
        txt.configure(state="disabled")

        btn_box = ttk.Frame(win, padding=(12, 8))
        btn_box.pack(fill=tk.X)
        ttk.Button(btn_box, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)

    def _show_about(self):
        """バージョン情報ダイアログの表示"""
        messagebox.showinfo(
            "バージョン情報",
            "StockMatch_Miya-v1.0.0\n"
            "注文品納期 × 管理表（現品確認日・請求日）超過照合システム\n\n"
            "照合規則: ファイル名先頭8桁の管理No完全照合\n"
            "対応形式: Excel (.xlsx, .xlsm) / CSV\n"
            "探索範囲: 指定フォルダおよび全サブフォルダ\n"
            "設定保存先: C:\\StockMatch\\config.json 専用\n\n"
            "© StockMatch_Miya System All Rights Reserved.",
        )

    def _open_config_folder(self):
        """設定ファイル保存先フォルダ（C:\StockMatch）をエクスプローラーで開く"""
        target_dir = self.ALLOWED_CONFIG_DIR
        if not os.path.exists(target_dir):
            try:
                os.makedirs(target_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("フォルダ参照エラー", f"C:\\StockMatch フォルダを作成できませんでした:\n{e}")
                return
        try:
            os.startfile(target_dir)
        except Exception as e:
            messagebox.showinfo("フォルダパス", f"設定ファイル保管場所:\n{target_dir}\n\n({e})")

    # =========================================================================
    # 設定ファイル（C:\StockMatch\config.json 専用・他フォルダ保存厳重遮断ガード）
    # =========================================================================
    def _is_strictly_allowed_path(self, target_filepath: str) -> bool:
        """対象パスが厳密に 'C:\\StockMatch\\config.json' であるかを検証"""
        try:
            norm_target = os.path.normcase(os.path.realpath(os.path.abspath(target_filepath)))
            norm_allowed = os.path.normcase(os.path.realpath(os.path.abspath(self.CONFIG_FILE_PATH)))
            return norm_target == norm_allowed
        except Exception:
            return False

    def _save_config(self, show_msg: bool = False):
        """指定された管理表ファイルパスと注文書フォルダパスを C:\\StockMatch\\config.json に保存"""
        mgmt = self.mgmt_file_path.get().strip().strip('"')
        order = self.order_dir_path.get().strip().strip('"')

        data = {
            "mgmt_file_path": mgmt,
            "order_dir_path": order,
            "last_updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        target_dir = self.ALLOWED_CONFIG_DIR
        target_file = self.CONFIG_FILE_PATH

        if not self._is_strictly_allowed_path(target_file):
            err_msg = f"セキュリティ違反: 許可されていない保存先への書き込みが要求されました。\n対象: {target_file}"
            if show_msg:
                messagebox.showerror("厳重ブロック", err_msg)
            return

        try:
            if not os.path.exists(target_dir):
                os.makedirs(target_dir, exist_ok=True)
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    "保存エラー",
                    f"C:\\StockMatch フォルダを作成できませんでした:\n{e}",
                )
            return

        try:
            with open(target_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if show_msg:
                messagebox.showinfo("設定保存完了", f"パス設定を保存しました:\n{target_file}")
        except Exception as e:
            if show_msg:
                messagebox.showerror("保存エラー", f"C:\\StockMatch\\config.json への保存に失敗しました:\n{e}")

    def _save_config_manual(self):
        """「パス設定を保存」ボタン押下時の処理"""
        self._save_config(show_msg=True)

    def _load_config_and_autoload(self):
        """起動時に C:\\StockMatch\\config.json のみを読み込み"""
        cfg_path = self.CONFIG_FILE_PATH
        if os.path.exists(cfg_path) and self._is_strictly_allowed_path(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                mgmt = data.get("mgmt_file_path", "")
                order = data.get("order_dir_path", "")
                if mgmt and os.path.isfile(mgmt):
                    self.mgmt_file_path.set(mgmt)
                if order and os.path.isdir(order):
                    self.order_dir_path.set(order)
            except Exception:
                pass

        if not self.mgmt_file_path.get() or not self.order_dir_path.get():
            self._autoload_default_files()

    def _autoload_default_files(self):
        """カレントディレクトリにある管理表ファイルと注文書フォルダを自動補完"""
        cwd = os.getcwd()
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
        """① 注文品管理表ファイル（単一ファイル）の選択ダイアログ"""
        initial_dir = os.path.dirname(self.mgmt_file_path.get().strip().strip('"'))
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.getcwd()

        path = filedialog.askopenfilename(
            title="注文品管理表のExcelファイルを選択",
            initialdir=initial_dir,
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls"), ("All Files", "*.*")],
        )
        if path:
            self.mgmt_file_path.set(path)

    def _select_order_dir(self):
        """② 注文書フォルダの選択ダイアログ"""
        initial_dir = self.order_dir_path.get().strip().strip('"')
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.getcwd()
        folder = filedialog.askdirectory(title="注文書が入っているフォルダを選択", initialdir=initial_dir)
        if folder:
            self.order_dir_path.set(folder)

    def _report_progress(self, percent: float, message: str):
        """バックグラウンドスレッドから進捗バーとステータスメッセージを更新キューへ送る"""
        self.ui_queue.put(("PROGRESS", (percent, message)))

    def _collect_valid_order_files(
        self, folder_path: str, mgmt_filepath: str, progress_cb=None
    ) -> tuple[list[dict], int, int]:
        """注文書フォルダ内（サブフォルダ含む全階層）を再帰的に走査して有効な注文書明細を抽出"""
        order_items = []
        valid_file_count = 0
        ignored_file_count = 0

        mgmt_abspath = os.path.normcase(os.path.realpath(os.path.abspath(mgmt_filepath))) if mgmt_filepath else ""

        file_list = []
        for root_dir, dirs, files in os.walk(folder_path):
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

            try:
                extracted = self._parse_order_book(full_path)
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
    def _parse_flexible_date(val, default_year: int = None) -> datetime.date | None:
        """Excel内の多様な日付表現を安全に datetime.date に変換"""
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
        if not s or s in ("-", "ー", "未", "未定", "None", "nan", "NaT"):
            return None

        s = unicodedata.normalize("NFKC", s)
        for ch in ("/", "-", ".", "年", "月"):
            s = s.replace(ch, "-")
        s = s.replace("日", " ")
        s = re.sub(r"\(.*?\)", "", s).strip()

        m_full = re.search(r"(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})", s)
        if m_full:
            try:
                return datetime.date(int(m_full.group(1)), int(m_full.group(2)), int(m_full.group(3)))
            except ValueError:
                return None

        m_short_yr = re.search(r"^(\d{2})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})$", s)
        if m_short_yr:
            try:
                year = 2000 + int(m_short_yr.group(1))
                return datetime.date(year, int(m_short_yr.group(2)), int(m_short_yr.group(3)))
            except ValueError:
                return None

        m_md = re.search(r"^(\d{1,2})\s*-\s*(\d{1,2})", s)
        if m_md:
            year = default_year or datetime.date.today().year
            try:
                return datetime.date(year, int(m_md.group(1)), int(m_md.group(2)))
            except ValueError:
                return None

        m_8digit = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
        if m_8digit:
            try:
                return datetime.date(int(m_8digit.group(1)), int(m_8digit.group(2)), int(m_8digit.group(3)))
            except ValueError:
                return None

        return None

    @staticmethod
    def _is_filled_confirmation(raw_val) -> bool:
        """現品確認日や請求日のセルに記入があるか判定"""
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
        """照合比較用に文字列を正規化（スペース・中点・記号・法人格を除去）"""
        if not text or pd.isna(text):
            return ""
        s = str(text).strip()
        if s.lower() in ("none", "nan", "nat"):
            return ""
        s = unicodedata.normalize("NFKC", s)
        for corp in ("株式会社", "有限会社", "(株)", "(有)", "御中", "様"):
            s = s.replace(corp, "")

        s = re.sub(r"[^0-9A-Za-z一-龥ぁ-んァ-ン]", "", s)
        return s.upper()

    @classmethod
    def _make_bigrams(cls, norm_str: str) -> frozenset:
        """2文字N-gram集合を作成"""
        if len(norm_str) < 2:
            return frozenset()
        return frozenset(norm_str[i : i + 2] for i in range(len(norm_str) - 1))

    @classmethod
    def _fast_similarity(
        cls, na: str, bg_a: frozenset, nb: str, bg_b: frozenset
    ) -> float:
        """高速あいまい一致度計算"""
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
            if bigram_score >= 0.45:
                seq_ratio = difflib.SequenceMatcher(None, na, nb).quick_ratio()
                return max(bigram_score, seq_ratio)
            return bigram_score

        return difflib.SequenceMatcher(None, na, nb).quick_ratio()

    @staticmethod
    def _extract_Leading_number(text: str) -> str:
        """数値部分を抽出して正規化"""
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
    def _clean_cell_str(v) -> str:
        """セル値をクリーンな文字列に変換"""
        if v is None or pd.isna(v):
            return ""
        s = str(v).strip()
        if s.lower() in ("none", "nan", "nat"):
            return ""
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        return s

    # =========================================================================
    # Excel解析ロジック（①注文品管理表 & ②注文書）
    # =========================================================================
    def _parse_management_book(self, filepath: str, progress_cb=None) -> list[dict]:
        """『注文品管理表.xlsx』の全シートを高速走査"""
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
            elif any(k in s_clean for k in ("外注加工", "外注", "外依")):
                sheet_cat = "OUTSOURCE"
            else:
                sheet_cat = "OTHER"

            # シート名に含まれる数字（例: 22609030）をシートのデフォルト管理No候補とする
            m_sheet_no = re.search(r"(\d{8})", sheet_name)
            sheet_default_order_no = m_sheet_no.group(1) if m_sheet_no else ""

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

            for list_i, (r_idx, row_tuple) in enumerate(raw_rows[:15]):
                if r_idx > 20:
                    break
                prev_tuple = raw_rows[list_i - 1][1] if list_i > 0 and raw_rows[list_i - 1][0] == r_idx - 1 else ()
                row_vals = []
                for c_idx in range(1, len(row_tuple) + 1):
                    v_curr = str(row_tuple[c_idx - 1] or "").strip()
                    v_above = str(prev_tuple[c_idx - 1] or "").strip() if c_idx - 1 < len(prev_tuple) else ""
                    combined = re.sub(r"\s+", "", f"{v_above}{v_curr}")
                    row_vals.append((c_idx, combined, re.sub(r"\s+", "", v_curr)))

                joined = " ".join(v[1] for v in row_vals)
                if "品名" in joined and (
                    "図面" in joined or "品番" in joined or "仕様" in joined or "現品" in joined or "メーカー" in joined
                ):
                    header_pos_in_list = list_i
                    for c_idx, v_comb, v_curr in row_vals:
                        if not v_comb:
                            continue
                        if "現品確認" in v_comb or (c_idx == 1 and "現品" in v_comb):
                            col_map.setdefault("actual_date", c_idx)
                        elif "請求日" in v_comb or (c_idx >= 30 and "請求" in v_comb):
                            col_map.setdefault("invoice_date", c_idx)
                        elif "管理No" in v_comb or "管理番号" in v_comb or "注文番号" in v_comb or "発注番号" in v_comb or "手配番号" in v_comb:
                            col_map.setdefault("order_no", c_idx)
                        elif "発注日" in v_comb or "注文日" in v_comb:
                            col_map.setdefault("order_date", c_idx)
                        elif "図面番号" in v_comb or "図番" in v_comb:
                            col_map.setdefault("drawing_no", c_idx)
                        elif "品番" in v_comb or "型番" in v_comb or "部品番号" in v_comb:
                            col_map.setdefault("part_no", c_idx)
                        elif "品名" in v_comb or "名称" in v_comb or "商品名" in v_comb:
                            col_map.setdefault("item_name", c_idx)
                        elif "メーカー" in v_comb or "発注先" in v_comb or "仕入先" in v_comb or "業者" in v_comb:
                            col_map.setdefault("maker", c_idx)
                        elif "仕様" in v_comb or "規格" in v_comb or "材質" in v_comb or "寸法" in v_comb:
                            col_map.setdefault("spec", c_idx)
                        elif "購入数" in v_comb or "発注数" in v_comb or v_curr in ("数量", "数", "購入数"):
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

            if "drawing_no" in col_map:
                base_c = col_map["drawing_no"]
                col_map.setdefault("part_no", base_c + 1)
                col_map.setdefault("item_name", base_c + 2)
                col_map.setdefault("maker", base_c + 3)
                col_map.setdefault("spec", base_c + 4)
                col_map.setdefault("qty", base_c + 5)
                col_map.setdefault("unit_price", base_c + 6)
                col_map.setdefault("amount", base_c + 7)
            elif "item_name" in col_map and col_map["item_name"] >= 3:
                ic = col_map["item_name"]
                col_map.setdefault("drawing_no", ic - 2)
                col_map.setdefault("part_no", ic - 1)
                col_map.setdefault("maker", ic + 1)
                col_map.setdefault("spec", ic + 2)
                col_map.setdefault("qty", ic + 3)
                col_map.setdefault("unit_price", ic + 4)
                col_map.setdefault("amount", ic + 5)

            col_map.setdefault("actual_date", 1)
            col_map.setdefault("invoice_date", 35)

            def _get_col(rtuple: tuple, col_1based: int | None):
                if not col_1based or col_1based < 1 or col_1based > len(rtuple):
                    return None
                return rtuple[col_1based - 1]

            for r_idx, row_tuple in raw_rows[header_pos_in_list + 1 :]:
                drawing_val = _get_col(row_tuple, col_map.get("drawing_no"))
                part_val = _get_col(row_tuple, col_map.get("part_no"))
                item_val = _get_col(row_tuple, col_map.get("item_name"))
                maker_val = _get_col(row_tuple, col_map.get("maker"))
                spec_val = _get_col(row_tuple, col_map.get("spec"))
                order_no_val = _get_col(row_tuple, col_map.get("order_no"))

                if not any([drawing_val, part_val, item_val, spec_val, order_no_val]):
                    continue

                order_date_raw = _get_col(row_tuple, col_map.get("order_date"))
                order_date = self._parse_flexible_date(order_date_raw)
                inferred_year = order_date.year if order_date else datetime.date.today().year

                a_col_val = _get_col(row_tuple, 1)
                mapped_actual_val = _get_col(row_tuple, col_map.get("actual_date"))
                actual_raw = a_col_val if self._is_filled_confirmation(a_col_val) else mapped_actual_val

                ai_col_val = _get_col(row_tuple, 35)
                mapped_inv_val = _get_col(row_tuple, col_map.get("invoice_date"))
                invoice_raw = ai_col_val if self._is_filled_confirmation(ai_col_val) else mapped_inv_val

                due_raw = _get_col(row_tuple, col_map.get("due_date"))

                actual_date = self._parse_flexible_date(actual_raw, default_year=inferred_year)
                actual_filled = actual_date is not None or self._is_filled_confirmation(actual_raw)

                invoice_date = self._parse_flexible_date(invoice_raw, default_year=inferred_year)
                invoice_filled = invoice_date is not None or self._is_filled_confirmation(invoice_raw)

                mgmt_due_date = self._parse_flexible_date(due_raw, default_year=inferred_year)

                drawing_str = self._clean_cell_str(drawing_val)
                part_str = self._clean_cell_str(part_val)
                item_str = self._clean_cell_str(item_val)
                maker_str = self._clean_cell_str(maker_val)
                spec_str = self._clean_cell_str(spec_val)
                qty_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("qty")))
                unit_price_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("unit_price")))
                amount_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("amount")))

                raw_order_no_str = self._clean_cell_str(order_no_val)
                order_no_str = raw_order_no_str or sheet_default_order_no
                person_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("person")))

                model_display_parts = [p for p in (drawing_str, part_str, spec_str) if p]
                model_spec_display = " / ".join(dict.fromkeys(model_display_parts))

                norm_drawing = self._normalize_key(drawing_str)
                norm_part = self._normalize_key(part_str)
                norm_item = self._normalize_key(item_str)
                norm_spec = self._normalize_key(spec_str)
                norm_maker = self._normalize_key(maker_str)
                norm_combined = self._normalize_key(f"{drawing_str} {part_str} {item_str} {spec_str}")

                # 管理No（8桁）の正規化抽出
                m_ord = re.search(r"(\d{8})", order_no_str)
                exact_order_no_8 = m_ord.group(1) if m_ord else (order_no_str if len(order_no_str) == 8 and order_no_str.isdigit() else "")

                mgmt_rows.append(
                    {
                        "sheet": sheet_name,
                        "sheet_cat": sheet_cat,
                        "row": r_idx,
                        "order_no": order_no_str,
                        "order_no_8": exact_order_no_8,
                        "norm_order_no": self._normalize_key(order_no_str),
                        "order_date": order_date,
                        "drawing_no": drawing_str,
                        "part_no": part_str,
                        "item_name": item_str,
                        "maker": maker_str,
                        "supplier": "",
                        "spec": spec_str,
                        "model_spec": model_spec_display,
                        "qty": qty_str,
                        "qty_num": self._extract_Leading_number(qty_str),
                        "unit_price": unit_price_str,
                        "price_num": self._extract_Leading_number(unit_price_str),
                        "amount": amount_str,
                        "amt_num": self._extract_Leading_number(amount_str),
                        "actual_date": actual_date,
                        "actual_raw": self._clean_cell_str(actual_raw),
                        "actual_filled": actual_filled,
                        "invoice_date": invoice_date,
                        "invoice_raw": self._clean_cell_str(invoice_raw),
                        "invoice_filled": invoice_filled,
                        "mgmt_due_date": mgmt_due_date,
                        "person": person_str,
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
                        "matched": False,
                    }
                )

        wb.close()
        return mgmt_rows

    def _parse_order_book(self, filepath: str) -> list[dict]:
        """注文書Excelファイルから明細を抽出（ファイル名先頭8桁を管理Noとして取得）"""
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        order_items = []
        filename = os.path.basename(filepath)

        # 注文書ファイル名の頭から8桁を管理Noとして取得
        file_order_no_8 = ""
        prefix_match = re.match(r"^(\d{8})", filename)
        if prefix_match:
            file_order_no_8 = prefix_match.group(1)
        else:
            m_any_8 = re.search(r"(\d{8})", filename)
            if m_any_8:
                file_order_no_8 = m_any_8.group(1)

        if "購依" in filename or "購入" in filename:
            req_type_code = "PURCHASE"
            req_type_label = "購入依頼"
        elif "外依" in filename or "外注" in filename:
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

            sheet_order_no = file_order_no_8
            sheet_supplier = file_supplier
            sheet_order_date = None
            default_year = datetime.date.today().year

            for r_idx, row_tuple in raw_rows:
                if r_idx > 16:
                    break
                for c_idx in range(1, min(21, len(row_tuple) + 1)):
                    cell_val = str(row_tuple[c_idx - 1] or "").strip()
                    if not cell_val:
                        continue

                    if not sheet_order_no and ("注文番号" in cell_val or "発注番号" in cell_val or "管理No" in cell_val or "No." in cell_val):
                        m_in = re.search(r"(\d{8})", cell_val)
                        if m_in:
                            sheet_order_no = m_in.group(1)
                        else:
                            for offset in range(1, 9):
                                if c_idx - 1 + offset >= len(row_tuple):
                                    break
                                right_val = str(row_tuple[c_idx - 1 + offset] or "").strip()
                                right_val = re.sub(r"\.0$", "", right_val)
                                m_r8 = re.search(r"(\d{8})", right_val)
                                if m_r8:
                                    sheet_order_no = m_r8.group(1)
                                    break

                    if "注文日" in cell_val or "発注日" in cell_val:
                        for offset in range(1, 9):
                            if c_idx - 1 + offset >= len(row_tuple):
                                break
                            d_candidate = self._parse_flexible_date(row_tuple[c_idx - 1 + offset])
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
                        ("株式会社" in cell_val or "有限会社" in cell_val or "(株)" in cell_val or "（株）" in cell_val or "(有)" in cell_val or "（有）" in cell_val)
                        and r_idx <= 10
                        and "丸菱" not in cell_val
                    ):
                        cleaned_sup = re.sub(r"(御中|様|殿)$", "", cell_val).strip()
                        if cleaned_sup and not sheet_supplier:
                            sheet_supplier = cleaned_sup

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
                        v_clean = re.sub(r"\s+", "", val)
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
                        elif "購入数" in v_clean or "数量" in v_clean or v_clean == "数":
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
                    for stop_word in ("合計", "小計", "消費税", "総合計", "特記事項", "以上")
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

                unit_price_str = self._clean_cell_str(unit_price_val)
                amount_str = self._clean_cell_str(amount_val)

                model_parts = [p for p in (drawing_str, part_str, spec_str) if p]
                model_spec_display = " / ".join(dict.fromkeys(model_parts))

                norm_drawing = self._normalize_key(drawing_str)
                norm_part = self._normalize_key(part_str)
                norm_item = self._normalize_key(item_str)
                norm_spec = self._normalize_key(spec_str)
                norm_maker = self._normalize_key(maker_str)
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
                        "order_no_8": file_order_no_8 or (sheet_order_no if len(sheet_order_no) == 8 and sheet_order_no.isdigit() else ""),
                        "norm_order_no": self._normalize_key(sheet_order_no),
                        "order_date": sheet_order_date,
                        "drawing_no": drawing_str,
                        "part_no": part_str,
                        "item_name": item_str,
                        "maker": maker_str,
                        "supplier": sheet_supplier or file_supplier,
                        "spec": spec_str,
                        "model_spec": model_spec_display,
                        "qty": qty_str,
                        "qty_num": self._extract_Leading_number(qty_str),
                        "unit_price": unit_price_str,
                        "price_num": self._extract_Leading_number(unit_price_str),
                        "amount": amount_str,
                        "amt_num": self._extract_Leading_number(amount_str),
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
    # 照合・遅延判定メイン処理
    # =========================================================================
    def run_analysis(self):
        """照合処理をバックグラウンドスレッドで開始"""
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
            messagebox.showerror("日付エラー", "判定基準日は YYYY-MM-DD 形式（例: 2026-10-30）で入力してください。")
            return

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED, text="照合処理中...")
        self.export_btn.configure(state=tk.DISABLED)
        self.progress_var.set(2.0)
        self.progress_text_var.set("解析を開始しています...")
        self.status_var.set("実行中: ファイルを読み込んでいます。しばらくお待ちください...")

        worker = threading.Thread(
            target=self._analysis_worker,
            args=(mgmt_path, order_dir, base_date),
            daemon=True,
        )
        worker.start()
        self.root.after(40, self._poll_ui_queue)

    def _analysis_worker(self, mgmt_path: str, order_dir: str, base_date: datetime.date):
        """バックグラウンドスレッドで実行されるExcel読込・照合処理"""
        try:
            self._report_progress(5.0, "[1/3] 注文品管理表を開いています...")
            mgmt_rows = self._parse_management_book(mgmt_path, progress_cb=self._report_progress)

            self._report_progress(35.0, "[2/3] 注文書フォルダ（サブフォルダ含む）を走査中...")
            order_items, valid_order_files, ignored_files = self._collect_valid_order_files(
                order_dir, mgmt_path, progress_cb=self._report_progress
            )

            self._report_progress(75.0, f"[3/3] 明細の照合と超過判定を実行中 ({len(order_items)} 件)...")
            analyzed_records = self._match_and_evaluate(
                mgmt_rows, order_items, base_date, progress_cb=self._report_progress
            )

            self.ui_queue.put(
                (
                    "DONE",
                    {
                        "records": analyzed_records,
                        "base_date": base_date,
                        "valid_order_files": valid_order_files,
                        "ignored_files": ignored_files,
                        "order_items_count": len(order_items),
                        "mgmt_rows_count": len(mgmt_rows),
                    },
                )
            )
        except Exception as e:
            self.ui_queue.put(("ERROR", str(e)))

    def _poll_ui_queue(self):
        """ワーカースレッドからの進捗・完了通知を受け取ってUIを更新"""
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
                    self.export_btn.configure(state=tk.NORMAL)

                    self.analyzed_records = payload["records"]
                    base_date = payload["base_date"]
                    valid_order_files = payload["valid_order_files"]
                    ignored_files = payload["ignored_files"]
                    order_items_count = payload["order_items_count"]
                    mgmt_rows_count = payload["mgmt_rows_count"]

                    self.column_filters.clear()
                    self._update_all_headings()

                    self._update_kpis()
                    self._refresh_table()

                    delay_cnt = sum(1 for r in self.analyzed_records if r["is_any_delay"])
                    unmatched_cnt = sum(1 for r in self.analyzed_records if r["is_unmatched"])
                    self.progress_var.set(100.0)
                    self.progress_text_var.set(
                        f"完了！ (注文書 {valid_order_files}ファイル・{order_items_count}明細を照合 / 管理表 {mgmt_rows_count}行)"
                    )
                    self.status_var.set(
                        f"照合完了 (基準日: {base_date.strftime('%Y-%m-%d')}) — "
                        f"注文書明細: {order_items_count}件 / 超過(遅れ): {delay_cnt}件 / 不一致(未記載): {unmatched_cnt}件"
                    )
                    return
                elif msg_type == "ERROR":
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL, text="照合・遅延チェック実行")
                    self.export_btn.configure(state=tk.NORMAL)
                    self.progress_var.set(0.0)
                    self.progress_text_var.set("エラーが発生しました")
                    self.status_var.set(f"エラー: {payload}")
                    messagebox.showerror("解析エラー", f"処理中にエラーが発生しました:\n{payload}")
                    return
        except queue.Empty:
            pass

        if self.is_running:
            self.root.after(40, self._poll_ui_queue)

    def _match_and_evaluate(
        self,
        mgmt_rows: list[dict],
        order_items: list[dict],
        base_date: datetime.date,
        progress_cb=None,
    ) -> list[dict]:
        """
        注文書の明細と注文品管理表を照合する。
        【重要規則】注文書のファイル名先頭8桁が管理Noとなる。
        注文品管理表の中にこの管理Noと一致するデータが存在しない場合は、無条件で不一致判定とする。
        """
        results = []
        total_orders = len(order_items)

        # 管理表に存在する全管理No（8桁）の集合を事前構築
        all_mgmt_order_nos_8 = set()
        for m in mgmt_rows:
            if m.get("order_no_8"):
                all_mgmt_order_nos_8.add(m["order_no_8"])
            elif m.get("norm_order_no"):
                all_mgmt_order_nos_8.add(m["norm_order_no"])

        for idx, o_item in enumerate(order_items, start=1):
            if progress_cb and (idx == 1 or idx % 5 == 0 or idx == total_orders):
                pct = 75.0 + (23.0 * idx / max(1, total_orders))
                progress_cb(
                    pct,
                    f"[3/3] 注文書明細と管理表を照合中 ({idx}/{total_orders} 件目)...",
                )

            req_code = o_item.get("req_type_code", "UNKNOWN")
            o_order_no_8 = o_item.get("order_no_8", "")

            # -----------------------------------------------------------------
            # 必須条件チェック: ファイル名先頭8桁の管理Noが管理表に存在するか？
            # 一致する管理Noが管理表に存在しない場合は直ちに不一致判定とする
            # -----------------------------------------------------------------
            if not o_order_no_8 or o_order_no_8 not in all_mgmt_order_nos_8:
                results.append(
                    self._evaluate_single_record(
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=o_order_no_8 or o_item.get("order_no", ""),
                        supplier=o_item.get("supplier", ""),
                        maker=o_item.get("maker", ""),
                        item_name=o_item.get("item_name", ""),
                        model_spec=o_item.get("model_spec", ""),
                        qty=o_item.get("qty", ""),
                        due_date=o_item.get("due_date"),
                        actual_date=None,
                        actual_raw="",
                        actual_filled=False,
                        invoice_date=None,
                        invoice_raw="",
                        invoice_filled=False,
                        person="【不一致】",
                        source_info=f"【不一致・管理No未登録】注文書:{o_item['file']}(行{o_item['row']})",
                        base_date=base_date,
                        is_unmatched=True,
                    )
                )
                continue

            # 管理Noが一致する管理表行のみを厳密な照合対象とする
            matched_no_mgmt = [
                m for m in mgmt_rows
                if not m["matched"] and (m.get("order_no_8") == o_order_no_8 or m.get("norm_order_no") == o_order_no_8)
            ]

            primary_mgmt = []
            fallback_mgmt = []
            for m in matched_no_mgmt:
                if req_code == "PURCHASE":
                    if m.get("sheet_cat") == "PURCHASE":
                        primary_mgmt.append(m)
                    else:
                        fallback_mgmt.append(m)
                elif req_code == "OUTSOURCE":
                    if m.get("sheet_cat") == "OUTSOURCE":
                        primary_mgmt.append(m)
                    else:
                        fallback_mgmt.append(m)
                else:
                    primary_mgmt.append(m)

            search_targets = primary_mgmt if primary_mgmt else fallback_mgmt

            # 管理No一致行がすべて照合済み、または対象シートに見つからない場合
            if not search_targets:
                results.append(
                    self._evaluate_single_record(
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=o_order_no_8 or o_item.get("order_no", ""),
                        supplier=o_item.get("supplier", ""),
                        maker=o_item.get("maker", ""),
                        item_name=o_item.get("item_name", ""),
                        model_spec=o_item.get("model_spec", ""),
                        qty=o_item.get("qty", ""),
                        due_date=o_item.get("due_date"),
                        actual_date=None,
                        actual_raw="",
                        actual_filled=False,
                        invoice_date=None,
                        invoice_raw="",
                        invoice_filled=False,
                        person="【不一致】",
                        source_info=f"【不一致・管理No該当なし】注文書:{o_item['file']}(行{o_item['row']})",
                        base_date=base_date,
                        is_unmatched=True,
                    )
                )
                continue

            o_no_norm = o_item["norm_order_no"]
            o_qty_num = o_item["qty_num"]
            o_drawing_norm = o_item["norm_drawing"]
            o_part_norm = o_item["norm_part"]

            best_mgmt = None
            best_score = -1.0

            # 1. 管理No一致行内での図面番号・品番による最優先照合
            has_identifying_code = bool(o_drawing_norm or o_part_norm)

            if has_identifying_code:
                candidate_matches = []
                for m_row in search_targets:
                    if m_row["matched"]:
                        continue

                    m_drawing_norm = m_row["norm_drawing"]
                    m_part_norm = m_row["norm_part"]

                    exact_drawing = bool(o_drawing_norm and m_drawing_norm and o_drawing_norm == m_drawing_norm)
                    exact_part = bool(o_part_norm and m_part_norm and o_part_norm == m_part_norm)
                    cross_dp = bool(o_drawing_norm and m_part_norm and o_drawing_norm == m_part_norm)
                    cross_pd = bool(o_part_norm and m_drawing_norm and o_part_norm == m_drawing_norm)

                    if exact_drawing or exact_part or cross_dp or cross_pd:
                        code_score = 1000.0
                        if exact_drawing and exact_part:
                            code_score += 500.0
                        if o_no_norm and m_row["norm_order_no"] and o_no_norm == m_row["norm_order_no"]:
                            code_score += 100.0
                        if o_qty_num and m_row["qty_num"] and o_qty_num == m_row["qty_num"]:
                            code_score += 50.0
                        candidate_matches.append((code_score, m_row))
                        continue

                    part_match = False
                    if o_drawing_norm and m_drawing_norm and (o_drawing_norm in m_drawing_norm or m_drawing_norm in o_drawing_norm):
                        part_match = True
                    elif o_part_norm and m_part_norm and (o_part_norm in m_part_norm or m_part_norm in o_part_norm):
                        part_match = True
                    elif o_drawing_norm and m_part_norm and (o_drawing_norm in m_part_norm or m_part_norm in o_drawing_norm):
                        part_match = True
                    elif o_part_norm and m_drawing_norm and (o_part_norm in m_drawing_norm or m_drawing_norm in o_part_norm):
                        part_match = True

                    if part_match:
                        code_score = 500.0
                        if o_no_norm and m_row["norm_order_no"] and o_no_norm == m_row["norm_order_no"]:
                            code_score += 50.0
                        if o_qty_num and m_row["qty_num"] and o_qty_num == m_row["qty_num"]:
                            code_score += 30.0
                        candidate_matches.append((code_score, m_row))

                if candidate_matches:
                    candidate_matches.sort(key=lambda x: x[0], reverse=True)
                    best_mgmt = candidate_matches[0][1]

            # 2. フォールバック照合（図番・品番が未記入の場合）
            if best_mgmt is None and not has_identifying_code:
                o_price_num = o_item["price_num"]
                o_amt_num = o_item["amt_num"]

                for m_row in search_targets:
                    if m_row["matched"]:
                        continue

                    score = 0.0
                    matched_key_count = 0

                    if o_item["norm_item"] and m_row["norm_item"]:
                        name_sim = self._fast_similarity(
                            o_item["norm_item"], o_item["bg_item"], m_row["norm_item"], m_row["bg_item"]
                        )
                        if o_item["norm_item"] == m_row["norm_item"] or o_item["norm_item"] in m_row["norm_item"] or m_row["norm_item"] in o_item["norm_item"]:
                            name_sim = max(name_sim, 0.95)
                        if name_sim >= 0.45:
                            score += 50.0 * name_sim
                            matched_key_count += 1

                    if o_item["norm_spec"] and m_row["norm_spec"]:
                        spec_sim = self._fast_similarity(
                            o_item["norm_spec"], o_item["bg_spec"], m_row["norm_spec"], m_row["bg_spec"]
                        )
                        if o_item["norm_spec"] == m_row["norm_spec"] or o_item["norm_spec"] in m_row["norm_spec"] or m_row["norm_spec"] in o_item["norm_spec"]:
                            spec_sim = max(spec_sim, 0.95)
                        if spec_sim >= 0.45:
                            score += 40.0 * spec_sim
                            matched_key_count += 1

                    if o_qty_num and m_row["qty_num"] and o_qty_num == m_row["qty_num"]:
                        score += 20.0
                    if o_price_num and m_row["price_num"] and o_price_num == m_row["price_num"]:
                        score += 15.0
                    if o_amt_num and m_row["amt_num"] and o_amt_num == m_row["amt_num"]:
                        score += 15.0

                    if matched_key_count >= 1 and score > best_score and score >= 25.0:
                        best_score = score
                        best_mgmt = m_row

            if best_mgmt:
                best_mgmt["matched"] = True
                final_supplier = o_item.get("supplier") or best_mgmt.get("supplier", "")
                final_maker = o_item.get("maker") or best_mgmt.get("maker", "")
                record = self._evaluate_single_record(
                    req_type=o_item.get("req_type_label", "注文"),
                    order_no=best_mgmt.get("order_no") or o_order_no_8 or o_item.get("order_no", ""),
                    supplier=final_supplier,
                    maker=final_maker,
                    item_name=o_item.get("item_name") or best_mgmt.get("item_name", ""),
                    model_spec=o_item.get("model_spec") or best_mgmt.get("model_spec", ""),
                    qty=o_item.get("qty") or best_mgmt.get("qty", ""),
                    due_date=o_item.get("due_date") or best_mgmt.get("mgmt_due_date"),
                    actual_date=best_mgmt.get("actual_date"),
                    actual_raw=best_mgmt.get("actual_raw", ""),
                    actual_filled=best_mgmt.get("actual_filled", False),
                    invoice_date=best_mgmt.get("invoice_date"),
                    invoice_raw=best_mgmt.get("invoice_raw", ""),
                    invoice_filled=best_mgmt.get("invoice_filled", False),
                    person=best_mgmt.get("person", ""),
                    source_info=f"[{o_item.get('req_type_label', '')}] 注文書:{o_item['file']} ⇔ 管理表:{best_mgmt['sheet']}(行{best_mgmt['row']})",
                    base_date=base_date,
                    is_unmatched=False,
                )
            else:
                # 管理Noはあるが明細が見当たらない場合（不一致判定）
                record = self._evaluate_single_record(
                    req_type=o_item.get("req_type_label", "注文"),
                    order_no=o_order_no_8 or o_item.get("order_no", ""),
                    supplier=o_item.get("supplier", ""),
                    maker=o_item.get("maker", ""),
                    item_name=o_item.get("item_name", ""),
                    model_spec=o_item.get("model_spec", ""),
                    qty=o_item.get("qty", ""),
                    due_date=o_item.get("due_date"),
                    actual_date=None,
                    actual_raw="",
                    actual_filled=False,
                    invoice_date=None,
                    invoice_raw="",
                    invoice_filled=False,
                    person="【不一致】",
                    source_info=f"【不一致・管理表未記載】注文書:{o_item['file']}(行{o_item['row']})",
                    base_date=base_date,
                    is_unmatched=True,
                )
            results.append(record)

        # 並び替え: 不一致および超過を上位に表示
        results.sort(
            key=lambda r: (
                not (r["is_unmatched"] or r["is_any_delay"]),
                not r["is_unmatched"],
                -(r["overdue_days_num"] or -9999),
                r["order_no"],
            )
        )
        return results

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
        is_unmatched: bool = False,
    ) -> dict:
        """1明細に対する現品・請求の判定処理"""
        actual_overdue = False
        invoice_overdue = False
        overdue_days_num = None
        days_overdue_str = "-"

        if due_date:
            delta_from_today = (base_date - due_date).days
            if delta_from_today > 0:
                overdue_days_num = delta_from_today
                days_overdue_str = (
                    f"{delta_from_today}日超過"
                    if not is_fallback_order_date
                    else f"発注後{delta_from_today}日経過"
                )
            elif delta_from_today == 0:
                overdue_days_num = 0
                days_overdue_str = "本日納期"
            else:
                overdue_days_num = delta_from_today
                days_overdue_str = f"残り{-delta_from_today}日"

        # 管理表に一致しない項目（管理No不一致または明細未記載）の場合の判定表示
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
            # 1) 現品確認日（A列：物）の超過判定
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

            # 2) 請求日（AI列：伝票）の超過判定
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

            # 総合ステータスバッジ
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
            due_str = f"({due_date.strftime('%Y-%m-%d')}発注)"

        return {
            "status_badge": status_badge,
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
        }

    def _update_kpis(self):
        """上部KPI集計カードの数値更新"""
        total = len(self.analyzed_records)
        unmatched_cnt = sum(1 for r in self.analyzed_records if r.get("is_unmatched"))
        both_cnt = sum(1 for r in self.analyzed_records if not r.get("is_unmatched") and r["actual_overdue"] and r["invoice_overdue"])
        actual_cnt = sum(1 for r in self.analyzed_records if not r.get("is_unmatched") and r["actual_overdue"] and not r["invoice_overdue"])
        invoice_cnt = sum(1 for r in self.analyzed_records if not r.get("is_unmatched") and not r["actual_overdue"] and r["invoice_overdue"])
        ok_cnt = sum(1 for r in self.analyzed_records if not r.get("is_unmatched") and not r["actual_overdue"] and not r["invoice_overdue"])

        self.kpi_vars["total"].set(f"{total} 件")
        self.kpi_vars["unmatched"].set(f"{unmatched_cnt} 件")
        self.kpi_vars["both_delay"].set(f"{both_cnt} 件")
        self.kpi_vars["actual_delay"].set(f"{actual_cnt} 件")
        self.kpi_vars["invoice_delay"].set(f"{invoice_cnt} 件")
        self.kpi_vars["ok"].set(f"{ok_cnt} 件")

    def _refresh_table(self):
        """フィルタ条件・検索キーワードに基づいてTreeviewを再描画"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        mode = self.filter_mode.get()
        kw = self.search_keyword.get().strip().lower()

        for rec in self.analyzed_records:
            if mode == "ALL_DELAY":
                # すべての異常（不一致または超過）
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

            # 各列の個別フィルター
            skip_by_col = False
            for col_id, allowed_values in self.column_filters.items():
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

            self.tree.insert(
                "",
                tk.END,
                values=(
                    rec["status_badge"],
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

    def _update_heading_display(self, col_id: str):
        """列ヘッダーのテキストを更新"""
        orig_title = self.col_titles.get(col_id, col_id)
        if col_id in self.column_filters:
            display_title = f"{orig_title} [▼]"
        else:
            display_title = f"{orig_title} ▾"
        self.tree.heading(col_id, text=display_title)

    def _update_all_headings(self):
        """全列のヘッダー表示を更新"""
        for col_id, _, _, _ in self.col_configs:
            self._update_heading_display(col_id)

    def _on_tree_click(self, event):
        """ヘッダークリックでExcel風フィルターを開く"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col = self.tree.identify_column(event.x)
            col_index = int(col.replace("#", "")) - 1
            if 0 <= col_index < len(self.col_configs):
                col_id = self.col_configs[col_index][0]
                self._open_excel_column_filter(col_id, event.x_root, event.y_root)

    def _open_excel_column_filter(self, col_id: str, x: int, y: int):
        """Excel風ドロップダウンフィルター・ソートウィンドウを表示"""
        if not self.analyzed_records:
            return

        all_unique_vals = set()
        for rec in self.analyzed_records:
            val = str(rec.get(col_id, "") or "").strip()
            if not val:
                val = "(空白)"
            all_unique_vals.add(val)

        sorted_vals = sorted(all_unique_vals, key=lambda s: (s == "(空白)", s))
        currently_selected = set(self.column_filters.get(col_id, all_unique_vals))

        win = tk.Toplevel(self.root)
        win.title(f"フィルター: {self.col_titles.get(col_id, col_id)}")

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
            command=lambda: [self._sort_tree(col_id, False), win.destroy()],
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
            command=lambda: [self._sort_tree(col_id, True), win.destroy()],
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

        def _on_frame_configure(e):
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
                self.column_filters.pop(col_id, None)
            else:
                self.column_filters[col_id] = selected
            self._update_heading_display(col_id)
            self._refresh_table()
            _cleanup_and_destroy()

        def _clear_filter():
            self.column_filters.pop(col_id, None)
            self._update_heading_display(col_id)
            self._refresh_table()
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

    def _clear_all_column_filters(self):
        """すべての列フィルターを一括解除"""
        self.column_filters.clear()
        self._update_all_headings()
        self._refresh_table()

    def _sort_tree(self, col: str, reverse: bool):
        """Treeview列ヘッダークリック時のソート処理"""
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            self.tree.move(k, "", index)

    def _export_results(self):
        """検査結果をExcelまたはCSVに出力する"""
        if not self.analyzed_records:
            messagebox.showinfo("出力対象なし", "先に「照合・遅延チェック実行」を行ってください。")
            return

        save_path = filedialog.asksaveasfilename(
            title="超過一覧レポートの保存",
            defaultextension=".xlsx",
            initialfile=f"納期超過チェック結果_{self.base_date_str.get()}.xlsx",
            filetypes=[("Excelファイル", "*.xlsx"), ("CSVファイル", "*.csv")],
        )
        if not save_path:
            return

        export_rows = []
        for r in self.analyzed_records:
            export_rows.append(
                {
                    "総合判定": r["status_badge"],
                    "依頼区分": r["req_type"],
                    "管理No(注文番号)": r["order_no"],
                    "発注先": r["supplier"],
                    "メーカー": r["maker"],
                    "品名": r["item_name"],
                    "型番・規格": r["model_spec"],
                    "数量": r["qty"],
                    "注文書納期": r["due_date_str"],
                    "基準日超過": r["days_overdue_str"],
                    "現品確認日(物)": r["actual_date_str"],
                    "現品状況": r["actual_status"],
                    "請求日(伝票)": r["invoice_date_str"],
                    "請求状況": r["invoice_status"],
                    "現品超過フラグ": "超過" if r["actual_overdue"] else "正常",
                    "請求超過フラグ": "超過" if r["invoice_overdue"] else "正常",
                    "管理表不一致フラグ": "不一致" if r.get("is_unmatched") else "一致",
                    "担当": r["person"],
                    "照合元情報": r["source_info"],
                }
            )

        df = pd.DataFrame(export_rows)
        try:
            if save_path.lower().endswith(".csv"):
                df.to_csv(save_path, index=False, encoding="utf-8-sig")
            else:
                df.to_excel(save_path, index=False)
            messagebox.showinfo("出力完了", f"検査結果を保存しました:\n{save_path}")
        except Exception as e:
            messagebox.showerror("保存エラー", f"ファイルの保存に失敗しました:\n{e}")


def main():
    root = tk.Tk()
    app = StockMatchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()