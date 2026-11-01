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

try:
    from scipy.optimize import linear_sum_assignment
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False


class StockMatchApp:
    """
    StockMatch_Miya-v1.0.0
    注文品管理表（1つのファイル）と注文書フォルダ内（サブフォルダ含む全階層）の全注文書Excelを照合し、
    今日（判定基準日）時点で「現品確認日」および「請求日」の納期超過を検査・可視化するデスクトップアプリ
    ※設定ファイルはアプリが存在するフォルダ内の config.json に保存・読込されます。
    ※管理Noが指定の数字（デフォルト: 2、カンマ区切りで複数指定可）から始まる明細のみを対象とし、設定を保存できます。
    ※2行目に「見積」の記載がある見積依頼書は自動除外し、「注文」の記載がある注文書のみを照合します。
    """

    # アプリ配置フォルダおよび設定ファイルパスの定義
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    CONFIG_FILE_PATH = os.path.join(APP_DIR, "config.json")

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
        # 管理No頭文字フィルター設定（デフォルト: "2"）
        self.order_no_prefix = tk.StringVar(value="2")
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
        config_menu.add_command(label="現在の設定を保存 (アプリフォルダ直下)", command=self._save_config_manual)
        config_menu.add_command(label="アプリ配置フォルダを開く", command=self._open_config_folder)
        menubar.add_cascade(label="設定(S)", menu=config_menu)

        # 3. ヘルプメニュー
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使い方・操作ガイド(H)...", command=lambda: self._show_help_guide(tab_index=0), accelerator="F1")
        help_menu.add_command(label="照合・判定ロジックの詳細...", command=lambda: self._show_help_guide(tab_index=1))
        help_menu.add_command(label="ショートカット・便利機能一覧...", command=lambda: self._show_help_guide(tab_index=2))
        help_menu.add_separator()
        help_menu.add_command(label="バージョン情報...", command=self._show_about)
        menubar.add_cascade(label="ヘルプ(H)", menu=help_menu)

        self.root.config(menu=menubar)

        # ショートカットキー割り当て
        self.root.bind_all("<F1>", lambda e: self._show_help_guide(tab_index=0))
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
        ttk.Button(row2, text="パス・設定を保存", style="Action.TButton", command=self._save_config_manual).pack(
            side=tk.LEFT
        )

        # ③ 基準日・管理No頭文字・実行ボタン行
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
        ).pack(side=tk.LEFT, padx=(0, 16))

        # 管理No頭文字フィルター設定入力欄（説明書きをわかりやすく改善）
        ttk.Label(
            row3,
            text="照合管理No頭文字:",
            style="Card.TLabel",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 4))
        prefix_entry = ttk.Entry(row3, textvariable=self.order_no_prefix, width=8, font=("Meiryo UI", 9, "bold"))
        prefix_entry.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Label(
            row3,
            text="※カンマ区切りで複数指定可 (例: 2, 3) ／ 空欄で全件照合",
            style="Card.TLabel",
            foreground="#475569",
            font=("Meiryo UI", 8),
        ).pack(side=tk.LEFT, padx=(0, 12))

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

        self._create_kpi_card(kpi_frame, "検査対象注文明細 総数", self.kpi_vars["total"], "#334155", "#f1f5f9").pack(
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
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_keyword, width=20)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind("<KeyRelease>", lambda e: self._refresh_table())

        self.clear_col_filters_btn = ttk.Button(
            filter_frame,
            text="列フィルター解除 (Excel風)",
            style="Action.TButton",
            command=self._clear_all_column_filters,
        )
        self.clear_col_filters_btn.pack(side=tk.RIGHT, padx=(6, 0))

        self.inspect_btn = ttk.Button(
            filter_frame,
            text="🔍 選択行の不一致・照合詳細を表示",
            style="Action.TButton",
            command=self._open_selected_row_detail,
        )
        self.inspect_btn.pack(side=tk.RIGHT, padx=(6, 0))

        # 5. メイン結果テーブル（Treeview）
        table_container = ttk.Frame(self.root, padding=(16, 0, 16, 12))
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

        self.tree = ttk.Treeview(table_container, columns=columns, show="headings", selectmode="browse")

        self.col_configs = [
            ("status_badge", "総合判定", 155, "center"),
            ("mismatch_reason", "不一致理由・照合診断", 340, "w"),
            ("req_type", "依頼区分", 80, "center"),
            ("order_no", "管理No(注文番号)", 115, "center"),
            ("supplier", "発注先", 130, "w"),
            ("maker", "メーカー", 100, "w"),
            ("item_name", "品名", 145, "w"),
            ("model_spec", "型番・規格(図番)", 165, "w"),
            ("qty", "数量", 60, "center"),
            ("due_date", "注文書納期", 95, "center"),
            ("days_overdue", "基準日超過", 90, "center"),
            ("actual_date", "現品確認日(物)", 110, "center"),
            ("actual_status", "現品状況", 135, "center"),
            ("invoice_date", "請求日(伝票)", 110, "center"),
            ("invoice_status", "請求状況", 135, "center"),
            ("person", "担当", 70, "center"),
            ("source_info", "照合元情報", 190, "w"),
        ]

        self.col_titles = {col_id: title for col_id, title, _, _ in self.col_configs}

        for col_id, heading_text, width, anchor in self.col_configs:
            self._update_heading_display(col_id)
            self.tree.column(col_id, width=width, anchor=anchor)

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # 行の背景色タグ設定
        self.tree.tag_configure("UNMATCHED", background="#ede9fe", foreground="#4c1d95")  # 紫系（不一致）
        self.tree.tag_configure("BOTH_DELAY", background="#fee2e2", foreground="#7f1d1d")
        self.tree.tag_configure("ACTUAL_DELAY", background="#ffedd5", foreground="#7c2d12")
        self.tree.tag_configure("INVOICE_DELAY", background="#fef9c3", foreground="#713f12")
        self.tree.tag_configure("OK", background="#ffffff", foreground="#334155")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set)

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
    # 不一致・照合詳細の比較インスペクタ（ダブルクリック / ボタン表示）
    # =========================================================================
    def _on_tree_double_click(self, event):
        """データ行のダブルクリックで詳細比較ウィンドウを表示"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            return
        self._open_selected_row_detail()

    def _open_selected_row_detail(self):
        """現在選択されている行の不一致理由・項目比較ダイアログを開く"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("行未選択", "一覧テーブルから確認したい行をクリックして選択してください。")
            return

        item_id = sel[0]
        rec_idx_str = item_id.replace("rec_", "")
        if not rec_idx_str.isdigit():
            return
        rec_idx = int(rec_idx_str)
        if not (0 <= rec_idx < len(self.analyzed_records)):
            return

        rec = self.analyzed_records[rec_idx]
        diag = rec.get("diagnostic_data", {})

        win = tk.Toplevel(self.root)
        win.title(f"不一致・照合詳細比較 — 管理No: {rec['order_no']} / 品名: {rec['item_name']}")
        win.geometry("920x660")
        win.minsize(760, 520)
        win.transient(self.root)

        top_banner = tk.Frame(
            win,
            bg="#4c1d95" if rec.get("is_unmatched") else "#1e293b",
            padx=14,
            pady=10,
        )
        top_banner.pack(fill=tk.X)

        tk.Label(
            top_banner,
            text=f"判定: {rec['status_badge']}   |   {rec['mismatch_reason']}",
            bg=top_banner["bg"],
            fg="#ffffff",
            font=("Meiryo UI", 10, "bold"),
            anchor="w",
        ).pack(fill=tk.X)

        body = ttk.Frame(win, padding=14)
        body.pack(fill=tk.BOTH, expand=True)

        # 項目別比較テーブル
        lbl_cmp = ttk.Label(
            body,
            text="■ 注文書 と 注文品管理表（最有力候補または照合行）の項目別比較（※単価・金額は照合対象外）",
            font=("Meiryo UI", 9, "bold"),
        )
        lbl_cmp.pack(anchor="w", pady=(0, 6))

        cmp_cols = ("field", "order_val", "mgmt_val", "judge")
        cmp_tree = ttk.Treeview(body, columns=cmp_cols, show="headings", height=10)
        cmp_tree.heading("field", text="比較項目")
        cmp_tree.heading("order_val", text="注文書の記載内容")
        cmp_tree.heading("mgmt_val", text="注文品管理表の記載内容 (最有力候補)")
        cmp_tree.heading("judge", text="照合判定 (単価・金額除外)")

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

        # 詳細診断レポートテキスト
        lbl_rep = ttk.Label(
            body,
            text="■ 詳細診断レポート（なぜ不一致になったか / 管理表側の検索状況）",
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
    # ヘルプ・ガイドダイアログ（リッチテキスト＆タブ形式で視認性を大幅向上）
    # =========================================================================
    def _configure_rich_text_tags(self, txt: tk.Text):
        """ヘルプ表示用のTextウィジェットに見出し・強調・ボックス等のスタイルタグを設定"""
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
        """タブ形式で整理された見やすいヘルプ・操作ガイドウィンドウを表示"""
        win = tk.Toplevel(self.root)
        win.title("StockMatch_Miya-v1.0.0 - ヘルプ＆操作ガイド")
        win.geometry("880x680")
        win.minsize(720, 520)
        win.transient(self.root)

        # 上部ヘッダーバナー
        banner = tk.Frame(win, bg="#1e293b", padx=18, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text="📖 StockMatch_Miya-v1.0.0  取扱説明書・操作ガイド",
            bg="#1e293b",
            fg="#ffffff",
            font=("Meiryo UI", 12, "bold"),
            anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            banner,
            text="注文書フォルダと注文品管理表の自動照合・納期超過チェック・不一致診断ツールの使い方",
            bg="#1e293b",
            fg="#cbd5e1",
            font=("Meiryo UI", 9),
            anchor="w",
        ).pack(fill=tk.X, pady=(2, 0))

        # タブコントロール（Notebook）
        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=14, pady=(12, 6))

        # --- タブ1: 基本的な使い方 ---
        tab1 = ttk.Frame(notebook, padding=8)
        notebook.add(tab1, text="  ① 基本的な使い方・設定  ")
        txt1 = self._create_scrollable_text(tab1)

        txt1.insert(tk.END, "  STEP 1 : ファイルとフォルダの指定  \n", "h1")
        txt1.insert(tk.END, "① 注文品管理表ファイルの選択\n", "h2")
        txt1.insert(tk.END, "・『ファイルを参照...』ボタンから、基準となる管理表Excel（例: 注文品管理表.xlsx）を選択します。\n", "body")
        txt1.insert(tk.END, "・管理表の15列目（O列）付近にある個別の「品名」列を正確に参照します（D列の案件名「品名、ライン名等」とは混同しません）。\n", "sub_note")

        txt1.insert(tk.END, "② 注文書フォルダパスの選択（サブフォルダ自動対応）\n", "h2")
        txt1.insert(tk.END, "・『フォルダを参照...』ボタンから、注文書Excelが入っている親フォルダを選択します。\n", "body")
        txt1.insert(tk.END, "★【サブフォルダ一括探索】指定したフォルダ直下だけでなく、配下のすべてのサブフォルダ（月別・案件別フォルダ等）も自動的に走査します。\n", "point")
        txt1.insert(tk.END, "★【見積依頼書の自動除外】Excelシートの2行目を自動チェックし、「見積」と書かれた見積依頼書は除外して「注文」のシートのみを照合します。\n", "point")

        txt1.insert(tk.END, "\n  STEP 2 : 照合管理No頭文字・判定基準日の設定  \n", "h1")
        txt1.insert(tk.END, "■ 照合管理No頭文字（カンマ区切りで複数指定が可能！）\n", "h2")
        txt1.insert(tk.END, "・注文書ファイル名の先頭8桁（管理No）が、指定した数字で始まるものだけを照合対象に絞り込みます。\n", "body")
        txt1.insert(tk.END, "  ▸ 単独指定の例 :  「 2 」と入力  ⇒  2から始まる管理No（例: 22609030 等）のみを照合します（初期設定）。\n", "point")
        txt1.insert(tk.END, "  ▸ 複数指定の例 :  「 2, 3 」や「 126, 226 」とカンマ区切りで入力  ⇒  該当する複数の頭文字すべてを同時に照合します。\n", "point")
        txt1.insert(tk.END, "  ▸ 全件照合の例 :  入力欄を「 空欄 」にする  ⇒  頭文字の制限を行わず、すべての管理Noを照合します。\n", "body")

        txt1.insert(tk.END, "■ 判定基準日（今日）と設定の保存\n", "h2")
        txt1.insert(tk.END, "・納期超過の基準となる日付（初期値は当日）を YYYY-MM-DD 形式で指定します。\n", "body")
        txt1.insert(tk.END, "・『パス・設定を保存』ボタンを押すと、「管理表パス」「注文書フォルダパス」「照合管理No頭文字」がアプリと同じフォルダ内の config.json に保存され、次回起動時に自動で復元されます。\n", "body")
        txt1.insert(tk.END, f"  （現在の保存先パス: {self.CONFIG_FILE_PATH}）\n", "sub_note")

        txt1.insert(tk.END, "\n  STEP 3 : 照合実行と結果の確認  \n", "h1")
        txt1.insert(tk.END, "■ 『照合・遅延チェック実行』ボタン（または F5 キー）\n", "h2")
        txt1.insert(tk.END, "・ボタンを押すと照合を開始し、進捗バーでリアルタイムに進行状況を表示します。\n", "body")
        txt1.insert(tk.END, "・一覧の行を【ダブルクリック】すると、注文書と管理表の項目を左右に並べた「詳細比較ウィンドウ」が開き、不一致箇所が一目で分かります。\n", "point")
        txt1.configure(state="disabled")

        # --- タブ2: 照合・判定ロジックの詳細 ---
        tab2 = ttk.Frame(notebook, padding=8)
        notebook.add(tab2, text="  ② 照合ルール・不一致診断の仕組み  ")
        txt2 = self._create_scrollable_text(tab2)

        txt2.insert(tk.END, "  1. 注文書と管理表の照合ルール（高精度マッチング）  \n", "h1")
        txt2.insert(tk.END, "■ 単価・金額は照合対象から除外（誤判定の防止）\n", "h2")
        txt2.insert(tk.END, "★【重要】管理表側で単価が空欄だったり金額が 0 円になっているケースや、税抜/税込の表記差による誤った不一致を防ぐため、「単価」と「金額」は照合判定から完全に除外しています。\n", "point")
        txt2.insert(tk.END, "・照合に使用する項目は『①管理No(8桁)』『②図面番号』『③品番』『④品名(O列)』『⑤仕様・規格』『⑥発注先・メーカー』『⑦購入数(数量)』の7項目です。\n", "body")

        txt2.insert(tk.END, "■ 購入依頼（購依）と外注依頼（外依）の自動判別\n", "h2")
        txt2.insert(tk.END, "・注文書ファイル名に「購依」が含まれる場合は管理表の『購入品リスト』を優先参照し、「外依」が含まれる場合は『外注加工リスト』を優先参照します。\n", "body")

        txt2.insert(tk.END, "■ 全体最適マッチング（別業者の誤消費ブロック）\n", "h2")
        txt2.insert(tk.END, "・発注先やメーカー名の一致を評価し、別会社の注文書が他社の管理表行を誤って先取り消費してしまう現象を防止しています。\n", "body")
        txt2.insert(tk.END, "・同じ管理No内の全明細と管理表行のスコアを全体計算し、最も整合性の高い1対1ペアを自動決定します。\n", "body")

        txt2.insert(tk.END, "\n  2. 不一致理由・照合診断の見方  \n", "h1")
        txt2.insert(tk.END, "■ 【管理No未登録 / 管理No不一致】と表示される場合\n", "h2")
        txt2.insert(tk.END, "・注文書ファイル名先頭8桁の管理Noが、注文品管理表のどのシートにも1件も登録されていない状態です。\n", "warn")
        txt2.insert(tk.END, "・管理表の別管理Noに同じ品目がある場合は、そのシート名・行番号・管理Noも自動案内します。\n", "sub_note")

        txt2.insert(tk.END, "■ 【項目不一致】と表示される場合\n", "h2")
        txt2.insert(tk.END, "・管理Noは管理表に存在するものの、図番・品番・品名・仕様・数量のいずれかが食い違っている状態です。\n", "warn")
        txt2.insert(tk.END, "・同じ管理No内で最も近い候補行を特定し、「どの一致していて、どの項目が違ったのか（例: 一致:図番,品名 / 不一致:品番）」を一覧に表示します。\n", "body")

        txt2.insert(tk.END, "\n  3. 現品確認日（物：A列）と請求日（伝票：AI列）の超過判定  \n", "h1")
        txt2.insert(tk.END, "■ 現品確認日（管理表 A列）と 請求日（管理表 AI列）\n", "h2")
        txt2.insert(tk.END, "・セルに日付または確認済みの文字が入力されていれば『確認済 (正常)』と判定します。\n", "badge_ok")
        txt2.insert(tk.END, "・空欄（未入力）のまま注文書納期が判定基準日を過ぎている場合は『未確認 (超過X日)』として警告表示します。\n", "warn")
        txt2.configure(state="disabled")

        # --- タブ3: フィルター・ショートカット一覧 ---
        tab3 = ttk.Frame(notebook, padding=8)
        notebook.add(tab3, text="  ③ 便利機能・ショートカットキー  ")
        txt3 = self._create_scrollable_text(tab3)

        txt3.insert(tk.END, "  Excel風 列フィルター＆ソート機能  \n", "h1")
        txt3.insert(tk.END, "■ 列ヘッダーをクリックして絞り込み・並べ替え\n", "h2")
        txt3.insert(tk.END, "・一覧テーブルの各列タイトル（総合判定、依頼区分、管理No、発注先、メーカー 等）をクリックすると、Excelと同じフィルターメニューが開きます。\n", "body")
        txt3.insert(tk.END, "・値の検索やチェックボックスで特定の業者・管理No・判定結果だけに絞り込むことができます。\n", "body")
        txt3.insert(tk.END, "・フィルター適用中の列には『 [▼] 』マークが付きます。画面右上の『列フィルター解除 (Excel風)』ボタンでいつでも全解除できます。\n", "point")

        txt3.insert(tk.END, "\n  行ダブルクリック詳細比較（インスペクタ機能）  \n", "h1")
        txt3.insert(tk.END, "■ 注文書と管理表の値を左右に並べて比較\n", "h2")
        txt3.insert(tk.END, "・一覧の任意の行を【ダブルクリック】（または選択して『🔍 選択行の不一致・照合詳細を表示』をクリック）すると、項目別の比較画面が開きます。\n", "body")
        txt3.insert(tk.END, "・管理No、図面番号、品番、品名、仕様、メーカー、購入数の7項目について「◎完全一致 / △部分一致 / ×不一致」が色分け表示されます。\n", "body")

        txt3.insert(tk.END, "\n  キーボードショートカット一覧  \n", "h1")
        txt3.insert(tk.END, "・ F1 キー        :  このヘルプ＆操作ガイドを開く\n", "body")
        txt3.insert(tk.END, "・ F5 キー        :  照合・遅延チェックを実行する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + O       :  ① 注文品管理表ファイルを選択する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + D       :  ② 注文書フォルダを選択する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + S       :  検査結果を CSV / Excel ファイルに出力する\n", "body")
        txt3.insert(tk.END, "・ Alt + F4       :  アプリケーションを終了する\n", "body")
        txt3.configure(state="disabled")

        # 指定タブを選択状態にする
        try:
            notebook.select(tab_index)
        except Exception:
            pass

        btn_box = ttk.Frame(win, padding=(14, 8))
        btn_box.pack(fill=tk.X)
        ttk.Button(btn_box, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)

    def _create_scrollable_text(self, parent) -> tk.Text:
        """スクロールバー付きの整形済みTextウィジェットを生成"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(
            frame,
            wrap="word",
            font=("Meiryo UI", 9),
            bg="#ffffff",
            fg="#1e293b",
            padx=14,
            pady=12,
            relief="solid",
            borderwidth=1,
        )
        vsb = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._configure_rich_text_tags(txt)
        return txt

    def _show_about(self):
        """バージョン情報ダイアログの表示"""
        messagebox.showinfo(
            "バージョン情報",
            "StockMatch_Miya-v1.0.0\n"
            "注文品納期 × 管理表（現品確認日・請求日）超過照合システム\n\n"
            "・照合方式: 全体最適1対1マッチング（※単価・金額は照合対象外）\n"
            "・頭文字指定: 管理No頭文字フィルター搭載（デフォルト: 2 ／ カンマ区切りで複数指定可）\n"
            "・自動分別: 2行目の記載により「見積依頼書」を除外し「注文書」のみ照合\n"
            "・品名参照: 管理表の個別「品名」列（O列）を正確に参照\n"
            "・探索範囲: 指定フォルダおよび配下の全サブフォルダを一括検査\n"
            f"・設定保存先: {self.CONFIG_FILE_PATH}\n\n"
            "© StockMatch_Miya System All Rights Reserved.",
        )

    def _open_config_folder(self):
        """アプリ配置フォルダをエクスプローラーで開く"""
        target_dir = self.APP_DIR
        try:
            os.startfile(target_dir)
        except Exception as e:
            messagebox.showinfo("フォルダパス", f"アプリ配置フォルダ:\n{target_dir}\n\n({e})")

    # =========================================================================
    # 設定ファイル（アプリ配置フォルダ直下 config.json）
    # =========================================================================
    def _is_strictly_allowed_path(self, target_filepath: str) -> bool:
        """対象パスがアプリフォルダ直下の config.json であるかを検証"""
        try:
            norm_target = os.path.normcase(os.path.realpath(os.path.abspath(target_filepath)))
            norm_allowed = os.path.normcase(os.path.realpath(os.path.abspath(self.CONFIG_FILE_PATH)))
            return norm_target == norm_allowed
        except Exception:
            return False

    def _save_config(self, show_msg: bool = False):
        """指定された管理表ファイルパス、注文書フォルダパス、管理No頭文字をアプリフォルダ内の config.json に保存"""
        mgmt = self.mgmt_file_path.get().strip().strip('"')
        order = self.order_dir_path.get().strip().strip('"')
        prefix = self.order_no_prefix.get().strip()

        data = {
            "mgmt_file_path": mgmt,
            "order_dir_path": order,
            "order_no_prefix": prefix,
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
                prefix_disp = prefix if prefix else "（指定なし・全件対象）"
                messagebox.showinfo(
                    "設定保存完了",
                    f"パスおよび設定を保存しました:\n{target_file}\n\n・照合管理No頭文字: {prefix_disp}",
                )
        except Exception as e:
            if show_msg:
                messagebox.showerror("保存エラー", f"設定ファイルへの保存に失敗しました:\n{e}")

    def _save_config_manual(self):
        """「パス・設定を保存」ボタン押下時の処理"""
        self._save_config(show_msg=True)

    def _load_config_and_autoload(self):
        """起動時にアプリフォルダ直下の config.json を読み込み"""
        cfg_path = self.CONFIG_FILE_PATH
        if os.path.exists(cfg_path) and self._is_strictly_allowed_path(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                mgmt = data.get("mgmt_file_path", "")
                order = data.get("order_dir_path", "")
                prefix = data.get("order_no_prefix", "2")
                if mgmt and os.path.isfile(mgmt):
                    self.mgmt_file_path.set(mgmt)
                if order and os.path.isdir(order):
                    self.order_dir_path.set(order)
                if prefix is not None:
                    self.order_no_prefix.set(str(prefix))
            except Exception:
                pass

        if not self.mgmt_file_path.get() or not self.order_dir_path.get():
            self._autoload_default_files()

    def _autoload_default_files(self):
        """カレントディレクトリにある管理表ファイルと注文書フォルダを自動補完"""
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
        """① 注文品管理表ファイル（単一ファイル）の選択ダイアログ"""
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
        """② 注文書フォルダの選択ダイアログ"""
        initial_dir = self.order_dir_path.get().strip().strip('"')
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = self.APP_DIR
        folder = filedialog.askdirectory(title="注文書が入っているフォルダを選択", initialdir=initial_dir)
        if folder:
            self.order_dir_path.set(folder)

    def _report_progress(self, percent: float, message: str):
        """バックグラウンドスレッドから進捗バーとステータスメッセージを更新キューへ送る"""
        self.ui_queue.put(("PROGRESS", (percent, message)))

    def _get_target_prefixes(self) -> list[str]:
        """設定された照合管理No頭文字のリストを取得（カンマ区切り・読点・空白区切り対応、全角半角正規化）"""
        raw = self.order_no_prefix.get().strip()
        if not raw:
            return []
        norm = unicodedata.normalize("NFKC", raw).replace("、", ",").replace("，", ",")
        prefixes = [p.strip() for p in norm.split(",") if p.strip()]
        return prefixes

    def _is_order_no_targeted(self, order_no: str, prefixes: list[str]) -> bool:
        """管理Noが指定された頭文字に合致するか判定"""
        if not prefixes:
            return True
        if not order_no:
            return False
        clean_no = str(order_no).strip()
        return any(clean_no.startswith(p) for p in prefixes)

    def _collect_valid_order_files(
        self, folder_path: str, mgmt_filepath: str, prefixes: list[str], progress_cb=None
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

            # ファイル名の先頭8桁から管理Noをチェック（指定頭文字で始まらないファイルは即座に除外）
            norm_entry = unicodedata.normalize("NFKC", entry_name)
            m_pre = re.match(r"^\s*(\d{8})", norm_entry)
            if m_pre:
                file_no_8 = m_pre.group(1)
                if not self._is_order_no_targeted(file_no_8, prefixes):
                    ignored_file_count += 1
                    continue

            try:
                extracted = self._parse_order_book(full_path, prefixes)
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
        """照合比較用に文字列を正規化（スペース・中点・記号・法人格を除去、×やxも統一）"""
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
        """会社名比較用の簡易正規化（法人格や不要記号を除去）"""
        if not text:
            return ""
        s = unicodedata.normalize("NFKC", str(text)).strip()
        for corp in ("株式会社", "有限会社", "(株)", "（株）", "(有)", "（有）", "御中", "様", "殿", "本社", "営業所", "工場"):
            s = s.replace(corp, "")
        s = re.sub(r"[\s\-_・/]", "", s)
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
            if bigram_score >= 0.40:
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

    @staticmethod
    def _extract_8digit_no(text: str) -> str:
        """文字列から8桁の管理Noを抽出（ハイフン区切り等にも対応）"""
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
    # Excel解析ロジック（①注文品管理表 & ②注文書）
    # =========================================================================
    def _parse_management_book(self, filepath: str, progress_cb=None) -> list[dict]:
        """
        『注文品管理表.xlsx』の全シートを高速走査
        ※「品名、ライン名等」（案件名）ではなく、個別の「品名」列（15列目/O列）を正確に品名として取得します。
        """
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
                if "品名" in joined and (
                    "図面" in joined or "品番" in joined or "仕様" in joined or "現品" in joined or "メーカー" in joined
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
                        # 個別品名列の特定（「品名、ライン名等」などの案件総称列を除外し、純粋な「品名」列を取得）
                        elif "ライン名" in v_comb or "品名、ライン名" in v_comb or "品名・ライン名" in v_comb:
                            col_map["line_name"] = c_idx
                        elif ("品名" in v_comb or "名称" in v_comb or "商品名" in v_comb) and "品名、ライン名" not in v_comb:
                            if c_idx >= 10:
                                col_map["item_name"] = c_idx
                            elif "item_name" not in col_map:
                                col_map["item_name"] = c_idx
                        elif "メーカー" in v_comb or "発注先" in v_comb or "仕入先" in v_comb or "業者" in v_comb:
                            col_map.setdefault("maker", c_idx)
                        elif "仕様" in v_comb or "規格" in v_comb or "材質" in v_comb or "寸法" in v_comb:
                            col_map.setdefault("spec", c_idx)
                        elif "購入数" in v_comb or "発注数" in v_comb or v_curr in ("数量", "数", "購入数", "必要数"):
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

            # 管理表標準構造（M列:図番=13, N列:品番=14, O列:品名=15, P列:メーカー=16, Q列:仕様=17, R列:購入数=18）の補完
            if "drawing_no" in col_map:
                base_c = col_map["drawing_no"]
                col_map.setdefault("part_no", base_c + 1)
                col_map.setdefault("item_name", base_c + 2)  # M(13)+2 = O(15)「品名」
                col_map.setdefault("maker", base_c + 3)
                col_map.setdefault("spec", base_c + 4)
                col_map.setdefault("qty", base_c + 5)
                col_map.setdefault("unit_price", base_c + 6)
                col_map.setdefault("amount", base_c + 7)
            elif "part_no" in col_map:
                pc = col_map["part_no"]
                col_map.setdefault("drawing_no", pc - 1)
                col_map.setdefault("item_name", pc + 1)
                col_map.setdefault("maker", pc + 2)
                col_map.setdefault("spec", pc + 3)
                col_map.setdefault("qty", pc + 4)
                col_map.setdefault("unit_price", pc + 5)
                col_map.setdefault("amount", pc + 6)

            if col_map.get("item_name") and col_map["item_name"] < 10 and col_map.get("drawing_no"):
                col_map["item_name"] = col_map["drawing_no"] + 2

            col_map.setdefault("actual_date", 1)
            col_map.setdefault("invoice_date", 35)

            data_slice = raw_rows[header_pos_in_list + 1 :]

            if "order_no" not in col_map and data_slice:
                max_scan_col = min(13, col_map.get("drawing_no", 13))
                best_c = None
                best_hits = 0
                for c_try in range(2, max_scan_col):
                    hits = 0
                    for _, rtuple in data_slice[:80]:
                        if c_try - 1 < len(rtuple):
                            v_try = self._extract_8digit_no(rtuple[c_try - 1])
                            if v_try and not v_try.startswith("202"):
                                hits += 1
                    if hits > best_hits:
                        best_hits = hits
                        best_c = c_try
                if best_c and best_hits >= 1:
                    col_map["order_no"] = best_c

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

                if not any([drawing_val, part_val, item_val, spec_val]):
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
                line_str = self._clean_cell_str(line_val)
                qty_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("qty")))
                unit_price_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("unit_price")))
                amount_str = self._clean_cell_str(_get_col(row_tuple, col_map.get("amount")))

                raw_order_no_str = self._clean_cell_str(order_no_val)
                exact_order_no_8 = self._extract_8digit_no(raw_order_no_str)

                if not exact_order_no_8:
                    max_c = min(13, col_map.get("drawing_no", 13))
                    for c_scan in range(2, max_c):
                        if c_scan == col_map.get("order_date"):
                            continue
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
                        "supplier": "",
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

    def _parse_order_book(self, filepath: str, prefixes: list[str]) -> list[dict]:
        """
        注文書Excelファイルから明細を抽出
        ※2行目を確認し、「見積」や「見積依頼」などの記載がある場合は見積依頼書として完全にスキップし、
          「注文」や「発注」などの記載がある注文書シートのみを照合対象として抽出します。
        ※管理No頭文字フィルターに合致しない注文書は除外します。
        """
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        order_items = []
        filename = os.path.basename(filepath)
        norm_fname = unicodedata.normalize("NFKC", filename)

        file_order_no_8 = ""
        prefix_match = re.match(r"^\s*(\d{8})", norm_fname)
        if prefix_match:
            file_order_no_8 = prefix_match.group(1)
        else:
            file_order_no_8 = self._extract_8digit_no(norm_fname)

        # ファイル名由来の管理Noが指定の頭文字で始まらない場合はスキップ
        if file_order_no_8 and not self._is_order_no_targeted(file_order_no_8, prefixes):
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

            # =================================================================
            # 【重要】2行目における「注文」vs「見積」の分別判定
            # =================================================================
            row2_text = ""
            for r_idx, row_tuple in raw_rows:
                if r_idx == 2:
                    row2_text = " ".join(
                        unicodedata.normalize("NFKC", str(v or "")).strip() for v in row_tuple
                    )
                    break

            # 2行目に「見積」の文字が含まれる場合は見積依頼書とみなし、このシートは完全にスキップ
            if "見積" in row2_text:
                continue

            # 2行目に「注文」「発注」「注文書」「発注書」等の記載がなく、かつファイル名やシートにも見積の気配がある場合はスキップ
            if row2_text and not any(k in row2_text for k in ("注文", "発注")):
                if "見積" in row2_text or "見積" in filename or "見積" in sheet_name:
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

            # シート内で特定された管理Noが指定の頭文字でない場合はスキップ
            if sheet_order_no and not self._is_order_no_targeted(sheet_order_no, prefixes):
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

                # 注記行や非明細行の除外（数量と納期が両方とも空欄の行は除外）
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
                        "order_no_8": file_order_no_8
                        or (sheet_order_no if len(sheet_order_no) == 8 and sheet_order_no.isdigit() else ""),
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
    # 照合・遅延判定＋不一致理由の診断ロジック
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

        target_prefixes = self._get_target_prefixes()

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED, text="照合処理中...")
        self.export_btn.configure(state=tk.DISABLED)
        self.progress_var.set(2.0)
        self.progress_text_var.set("解析を開始しています...")
        prefix_info = f" (対象頭文字: {', '.join(target_prefixes)})" if target_prefixes else " (全件対象)"
        self.status_var.set(f"実行中: ファイルを読み込んでいます{prefix_info}。しばらくお待ちください...")

        worker = threading.Thread(
            target=self._analysis_worker,
            args=(mgmt_path, order_dir, base_date, target_prefixes),
            daemon=True,
        )
        worker.start()
        self.root.after(40, self._poll_ui_queue)

    def _analysis_worker(
        self, mgmt_path: str, order_dir: str, base_date: datetime.date, prefixes: list[str]
    ):
        """バックグラウンドスレッドで実行されるExcel読込・照合処理"""
        try:
            self._report_progress(5.0, "[1/3] 注文品管理表を開いています...")
            mgmt_rows = self._parse_management_book(mgmt_path, progress_cb=self._report_progress)

            self._report_progress(35.0, "[2/3] 注文書フォルダ（サブフォルダ含む）を走査中...")
            order_items, valid_order_files, ignored_files = self._collect_valid_order_files(
                order_dir, mgmt_path, prefixes, progress_cb=self._report_progress
            )

            self._report_progress(75.0, f"[3/3] 明細の最適照合と不一致診断を実行中 ({len(order_items)} 件)...")
            analyzed_records = self._match_and_evaluate(
                mgmt_rows, order_items, base_date, prefixes, progress_cb=self._report_progress
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
                        "prefixes": prefixes,
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
                    order_items_count = payload["order_items_count"]
                    mgmt_rows_count = payload["mgmt_rows_count"]
                    prefixes = payload.get("prefixes", [])
                    p_info = f" (対象頭文字: {', '.join(prefixes)})" if prefixes else " (頭文字制限なし)"

                    self.column_filters.clear()
                    self._update_all_headings()

                    self._update_kpis()
                    self._refresh_table()

                    delay_cnt = sum(1 for r in self.analyzed_records if r["is_any_delay"])
                    unmatched_cnt = sum(1 for r in self.analyzed_records if r["is_unmatched"])
                    self.progress_var.set(100.0)
                    self.progress_text_var.set(
                        f"完了！ (注文書 {valid_order_files}ファイル・{order_items_count}明細を照合 / 管理表 {mgmt_rows_count}行){p_info}"
                    )
                    self.status_var.set(
                        f"照合完了 (基準日: {base_date.strftime('%Y-%m-%d')}{p_info}) — "
                        f"注文書明細: {order_items_count}件 / 超過(遅れ): {delay_cnt}件 / 不一致(未記載): {unmatched_cnt}件 (※行をダブルクリックで不一致箇所の詳細比較を表示)"
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

    def _score_candidate_pair(self, o_item: dict, m_row: dict) -> float:
        """
        同一管理No内の注文書1明細と管理表1行の総合一致スコアを算出する。
        ※単価・金額は照合から除外し、図番・品番・品名・仕様・数量・メーカーの一致度のみで評価します。
        """
        o_sup = o_item.get("norm_supplier", "")
        o_mkr = o_item.get("norm_maker", "")
        m_mkr = m_row.get("norm_maker", "")

        # 発注先/メーカーの判定
        sup_matched = False
        if o_sup and m_mkr:
            if o_sup == m_mkr or o_sup in m_mkr or m_mkr in o_sup:
                sup_matched = True
        if o_mkr and m_mkr:
            if o_mkr == m_mkr or o_mkr in m_mkr or m_mkr in o_mkr:
                sup_matched = True

        # 発注先が特定されており、管理表のメーカー・業者と全く異なる業者名の場合は大幅減点
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

        # 図面番号の一致判定（完全一致のみ強力加点。先頭が似ていても末尾が異なる場合は不一致ペナルティ）
        if o_d and m_d:
            if o_d == m_d:
                score += 500.0
                strong_evidence = True
            else:
                score -= 300.0

        # 品番の一致判定
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

        # 型番・コードのクロス照合（注文書の仕様 ⇔ 管理表の品番 など）
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

        # 品名の一致判定
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

        # 仕様の一致判定
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

        # 数量の一致判定
        if o_item["qty_num"] and m_row["qty_num"]:
            if o_item["qty_num"] == m_row["qty_num"]:
                score += 150.0
                strong_evidence = True
            else:
                score -= 40.0

        # 図番または品番が完全一致していれば十分な証拠とする
        if (o_d and m_d and o_d == m_d) or (o_p and m_p and o_p == m_p):
            strong_evidence = True

        if not strong_evidence:
            return -1.0

        return score

    def _build_field_comparison_data(
        self, o_item: dict, m_row: dict | None, reason_summary: str, extra_notes: str = ""
    ) -> dict:
        """
        注文書の項目と管理表（照合行または最有力候補行）を1項目ずつ比較し、
        どこが一致していてどこが不一致なのかを診断データとして生成する。
        ※単価と金額は照合対象外（IGNORED）として明示。
        """
        rows = []

        def _cmp_field(label: str, o_raw: str, m_raw: str, o_norm: str, m_norm: str):
            o_disp = o_raw if o_raw else "(空欄)"
            m_disp = m_raw if m_raw else "(空欄)"
            if not m_row:
                return (label, o_disp, "— (該当行なし)", "対象なし", "DIFF")
            if not o_norm and not m_norm:
                return (label, o_disp, m_disp, "両方空欄 (-)", "INFO")
            if o_norm == m_norm:
                return (label, o_disp, m_disp, "◎ 完全一致", "MATCH")
            if o_norm and m_norm and (o_norm in m_norm or m_norm in o_norm):
                return (label, o_disp, m_disp, "△ 部分一致", "PARTIAL")
            return (label, o_disp, m_disp, "× 不一致", "DIFF")

        o_no = o_item.get("order_no_8") or o_item.get("order_no", "")
        m_no = (m_row.get("order_no_8") or m_row.get("order_no", "")) if m_row else ""
        rows.append(_cmp_field("① 管理No (8桁)", o_no, m_no, o_no, m_no))

        rows.append(
            _cmp_field(
                "② 図面番号",
                o_item.get("drawing_no", ""),
                m_row.get("drawing_no", "") if m_row else "",
                o_item.get("norm_drawing", ""),
                m_row.get("norm_drawing", "") if m_row else "",
            )
        )
        rows.append(
            _cmp_field(
                "③ 品番",
                o_item.get("part_no", ""),
                m_row.get("part_no", "") if m_row else "",
                o_item.get("norm_part", ""),
                m_row.get("norm_part", "") if m_row else "",
            )
        )
        rows.append(
            _cmp_field(
                "④ 品名",
                o_item.get("item_name", ""),
                m_row.get("item_name", "") if m_row else "",
                o_item.get("norm_item", ""),
                m_row.get("norm_item", "") if m_row else "",
            )
        )
        rows.append(
            _cmp_field(
                "⑤ 仕様・規格",
                o_item.get("spec", ""),
                m_row.get("spec", "") if m_row else "",
                o_item.get("norm_spec", ""),
                m_row.get("norm_spec", "") if m_row else "",
            )
        )
        rows.append(
            _cmp_field(
                "⑥ メーカー",
                o_item.get("maker", ""),
                m_row.get("maker", "") if m_row else "",
                o_item.get("norm_maker", ""),
                m_row.get("norm_maker", "") if m_row else "",
            )
        )
        rows.append(
            _cmp_field(
                "⑦ 購入数 (数量)",
                o_item.get("qty", ""),
                m_row.get("qty", "") if m_row else "",
                o_item.get("qty_num", ""),
                m_row.get("qty_num", "") if m_row else "",
            )
        )

        # 単価と金額は照合対象外（参考表示）
        o_p_disp = o_item.get("unit_price", "") or "(空欄)"
        m_p_disp = (m_row.get("unit_price", "") if m_row else "") or "(空欄)"
        rows.append(("⑧ 単価", o_p_disp, m_p_disp, "— (照合対象外)", "IGNORED"))

        o_a_disp = o_item.get("amount", "") or "(空欄)"
        m_a_disp = (m_row.get("amount", "") if m_row else "") or "(空欄)"
        rows.append(("⑨ 金額", o_a_disp, m_a_disp, "— (照合対象外)", "IGNORED"))

        cross_notes = []
        if m_row:
            if o_item["norm_spec"] and m_row["norm_part"] and o_item["norm_spec"] == m_row["norm_part"]:
                cross_notes.append("・注文書の『仕様』と管理表の『品番』が一致しています。")
            if o_item["norm_part"] and m_row["norm_spec"] and o_item["norm_part"] == m_row["norm_spec"]:
                cross_notes.append("・注文書の『品番』と管理表の『仕様』が一致しています。")
            if m_row.get("line_name"):
                cross_notes.append(f"・管理表D列の『品名、ライン名等』: {m_row['line_name']}")

        detail_lines = [
            f"【診断結果サマリー】\n  {reason_summary}",
            "",
            f"【注文書データ元】\n  ファイル: {o_item['file']} (シート: {o_item['sheet']} / 行: {o_item['row']}) / 依頼区分: {o_item['req_type_label']}",
        ]
        if m_row:
            detail_lines.append(
                f"【管理表の比較対象行】\n  シート: {m_row['sheet']} (行: {m_row['row']}) / 管理No: {m_row.get('order_no_8') or m_row.get('order_no')}"
            )
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
        all_mgmt_order_nos_8: set,
    ) -> tuple[str, dict]:
        """なぜ注文書明細が管理表と一致しなかったのかを診断"""
        o_no = o_item.get("order_no_8", "")

        if not o_no:
            reason = f"【管理No取得不可】ファイル名『{o_item['file']}』の先頭から8桁の管理Noが検出できません"
            diag = self._build_field_comparison_data(
                o_item,
                None,
                reason,
                "注文書ファイル名の先頭8桁が数字（例: 22609030...）になっているか確認してください。",
            )
            return reason, diag

        if o_no not in all_mgmt_order_nos_8:
            best_other_row = None
            best_other_sc = -999.0
            for m in mgmt_rows:
                sc = self._score_candidate_pair(o_item, m)
                if sc > best_other_sc:
                    best_other_sc = sc
                    best_other_row = m

            similar_nos = difflib.get_close_matches(o_no, sorted(all_mgmt_order_nos_8), n=3, cutoff=0.6)
            sim_str = f" (管理表内の類似No: {', '.join(similar_nos)})" if similar_nos else ""

            if best_other_row and best_other_sc >= 150.0:
                other_no = best_other_row.get("order_no_8") or best_other_row.get("order_no") or "空欄"
                reason = (
                    f"【管理No不一致】注文書の管理No[{o_no}]が管理表に未登録{sim_str} "
                    f"※管理表『{best_other_row['sheet']}』行{best_other_row['row']}に同一品がありますが管理Noが[{other_no}]です"
                )
                extra = (
                    f"・注文書ファイル名の管理Noは [{o_no}] ですが、管理表のどのシートにも [{o_no}] が存在しません。\n"
                    f"・管理表の『{best_other_row['sheet']}』シート 行{best_other_row['row']} に同じ品目のデータが見つかりました（管理表側の管理Noは [{other_no}] です）。"
                )
                diag = self._build_field_comparison_data(o_item, best_other_row, reason, extra)
            else:
                sample_nos = ", ".join(sorted(list(all_mgmt_order_nos_8))[:15])
                reason = f"【管理No未登録】管理No[{o_no}]が管理表(全シート)に1件も存在しません{sim_str}"
                extra = (
                    f"・注文書の管理No [{o_no}] を注文品管理表の全シートから検索しましたが、1行も見つかりませんでした。\n"
                    f"・管理表側で検出された管理No一覧（抜粋）: {sample_nos}\n"
                    f"・管理表に該当の管理No（{o_no}）が記入されているか確認してください。"
                )
                diag = self._build_field_comparison_data(o_item, None, reason, extra)
            return reason, diag

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
            _check_diff("数量", o_item["qty"], best_cand["qty"], o_item["qty_num"], best_cand["qty_num"])

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
                    f"   {idx_c}. {r_c['sheet']}(行{r_c['row']}){used_mark} | 業者:{r_c.get('maker')} | 図番:{r_c['drawing_no']} | 品番:{r_c['part_no']} | 品名:{r_c['item_name']} | 仕様:{r_c['spec']} | 数量:{r_c['qty']}"
                )

            diag = self._build_field_comparison_data(
                o_item, best_cand, reason, "\n".join(cand_list_lines)
            )
            return reason, diag

        reason = f"【明細未記載】管理No[{o_no}]はありますが一致する明細行がありません"
        diag = self._build_field_comparison_data(o_item, None, reason)
        return reason, diag

    def _match_and_evaluate(
        self,
        mgmt_rows: list[dict],
        order_items: list[dict],
        base_date: datetime.date,
        prefixes: list[str],
        progress_cb=None,
    ) -> list[dict]:
        """
        同一管理No内で全明細と管理表行を同時に評価し、全体最適ペアを決定する。
        ※指定頭文字フィルター（prefixes）に適合する管理Noのみを照合対象とします。
        """
        results = []

        all_mgmt_order_nos_8 = set()
        mgmt_by_no = {}
        for m in mgmt_rows:
            no8 = m.get("order_no_8") or m.get("norm_order_no")
            if no8:
                all_mgmt_order_nos_8.add(no8)
                mgmt_by_no.setdefault(no8, []).append(m)

        # 注文書明細を管理Noごとにグループ化（指定頭文字に合致するもののみ）
        orders_by_no = {}
        unmatched_orders_no_invalid = []

        for o_item in order_items:
            o_no = o_item.get("order_no_8", "")
            if not self._is_order_no_targeted(o_no, prefixes):
                continue
            if not o_no or o_no not in all_mgmt_order_nos_8:
                unmatched_orders_no_invalid.append(o_item)
            else:
                orders_by_no.setdefault(o_no, []).append(o_item)

        total_orders = sum(len(l) for l in orders_by_no.values()) + len(unmatched_orders_no_invalid)
        processed_orders = 0

        # 管理Noごとに全体最適マッチングを実施
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
                    # 依頼区分（購入 vs 外注）の一致によるボーナス
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

            # 結果レコードの生成
            for r_idx, o_item in enumerate(o_list):
                processed_orders += 1
                if progress_cb and (processed_orders == 1 or processed_orders % 5 == 0 or processed_orders == total_orders):
                    pct = 75.0 + (23.0 * processed_orders / max(1, total_orders))
                    progress_cb(pct, f"[3/3] 注文書明細と管理表を照合・診断中 ({processed_orders}/{total_orders} 件目)...")

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
                        mismatch_reason=match_summary,
                        diagnostic_data=diag_data,
                    )
                else:
                    reason_txt, diag_data = self._diagnose_unmatched_item(o_item, mgmt_rows, all_mgmt_order_nos_8)
                    rec = self._evaluate_single_record(
                        req_type=o_item.get("req_type_label", "注文"),
                        order_no=o_no,
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
                        source_info=f"【不一致】注文書:{o_item['file']}(行{o_item['row']})",
                        base_date=base_date,
                        is_unmatched=True,
                        mismatch_reason=reason_txt,
                        diagnostic_data=diag_data,
                    )
                results.append(rec)

        # 管理No自体が見つからなかった明細の処理
        for o_item in unmatched_orders_no_invalid:
            processed_orders += 1
            reason_txt, diag_data = self._diagnose_unmatched_item(o_item, mgmt_rows, all_mgmt_order_nos_8)
            rec = self._evaluate_single_record(
                req_type=o_item.get("req_type_label", "注文"),
                order_no=o_item.get("order_no_8") or o_item.get("order_no", ""),
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
                source_info=f"【不一致】注文書:{o_item['file']}(行{o_item['row']})",
                base_date=base_date,
                is_unmatched=True,
                mismatch_reason=reason_txt,
                diagnostic_data=diag_data,
            )
            results.append(rec)

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
        mismatch_reason: str = "",
        diagnostic_data: dict | None = None,
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
            due_str = f"({due_date.strftime('%Y-%m-%d')}発注)"

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
        """上部KPI集計カードの数値更新"""
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

            self.tree.insert(
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
                    "不一致理由・照合診断": r["mismatch_reason"],
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