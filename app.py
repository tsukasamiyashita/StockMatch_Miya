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
    StockMatch_Miya-v1.4.0
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

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("StockMatch_Miya-v1.4.0 - 注文品納期・現品/請求照合システム")

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
        """一覧から除外する管理Noを専用ダイアログで行区切り編集する（Excel複数セル貼付対応）"""
        win = tk.Toplevel(self.root)
        win.title("一覧から除外する管理Noの設定")
        self._apply_window_icon(win)
        win.geometry("540x640")
        win.minsize(460, 500)
        win.transient(self.root)
        win.grab_set()

        # 画面中央に配置
        try:
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            x = rx + max(0, (rw - 540) // 2)
            y = ry + max(0, (rh - 640) // 2)
            win.geometry(f"540x640+{x}+{y}")
        except Exception:
            pass

        # 1. 上部ヘッダーバナー
        banner = tk.Frame(win, bg="#1e293b", padx=16, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text="🚫 一覧から除外する管理Noの設定（行区切り入力）",
            bg="#1e293b",
            fg="#ffffff",
            font=("Meiryo UI", 11, "bold"),
            anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            banner,
            text="照合結果や注文書未検出一覧から除外したい管理Noを、1行に1つずつ入力してください。\n★ Excelでコピーした複数セル（列全体や複数行）をそのまま貼り付け（Ctrl+V）できます。",
            bg="#1e293b",
            fg="#cbd5e1",
            font=("Meiryo UI", 9),
            anchor="w",
            justify="left",
        ).pack(fill=tk.X, pady=(4, 0))

        # 2. クイック操作ツールバー
        toolbar = tk.Frame(win, bg="#f1f5f9", padx=14, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
        toolbar.pack(fill=tk.X)

        # 3. テキスト入力エリア
        txt_frame = tk.Frame(win, padx=14, pady=10, bg="#f8fafc")
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(
            txt_frame,
            wrap="none",
            font=("Consolas", 10),
            padx=10,
            pady=8,
            relief="solid",
            borderwidth=1,
            undo=True,
        )
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        hsb = ttk.Scrollbar(txt_frame, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        txt.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        txt_frame.grid_rowconfigure(0, weight=1)
        txt_frame.grid_columnconfigure(0, weight=1)

        # 現在の値を改行区切りでセット
        current_list = self._get_excluded_order_nos_list()
        if current_list:
            txt.insert("1.0", "\n".join(current_list) + "\n")

        # 4. リアルタイム件数表示バー
        info_bar = tk.Frame(win, bg="#f8fafc", padx=16, pady=4)
        info_bar.pack(fill=tk.X)
        lbl_info = tk.Label(
            info_bar,
            text=f"現在の有効な管理No: {len(current_list)} 件",
            bg="#f8fafc",
            fg="#1d4ed8",
            font=("Meiryo UI", 9, "bold"),
            anchor="w",
        )
        lbl_info.pack(side=tk.LEFT)

        def _update_count(*_):
            content = txt.get("1.0", tk.END)
            items = self._parse_lines_to_order_nos(content)
            lbl_info.configure(text=f"現在の有効な管理No: {len(items)} 件（重複除外後）")

        txt.bind("<KeyRelease>", _update_count)

        # ツールバーのアクション
        def _cleanup_text():
            content = txt.get("1.0", tk.END)
            items = self._parse_lines_to_order_nos(content)
            txt.delete("1.0", tk.END)
            if items:
                txt.insert("1.0", "\n".join(items) + "\n")
            _update_count()

        def _paste_clipboard():
            try:
                clip = win.clipboard_get()
                if clip:
                    cur = txt.get("1.0", tk.END).rstrip()
                    if cur:
                        txt.insert(tk.END, "\n" + clip.strip() + "\n")
                    else:
                        txt.insert("1.0", clip.strip() + "\n")
                    _cleanup_text()
            except Exception as e:
                messagebox.showwarning("貼付エラー", f"クリップボードの読み取りに失敗しました:\n{e}", parent=win)

        def _clear_all():
            txt.delete("1.0", tk.END)
            _update_count()

        ttk.Button(toolbar, text="📋 クリップボードから貼り付け", style="Action.TButton", command=_paste_clipboard).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar, text="✨ 重複・空白を整理", style="Action.TButton", command=_cleanup_text).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar, text="🗑 すべてクリア", style="Action.TButton", command=_clear_all).pack(
            side=tk.LEFT
        )

        # 5. 下部ボタンバー
        btn_bar = tk.Frame(win, bg="#f1f5f9", padx=16, pady=10, highlightbackground="#cbd5e1", highlightthickness=1)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        def _apply_and_save():
            content = txt.get("1.0", tk.END)
            items = self._parse_lines_to_order_nos(content)
            # 改行区切りで格納
            self.excluded_order_nos_str.set("\n".join(items))
            self._update_excluded_summary_badge()
            self._update_kpis()
            self._refresh_main_table()
            self._refresh_missing_table()
            # 設定ファイルへの保存も自動実行
            self._save_config(show_msg=False)
            win.destroy()
            if items:
                self.status_var.set(
                    f"除外管理No ({len(items)}件) を設定・保存しました。一覧および件数サマリーに反映済みです。"
                )
                messagebox.showinfo(
                    "設定完了",
                    f"除外する管理No（計 {len(items)} 件）を設定し、設定ファイル(config.json)に保存しました。\n"
                    "一覧および件数サマリーから該当の管理Noが除外されています。",
                    parent=self.root,
                )
            else:
                self.status_var.set("除外管理Noの設定をすべて解除・保存しました。")
                messagebox.showinfo(
                    "解除完了",
                    "除外管理Noの設定をすべて解除し、設定ファイル(config.json)に保存しました。",
                    parent=self.root,
                )

        ttk.Button(btn_bar, text="✔ 適用して設定を保存", style="Primary.TButton", command=_apply_and_save).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        ttk.Button(btn_bar, text="キャンセル", style="Action.TButton", command=win.destroy).pack(side=tk.RIGHT)

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
            label="現在の設定を保存 (パス・超過日数・経過除外月数・指示書フィルタ・除外管理No)",
            command=self._save_config_manual,
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

        title_lbl = ttk.Label(header_frame, text="StockMatch_Miya-v1.4.0", style="Header.TLabel")
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

        self.col_titles_main = {col_id: title for col_id, title, _, _ in self.col_configs_main}

        for col_id, _, width, anchor in self.col_configs_main:
            self._update_heading_display_main(col_id)
            self.tree_main.column(col_id, width=width, anchor=anchor)

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
    def _open_excel_at_location(
        self,
        filepath: str,
        sheet_name: str = "",
        row_num: int | None = None,
        label_for_status: str = "Excelファイル",
        read_only_strict: bool = False,
    ):
        """
        指定されたExcelファイルを開き、該当シート・該当行を選択状態にして前面表示する。
        read_only_strict=True（注文書）の場合は以下の多重誤保存ブロック機構を適用する:
          1. Excel COM の Workbooks.Open で ReadOnly=True, IgnoreReadOnlyRecommended=True,
             UpdateLinks=0, Notify=False, AddToMru=False を指定して強制的に読み取り専用で開く。
          2. すでにExcel上で該当注文書が通常モード（編集可能）で開かれていた場合でも、
             ChangeFileAccess(xlReadOnly=3) により読み取り専用モードへ強制切り替えする。
          3. 開いた後のメモリ上の全シートおよびブック構造に保護ロック（Protect）をかけ、
             万が一のキー入力・セル書き換え操作自体をブロックする（行選択・スクロール・コピーは可能）。
          4. target_wb.Saved = True をセットし、=TODAY()等の自動再計算による「変更を保存しますか？」
             確認ダイアログ自体の発生を抑止する。
          5. os.startfile による通常モード（編集可能）でのフォールバック起動を完全に禁止する。
        """
        clean_path = filepath.strip().strip('"') if filepath else ""
        if not clean_path or not os.path.isfile(clean_path):
            messagebox.showwarning(
                "ファイル未検出",
                f"対象の{label_for_status}が見つかりません:\n{clean_path or '(未指定)'}",
            )
            return

        abs_path = os.path.abspath(clean_path)
        norm_target = os.path.normcase(os.path.realpath(abs_path))

        if not HAS_WIN32COM:
            if read_only_strict:
                messagebox.showerror(
                    "読み取り専用保護エラー",
                    "Excel連携コンポーネント(pywin32)が利用できないため、注文書を安全な『読み取り専用・保存ブロックモード』で開けません。\n\n"
                    "誤って上書き保存してしまう事故を厳重に防止するため、通常モードでの直接起動をブロックしました。\n"
                    "『🔍 照合詳細・プレビュー』画面内の『② 📄 注文書シート参照 (プレビュー)』タブから安全に内容を確認してください。",
                )
                return
            try:
                os.startfile(abs_path)
            except Exception as e:
                messagebox.showerror("起動エラー", f"ファイルを開けませんでした:\n{e}")
            return

        ro_badge = " 【読み取り専用・保存ブロック】" if read_only_strict else ""
        loc_desc = f" (シート:{sheet_name}" + (f" / 行:{row_num})" if row_num else ")") if sheet_name else ""
        self.status_var.set(f"{label_for_status}{ro_badge}をExcelで開いています: {os.path.basename(abs_path)}{loc_desc} ...")

        def _worker():
            try:
                pythoncom.CoInitialize()
                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                except Exception:
                    excel = win32com.client.Dispatch("Excel.Application")

                excel.Visible = True

                target_wb = None
                try:
                    for wb_cand in excel.Workbooks:
                        try:
                            cand_full = os.path.normcase(os.path.realpath(os.path.abspath(wb_cand.FullName)))
                            if cand_full == norm_target:
                                target_wb = wb_cand
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

                if target_wb is None:
                    if read_only_strict:
                        # UpdateLinks=0, ReadOnly=True, Format=None, Password=None, WriteResPassword=None,
                        # IgnoreReadOnlyRecommended=True, Origin=None, Delimiter=None, Editable=False,
                        # Notify=False, Converter=None, AddToMru=False
                        target_wb = excel.Workbooks.Open(
                            abs_path,
                            0,
                            True,
                            None,
                            None,
                            None,
                            True,
                            None,
                            None,
                            False,
                            False,
                            None,
                            False,
                        )
                    else:
                        target_wb = excel.Workbooks.Open(abs_path)
                else:
                    # すでにExcelで開かれているブックが読み取り専用でない場合、注文書は強制的に読み取り専用へ切り替える
                    if read_only_strict:
                        try:
                            if not bool(target_wb.ReadOnly):
                                prev_alerts = excel.DisplayAlerts
                                excel.DisplayAlerts = False
                                try:
                                    target_wb.Saved = True
                                    target_wb.ChangeFileAccess(3)  # xlReadOnly = 3
                                finally:
                                    excel.DisplayAlerts = prev_alerts
                        except Exception:
                            pass

                target_wb.Activate()

                target_ws = None
                if sheet_name:
                    try:
                        target_ws = target_wb.Sheets(sheet_name)
                        target_ws.Activate()
                    except Exception:
                        target_ws = target_wb.ActiveSheet
                else:
                    target_ws = target_wb.ActiveSheet

                if target_ws is not None and row_num and int(row_num) >= 1:
                    r_int = int(row_num)
                    try:
                        target_ws.Rows(r_int).Select()
                        excel.ActiveWindow.ScrollRow = max(1, r_int - 6)
                    except Exception:
                        pass

                # 注文書の場合はメモリ上の全シートに誤編集・誤保存ブロック保護を適用し、Saved=Trueで保存確認ダイアログも抑止
                if read_only_strict:
                    try:
                        for ws_i in target_wb.Worksheets:
                            try:
                                # EnableSelection = 0 (xlNoRestrictions: セル選択・行ジャンプ・コピーは許可しつつ編集のみ禁止)
                                ws_i.EnableSelection = 0
                                ws_i.Protect(
                                    Password="StockMatch_ReadOnly_Guard",
                                    DrawingObjects=True,
                                    Contents=True,
                                    Scenarios=True,
                                    UserInterfaceOnly=False,
                                    AllowFormattingCells=False,
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
                        # メモリ上の保護設定や=TODAY()再計算によるDirtyフラグをクリアし、閉じる際の「保存しますか？」ポップアップ自体を抑止
                        target_wb.Saved = True
                    except Exception:
                        pass

                try:
                    if excel.WindowState == -4140:  # xlMinimized
                        excel.WindowState = -4137  # xlMaximized
                except Exception:
                    pass

                self.root.after(
                    0,
                    lambda: self.status_var.set(
                        f"{label_for_status}{ro_badge}を表示しました: {os.path.basename(abs_path)}{loc_desc}"
                    ),
                )
            except Exception as e:
                if read_only_strict:
                    # 注文書の場合は os.startfile による通常モード（編集可能）起動へのフォールバックを絶対に行わない
                    self.root.after(
                        0,
                        lambda err=str(e): messagebox.showerror(
                            "注文書 読み取り専用起動エラー",
                            f"注文書を安全な『読み取り専用・保存ブロックモード』で開けませんでした:\n{err}\n\n"
                            "※誤って上書き保存してしまう事故を防止するため、通常モードでの起動を厳重にブロックしました。\n"
                            "詳細ウィンドウの『② 📄 注文書シート参照 (プレビュー)』タブから安全に確認できます。",
                        ),
                    )
                else:
                    try:
                        os.startfile(abs_path)
                    except Exception as e2:
                        self.root.after(
                            0,
                            lambda err=str(e2): messagebox.showerror(
                                "Excel起動エラー", f"ファイルを開けませんでした:\n{err}"
                            ),
                        )
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=_worker, daemon=True).start()

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

    def _open_selected_row_detail(self, is_missing_tab: bool):
        """選択されている行の不一致理由・項目比較および注文書/管理表プレビューダイアログを開く"""
        rec = self._get_selected_record(is_missing_tab=is_missing_tab)
        if not rec:
            return
        diag = rec.get("diagnostic_data", {})
        order_path = diag.get("order_full_path", "")
        order_file = diag.get("order_file", "")
        order_sheet = diag.get("order_sheet", "")
        order_row = diag.get("order_row")
        mgmt_sheet = diag.get("mgmt_sheet") or rec.get("mgmt_sheet", "")
        mgmt_row = diag.get("mgmt_row") or (
            int(rec["mgmt_row"]) if str(rec.get("mgmt_row", "")).isdigit() else None
        )
        order_no_val = rec.get("order_no", "")

        win = tk.Toplevel(self.root)
        win.title(f"詳細・データ参照 — 管理No: {order_no_val} / 品名: {rec['item_name']}")
        self._apply_window_icon(win)
        win.geometry("1020x720")
        win.minsize(820, 560)
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

        # クイック参照ツールバー（Excelで該当行へジャンプ・フォルダ参照）
        ref_bar = tk.Frame(win, bg="#f1f5f9", padx=14, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
        ref_bar.pack(fill=tk.X)

        tk.Label(
            ref_bar,
            text="直接参照:",
            bg="#f1f5f9",
            fg="#1e293b",
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 8))

        if order_path and os.path.isfile(order_path):
            ttk.Button(
                ref_bar,
                text=f"📄 注文書をExcelで開く 【読取専用・保存禁止】 ({order_file} 行{order_row})",
                style="Action.TButton",
                command=lambda: self._open_excel_at_location(
                    order_path,
                    sheet_name=order_sheet,
                    row_num=order_row,
                    label_for_status="注文書",
                    read_only_strict=True,
                ),
            ).pack(side=tk.LEFT, padx=(0, 8))
            ttk.Button(
                ref_bar,
                text="📂 注文書の保存フォルダを開く",
                style="Action.TButton",
                command=lambda: self._open_order_file_parent_folder(order_path),
            ).pack(side=tk.LEFT, padx=(0, 8))

        mgmt_btn_label = (
            f"📋 注文品管理表をExcelで開く ({mgmt_sheet} 行{mgmt_row})"
            if mgmt_sheet and mgmt_row
            else "📋 注文品管理表をExcelで開く"
        )
        ttk.Button(
            ref_bar,
            text=mgmt_btn_label,
            style="Action.TButton",
            command=lambda: self._open_selected_row_mgmt(is_missing_tab=is_missing_tab),
        ).pack(side=tk.LEFT, padx=(0, 8))

        # 3つのタブで「項目比較」「注文書プレビュー」「管理表プレビュー」を切り替え可能にする
        detail_nb = ttk.Notebook(win)
        detail_nb.pack(fill=tk.BOTH, expand=True, padx=14, pady=(10, 6))

        # --- タブ1: 項目別比較・診断レポート ---
        tab_cmp = ttk.Frame(detail_nb, padding=10)
        detail_nb.add(tab_cmp, text="  ① 🔍 項目比較・診断レポート  ")

        lbl_cmp = ttk.Label(
            tab_cmp,
            text="■ 注文書 と 注文品管理表 の項目別詳細（単価・金額は照合対象外）",
            font=("Meiryo UI", 9, "bold"),
        )
        lbl_cmp.pack(anchor="w", pady=(0, 6))

        cmp_cols = ("field", "order_val", "mgmt_val", "judge")
        cmp_tree = ttk.Treeview(tab_cmp, columns=cmp_cols, show="headings", height=10)
        cmp_tree.heading("field", text="比較項目")
        cmp_tree.heading("order_val", text="注文書の記載内容")
        cmp_tree.heading("mgmt_val", text="注文品管理表の記載内容")
        cmp_tree.heading("judge", text="照合判定")

        cmp_tree.column("field", width=160, anchor="w")
        cmp_tree.column("order_val", width=310, anchor="w")
        cmp_tree.column("mgmt_val", width=310, anchor="w")
        cmp_tree.column("judge", width=140, anchor="center")

        cmp_tree.tag_configure("MATCH", background="#dcfce7", foreground="#166534")
        cmp_tree.tag_configure("PARTIAL", background="#fef9c3", foreground="#854d0e")
        cmp_tree.tag_configure("DIFF", background="#fee2e2", foreground="#991b1b")
        cmp_tree.tag_configure("INFO", background="#f1f5f9", foreground="#334155")
        cmp_tree.tag_configure("IGNORED", background="#f8fafc", foreground="#64748b")

        cmp_rows = diag.get("comparison_rows", [])
        for f_name, o_v, m_v, j_txt, tag in cmp_rows:
            cmp_tree.insert("", tk.END, values=(f_name, o_v, m_v, j_txt), tags=(tag,))

        cmp_tree.pack(fill=tk.X, pady=(0, 10))

        lbl_rep = ttk.Label(
            tab_cmp,
            text="■ 詳細診断レポート",
            font=("Meiryo UI", 9, "bold"),
        )
        lbl_rep.pack(anchor="w", pady=(0, 4))

        txt_frame = ttk.Frame(tab_cmp)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(txt_frame, wrap="word", font=("Meiryo UI", 9), padx=10, pady=8, relief="solid", borderwidth=1)
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        txt.insert("1.0", diag.get("detail_text", rec.get("mismatch_reason", "")))
        txt.configure(state="disabled")

        # --- タブ2: 注文書シート参照 (アプリ内プレビュー) ---
        tab_order_prev = ttk.Frame(detail_nb, padding=10)
        detail_nb.add(tab_order_prev, text="  ② 📄 注文書シート参照 (プレビュー)  ")
        self._build_order_sheet_preview_tab(tab_order_prev, order_path, order_sheet, order_row)

        # --- タブ3: 注文品管理表参照 (アプリ内プレビュー) ---
        tab_mgmt_prev = ttk.Frame(detail_nb, padding=10)
        detail_nb.add(tab_mgmt_prev, text="  ③ 📋 注文品管理表参照 (同一管理No・周辺行)  ")
        self._build_mgmt_sheet_preview_tab(tab_mgmt_prev, order_no_val, mgmt_sheet, mgmt_row, is_missing_tab)

        btn_bar = ttk.Frame(win, padding=(14, 8))
        btn_bar.pack(fill=tk.X)
        ttk.Button(btn_bar, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)

    def _build_order_sheet_preview_tab(
        self, parent: ttk.Frame, order_path: str, order_sheet: str, target_row: int | None
    ):
        """注文書Excelシートの内容をアプリ内のテーブルでそのまま参照・確認できるプレビュータブ"""
        if not order_path or not os.path.isfile(order_path):
            ttk.Label(
                parent,
                text="この明細は『注文書未検出（管理表のみ記載）』のため、対応する注文書ファイルが存在しません。",
                font=("Meiryo UI", 10, "bold"),
                foreground="#92400e",
            ).pack(anchor="w", pady=20)
            return

        info_lbl = ttk.Label(
            parent,
            text=f"■ 注文書ファイル: {os.path.basename(order_path)}   |   シート: {order_sheet}   |   該当明細行: 行{target_row or '-'} （★黄色ハイライト行）",
            font=("Meiryo UI", 9, "bold"),
        )
        info_lbl.pack(anchor="w", pady=(0, 6))

        tbl_frame = ttk.Frame(parent)
        tbl_frame.pack(fill=tk.BOTH, expand=True)

        max_preview_cols = 14
        col_ids = ["row_no"] + [f"c{i}" for i in range(1, max_preview_cols + 1)]
        tree = ttk.Treeview(tbl_frame, columns=col_ids, show="headings", selectmode="browse")
        tree.heading("row_no", text="行")
        tree.column("row_no", width=55, anchor="center")

        for i in range(1, max_preview_cols + 1):
            col_letter = chr(64 + i)
            tree.heading(f"c{i}", text=f"{col_letter}列")
            tree.column(f"c{i}", width=110, anchor="w")

        tree.tag_configure("TARGET_ROW", background="#fef08a", foreground="#713f12")
        tree.tag_configure("ORDER_DATE_ROW", background="#e0f2fe", foreground="#0369a1")
        tree.tag_configure("HEADER_ROW", background="#e2e8f0", foreground="#0f172a")

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.grid_rowconfigure(0, weight=1)
        tbl_frame.grid_columnconfigure(0, weight=1)

        wb = None
        try:
            wb = openpyxl.load_workbook(order_path, data_only=True, read_only=True)
            ws = wb[order_sheet] if order_sheet in wb.sheetnames else wb.active
            target_iid = None
            for r_idx, row_tuple in enumerate(ws.iter_rows(max_row=120, max_col=30, values_only=True), start=1):
                if not row_tuple or not any(v is not None and str(v).strip() != "" for v in row_tuple):
                    continue
                non_empty_vals = [
                    self._clean_cell_str(v)
                    for v in row_tuple
                    if v is not None and str(v).strip() != ""
                ]
                if not non_empty_vals:
                    continue
                padded = (non_empty_vals + [""] * max_preview_cols)[:max_preview_cols]
                row_label = f"★ {r_idx}" if target_row and r_idx == int(target_row) else str(r_idx)
                tags = ()
                if target_row and r_idx == int(target_row):
                    tags = ("TARGET_ROW",)
                elif r_idx == 4:
                    tags = ("ORDER_DATE_ROW",)
                elif any("品名" in v or "図面" in v or "数量" in v for v in non_empty_vals):
                    tags = ("HEADER_ROW",)

                iid = f"ord_r_{r_idx}"
                tree.insert("", tk.END, iid=iid, values=[row_label] + padded, tags=tags)
                if target_row and r_idx == int(target_row):
                    target_iid = iid
            if target_iid:
                tree.selection_set(target_iid)
                tree.see(target_iid)
        except Exception as e:
            tree.insert("", tk.END, values=["エラー", f"注文書プレビューの読み込みに失敗しました: {e}"] + [""] * (max_preview_cols - 1))
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    def _build_mgmt_sheet_preview_tab(
        self,
        parent: ttk.Frame,
        order_no_val: str,
        mgmt_sheet: str,
        target_row: int | None,
        is_missing_tab: bool,
    ):
        """注文品管理表の同一管理No行および周辺行をアプリ内で即座に一覧参照できるプレビュータブ"""
        info_lbl = ttk.Label(
            parent,
            text=(
                f"■ 注文品管理表 抽出プレビュー   |   対象管理No: {order_no_val or '-'}   |   "
                f"照合先: {mgmt_sheet or '-'} (行{target_row or '-'})  ※ダブルクリックでExcelの該当行を開きます"
            ),
            font=("Meiryo UI", 9, "bold"),
        )
        info_lbl.pack(anchor="w", pady=(0, 6))

        tbl_frame = ttk.Frame(parent)
        tbl_frame.pack(fill=tk.BOTH, expand=True)

        cols = (
            "mark",
            "sheet",
            "row",
            "order_no",
            "instruction_disp",
            "ship_to",
            "line_name",
            "product_detail",
            "actual_raw",
            "drawing_no",
            "part_no",
            "item_name",
            "maker",
            "spec",
            "qty",
            "mgmt_due_raw",
            "invoice_raw",
            "person",
        )
        tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="browse")
        col_defs = [
            ("mark", "対象", 65, "center"),
            ("sheet", "シート", 85, "center"),
            ("row", "行", 55, "center"),
            ("order_no", "管理No", 95, "center"),
            ("instruction_disp", "指示書(J列)", 80, "center"),
            ("ship_to", "出荷先", 100, "w"),
            ("line_name", "品名・ライン名等", 130, "w"),
            ("product_detail", "製品詳細", 130, "w"),
            ("actual_raw", "現品確認(A列)", 95, "center"),
            ("drawing_no", "図面番号", 110, "w"),
            ("part_no", "品番", 110, "w"),
            ("item_name", "品名", 140, "w"),
            ("maker", "メーカー/業者", 110, "w"),
            ("spec", "仕様・規格", 130, "w"),
            ("qty", "購入数(AB列)", 85, "center"),
            ("mgmt_due_raw", "管理表納期", 90, "center"),
            ("invoice_raw", "請求日(AI列)", 95, "center"),
            ("person", "担当", 65, "center"),
        ]
        for c_id, title, w, anc in col_defs:
            tree.heading(c_id, text=title)
            tree.column(c_id, width=w, anchor=anc)

        tree.tag_configure("TARGET_ROW", background="#fef08a", foreground="#713f12")
        tree.tag_configure("SAME_NO", background="#f0fdf4", foreground="#166534")

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.grid_rowconfigure(0, weight=1)
        tbl_frame.grid_columnconfigure(0, weight=1)

        matched_rows = []
        for m in self.cached_mgmt_rows:
            same_no = order_no_val and (
                m.get("order_no_8") == order_no_val or m.get("order_no") == order_no_val
            )
            near_row = (
                mgmt_sheet
                and target_row
                and m.get("sheet") == mgmt_sheet
                and abs(int(m.get("row", 0)) - int(target_row)) <= 10
            )
            if same_no or near_row:
                matched_rows.append(m)

        target_iid = None
        for idx, m in enumerate(matched_rows):
            is_target = (
                mgmt_sheet
                and target_row
                and m.get("sheet") == mgmt_sheet
                and int(m.get("row", -1)) == int(target_row)
            )
            mark_str = "★ 該当行" if is_target else ("同一No" if m.get("order_no_8") == order_no_val else "周辺行")
            tag = "TARGET_ROW" if is_target else ("SAME_NO" if m.get("order_no_8") == order_no_val else "")
            iid = f"mprev_{idx}"
            tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    mark_str,
                    m.get("sheet", ""),
                    m.get("row", ""),
                    m.get("order_no_8") or m.get("order_no", ""),
                    m.get("instruction_disp") or m.get("instruction_raw") or "(空白)",
                    m.get("ship_to", ""),
                    m.get("line_name", ""),
                    m.get("product_detail", ""),
                    m.get("actual_raw") or (m["actual_date"].strftime("%Y-%m-%d") if m.get("actual_date") else "未入力"),
                    m.get("drawing_no", ""),
                    m.get("part_no", ""),
                    m.get("item_name", ""),
                    m.get("maker", ""),
                    m.get("spec", ""),
                    m.get("qty", ""),
                    m.get("mgmt_due_raw") or (m["mgmt_due_date"].strftime("%Y-%m-%d") if m.get("mgmt_due_date") else ""),
                    m.get("invoice_raw") or (m["invoice_date"].strftime("%Y-%m-%d") if m.get("invoice_date") else "未入力"),
                    m.get("person", ""),
                ),
                tags=(tag,) if tag else (),
            )
            if is_target:
                target_iid = iid

        if target_iid:
            tree.selection_set(target_iid)
            tree.see(target_iid)

        def _on_mgmt_prev_dblclick(_e):
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], "values")
            if len(vals) >= 3:
                sh_name = str(vals[1])
                r_num = int(vals[2]) if str(vals[2]).isdigit() else None
                mgmt_path = self.mgmt_file_path.get().strip().strip('"')
                self._open_excel_at_location(mgmt_path, sheet_name=sh_name, row_num=r_num, label_for_status="注文品管理表")

        tree.bind("<Double-1>", _on_mgmt_prev_dblclick)

    # =========================================================================
    # ヘルプ・ガイドダイアログ
    # =========================================================================
    def _create_scrollable_text(self, parent) -> tk.Text:
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(
            container,
            wrap="char",
            font=("Meiryo UI", 9),
            padx=16,
            pady=14,
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
            font=("Meiryo UI", 11, "bold"),
            foreground="#0f172a",
            background="#e2e8f0",
            spacing1=10,
            spacing3=8,
            lmargin1=10,
            lmargin2=10,
            rmargin=10,
        )
        txt.tag_configure(
            "h2",
            font=("Meiryo UI", 10, "bold"),
            foreground="#1d4ed8",
            spacing1=12,
            spacing3=4,
            lmargin1=8,
            lmargin2=8,
            rmargin=10,
        )
        txt.tag_configure(
            "body",
            font=("Meiryo UI", 9),
            foreground="#1e293b",
            spacing1=2,
            spacing2=3,
            spacing3=3,
            lmargin1=18,
            lmargin2=32,
            rmargin=10,
        )
        txt.tag_configure(
            "point",
            font=("Meiryo UI", 9, "bold"),
            foreground="#0f172a",
            background="#fef9c3",
            spacing1=4,
            spacing2=3,
            spacing3=4,
            lmargin1=18,
            lmargin2=32,
            rmargin=10,
        )
        txt.tag_configure(
            "warn",
            font=("Meiryo UI", 9, "bold"),
            foreground="#991b1b",
            background="#fee2e2",
            spacing1=4,
            spacing2=3,
            spacing3=4,
            lmargin1=18,
            lmargin2=32,
            rmargin=10,
        )
        txt.tag_configure(
            "badge_ok",
            font=("Meiryo UI", 9, "bold"),
            foreground="#166534",
            background="#dcfce7",
            spacing1=3,
            spacing2=3,
            spacing3=3,
            lmargin1=18,
            lmargin2=32,
            rmargin=10,
        )
        txt.tag_configure(
            "sub_note",
            font=("Meiryo UI", 9),
            foreground="#475569",
            spacing1=2,
            spacing2=2,
            spacing3=4,
            lmargin1=30,
            lmargin2=42,
            rmargin=10,
        )

    def _show_help_guide(self, tab_index: int = 0):
        win = tk.Toplevel(self.root)
        win.title("StockMatch_Miya-v1.4.0 - ヘルプ＆操作ガイド")
        self._apply_window_icon(win)
        win.geometry("960x740")
        win.minsize(780, 560)
        win.transient(self.root)

        banner = tk.Frame(win, bg="#1e293b", padx=18, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text="📖 StockMatch_Miya-v1.4.0  取扱説明書・操作ガイド",
            bg="#1e293b",
            fg="#ffffff",
            font=("Meiryo UI", 12, "bold"),
            anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            banner,
            text="注文書と注文品管理表の双方向自動照合・タブ分離一覧表示・納期超過チェックツールの詳しい使い方",
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

        txt1.insert(tk.END, "  STEP 1 : 注文品管理表と注文書フォルダの指定  \n", "h1")
        txt1.insert(tk.END, "■ ① 注文品管理表ファイルの指定\n", "h2")
        txt1.insert(tk.END, "・画面上部の『① 注文品管理表ファイル』行にある『ファイルを参照...』ボタン（または Ctrl+O）を押し、管理表Excel（例: 注文品管理表.xlsx）を選択します。\n", "body")
        txt1.insert(tk.END, "・右端の『📋 管理表を開く』ボタンを押すと、現在指定されている管理表ファイルをいつでも直接Excelで開いて確認できます。\n", "body")
        txt1.insert(tk.END, "★【自動抽出＆列構造の自動認識】管理表内のすべてのシートから「管理No（8桁の数字）」を自動抽出し、A列（現品確認日）、AB列（28列目：購入数）、AI列（請求日）などの列位置を正確に読み取ります。\n", "point")
        txt1.insert(tk.END, "※管理表がExcelで開かれたままの状態でも、アプリが一時コピーを作成して安全に読み取るため、Excelを閉じずにそのまま照合を実行できます。\n", "sub_note")

        txt1.insert(tk.END, "■ ② 注文書フォルダパスの指定（サブフォルダ一括探索・高速スキャン）\n", "h2")
        txt1.insert(tk.END, "・『② 注文書フォルダパス』行にある『フォルダを参照...』ボタン（または Ctrl+D）を押し、注文書Excelが保管されている親フォルダ（例: 依頼ファイル）を選択します。\n", "body")
        txt1.insert(tk.END, "・右端の『📂 フォルダを開く』ボタンを押すと、指定した注文書フォルダをWindowsエクスプローラーで直接開きます。\n", "body")
        txt1.insert(tk.END, "★【管理Noフォルダ自動スキップ（高速枝刈りスキャン）】親フォルダ配下の多数のサブフォルダを探索する際、サブフォルダ名に8桁の管理Noが含まれており、その管理Noが「注文品管理表に存在しない」または「除外設定されている」場合は、そのサブフォルダの中身を一切開かずに丸ごとスキップします。何千個ものフォルダがある環境でも数秒〜数十秒で高速に探索が完了します。\n", "point")
        txt1.insert(tk.END, "★【マルチスレッド並列読込＆メモリキャッシュ】対象となった注文書Excelは並列処理で一斉に読み込まれます。さらに一度読み込んだファイルは更新日時・ファイルサイズとともにメモリにキャッシュされるため、2回目以降の再照合は瞬時に完了します。\n", "point")
        txt1.insert(tk.END, "★【見積依頼書の自動除外】ファイル名およびExcelシートの2行目タイトルを自動検査し、「見積」が含まれるシートは自動的に除外して「注文」シートのみを照合対象にします。\n", "point")
        txt1.insert(tk.END, "※Googleドライブ（G:ドライブ等）や社内ネットワーク上の共有フォルダ（\\\\サーバー名\\共有名 等）にあるファイルでも問題なく高速に動作します。\n", "sub_note")

        txt1.insert(tk.END, "\n  STEP 2 : 判定基準日・納期未記載時の日数・経過月数除外・設定の保存  \n", "h1")
        txt1.insert(tk.END, "■ ③ 判定基準日（今日）・納期未記載時の経過日数・注文日からの経過除外設定\n", "h2")
        txt1.insert(tk.END, "・『③ 判定基準日 (今日)』には、納期超過を判定する基準日（初期値はアプリ起動当日の日付）を YYYY-MM-DD 形式で入力します。過去や未来の日付を入力してシミュレーションすることも可能で、『今日に戻す』ボタンでいつでも当日の日付に戻せます。\n", "body")
        txt1.insert(tk.END, "・『納期未記載時: 注文日(4行目)から [ 7 ] 日経過で超過』には、注文書の明細行に納期が書かれていない場合の猶予日数（初期値: 7日）を指定します。\n", "body")
        txt1.insert(tk.END, "★【納期空欄明細の自動算出】注文書明細の納期欄が空欄（または「別途」等の日付変換できない文字のみ）の場合、注文書シート4行目の「注文日」からこの指定日数を加算した日を「みなし納期」として超過判定を行います。\n", "point")
        txt1.insert(tk.END, "・『経過除外: 注文日から [ 12 ] ヶ月経過で除外』には、注文日から何ヶ月経過した古いデータを一覧や集計から除外するか（初期値: 12ヶ月）を指定します。\n", "body")
        txt1.insert(tk.END, "★【古い注文データの自動除外】判定基準日時点で、注文日（注文書4行目等）から指定ヶ月数（デフォルト12ヶ月）以上が経過している明細は、結果一覧・件数サマリー・Excel出力から自動的に除外されます（※0を入力すると経過月数による除外を行わず全期間を表示します）。\n", "point")

        txt1.insert(tk.END, "■ 設定の保存（💾 設定を保存）と保存先フォルダの変更（⚙ 保存先設定...）\n", "h2")
        txt1.insert(tk.END, "・『③ 判定基準日』行の右端にある『💾 設定を保存』ボタンを押すと、現在指定している「①管理表ファイルパス」「②注文書フォルダパス」「③納期未記載時の超過日数」「③注文日からの経過除外月数」「④除外する管理Noリスト」「④管理表 指示書列(J列)フィルタ設定」が設定ファイル（config.json）に保存され、次回アプリ起動時に自動で読み込まれます。\n", "point")
        txt1.insert(tk.END, "・『💾 設定を保存』の左隣にある『⚙ 保存先設定...』ボタン（または上部メニュー『設定(S)』→『設定ファイル(config.json)の保存先設定...』）を押すと、config.json を保存するフォルダを自由に変更できます。\n", "body")
        txt1.insert(tk.END, "★【保存先の変更とワンクリック初期値復帰】共有フォルダやGoogleドライブ等を保存先に指定すれば複数環境で設定を共有できます。また、保存先設定ウィンドウ内の『🏠 アプリと同じ場所（初期値）に戻す』ボタンを押すだけで、いつでもボタン一つでアプリ本体と同じフォルダ（初期状態）に戻すことができます。\n", "point")
        txt1.insert(tk.END, "※保存先フォルダの場所設定自体は、Windowsのユーザー設定領域（%APPDATA%\\StockMatch_Miya\\config_location.txt）に記憶されるため、次回起動時も指定した保存先の config.json が自動的に参照されます。\n", "sub_note")

        txt1.insert(tk.END, "\n  STEP 3 : 一覧から除外する管理Noの設定（専用ウィンドウ・Excel貼り付け対応）  \n", "h1")
        txt1.insert(tk.END, "■ ④ 一覧から除外する管理Noの登録・編集\n", "h2")
        txt1.insert(tk.END, "・テスト案件やキャンセル案件、すでに完了扱いで一覧に出したくない管理Noがある場合は、『④ 一覧から除外する管理No』行の『🚫 除外する管理Noを設定・編集...』ボタンをクリックします。\n", "body")
        txt1.insert(tk.END, "・専用の設定ウィンドウが開き、除外したい管理Noを「1行に1つずつ（改行区切り）」で分かりやすく入力・確認・編集できます。\n", "point")
        txt1.insert(tk.END, "★【Excelから複数セルをコピーして一括貼り付け】Excel上で複数の管理Noセル（縦一列や複数行・複数列の範囲）を選択してコピー（Ctrl+C）し、設定ウィンドウの『📋 クリップボードから貼り付け』ボタンを押す（またはテキスト欄で Ctrl+V）だけで、自動的に1行ずつの管理Noに分解されて一括登録されます。\n", "point")
        txt1.insert(tk.END, "・『✨ 重複・空白を整理』ボタンを押すと、空行や重複している管理No、前後の余分な空白・カンマ等を自動で綺麗に整理します。\n", "body")
        txt1.insert(tk.END, "・『✔ 適用して設定を保存』ボタンを押すと、設定ファイル（config.json）へ自動保存されると同時に、すでに表示中の照合結果一覧や件数サマリーからも即座に除外が反映されます（再スキャン不要）。\n", "point")
        txt1.insert(tk.END, "・メイン画面の④行には現在の登録件数と除外中の管理Noプレビューがバッジ表示され、右端の『全解除』ボタンからワンクリックで除外設定をすべて解除することもできます。\n", "sub_note")

        txt1.insert(tk.END, "■ ④ 注文品管理表の『指示書列(J列)フィルタ』機能\n", "h2")
        txt1.insert(tk.END, "・『④ 一覧から除外する管理No』行の右側にある『📋 管理表 指示書列(J列)フィルタ』ドロップダウンから、注文品管理表のJ列（指示書列）の値に応じた絞り込み条件を選択できます。\n", "body")
        txt1.insert(tk.END, "★【指示書列(J列)によるワンクリック絞り込み（初期値: 「有り」のみ）】\n  ・『「有り」のみ』（初期値）: J列に「有り（有・あり・○等）」と記載されている案件のみを表示します。\n  ・『有り または 空白 のみ』: J列が「有り」または「空白（未記入）」の案件のみを表示し、「無し・不要・-」等の案件を自動除外します。\n  ・『「空白」のみ』: J列が空白（空欄）の案件のみを表示します。\n  ・『すべて表示 (フィルタなし)』: J列の値に関わらずすべての案件を表示します。\n", "point")
        txt1.insert(tk.END, "・ドロップダウンを切り替えると再スキャン不要で即座に一覧・件数サマリーに反映され、『💾 設定を保存』で次回起動時にも引き継がれます。\n", "sub_note")

        txt1.insert(tk.END, "\n  STEP 4 : 照合・遅延チェックの実行と処理の中止  \n", "h1")
        txt1.insert(tk.END, "■ ⑤ 進捗状況の確認と『▶ 照合・遅延チェック実行 (F5)』\n", "h2")
        txt1.insert(tk.END, "・『⑤ 進捗状況』行の右端にある青い『▶ 照合・遅延チェック実行 (F5)』ボタンを押す（またはキーボードの F5 キーを押す）と、管理表と注文書フォルダの照合・納期超過チェックが開始されます。\n", "point")
        txt1.insert(tk.END, "・実行中はプログレスバーと進捗メッセージに「現在の処理ステップ（[1/3]管理表の読込 → [2/3]注文書フォルダの探索・読込 → [3/3]明細の突合・超過判定）」と処理件数・スキップされたフォルダ数がリアルタイム表示されます。\n", "body")
        txt1.insert(tk.END, "★【実行中の安全な中止（Escキー対応）】照合処理の実行中は、実行ボタンの左隣にある赤い『⏹ 中止 (Esc)』ボタンを押すか、キーボードの Esc キーを押すことで、いつでも安全かつ即座に処理を中断できます。\n", "point")

        txt1.insert(tk.END, "\n  STEP 5 : 8つの件数サマリーカードと2つの独立タブによる結果確認  \n", "h1")
        txt1.insert(tk.END, "■ 照合件数サマリー ＆ ワンクリック表示フィルタ切替（8枚のカード）\n", "h2")
        txt1.insert(tk.END, "・画面中段には照合結果を集計した8枚のサマリーカードが並びます。カード本体またはカード下部のボタンをクリックするだけで、見たい条件の明細だけに一瞬で絞り込まれます（現在選択中のカードは太枠と『✔ フィルター表示中』で強調されます）。\n", "point")
        txt1.insert(tk.END, "  1. 【注文書明細 総数】 : 注文書から検出された全明細を表示します（『▼ 全件を表示』）。\n", "body")
        txt1.insert(tk.END, "  2. 【すべての異常(超過・不一致)】 : 納期超過・伝票未着・不一致・管理表未記載など、確認が必要な異常案件をすべてまとめて表示します（『▼ 異常をすべて表示』）。\n", "body")
        txt1.insert(tk.END, "  3. 【警告：現品・請求の両方超過】 : 納期を過ぎても「現品確認日(A列)」と「請求日(AI列)」の両方が未記入の最重要確認案件を表示します（※照合完了時の初期表示フィルタ）。\n", "warn")
        txt1.insert(tk.END, "  4. 【現品超過：現品確認日のみ超過】 : 請求書（伝票）は処理済みですが、現品確認日(A列)が未記入のまま納期を過ぎている案件を表示します。\n", "body")
        txt1.insert(tk.END, "  5. 【請求超過・未着：請求日(伝票)超過】 : 現品は届いていますが請求日(AI列)が未記入のまま納期を過ぎている案件、または現品確認日から長期間伝票が届いていない案件を表示します。\n", "body")
        txt1.insert(tk.END, "  6. 【不一致：管理表との不一致】 : 注文書と管理表で品名・品番・図番・数量などが食い違っている案件、および注文書はあるが管理表に記載がない案件を表示します。\n", "body")
        txt1.insert(tk.END, "  7. 【正常：納期前・確認済】 : 現品・請求ともに確認済みの案件、またはまだ納期前の正常な案件を表示します。\n", "badge_ok")
        txt1.insert(tk.END, "  8. 【別タブ：注文書未検出】 : 管理表に記載があるのに注文書Excelが見つからない案件数を表示します。クリックすると自動的に『タブ②（未発注・注文書なし一覧）』へ切り替わります。\n", "point")

        txt1.insert(tk.END, "■ 2つの独立タブ（①納期超過・照合結果一覧 ／ ②未発注・注文書なし一覧）\n", "h2")
        txt1.insert(tk.END, "・一覧表示エリアは用途に応じて2つの独立したタブに分かれています。現在開いているタブには紺色で『【選択中】▶』マークが表示されます。\n", "body")
        txt1.insert(tk.END, "・『① 納期超過・照合結果 一覧』タブ : 注文書が存在する案件の納期超過チェック結果や、管理表との項目照合結果を確認するメインの一覧です。\n", "point")
        txt1.insert(tk.END, "・『② 【未発注・注文書なし】管理表のみ記載 一覧』タブ : 管理表に載っているが注文書ファイルがまだ存在しない案件（未発注または注文書未格納）だけを分離して確認できる専用一覧です（購入数が0の在庫充当品等は自動で除外されます）。\n", "point")

        txt1.insert(tk.END, "\n  STEP 6 : 検索・詳細プレビュー・元ファイル参照・Excel直接出力  \n", "h1")
        txt1.insert(tk.END, "■ ツールバーの操作ボタンとキーワード検索\n", "h2")
        txt1.insert(tk.END, "・『全件表示に切替』ボタン : サマリーカードによる絞り込みを解除し、現在のタブの全データを表示します。\n", "body")
        txt1.insert(tk.END, "・『キーワード検索』欄 : 管理No、発注先、メーカー、品名、型番・図番、不一致理由などの一部を入力すると、入力と同時にリアルタイムで該当行だけを絞り込みます（『クリア』ボタンで解除）。\n", "body")
        txt1.insert(tk.END, "・『🔍 照合詳細・プレビュー』ボタン（または行をダブルクリック） : 選択した行の「注文書 vs 管理表の項目別比較表」と「注文書・管理表のシートプレビュー」を専用ウィンドウで詳しく確認できます。\n", "point")
        txt1.insert(tk.END, "・『📄 注文書を参照 (読取専用)』ボタン : 選択した明細の元になっている注文書Excelを開き、該当シート・該当行を自動選択して表示します（誤って上書き保存されないよう厳重な読み取り専用ロックがかかります）。\n", "point")
        txt1.insert(tk.END, "・『📋 管理表を参照』ボタン : 注文品管理表をExcelで開き、該当する管理Noのシート・行をピンポイントで自動選択して表示します。\n", "point")
        txt1.insert(tk.END, "・『列フィルター解除』ボタン : 列ヘッダー（▼）で設定したExcel風の列絞り込みをすべて解除します。\n", "body")
        txt1.insert(tk.END, "・『📊 Excelで直接開く』ボタン（または Ctrl+S） : 現在の照合結果を、パソコン内に不要な一時ファイルを残さずにExcelの新規ブックとして直接立ち上げ、色分け・列幅調整・A4横印刷設定済みの見やすい表として出力します（出力中も左隣の『⏹ 中止』ボタンや Esc キーで中断可能です）。\n", "point")
        txt1.configure(state="disabled")

        tab2 = ttk.Frame(notebook, padding=8)
        notebook.add(tab2, text="  ② 照合ルール・タブ分離の仕組み  ")
        txt2 = self._create_scrollable_text(tab2)

        txt2.insert(tk.END, "  1. なぜ2つのタブに分離されているのか？（作業効率と視認性の向上）  \n", "h1")
        txt2.insert(tk.END, "■ タブ①『納期超過・照合結果 一覧』と タブ②『【未発注・注文書なし】一覧』の役割分担\n", "h2")
        txt2.insert(tk.END, "・注文書がすでに発行されている案件の「納品・請求確認（納期管理）」と、これから注文書を発行・格納する「未発注案件の確認」は業務上の確認目的が異なります。\n", "body")
        txt2.insert(tk.END, "・そこで、注文書が存在する案件はすべて『タブ①』に集約し、管理表にのみ記載があって注文書Excelが見つからない案件は『タブ②』へ完全に分離しています。\n", "point")
        txt2.insert(tk.END, "★【管理表AB列（購入数0）の自動除外ルール】管理表のAB列（28列目）「購入数」が『0』（または0個・0台等）になっている行は、在庫充当品など発注不要な明細であるため、注文書がなくても異常とせず『タブ②（未発注一覧）』から自動的に除外します。\n", "point")
        txt2.insert(tk.END, "★【管理表未記載の検出】逆に、注文書フォルダ内に注文書Excelが存在するのに管理表側に該当する明細行がない場合は、『タブ①』にて『【管理表未記載】注文書のみ存在』として警告表示し、管理表への記入漏れを防ぎます。\n", "warn")

        txt2.insert(tk.END, "\n  2. 注文書と管理表の高精度マッチング（突合）アルゴリズム  \n", "h1")
        txt2.insert(tk.END, "■ 単価・金額は照合対象外（誤判定の完全防止）\n", "h2")
        txt2.insert(tk.END, "★【重要】注文書発行時点で単価が未定（0円や空欄）のケースや、管理表側との端数・消費税表記の違いによって「不一致」と誤判定されるのを防ぐため、単価および金額は照合計算から完全に除外しています。\n", "point")

        txt2.insert(tk.END, "■ 7つの照合項目と全体最適1対1マッチング\n", "h2")
        txt2.insert(tk.END, "・同じ管理No（8桁）を持つ注文書明細と管理表明細の間で、以下の項目を組み合わせて類似度スコアを算出し、明細の並び順が違っていても最も適切なペアを1対1で自動紐付けします。\n", "body")
        txt2.insert(tk.END, "  ① 管理No（8桁の注文番号・管理番号）\n  ② 図面番号（図番）\n  ③ 品番・型番\n  ④ 品名\n  ⑤ 仕様・寸法・規格\n  ⑥ 発注先（業者名）・メーカー名\n  ⑦ 数量（注文書の数量 と 管理表AB列の購入数）\n", "point")
        txt2.insert(tk.END, "・全角／半角、大文字／小文字、ハイフン記号の長短、（株）や株式会社などの法人格表記ゆれは自動的に正規化して比較します。\n", "sub_note")
        txt2.insert(tk.END, "・紐付けられたペアの間で数量違いや型番・品名の相違がある場合は、『【不一致】管理表との不一致』に分類し、「不一致理由・照合診断」列にどの項目がどう異なっているか（例: 数量不一致: 注文書=5 / 管理表=3）を具体的にわかりやすく表示します。\n", "warn")

        txt2.insert(tk.END, "\n  3. 現品確認日（A列）と請求日（AI列）の超過判定ルール  \n", "h1")
        txt2.insert(tk.END, "■ 管理表の記入状況と判定基準日の比較\n", "h2")
        txt2.insert(tk.END, "・管理表の『A列（現品確認日／物）』および『AI列（請求日／伝票）』の記載内容を確認します。\n", "body")
        txt2.insert(tk.END, "・日付または「済」「確認済」「○」などの記入がある場合は『確認済 (正常)』と判定します。\n", "badge_ok")
        txt2.insert(tk.END, "・空欄（未確認）のまま、注文書の納期が画面上部③で指定した『判定基準日』を過ぎている場合（基準日 ＞ 納期）、『未確認 (超過X日)』として警告表示します。\n", "warn")
        txt2.insert(tk.END, "  ・現品(A列)も請求(AI列)も両方空欄で納期超過 → 『【警告】現品・請求の両方超過』（赤色で最優先警告）\n", "warn")
        txt2.insert(tk.END, "  ・現品(A列)のみ空欄で納期超過（請求は済） → 『【現品超過】現品確認日のみ超過』（オレンジ色で警告）\n", "point")
        txt2.insert(tk.END, "  ・請求(AI列)のみ空欄で納期超過（現品は済） → 『【請求超過】請求日(伝票)のみ超過』（黄色で警告）\n", "point")
        txt2.insert(tk.END, "  ・納期が空欄等で判定できなくても、現品確認日(A列)から一定日数経過して請求日(AI列)が空欄の場合 → 『【伝票未着】現品確認からX日経過』としてお知らせします。\n", "point")

        txt2.insert(tk.END, "\n  4. 注文書納期の曖昧表記（上旬・中旬・下旬等）と納期空欄時の自動変換ルール  \n", "h1")
        txt2.insert(tk.END, "■ 期間表記・旬表記の期日自動変換\n", "h2")
        txt2.insert(tk.END, "・注文書の納期欄に具体的な日付ではなく「上旬」「中旬」「下旬」「月末」などが書かれている場合、それぞれの期間の最終日を期日として自動変換し、その日を過ぎたら超過と判定します。\n", "body")
        txt2.insert(tk.END, "  ・『11/上』『11/上旬』『11月上旬』『11/初』『11月頭』など → その月の 10日（例: 11月10日）\n", "point")
        txt2.insert(tk.END, "  ・『11月中旬』『11/中』『11/中頃』『11月半ば』など → その月の 20日（例: 11月20日）\n", "point")
        txt2.insert(tk.END, "  ・『11/下』『11/下旬』『11月下旬』『11/末』『11月末』『11月内』など → その月の 末日（例: 11月30日、2月は28日/29日）\n", "point")
        txt2.insert(tk.END, "  ・『11月前半』 → その月の 15日 ／ 『11月後半』 → その月の 末日\n", "point")
        txt2.insert(tk.END, "  ・『11/10〜11/15』のような期間幅表記 → 後ろの期日（11月15日）を基準に判定します。\n", "body")
        txt2.insert(tk.END, "  ・令和（R6.11.10 等）の和暦表記や、年が省略された月日表記（注文日の年から自動補完、年末年始の年跨ぎも自動考慮）にも完全対応しています。\n", "sub_note")

        txt2.insert(tk.END, "■ 納期未記載（空欄）時の『注文日(4行目) ＋ 指定日数』自動算出\n", "h2")
        txt2.insert(tk.END, "・注文書明細の納期欄が空欄の場合、注文書シートの4行目に記載されている「注文日（発行日）」を自動で読み取り、画面上部③で設定した日数（初期値: 7日）を足した日付を期日として自動算出します。\n", "point")
        txt2.insert(tk.END, "・その場合、一覧の「注文書納期」列には『2026-11-09 (注文日+7日)』のように算出根拠が明記されるため、一目で納期未記載案件であることが分かります。\n", "body")
        txt2.configure(state="disabled")

        tab3 = ttk.Frame(notebook, padding=8)
        notebook.add(tab3, text="  ③ 便利機能・ショートカットキー  ")
        txt3 = self._create_scrollable_text(tab3)

        txt3.insert(tk.END, "  1. 照合詳細・アプリ内シートプレビュー機能（ダブルクリック）  \n", "h1")
        txt3.insert(tk.END, "■ 専用詳細ウィンドウの3つのタブで原因を即座に特定\n", "h2")
        txt3.insert(tk.END, "・結果一覧の行を【ダブルクリック】するか、行を選択して『🔍 照合詳細・プレビュー』ボタンを押すと、詳細確認ウィンドウが開きます。\n", "point")
        txt3.insert(tk.END, "  ・タブ①『項目比較・診断レポート』 : 注文書と管理表の各項目（管理No・発注先・メーカー・品名・型番・数量・納期・確認日等）を左右に並べて比較し、一致／不一致の判定結果とファイルパス・行番号を一覧表示します。\n", "body")
        txt3.insert(tk.END, "  ・タブ②『注文書シート参照 (プレビュー)』 : Excelを立ち上げなくても、アプリ内で実際の注文書シートの内容を表形式でそのまま閲覧でき、該当する明細行が黄色くハイライト表示されます。\n", "point")
        txt3.insert(tk.END, "  ・タブ③『注文品管理表参照 (プレビュー)』 : アプリ内で注文品管理表の該当シートを表形式で閲覧でき、該当行へ自動でスクロール＆ハイライト表示されます。\n", "point")

        txt3.insert(tk.END, "\n  2. 元ファイルのピンポイントExcel参照と『注文書の誤保存完全ブロック』  \n", "h1")
        txt3.insert(tk.END, "■ 該当シート・該当行を自動選択してExcelで開く\n", "h2")
        txt3.insert(tk.END, "・一覧で明細を選択して『📄 注文書を参照 (読取専用)』または『📋 管理表を参照』ボタンを押す（または行を右クリックしてメニューから選ぶ）と、Excelが起動して該当ファイルの該当シート・該当行がハイライト選択された状態で瞬時に開きます。\n", "point")
        txt3.insert(tk.END, "★【注文書の誤保存・誤編集を防ぐ5重の安全保護】注文書を参照する際は、大事な原本ファイルを誤って書き換えてしまう事故を防ぐため、以下の5重の保護が自動的に適用されます。\n", "warn")
        txt3.insert(tk.END, "  ① 読み取り専用モード（ReadOnly）での強制起動\n  ② すでにExcelで通常モードで開かれていた場合も、自動的に読み取り専用へ切り替えて開き直す保護\n  ③ 開いたブックの全シートおよびブック構造へのメモリ上ロック適用（セルの書き換えを禁止しつつ、行選択やCtrl+Cでの文字コピーは自由に可能）\n  ④ Excelを閉じる際に「変更を保存しますか？」という紛らわしい確認ダイアログを出さずに安全に閉じられる保護\n  ⑤ 万が一制御機能が利用できない環境でも、通常モードでの直接起動を遮断して原本を守る安全設計\n", "body")

        txt3.insert(tk.END, "\n  3. Excel風 列フィルター＆ソート機能・Excel新規ブック直接出力  \n", "h1")
        txt3.insert(tk.END, "■ 列ヘッダー（▼）クリックによるExcel風フィルターと並べ替え\n", "h2")
        txt3.insert(tk.END, "・一覧表の各列タイトル（『照合判定 ▼』『発注先 ▼』『注文書納期 ▼』など）をクリックすると、Excelと同じようなフィルターメニューが開きます。\n", "point")
        txt3.insert(tk.END, "・昇順／降順での並べ替え（数値・日付・文字を自動判別）ができるほか、検索ボックスやチェックボックスを使って特定の業者や特定の管理No、特定の日付だけを自由に絞り込めます。\n", "body")
        txt3.insert(tk.END, "・フィルターが適用されている列はヘッダーに『[絞込]』マークが付き、ツールバーの『列フィルター解除』ボタンでいつでも一括解除できます。\n", "sub_note")

        txt3.insert(tk.END, "■ 『📊 Excelで直接開く』による一時ファイル不要のレポート出力\n", "h2")
        txt3.insert(tk.END, "・ツールバー右端の緑色ボタン『📊 Excelで直接開く』を押すと、PC内に余計な一時ファイルを一切作らずにExcelの新規ブック（Book1等）を直接立ち上げ、結果を綺麗に整形して書き出します。\n", "point")
        txt3.insert(tk.END, "・『照合結果_現在表示』『照合結果_全件』『注文書未検出_管理表のみ』の3シート構成で出力され、異常・警告に応じた行の色分け、見出し行の固定、オートフィルター、A4横1ページ幅の印刷設定まで全自動で完了します（必要な場合のみExcel側で名前を付けて保存してください）。\n", "body")

        txt3.insert(tk.END, "\n  4. 上部メニューバー・マウス操作・ショートカットキー一覧  \n", "h1")
        txt3.insert(tk.END, "■ マウス操作の便利な使い方\n", "h2")
        txt3.insert(tk.END, "・一覧の行を【ダブルクリック】 : 『🔍 照合詳細・プレビュー』ウィンドウを開く\n", "body")
        txt3.insert(tk.END, "・一覧の行を【右クリック】     : コンテキストメニューを開く（詳細表示／注文書参照／管理表参照／注文書フォルダを開く／管理Noをコピー／行データをコピー）\n", "body")
        txt3.insert(tk.END, "・列ヘッダーを【左クリック】   : 列ごとのExcel風フィルター・ソートメニューを開く\n", "body")

        txt3.insert(tk.END, "■ キーボードショートカット一覧\n", "h2")
        txt3.insert(tk.END, "・ F1 キー        :  この『ヘルプ＆操作ガイド』を開く\n", "body")
        txt3.insert(tk.END, "・ F5 キー        :  『▶ 照合・遅延チェック実行』を開始する\n", "body")
        txt3.insert(tk.END, "・ Esc キー       :  実行中の処理（照合・遅延チェック または Excel直接展開）を安全に中止する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + O       :  ① 注文品管理表ファイルを選択する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + D       :  ② 注文書フォルダを選択する\n", "body")
        txt3.insert(tk.END, "・ Ctrl + S       :  照合結果を『📊 Excelで直接開く』（新規ブック展開）\n", "body")
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
            "StockMatch_Miya-v1.4.0\n"
            "注文書 ⇔ 管理表 双方向照合・現品/請求超過システム\n\n"
            "・Excel直接起動: 一時ファイルなしで未保存新規ブック展開（A4横・1ページ収まり印刷設定済）\n"
            "・注文書保護: 注文書参照時は読み取り専用＋シート保護ロック＋保存確認抑止で誤保存を完全ブロック\n"
            "・タブ完全分離: 「①納期超過・照合一覧」と「②注文書未検出(管理表のみ)」を別一覧で確認可能\n"
            "・未発注一覧フィルタ: 注文品管理表のAB列「購入数」が0の明細は自動除外\n"
            "・指示書列(J列)フィルタ: 管理表J列が「有り または 空白 のみ」「有りのみ」「空白のみ」等で絞り込み可能\n"
            "・除外管理No設定: 一覧から除外したい管理No（複数可）を指定可能＆設定保存対応\n"
            "・経過月数除外: 注文日から指定ヶ月数(初期値12ヶ月)経過した古い案件を自動除外＆設定保存対応\n"
            "・納期未記載対応: 注文書納期が空欄時、注文日(4行目)から指定日数(初期値7日)経過で超過判定\n"
            "・照合方式: 全体最適1対1マッチング（※単価・金額は照合対象外）\n"
            "・自動分別: 2行目の記載により「見積依頼書」を除外し「注文書」のみ照合\n"
            "・高速探索: 管理表にない管理Noフォルダの丸ごとスキップ＋並列読込＋キャッシュ対応\n"
            "・保存先変更: 設定ファイル(config.json)の保存先を自由に指定可能＆ワンクリックで初期値(アプリと同じ場所)へ復帰対応\n"
            f"・設定保存先: {self.get_config_file_path()}\n\n"
            "© StockMatch_Miya System All Rights Reserved.",
        )

    @classmethod
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
        """config.json の保存先フォルダを指定・変更するダイアログを開く"""
        win = tk.Toplevel(self.root)
        win.title("設定ファイル (config.json) の保存先設定")
        self._apply_window_icon(win)
        win.geometry("600x460")
        win.minsize(500, 420)
        win.transient(self.root)
        win.grab_set()

        try:
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            x = rx + max(0, (rw - 600) // 2)
            y = ry + max(0, (rh - 460) // 2)
            win.geometry(f"600x460+{x}+{y}")
        except Exception:
            pass

        # 1. 上部ヘッダーバナー
        banner = tk.Frame(win, bg="#1e293b", padx=16, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text="⚙ 設定ファイル (config.json) の保存先設定",
            bg="#1e293b",
            fg="#ffffff",
            font=("Meiryo UI", 11, "bold"),
            anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            banner,
            text="設定（管理表パス、注文書フォルダパス、超過設定、除外管理No）を保存する場所を指定します。\nデフォルトはアプリ本体と同じフォルダです。社内共有フォルダ等に変更して共有することも可能です。",
            bg="#1e293b",
            fg="#cbd5e1",
            font=("Meiryo UI", 9),
            anchor="w",
            justify="left",
        ).pack(fill=tk.X, pady=(4, 0))

        content_frame = tk.Frame(win, bg="#f8fafc", padx=16, pady=14)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 現在の保存先表示
        tk.Label(content_frame, text="■ 現在の設定ファイル保存先:", bg="#f8fafc", font=("Meiryo UI", 9, "bold")).pack(anchor="w")

        curr_cfg = self.get_config_file_path()
        is_default = (os.path.normcase(curr_cfg) == os.path.normcase(os.path.abspath(self.DEFAULT_CONFIG_FILE_PATH)))

        curr_disp_frame = tk.Frame(content_frame, bg="#ffffff", relief="solid", borderwidth=1, padx=8, pady=6)
        curr_disp_frame.pack(fill=tk.X, pady=(4, 6))

        tk.Label(
            curr_disp_frame,
            text=curr_cfg,
            bg="#ffffff",
            fg="#0f172a",
            font=("Consolas", 9),
            anchor="w",
            justify="left",
        ).pack(fill=tk.X)

        status_tag = "✔ 現在: 初期値（アプリと同じ場所）に設定中" if is_default else "★ 現在: カスタム保存先に設定中"
        tag_fg = "#166534" if is_default else "#b45309"
        tk.Label(content_frame, text=status_tag, bg="#f8fafc", fg=tag_fg, font=("Meiryo UI", 9, "bold")).pack(anchor="w", pady=(0, 12))

        # 新しい保存先フォルダ選択
        tk.Label(content_frame, text="■ 保存先フォルダの変更:", bg="#f8fafc", font=("Meiryo UI", 9, "bold")).pack(anchor="w")

        sel_frame = tk.Frame(content_frame, bg="#f8fafc")
        sel_frame.pack(fill=tk.X, pady=(4, 10))

        curr_folder = os.path.dirname(curr_cfg)
        target_dir_var = tk.StringVar(value=curr_folder)
        dir_entry = ttk.Entry(sel_frame, textvariable=target_dir_var, font=("Meiryo UI", 9))
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        def _browse_folder():
            init_d = target_dir_var.get().strip().strip('"')
            if not init_d or not os.path.isdir(init_d):
                init_d = self.APP_DIR
            folder = filedialog.askdirectory(title="設定ファイル(config.json)の保存先フォルダを選択", initialdir=init_d, parent=win)
            if folder:
                target_dir_var.set(folder)

        ttk.Button(sel_frame, text="📁 フォルダを参照...", style="Action.TButton", command=_browse_folder).pack(side=tk.LEFT)

        # クイックアクションボタン行（「アプリと同じ場所に戻す」ボタンをわかりやすく配置！）
        action_row = tk.Frame(content_frame, bg="#f8fafc")
        action_row.pack(fill=tk.X, pady=(0, 12))

        def _set_default_folder_input():
            target_dir_var.set(self.APP_DIR)

        btn_to_default = tk.Button(
            action_row,
            text="🏠 アプリと同じ場所 (初期値) に戻す",
            bg="#f1f5f9",
            fg="#1d4ed8",
            activebackground="#e2e8f0",
            activeforeground="#1e40af",
            font=("Meiryo UI", 9, "bold"),
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=4,
            command=_set_default_folder_input,
        )
        btn_to_default.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(
            action_row,
            text="📂 現在のフォルダを開く",
            style="Action.TButton",
            command=lambda: os.startfile(os.path.dirname(self.get_config_file_path())),
        ).pack(side=tk.LEFT)

        # 既存設定の移行オプション
        copy_var = tk.BooleanVar(value=True)
        chk_copy = tk.Checkbutton(
            content_frame,
            text="変更前の設定内容を新しい保存先へ自動コピーして引き継ぐ (推奨)",
            variable=copy_var,
            bg="#f8fafc",
            activebackground="#f8fafc",
            font=("Meiryo UI", 9),
            anchor="w",
        )
        chk_copy.pack(fill=tk.X, pady=(0, 8))

        # 下部ボタンバー
        btn_bar = tk.Frame(win, bg="#f1f5f9", padx=16, pady=10, highlightbackground="#cbd5e1", highlightthickness=1)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        def _apply_location_change():
            sel_dir = target_dir_var.get().strip().strip('"')
            if not sel_dir:
                messagebox.showerror("入力エラー", "保存先フォルダを指定してください。", parent=win)
                return
            success = self._set_config_file_path(sel_dir, copy_existing=copy_var.get())
            if success:
                self._load_config_and_autoload()
                win.destroy()
                new_full = self.get_config_file_path()
                self.status_var.set(f"設定ファイルの保存先を変更しました: {new_full}")
                messagebox.showinfo(
                    "保存先変更完了",
                    f"設定ファイル (config.json) の保存先を変更しました:\n{new_full}\n\n"
                    "次回以降の「設定を保存」およびアプリ起動時の自動読み込みは、上記パスに対して行われます。",
                    parent=self.root,
                )

        ttk.Button(btn_bar, text="✔ この保存先を適用", style="Primary.TButton", command=_apply_location_change).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        ttk.Button(btn_bar, text="キャンセル", style="Action.TButton", command=win.destroy).pack(side=tk.RIGHT)

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

    def _collect_valid_order_files(
        self, folder_path: str, mgmt_filepath: str, valid_mgmt_order_nos: set[str], progress_cb=None
    ) -> tuple[list[dict], int, int]:
        """
        指定フォルダ配下を高速探索し、注文品管理表に存在する管理Noに対応する注文書のみを抽出する。
        1. サブフォルダ名に8桁の管理Noが含まれ、それが管理表に存在しない場合はフォルダごと丸ごとスキップ（枝刈り）
        2. ファイル名段階で拡張子・一時ファイル・見積書・管理No不一致を高速判定し、不要なI/Oを完全排除
        3. 未変更ファイルはメモリキャッシュから即座に復元し、新規・更新ファイルはマルチスレッド並列で高速解析
        """
        order_items = []
        valid_file_count = 0
        ignored_file_count = 0
        skipped_dir_count = 0

        mgmt_abspath = os.path.normcase(os.path.abspath(mgmt_filepath)) if mgmt_filepath else ""
        re_8digit = re.compile(r"(?<!\d)(\d{8})(?!\d)")

        candidate_files = []  # (full_path, rel_display, hint_order_no_8, mtime, size)

        # スタックを用いた os.scandir 枝刈り探索（不要なサブフォルダへのアクセス自体を遮断）
        scan_stack = [(folder_path, "")]
        scanned_dirs = 0

        while scan_stack:
            if self.cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

            curr_dir, inherited_no_8 = scan_stack.pop()
            scanned_dirs += 1

            if progress_cb and (scanned_dirs == 1 or scanned_dirs % 15 == 0):
                progress_cb(
                    36.0,
                    f"[2/3] フォルダ構造を高速フィルタリング中 (対象候補: {len(candidate_files)}件 / 不要フォルダスキップ: {skipped_dir_count}件)...",
                )

            try:
                with os.scandir(curr_dir) as it:
                    entries = list(it)
            except Exception:
                continue

            sub_dirs_to_push = []
            for entry in entries:
                if self.cancel_event.is_set():
                    raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

                entry_name = entry.name
                if entry_name.startswith("~$") or entry_name.startswith("."):
                    ignored_file_count += 1
                    continue

                try:
                    is_dir = entry.is_dir(follow_symlinks=False)
                except Exception:
                    is_dir = False

                if is_dir:
                    norm_dname = unicodedata.normalize("NFKC", entry_name).strip()
                    dir_8_list = re_8digit.findall(norm_dname)
                    matched_in_dir = [d for d in dir_8_list if d in valid_mgmt_order_nos]
                    non_date_in_dir = [d for d in dir_8_list if not d.startswith(("19", "20"))]

                    if matched_in_dir:
                        # 管理表に存在する管理Noのフォルダ → 探索対象として中へ入る
                        sub_dirs_to_push.append((entry.path, matched_in_dir[0]))
                    elif non_date_in_dir or dir_8_list:
                        # フォルダ名に8桁の管理Noがあるが管理表に一切存在しない → フォルダごと丸ごとスキップ！
                        skipped_dir_count += 1
                        ignored_file_count += 1
                        continue
                    else:
                        # 分類用フォルダ（8桁番号なし）→ 親の管理Noを引き継いで中へ入る
                        sub_dirs_to_push.append((entry.path, inherited_no_8))
                else:
                    if not entry_name.lower().endswith((".xlsx", ".xlsm")):
                        ignored_file_count += 1
                        continue

                    if "管理表" in entry_name or "納期超過チェック結果" in entry_name:
                        ignored_file_count += 1
                        continue

                    norm_fname = unicodedata.normalize("NFKC", entry_name).strip()

                    # ファイル名に「見積」が含まれ「注文」「発注」が含まれない場合は開かずに即スキップ
                    if "見積" in norm_fname and not any(k in norm_fname for k in ("注文", "発注")):
                        ignored_file_count += 1
                        continue

                    file_8_list = re_8digit.findall(norm_fname)
                    matched_in_file = [d for d in file_8_list if d in valid_mgmt_order_nos]
                    non_date_in_file = [d for d in file_8_list if not d.startswith(("19", "20"))]

                    if matched_in_file:
                        hint_no_8 = matched_in_file[0]
                    elif non_date_in_file or file_8_list:
                        # ファイル名に8桁番号があるが管理表に存在しない → Excelを開かずに即スキップ！
                        ignored_file_count += 1
                        continue
                    else:
                        hint_no_8 = inherited_no_8

                    full_path = entry.path
                    if os.path.normcase(os.path.abspath(full_path)) == mgmt_abspath:
                        ignored_file_count += 1
                        continue

                    try:
                        st = entry.stat(follow_symlinks=False)
                        mtime = st.st_mtime
                        fsize = st.st_size
                    except Exception:
                        mtime = 0.0
                        fsize = 0

                    rel_display = os.path.relpath(full_path, folder_path)
                    candidate_files.append((full_path, rel_display, hint_no_8, mtime, fsize))

            # フォルダ順を綺麗に保つため逆順でスタックに積む
            for sub_item in reversed(sorted(sub_dirs_to_push, key=lambda x: x[0])):
                scan_stack.append(sub_item)

        candidate_files.sort(key=lambda item: item[0])
        total_candidates = len(candidate_files)

        if total_candidates == 0:
            return [], 0, ignored_file_count

        # キャッシュヒット判定と新規解析対象の振り分け
        results_by_index: dict[int, list[dict]] = {}
        to_parse_tasks = []  # (idx, full_path, rel_display, hint_no_8, norm_path, mtime, fsize)

        for idx, (full_path, rel_display, hint_no_8, mtime, fsize) in enumerate(candidate_files):
            norm_path = os.path.normcase(os.path.abspath(full_path))
            cached_entry = self._order_file_cache.get(norm_path)
            if cached_entry is not None and mtime > 0 and cached_entry[0] == mtime and cached_entry[1] == fsize:
                cached_items = cached_entry[2]
                # キャッシュ内のアイテムのうち現在の管理表に含まれる管理Noのみ抽出
                valid_cached = [
                    dict(item)
                    for item in cached_items
                    if item.get("order_no_8") in valid_mgmt_order_nos
                ]
                results_by_index[idx] = valid_cached
            else:
                to_parse_tasks.append((idx, full_path, rel_display, hint_no_8, norm_path, mtime, fsize))

        # 未キャッシュ（初回または更新済み）の注文書Excelをマルチスレッド並列で高速読み込み
        if to_parse_tasks:
            total_to_parse = len(to_parse_tasks)
            completed_parse = 0

            def _worker_parse(task_tuple):
                t_idx, t_full_path, t_rel, t_hint_no, t_norm_path, t_mtime, t_fsize = task_tuple
                if self.cancel_event.is_set():
                    raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
                parsed = self._parse_order_book(
                    t_full_path, valid_mgmt_order_nos, folder_order_no_8=t_hint_no
                )
                return t_idx, t_rel, t_norm_path, t_mtime, t_fsize, parsed

            max_workers = min(6, max(1, total_to_parse))
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_task = {executor.submit(_worker_parse, t): t for t in to_parse_tasks}
                for future in as_completed(future_to_task):
                    if self.cancel_event.is_set():
                        for f in future_to_task:
                            f.cancel()
                        raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")

                    completed_parse += 1
                    try:
                        t_idx, t_rel, t_norm_path, t_mtime, t_fsize, parsed_items = future.result()
                        results_by_index[t_idx] = parsed_items
                        if t_mtime > 0:
                            with self._cache_lock:
                                self._order_file_cache[t_norm_path] = (t_mtime, t_fsize, parsed_items)
                    except AnalysisCancelledException:
                        for f in future_to_task:
                            f.cancel()
                        raise
                    except Exception:
                        t_orig = future_to_task[future]
                        results_by_index[t_orig[0]] = []

                    if progress_cb and (
                        completed_parse == 1
                        or completed_parse % 3 == 0
                        or completed_parse == total_to_parse
                    ):
                        pct = 38.0 + (36.0 * completed_parse / max(1, total_to_parse))
                        progress_cb(
                            pct,
                            f"[2/3] 対象の注文書を読み込み中 ({completed_parse}/{total_to_parse}件 / 不要フォルダ{skipped_dir_count}件スキップ済)",
                        )
        else:
            if progress_cb:
                progress_cb(
                    74.0,
                    f"[2/3] キャッシュから注文書データを高速復元しました ({total_candidates}件 / 不要フォルダ{skipped_dir_count}件スキップ済)",
                )

        # 元のファイル順序通りに結果を集約
        for idx in range(total_candidates):
            extracted = results_by_index.get(idx, [])
            if extracted:
                order_items.extend(extracted)
                valid_file_count += 1
            else:
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
        norm_path = os.path.normcase(os.path.abspath(filepath))
        mtime = 0.0
        fsize = 0
        try:
            st = os.stat(filepath)
            mtime = st.st_mtime
            fsize = st.st_size
        except Exception:
            pass

        if (
            self._mgmt_file_cache is not None
            and mtime > 0
            and self._mgmt_file_cache[0] == norm_path
            and self._mgmt_file_cache[1] == mtime
            and self._mgmt_file_cache[2] == fsize
            and (
                not self._mgmt_file_cache[3]
                or ("instruction_state" in self._mgmt_file_cache[3][0] and "ship_to" in self._mgmt_file_cache[3][0])
            )
        ):
            if progress_cb:
                progress_cb(33.0, "[1/3] 注文品管理表をメモリキャッシュから高速復元しました...")
            return [dict(r, matched=False, matched_by="") for r in self._mgmt_file_cache[3]]

        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        mgmt_rows = []
        sheet_names = wb.sheetnames
        total_sheets = len(sheet_names)

        for s_idx, sheet_name in enumerate(sheet_names, start=1):
            if self.cancel_event.is_set():
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

            sheet_default_order_no = self._extract_8digit_no(sheet_name)

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
                        elif "指示書" in v_comb:
                            col_map.setdefault("instruction_doc", c_idx)
                    break

            if header_pos_in_list is None:
                continue

            col_map.setdefault("actual_date", 1)       # A列: 現品確認日
            col_map.setdefault("instruction_doc", 10)  # J列: 指示書
            col_map.setdefault("order_no", 13)         # M列: 管理No
            col_map.setdefault("ship_to", 14)          # N列: 出荷先
            col_map.setdefault("line_name", 15)        # O列: 品名、ライン名等
            col_map.setdefault("product_detail", 16)   # P列: 製品詳細
            col_map.setdefault("item_name", 20)        # T列: 品名
            col_map.setdefault("maker", 21)            # U列: メーカー
            col_map.setdefault("spec", 23)             # W列: 仕様
            col_map["qty"] = 28                        # AB列: 購入数
            col_map.setdefault("invoice_date", 35)     # AI列: 請求日

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
                ship_to_val = _get_col(row_tuple, col_map.get("ship_to"))
                line_val = _get_col(row_tuple, col_map.get("line_name"))
                prod_detail_val = _get_col(row_tuple, col_map.get("product_detail"))
                order_no_val = _get_col(row_tuple, col_map.get("order_no"))

                qty_val = _get_col(row_tuple, 28)
                if qty_val is None:
                    qty_val = _get_col(row_tuple, col_map.get("qty"))

                if not any([drawing_val, part_val, item_val, spec_val, maker_val, line_val, prod_detail_val]):
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
                ship_to_str = self._clean_cell_str(ship_to_val)
                line_str = self._clean_cell_str(line_val)
                prod_detail_str = self._clean_cell_str(prod_detail_val)
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

                # J列 (10列目) の指示書列を読み取り
                j_col_val = _get_col(row_tuple, 10)
                mapped_inst_val = _get_col(row_tuple, col_map.get("instruction_doc"))
                inst_raw_val = j_col_val if (j_col_val is not None and str(j_col_val).strip() != "") else mapped_inst_val
                instruction_raw = self._clean_cell_str(inst_raw_val)
                instruction_state = self._classify_instruction_col(instruction_raw)
                instruction_disp = instruction_raw if instruction_raw else "(空白)"

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
                        "instruction_raw": instruction_raw,
                        "instruction_disp": instruction_disp,
                        "instruction_state": instruction_state,
                        "ship_to": ship_to_str,
                        "product_detail": prod_detail_str,
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
        if mtime > 0:
            self._mgmt_file_cache = (
                norm_path,
                mtime,
                fsize,
                [dict(r) for r in mgmt_rows],
            )
        return mgmt_rows

    def _parse_order_book(
        self, filepath: str, valid_mgmt_order_nos: set[str], folder_order_no_8: str = ""
    ) -> list[dict]:
        if self.cancel_event.is_set():
            raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        order_items = []
        filename = os.path.basename(filepath)
        norm_fname = unicodedata.normalize("NFKC", filename)

        prefix_match = re.match(r"^\s*(\d{8})", norm_fname)
        if prefix_match:
            file_order_no_8 = prefix_match.group(1)
        else:
            file_order_no_8 = self._extract_8digit_no(norm_fname)

        if not file_order_no_8 and folder_order_no_8:
            file_order_no_8 = folder_order_no_8

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
            if self.cancel_event.is_set():
                wb.close()
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
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

                # 2行目を読んだ瞬間に「見積依頼書」かどうかを判定し、見積なら3行目以降を一切読まず即座に打ち切る
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
                        d_in_cell = self._parse_flexible_date(cell_val, allow_period=False)
                        if d_in_cell:
                            sheet_order_date = d_in_cell
                            default_year = d_in_cell.year
                        else:
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

        m_inst_disp = (m_row.get("instruction_disp", "") if m_row else "") or "— (管理表なし)"
        rows.append(("⑩ 指示書 (管理表J列)", "— (注文書側なし)", m_inst_disp, f"状態: {m_inst_disp}", "INFO"))

        m_ship_disp = (m_row.get("ship_to", "") if m_row else "") or "(空欄)"
        rows.append(("⑪ 出荷先 (管理表N列)", "— (注文書側なし)", m_ship_disp, "管理表情報", "INFO"))

        m_line_disp = (m_row.get("line_name", "") if m_row else "") or "(空欄)"
        rows.append(("⑫ 品名・ライン名等 (管理表O列)", "— (注文書側なし)", m_line_disp, "管理表情報", "INFO"))

        m_prod_disp = (m_row.get("product_detail", "") if m_row else "") or "(空欄)"
        rows.append(("⑬ 製品詳細 (管理表P列)", "— (注文書側なし)", m_prod_disp, "管理表情報", "INFO"))

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
            diag["best_cand_instruction_raw"] = best_cand.get("instruction_raw", "")
            diag["best_cand_instruction_disp"] = best_cand.get("instruction_disp", "(空白)")
            diag["best_cand_instruction_state"] = best_cand.get("instruction_state", "BLANK")
            diag["best_cand_ship_to"] = best_cand.get("ship_to", "")
            diag["best_cand_line_name"] = best_cand.get("line_name", "")
            diag["best_cand_product_detail"] = best_cand.get("product_detail", "")
            return reason, diag

        supp_cand = same_no_rows[0] if same_no_rows else None
        reason = f"【明細未記載】管理No[{o_no}]は管理表にありますが一致する明細行がありません"
        diag = self._build_field_comparison_data(o_item, None, reason)
        diag["best_cand_instruction_raw"] = supp_cand.get("instruction_raw", "") if supp_cand else ""
        diag["best_cand_instruction_disp"] = supp_cand.get("instruction_disp", "(未記載)") if supp_cand else "(未記載)"
        diag["best_cand_instruction_state"] = supp_cand.get("instruction_state", "BLANK") if supp_cand else "BLANK"
        diag["best_cand_ship_to"] = supp_cand.get("ship_to", "") if supp_cand else ""
        diag["best_cand_line_name"] = supp_cand.get("line_name", "") if supp_cand else ""
        diag["best_cand_product_detail"] = supp_cand.get("product_detail", "") if supp_cand else ""
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
            if self.cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
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
                if self.cancel_event.is_set():
                    raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
                processed_orders += 1
                if progress_cb and (processed_orders == 1 or processed_orders % 5 == 0 or processed_orders == total_orders):
                    pct = 75.0 + (18.0 * processed_orders / max(1, total_orders))
                    progress_cb(pct, f"[3/3] 注文書明細と管理表を照合・診断中 ({processed_orders}/{total_orders} 件目)...")

                order_due = o_item.get("due_date")
                order_dt = o_item.get("order_date")
                if not order_dt and r_idx in matched_pairs:
                    order_dt = m_list[matched_pairs[r_idx][0]].get("order_date")

                is_fallback = False
                if order_due:
                    final_due = order_due
                elif order_dt:
                    final_due = order_dt + datetime.timedelta(days=grace_days)
                    is_fallback = True
                else:
                    final_due = m_list[matched_pairs[r_idx][0]].get("mgmt_due_date") if r_idx in matched_pairs else None

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
                        instruction_raw=best_mgmt.get("instruction_raw", ""),
                        instruction_disp=best_mgmt.get("instruction_disp", "(空白)"),
                        instruction_state=best_mgmt.get("instruction_state", "BLANK"),
                        ship_to=best_mgmt.get("ship_to", ""),
                        line_name=best_mgmt.get("line_name", ""),
                        product_detail=best_mgmt.get("product_detail", ""),
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
                        instruction_raw=diag_data.get("best_cand_instruction_raw", ""),
                        instruction_disp=diag_data.get("best_cand_instruction_disp", "(未記載)"),
                        instruction_state=diag_data.get("best_cand_instruction_state", "BLANK"),
                        ship_to=diag_data.get("best_cand_ship_to", ""),
                        line_name=diag_data.get("best_cand_line_name", ""),
                        product_detail=diag_data.get("best_cand_product_detail", ""),
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
            if self.cancel_event.is_set():
                raise AnalysisCancelledException("照合処理がユーザーによって中止されました。")
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
                "instruction_raw": m.get("instruction_raw", ""),
                "instruction_disp": m.get("instruction_disp", "(空白)"),
                "instruction_state": m.get("instruction_state", "BLANK"),
                "ship_to": m.get("ship_to", ""),
                "line_name": m.get("line_name", ""),
                "product_detail": m.get("product_detail", ""),
                "order_date": m.get("order_date"),
                "order_date_str": m["order_date"].strftime("%Y-%m-%d") if m.get("order_date") else "-",
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
            "instruction_raw": instruction_raw,
            "instruction_disp": instruction_disp,
            "instruction_state": instruction_state,
            "ship_to": ship_to,
            "line_name": line_name,
            "product_detail": product_detail,
            "order_date": order_date,
            "order_date_str": order_date.strftime("%Y-%m-%d") if order_date else "-",
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
            if 0 <= col_index < len(self.col_configs_main):
                col_id = self.col_configs_main[col_index][0]
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

        if self.is_running or self.is_excel_running:
            return

        if not HAS_WIN32COM:
            messagebox.showerror(
                "モジュール未検出",
                "Excelを直接起動するには 'pywin32' パッケージが必要です。\n"
                "ターミナルまたはコマンドプロンプトで以下を実行してください:\n"
                "pip install pywin32",
            )
            return

        self.is_excel_running = True
        self.excel_cancel_event.clear()
        self.excel_open_btn.configure(state=tk.DISABLED, text="Excel生成中...")
        self.excel_cancel_btn.configure(state=tk.NORMAL, text="⏹ 中止")
        self.cancel_btn.configure(state=tk.NORMAL, text="⏹ 中止 (Esc)")
        self.run_btn.configure(state=tk.DISABLED)
        self.progress_var.set(10.0)
        self.progress_text_var.set("Excelを起動しています...")
        self.status_var.set("Excelを起動し、新規ブックにデータを転送しています（一時ファイルなし）...")

        def _excel_worker():
            excel = None
            wb = None
            try:
                def _check_cancelled():
                    if self.excel_cancel_event.is_set():
                        if wb is not None:
                            try:
                                wb.Close(SaveChanges=False)
                            except Exception:
                                pass
                        if excel is not None:
                            try:
                                excel.ScreenUpdating = True
                                excel.DisplayAlerts = True
                                if excel.Workbooks.Count == 0:
                                    excel.Quit()
                            except Exception:
                                pass
                        raise AnalysisCancelledException("Excel直接展開処理がユーザーによって中止されました。")

                pythoncom.CoInitialize()
                _check_cancelled()

                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                except Exception:
                    excel = win32com.client.Dispatch("Excel.Application")

                _check_cancelled()
                excel.Visible = True
                excel.ScreenUpdating = False
                excel.DisplayAlerts = False

                self.ui_queue.put(("PROGRESS", (25.0, "新規Excelブックを生成しています...")))

                # 新規ブックの作成（一時ファイルはディスク上に作成されない）
                wb = excel.Workbooks.Add()
                _check_cancelled()

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
                self.ui_queue.put(("PROGRESS", (45.0, "シート1 (納期超過・照合結果) にデータを書き込み中...")))
                _check_cancelled()

                headers_main = [
                    "総合判定",
                    "不一致理由・照合診断",
                    "依頼区分",
                    "管理No(注文番号)",
                    "指示書(J列)",
                    "出荷先",
                    "品名・ライン名等",
                    "製品詳細",
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
                excluded_nos = self._get_excluded_order_nos()
                base_date = self._parse_flexible_date(self.base_date_str.get()) or datetime.date.today()
                exclude_months = self._get_exclude_months()
                inst_mode = self._get_instruction_filter_mode()
                for r in self.analyzed_records:
                    if self._is_record_excluded(r, excluded_nos, base_date, exclude_months, inst_mode):
                        continue
                    rows_main_data.append(
                        [
                            r["status_badge"],
                            r["mismatch_reason"],
                            r["req_type"],
                            r["order_no"],
                            r.get("instruction_disp", ""),
                            r.get("ship_to", ""),
                            r.get("line_name", ""),
                            r.get("product_detail", ""),
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

                _check_cancelled()

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

                    _check_cancelled()

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

                _check_cancelled()

                # -------------------------------------------------------------
                # 2. シート2: 注文書未検出(管理表のみ)
                # -------------------------------------------------------------
                self.ui_queue.put(("PROGRESS", (75.0, "シート2 (注文書未検出) にデータを書き込み中...")))
                _check_cancelled()

                headers_missing = [
                    "状態判定",
                    "管理表シート",
                    "行番号",
                    "管理No(8桁)",
                    "指示書(J列)",
                    "出荷先",
                    "品名・ライン名等",
                    "製品詳細",
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
                    if self._is_record_excluded(r, excluded_nos, base_date, exclude_months, inst_mode):
                        continue
                    rows_missing_data.append(
                        [
                            r["status_badge"],
                            r["mgmt_sheet"],
                            r["mgmt_row"],
                            r["order_no"],
                            r.get("instruction_disp", ""),
                            r.get("ship_to", ""),
                            r.get("line_name", ""),
                            r.get("product_detail", ""),
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

                _check_cancelled()

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

                    _check_cancelled()

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

                _check_cancelled()

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

            except AnalysisCancelledException:
                self.ui_queue.put(("EXCEL_OPENED_CANCELLED", None))
            except Exception as e:
                if excel is not None:
                    try:
                        excel.ScreenUpdating = True
                        excel.DisplayAlerts = True
                    except Exception:
                        pass
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
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("StockMatch.Miya.v1.4.0")
    except Exception:
        pass

    root = tk.Tk()
    _app = StockMatchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
