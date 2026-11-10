"""
StockMatch_Miya ダイアログ画面モジュール
表示列設定、除外No設定、保存先設定、照合詳細・シートプレビュー、Excel風列フィルター、ヘルプ、バージョン情報を集約
"""

import os
import re
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl

from .constants import APP_NAME, APP_VERSION, APP_TITLE, DEFAULT_VISIBLE_COLUMNS_MAIN
from .date_utils import parse_lines_to_order_nos, extract_8digit_no


def apply_window_icon(window: tk.Wm, icon_path: str | None):
    """ダイアログウィンドウにアプリアイコンを設定"""
    if icon_path and os.path.isfile(icon_path):
        try:
            window.iconbitmap(icon_path)
        except Exception:
            pass


def open_column_visibility_dialog(
    parent: tk.Tk | tk.Toplevel,
    col_configs_main: list,
    default_visible_cols: list,
    current_visible_cols: list,
    icon_path: str | None,
    on_apply,
):
    """照合結果一覧（タブ1）の表示・非表示列を設定するモーダルダイアログ"""
    win = tk.Toplevel(parent)
    win.title("照合結果一覧 - 表示列の設定")
    apply_window_icon(win, icon_path)
    win.geometry("520x640")
    win.minsize(460, 520)
    win.transient(parent)
    win.grab_set()

    header = tk.Frame(win, bg="#1e293b", padx=16, pady=12)
    header.pack(fill=tk.X)
    tk.Label(
        header,
        text="👁 照合結果一覧（タブ①）の表示列設定",
        font=("Meiryo UI", 11, "bold"),
        bg="#1e293b",
        fg="#ffffff",
    ).pack(anchor="w")
    tk.Label(
        header,
        text="一覧に表示したい列にチェックを入れてください。チェックを外した列は非表示になります。",
        font=("Meiryo UI", 9),
        bg="#1e293b",
        fg="#cbd5e1",
    ).pack(anchor="w", pady=(4, 0))

    action_bar = ttk.Frame(win, padding=(16, 8, 16, 4))
    action_bar.pack(fill=tk.X)

    check_vars = {}
    for col_id, title, _, _ in col_configs_main:
        is_checked = col_id in current_visible_cols
        check_vars[col_id] = tk.BooleanVar(value=is_checked)

    def _select_all():
        for var in check_vars.values():
            var.set(True)

    def _select_default():
        for cid, var in check_vars.items():
            var.set(cid in default_visible_cols)

    def _clear_all():
        for var in check_vars.values():
            var.set(False)

    ttk.Button(action_bar, text="すべて選択", style="Action.TButton", command=_select_all).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(action_bar, text=f"初期値に戻す ({len(default_visible_cols)}列)", style="Action.TButton", command=_select_default).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(action_bar, text="すべて解除", style="Action.TButton", command=_clear_all).pack(side=tk.LEFT)

    list_frame = ttk.Frame(win, padding=(16, 6, 16, 8))
    list_frame.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(list_frame, bg="#ffffff", highlightthickness=1, highlightbackground="#cbd5e1")
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    scrollable_content = tk.Frame(canvas, bg="#ffffff", padx=12, pady=8)

    scrollable_content.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas_window = canvas.create_window((0, 0), window=scrollable_content, anchor="nw")

    def _on_canvas_configure(event):
        canvas.itemconfig(canvas_window, width=event.width)

    canvas.bind("<Configure>", _on_canvas_configure)
    canvas.configure(yscrollcommand=scrollbar.set)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    for idx, (col_id, title, _, _) in enumerate(col_configs_main, start=1):
        row_bg = "#f8fafc" if idx % 2 == 0 else "#ffffff"
        item_frame = tk.Frame(scrollable_content, bg=row_bg, padx=8, pady=4)
        item_frame.pack(fill=tk.X, pady=1)

        is_default = col_id in default_visible_cols
        default_tag = " [初期表示]" if is_default else " (初期非表示)"
        tag_color = "#0369a1" if is_default else "#64748b"

        cb = tk.Checkbutton(
            item_frame,
            text=f"{idx:2d}. {title}",
            variable=check_vars[col_id],
            font=("Meiryo UI", 9, "bold" if is_default else "normal"),
            bg=row_bg,
            activebackground=row_bg,
            selectcolor="#ffffff",
            anchor="w",
        )
        cb.pack(side=tk.LEFT)

        tk.Label(
            item_frame,
            text=default_tag,
            font=("Meiryo UI", 8),
            bg=row_bg,
            fg=tag_color,
        ).pack(side=tk.LEFT, padx=(6, 0))

    footer = ttk.Frame(win, padding=(16, 12, 16, 14))
    footer.pack(fill=tk.X)

    def _get_selected_cols():
        selected = [cid for cid, var in check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("選択エラー", "少なくとも1つの列を選択してください。", parent=win)
            return None
        return selected

    def _apply_and_close(save_to_config=False):
        selected = _get_selected_cols()
        if not selected:
            return
        canvas.unbind_all("<MouseWheel>")
        win.destroy()
        if on_apply:
            on_apply(selected, save_to_config)

    def _on_cancel():
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", _on_cancel)

    ttk.Button(
        footer,
        text="💾 適用して保存",
        style="Primary.TButton",
        command=lambda: _apply_and_close(save_to_config=True),
    ).pack(side=tk.RIGHT, padx=(6, 0))

    ttk.Button(
        footer,
        text="✔ 適用 (保存なし)",
        style="Action.TButton",
        command=lambda: _apply_and_close(save_to_config=False),
    ).pack(side=tk.RIGHT, padx=(6, 0))

    ttk.Button(
        footer,
        text="キャンセル",
        style="Action.TButton",
        command=_on_cancel,
    ).pack(side=tk.RIGHT)


def open_exclude_dialog(
    parent: tk.Tk | tk.Toplevel,
    current_excluded_text: str,
    icon_path: str | None,
    on_apply,
):
    """除外する管理Noを1行ずつ編集・登録・Excel一括貼り付けする専用ダイアログ"""
    win = tk.Toplevel(parent)
    win.title("除外する管理Noの設定・編集")
    apply_window_icon(win, icon_path)
    win.geometry("540x620")
    win.minsize(460, 480)
    win.transient(parent)
    win.grab_set()

    header = tk.Frame(win, bg="#1e293b", padx=16, pady=12)
    header.pack(fill=tk.X)
    tk.Label(
        header,
        text="🚫 一覧から除外する管理No（テスト案件・完了案件）の設定",
        font=("Meiryo UI", 11, "bold"),
        bg="#1e293b",
        fg="#ffffff",
    ).pack(anchor="w")
    tk.Label(
        header,
        text="1行に1つの管理Noを入力してください。Excelのセル範囲をまとめてコピーして貼り付けも可能です。",
        font=("Meiryo UI", 9),
        bg="#1e293b",
        fg="#cbd5e1",
    ).pack(anchor="w", pady=(4, 0))

    cnt_label = ttk.Label(win, text="現在の登録件数: 0 件", font=("Meiryo UI", 9, "bold"), foreground="#0369a1")
    cnt_label.pack(anchor="w", padx=16, pady=(10, 4))

    txt_frame = ttk.Frame(win, padding=(16, 0, 16, 6))
    txt_frame.pack(fill=tk.BOTH, expand=True)

    txt = tk.Text(
        txt_frame,
        wrap="none",
        font=("Consolas", 10),
        undo=True,
        maxundo=50,
        highlightbackground="#cbd5e1",
        highlightcolor="#2563eb",
        highlightthickness=1,
    )
    vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
    hsb = ttk.Scrollbar(txt_frame, orient="horizontal", command=txt.xview)
    txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    txt.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    txt_frame.grid_rowconfigure(0, weight=1)
    txt_frame.grid_columnconfigure(0, weight=1)

    initial_items = parse_lines_to_order_nos(current_excluded_text)
    if initial_items:
        txt.insert("1.0", "\n".join(initial_items))

    def _update_count():
        content = txt.get("1.0", tk.END)
        items = parse_lines_to_order_nos(content)
        cnt_label.configure(text=f"現在の登録件数: {len(items)} 件 (重複・空行は自動整理されます)")

    _update_count()
    txt.bind("<KeyRelease>", lambda _e: _update_count())

    toolbar = ttk.Frame(win, padding=(16, 4))
    toolbar.pack(fill=tk.X)

    def _paste_clipboard():
        try:
            cb_text = win.clipboard_get()
            if not cb_text:
                return
            items = parse_lines_to_order_nos(cb_text)
            if not items:
                messagebox.showwarning("貼り付け", "クリップボード内に有効な管理Noが見つかりませんでした。", parent=win)
                return
            curr_text = txt.get("1.0", tk.END)
            curr_items = parse_lines_to_order_nos(curr_text)
            merged = []
            seen = set()
            for it in curr_items + items:
                if it not in seen:
                    seen.add(it)
                    merged.append(it)
            txt.delete("1.0", tk.END)
            txt.insert("1.0", "\n".join(merged))
            _update_count()
            messagebox.showinfo("貼り付け完了", f"クリップボードから {len(items)} 件の管理Noを追加しました。", parent=win)
        except Exception as e:
            messagebox.showerror("エラー", f"クリップボードの読み取りに失敗しました:\n{e}", parent=win)

    def _cleanup_text():
        content = txt.get("1.0", tk.END)
        items = parse_lines_to_order_nos(content)
        txt.delete("1.0", tk.END)
        if items:
            txt.insert("1.0", "\n".join(items))
        _update_count()

    def _clear_all():
        if messagebox.askyesno("全消去", "入力中の除外管理Noをすべてクリアしますか？", parent=win):
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

    btn_bar = tk.Frame(win, bg="#f1f5f9", padx=16, pady=10, highlightbackground="#cbd5e1", highlightthickness=1)
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _apply_and_save():
        content = txt.get("1.0", tk.END)
        items = parse_lines_to_order_nos(content)
        win.destroy()
        if on_apply:
            on_apply(items)

    ttk.Button(btn_bar, text="✔ 適用して設定を保存", style="Primary.TButton", command=_apply_and_save).pack(
        side=tk.RIGHT, padx=(8, 0)
    )
    ttk.Button(btn_bar, text="キャンセル", style="Action.TButton", command=win.destroy).pack(side=tk.RIGHT)


def open_config_location_dialog(
    parent: tk.Tk | tk.Toplevel,
    config_mgr,
    icon_path: str | None,
    on_success,
):
    """設定ファイル (config.json) の保存先パスを変更・初期値復帰するダイアログ"""
    win = tk.Toplevel(parent)
    win.title("設定ファイル (config.json) の保存先設定")
    apply_window_icon(win, icon_path)
    win.geometry("640x360")
    win.minsize(560, 320)
    win.transient(parent)
    win.grab_set()

    header = tk.Frame(win, bg="#1e293b", padx=16, pady=12)
    header.pack(fill=tk.X)
    tk.Label(
        header,
        text="⚙ 設定ファイル (config.json) の保存場所を指定・変更",
        font=("Meiryo UI", 11, "bold"),
        bg="#1e293b",
        fg="#ffffff",
    ).pack(anchor="w")
    tk.Label(
        header,
        text="共有サーバーやGoogleドライブ等の同期フォルダを指定することで、複数人で同じ除外設定等を共有できます。",
        font=("Meiryo UI", 9),
        bg="#1e293b",
        fg="#cbd5e1",
    ).pack(anchor="w", pady=(4, 0))

    body = ttk.Frame(win, padding=16)
    body.pack(fill=tk.BOTH, expand=True)

    curr_path = config_mgr.get_config_file_path()
    ttk.Label(body, text="現在の設定ファイル保存先パス:", font=("Meiryo UI", 9, "bold")).pack(anchor="w")
    curr_lbl = ttk.Label(body, text=curr_path, font=("Consolas", 9), foreground="#0369a1", wraplength=580)
    curr_lbl.pack(anchor="w", pady=(2, 12))

    new_path_var = tk.StringVar(value=curr_path)
    copy_existing_var = tk.BooleanVar(value=True)

    ttk.Label(body, text="新しい保存先フォルダまたはファイルパス:", font=("Meiryo UI", 9, "bold")).pack(anchor="w")
    input_frame = ttk.Frame(body)
    input_frame.pack(fill=tk.X, pady=(4, 8))

    entry = ttk.Entry(input_frame, textvariable=new_path_var, font=("Meiryo UI", 9))
    entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

    def _browse_dir():
        initial = os.path.dirname(new_path_var.get()) or config_mgr.app_dir
        d = filedialog.askdirectory(title="設定ファイルの保存先フォルダを選択", initialdir=initial, parent=win)
        if d:
            new_path_var.set(os.path.join(d, "config.json"))

    def _browse_file():
        initial = os.path.dirname(new_path_var.get()) or config_mgr.app_dir
        f = filedialog.askopenfilename(
            title="設定ファイル(config.json)を選択",
            initialdir=initial,
            filetypes=[("JSON 設定ファイル", "*.json"), ("すべてのファイル", "*.*")],
            parent=win,
        )
        if f:
            new_path_var.set(f)

    ttk.Button(input_frame, text="フォルダ参照...", style="Action.TButton", command=_browse_dir).pack(
        side=tk.LEFT, padx=(0, 4)
    )
    ttk.Button(input_frame, text="ファイル参照...", style="Action.TButton", command=_browse_file).pack(side=tk.LEFT)

    ttk.Checkbutton(
        body,
        text="変更前の設定ファイルの内容を、新しい保存先へ自動コピーして引き継ぐ",
        variable=copy_existing_var,
    ).pack(anchor="w", pady=(4, 10))

    btn_bar = ttk.Frame(win, padding=(16, 10))
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _do_reset():
        if messagebox.askyesno(
            "初期値に戻す",
            "設定ファイルの保存先をアプリ配置フォルダ（初期値）に戻しますか？",
            parent=win,
        ):
            ok, res = config_mgr.reset_to_default(copy_existing=False)
            if ok:
                win.destroy()
                if on_success:
                    on_success(res, is_reset=True)
            else:
                messagebox.showerror("エラー", res, parent=win)

    def _do_apply():
        target = new_path_var.get().strip()
        if not target:
            messagebox.showwarning("入力エラー", "保存先パスを入力してください。", parent=win)
            return
        ok, res = config_mgr.set_config_file_path(target, copy_existing=copy_existing_var.get())
        if ok:
            win.destroy()
            if on_success:
                on_success(res, is_reset=False)
        else:
            messagebox.showerror("設定エラー", res, parent=win)

    ttk.Button(btn_bar, text="🏠 アプリと同じ場所 (初期値) に戻す", style="Action.TButton", command=_do_reset).pack(
        side=tk.LEFT
    )
    ttk.Button(btn_bar, text="✔ 適用して変更", style="Primary.TButton", command=_do_apply).pack(
        side=tk.RIGHT, padx=(6, 0)
    )
    ttk.Button(btn_bar, text="キャンセル", style="Action.TButton", command=win.destroy).pack(side=tk.RIGHT)


def open_column_filter_dialog(
    parent: tk.Tk | tk.Toplevel,
    col_id: str,
    title: str,
    records: list[dict],
    filter_dict: dict,
    on_refresh,
    on_update_heading,
    on_sort,
    x: int,
    y: int,
):
    """列ヘッダークリックによるExcel風のフィルター＆ソートダイアログ"""
    win = tk.Toplevel(parent)
    win.title(f"フィルター: {title}")
    win.attributes("-topmost", True)
    win.resizable(False, False)

    win_w = 340
    win_h = 440
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    pos_x = min(x, screen_w - win_w - 20)
    pos_y = min(y, screen_h - win_h - 40)
    win.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    sort_frame = ttk.Frame(win, padding=8)
    sort_frame.pack(fill=tk.X)

    def _do_sort(rev):
        win.destroy()
        if on_sort:
            on_sort(col_id, rev)

    ttk.Button(sort_frame, text="▲ 昇順で並べ替え", command=lambda: _do_sort(False)).pack(
        side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4)
    )
    ttk.Button(sort_frame, text="▼ 降順で並べ替え", command=lambda: _do_sort(True)).pack(
        side=tk.LEFT, fill=tk.X, expand=True
    )

    ttk.Separator(win, orient="horizontal").pack(fill=tk.X, pady=4)

    search_frame = ttk.Frame(win, padding=(8, 4))
    search_frame.pack(fill=tk.X)
    search_var = tk.StringVar()
    ttk.Label(search_frame, text="検索:").pack(side=tk.LEFT)
    search_entry = ttk.Entry(search_frame, textvariable=search_var)
    search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

    val_counts = {}
    for r in records:
        v = str(r.get(col_id, "") or "").strip()
        if not v:
            v = "(空白)"
        val_counts[v] = val_counts.get(v, 0) + 1

    sorted_vals = sorted(val_counts.keys())
    currently_allowed = filter_dict.get(col_id, None)

    check_vars = {}
    for v in sorted_vals:
        init_state = True if currently_allowed is None else (v in currently_allowed)
        check_vars[v] = tk.BooleanVar(value=init_state)

    list_container = ttk.Frame(win, padding=8)
    list_container.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(list_container, bg="#ffffff", highlightthickness=1, highlightbackground="#cbd5e1")
    scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg="#ffffff")

    scroll_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    c_win = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(c_win, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    cb_widgets = []
    for v in sorted_vals:
        c_frame = tk.Frame(scroll_frame, bg="#ffffff")
        c_frame.pack(fill=tk.X, pady=1)
        cb = tk.Checkbutton(
            c_frame,
            text=f"{v} ({val_counts[v]})",
            variable=check_vars[v],
            bg="#ffffff",
            anchor="w",
            font=("Meiryo UI", 9),
        )
        cb.pack(fill=tk.X)
        cb_widgets.append((v, c_frame))

    def _filter_cb_list(*_args):
        query = search_var.get().strip().lower()
        for v, frame in cb_widgets:
            if not query or query in v.lower():
                frame.pack(fill=tk.X, pady=1)
            else:
                frame.pack_forget()

    search_var.trace_add("write", _filter_cb_list)

    btn_row = ttk.Frame(win, padding=(8, 2))
    btn_row.pack(fill=tk.X)

    def _select_all(state):
        for v, frame in cb_widgets:
            if frame.winfo_ismapped():
                check_vars[v].set(state)

    ttk.Button(btn_row, text="すべて選択", command=lambda: _select_all(True)).pack(side=tk.LEFT, padx=2)
    ttk.Button(btn_row, text="すべて解除", command=lambda: _select_all(False)).pack(side=tk.LEFT, padx=2)

    footer = ttk.Frame(win, padding=8)
    footer.pack(fill=tk.X)

    def _apply_filter():
        allowed = {v for v, var in check_vars.items() if var.get()}
        if len(allowed) == len(sorted_vals):
            filter_dict.pop(col_id, None)
        else:
            filter_dict[col_id] = allowed
        win.destroy()
        if on_update_heading:
            on_update_heading(col_id)
        if on_refresh:
            on_refresh()

    def _clear_filter():
        filter_dict.pop(col_id, None)
        win.destroy()
        if on_update_heading:
            on_update_heading(col_id)
        if on_refresh:
            on_refresh()

    ttk.Button(footer, text="フィルター解除", command=_clear_filter).pack(side=tk.LEFT)
    ttk.Button(footer, text="OK", style="Primary.TButton", command=_apply_filter).pack(side=tk.RIGHT, padx=(4, 0))
    ttk.Button(footer, text="キャンセル", command=win.destroy).pack(side=tk.RIGHT)


def open_detail_preview_dialog(
    parent: tk.Tk | tk.Toplevel,
    rec: dict,
    cached_mgmt_rows: list[dict],
    is_missing_tab: bool,
    icon_path: str | None,
    on_open_order,
    on_open_mgmt,
):
    """照合詳細・注文書シートプレビュー・管理表シートプレビューを表示する3タブダイアログ"""
    diag_data = rec.get("diagnostic_data", {})
    order_no = rec.get("order_no", "")
    item_title = rec.get("item_name") or rec.get("model_spec") or "明細詳細"

    win = tk.Toplevel(parent)
    win.title(f"照合詳細＆シートプレビュー - 管理No: {order_no} ({item_title})")
    apply_window_icon(win, icon_path)
    win.geometry("980x740")
    win.minsize(860, 600)

    header = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    header.pack(fill=tk.X)
    tk.Label(
        header,
        text=f"🔍 照合診断・シートプレビュー [管理No: {order_no}]",
        font=("Meiryo UI", 12, "bold"),
        bg="#1e293b",
        fg="#ffffff",
    ).pack(side=tk.LEFT)

    action_bar = tk.Frame(win, bg="#f1f5f9", padx=16, pady=6, highlightbackground="#cbd5e1", highlightthickness=1)
    action_bar.pack(fill=tk.X)

    if not is_missing_tab:
        ttk.Button(
            action_bar,
            text="📄 注文書をExcelで開く (読取専用保護)",
            style="Action.TButton",
            command=lambda: on_open_order(rec) if on_open_order else None,
        ).pack(side=tk.LEFT, padx=(0, 8))

    ttk.Button(
        action_bar,
        text="📋 管理表をExcelで開く (該当行ジャンプ)",
        style="Action.TButton",
        command=lambda: on_open_mgmt(rec) if on_open_mgmt else None,
    ).pack(side=tk.LEFT)

    notebook = ttk.Notebook(win, padding=8)
    notebook.pack(fill=tk.BOTH, expand=True)

    # タブ1: 項目比較・診断レポート
    tab_cmp = ttk.Frame(notebook, padding=8)
    notebook.add(tab_cmp, text="  ① 🔍 項目比較・診断レポート  ")

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

    cmp_vsb = ttk.Scrollbar(tab_cmp, orient="vertical", command=cmp_tree.yview)
    cmp_tree.configure(yscrollcommand=cmp_vsb.set)
    cmp_tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    cmp_vsb.pack(side=tk.RIGHT, fill=tk.Y, in_=tab_cmp)

    comparison_rows = diag_data.get("comparison_rows", [])
    for row_item in comparison_rows:
        label, o_disp, m_disp, judge_text, tag = row_item
        cmp_tree.insert("", tk.END, values=(label, o_disp, m_disp, judge_text), tags=(tag,))

    log_frame = ttk.LabelFrame(tab_cmp, text="詳細診断ログ・ファイル情報", padding=8)
    log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

    txt_log = tk.Text(log_frame, wrap="word", font=("Meiryo UI", 9), height=8, bg="#f8fafc")
    txt_log.pack(fill=tk.BOTH, expand=True)
    detail_text = diag_data.get("detail_text", "詳細情報なし")
    txt_log.insert("1.0", detail_text)
    txt_log.configure(state="disabled")

    # タブ2: 注文書シートプレビュー
    tab_order_sheet = ttk.Frame(notebook, padding=8)
    notebook.add(tab_order_sheet, text="  ② 📄 注文書シート参照 (プレビュー)  ")

    order_full_path = diag_data.get("order_full_path", "")
    order_sheet_name = diag_data.get("order_sheet", "")
    order_row_idx = diag_data.get("order_row")

    if order_full_path and os.path.isfile(order_full_path):
        _build_sheet_preview_tree(tab_order_sheet, order_full_path, order_sheet_name, target_row=order_row_idx)
    else:
        ttk.Label(
            tab_order_sheet,
            text="注文書ファイルが見つかりません（管理表のみに記載されている未発注案件、またはファイルが移動された可能性があります）。",
            foreground="#64748b",
            font=("Meiryo UI", 10),
        ).pack(expand=True)

    # タブ3: 注文品管理表シートプレビュー
    tab_mgmt_sheet = ttk.Frame(notebook, padding=8)
    notebook.add(tab_mgmt_sheet, text="  ③ 📋 注文品管理表参照 (同一管理No・周辺行)  ")

    mgmt_sheet_name = diag_data.get("mgmt_sheet", "")
    mgmt_row_idx = diag_data.get("mgmt_row")
    _build_mgmt_preview_tree(tab_mgmt_sheet, cached_mgmt_rows, order_no, mgmt_sheet_name, mgmt_row_idx)


def _build_sheet_preview_tree(parent, filepath: str, sheet_name: str, target_row: int | None = None):
    """Excelシートを読み込んでアプリ内プレビュー表を生成"""
    lbl_info = ttk.Label(
        parent,
        text=f"ファイル: {os.path.basename(filepath)}  /  シート: {sheet_name}  （黄色ハイライト = 対象明細行）",
        font=("Meiryo UI", 9, "bold"),
    )
    lbl_info.pack(anchor="w", pady=(0, 6))

    tbl_frame = ttk.Frame(parent)
    tbl_frame.pack(fill=tk.BOTH, expand=True)

    max_preview_cols = 16
    col_ids = ["row_no"] + [f"c{i}" for i in range(1, max_preview_cols + 1)]
    tree = ttk.Treeview(tbl_frame, columns=col_ids, show="headings", selectmode="browse")
    tree.heading("row_no", text="行")
    tree.column("row_no", width=55, anchor="center")

    for i in range(1, max_preview_cols + 1):
        col_letter = chr(64 + i)
        tree.heading(f"c{i}", text=f"{col_letter}列")
        tree.column(f"c{i}", width=110, anchor="w")

    tree.tag_configure("TARGET_ROW", background="#fef08a", foreground="#713f12")
    tree.tag_configure("HEADER_ROW", background="#e2e8f0", foreground="#0f172a")

    vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tbl_frame.grid_rowconfigure(0, weight=1)
    tbl_frame.grid_columnconfigure(0, weight=1)

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        for r_idx, row in enumerate(ws.iter_rows(max_row=120, max_col=max_preview_cols, values_only=True), start=1):
            if not any(v is not None for v in row):
                continue
            vals = [str(r_idx)]
            for c_val in row:
                vals.append(str(c_val) if c_val is not None else "")
            while len(vals) < len(col_ids):
                vals.append("")

            tag = ()
            if target_row and r_idx == target_row:
                tag = ("TARGET_ROW",)
            elif r_idx <= 6 and any("品名" in str(v or "") or "図面" in str(v or "") for v in row):
                tag = ("HEADER_ROW",)

            item_id = tree.insert("", tk.END, values=vals, tags=tag)
            if target_row and r_idx == target_row:
                tree.see(item_id)
        wb.close()
    except Exception as e:
        ttk.Label(parent, text=f"シート読込エラー: {e}", foreground="#dc2626").pack(anchor="w")


def _build_mgmt_preview_tree(parent, cached_mgmt_rows: list[dict], order_no: str, target_sheet: str, target_row: int | None):
    """管理表の同一管理No行および周辺行プレビュー表を生成"""
    lbl_info = ttk.Label(
        parent,
        text=f"管理表プレビュー: 管理No [{order_no}] に該当するデータ一覧 （黄色 = 対象行）",
        font=("Meiryo UI", 9, "bold"),
    )
    lbl_info.pack(anchor="w", pady=(0, 6))

    tbl_frame = ttk.Frame(parent)
    tbl_frame.pack(fill=tk.BOTH, expand=True)

    cols = (
        "mark", "sheet", "row", "order_no", "instruction_disp", "ship_to",
        "line_name", "product_detail", "actual_raw", "drawing_no", "part_no",
        "item_name", "maker", "spec", "qty", "due_date", "invoice_raw", "person",
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
        ("qty", "購入数", 60, "center"),
        ("due_date", "納期", 90, "center"),
        ("invoice_raw", "請求日(AI列)", 95, "center"),
        ("person", "担当", 70, "center"),
    ]
    for cid, text, width, anchor in col_defs:
        tree.heading(cid, text=text)
        tree.column(cid, width=width, anchor=anchor)

    tree.tag_configure("TARGET_ROW", background="#fef08a", foreground="#713f12")
    tree.tag_configure("NORMAL_ROW", background="#ffffff", foreground="#334155")

    vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tbl_frame.grid_rowconfigure(0, weight=1)
    tbl_frame.grid_columnconfigure(0, weight=1)

    matching_rows = [
        m for m in cached_mgmt_rows
        if m.get("order_no_8") == order_no or m.get("norm_order_no") == order_no
    ]

    target_item_id = None
    for m in matching_rows:
        is_target = (m.get("sheet") == target_sheet and m.get("row") == target_row)
        mark = "▶ 対象" if is_target else ""
        due_str = m.get("mgmt_due_raw") or (m["mgmt_due_date"].strftime("%Y-%m-%d") if m.get("mgmt_due_date") else "")
        vals = (
            mark,
            m.get("sheet", ""),
            str(m.get("row", "")),
            m.get("order_no", ""),
            m.get("instruction_disp", ""),
            m.get("ship_to", ""),
            m.get("line_name", ""),
            m.get("product_detail", ""),
            m.get("actual_raw", ""),
            m.get("drawing_no", ""),
            m.get("part_no", ""),
            m.get("item_name", ""),
            m.get("maker", ""),
            m.get("spec", ""),
            m.get("qty", ""),
            due_str,
            m.get("invoice_raw", ""),
            m.get("person", ""),
        )
        tag = ("TARGET_ROW",) if is_target else ("NORMAL_ROW",)
        iid = tree.insert("", tk.END, values=vals, tags=tag)
        if is_target:
            target_item_id = iid

    if target_item_id:
        tree.see(target_item_id)
        tree.selection_set(target_item_id)


def show_help_guide(parent: tk.Tk | tk.Toplevel, initial_tab: int = 0, icon_path: str | None = None):
    """取扱説明書・操作ガイドを表示するダイアログ"""
    win = tk.Toplevel(parent)
    win.title(f"{APP_NAME} - ヘルプ＆操作ガイド")
    apply_window_icon(win, icon_path)
    win.geometry("960x740")
    win.minsize(820, 600)

    header = tk.Frame(win, bg="#1e293b", padx=20, pady=12)
    header.pack(fill=tk.X)

    tk.Label(
        header,
        text=f"📖 {APP_NAME}  取扱説明書・操作ガイド",
        bg="#1e293b",
        fg="#ffffff",
        font=("Meiryo UI", 12, "bold"),
    ).pack(anchor="w")

    tk.Label(
        header,
        text="注文書と注文品管理表の双方向自動照合・タブ分離一覧表示・納期超過チェックツールの使い方",
        bg="#1e293b",
        fg="#94a3b8",
        font=("Meiryo UI", 9),
    ).pack(anchor="w", pady=(2, 0))

    notebook = ttk.Notebook(win, padding=8)
    notebook.pack(fill=tk.BOTH, expand=True)

    def _create_scrollable_text(p):
        txt = tk.Text(p, wrap="word", font=("Meiryo UI", 9), padx=12, pady=10, bg="#ffffff", fg="#0f172a")
        vsb = ttk.Scrollbar(p, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        txt.tag_configure("h1", font=("Meiryo UI", 11, "bold"), foreground="#1e293b", spacing1=12, spacing3=4)
        txt.tag_configure("h2", font=("Meiryo UI", 10, "bold"), foreground="#2563eb", spacing1=8, spacing3=2)
        txt.tag_configure("body", font=("Meiryo UI", 9), foreground="#334155", spacing1=2, spacing3=2)
        txt.tag_configure("point", font=("Meiryo UI", 9, "bold"), foreground="#0f172a", background="#fef08a")
        txt.tag_configure("warn", font=("Meiryo UI", 9, "bold"), foreground="#dc2626")
        txt.tag_configure("sub_note", font=("Meiryo UI", 8), foreground="#64748b")
        return txt

    # タブ1: 基本的な使い方・設定
    t1 = ttk.Frame(notebook, padding=8)
    notebook.add(t1, text="  ① 基本的な使い方・設定  ")
    txt1 = _create_scrollable_text(t1)

    txt1.insert(tk.END, "  STEP 1 : 注文品管理表ファイル と 注文書フォルダ の指定  \n", "h1")
    txt1.insert(tk.END, "■ ① 注文品管理表ファイルの指定\n", "h2")
    txt1.insert(tk.END, "・画面上部の『① 注文品管理表ファイル』行にある『ファイルを参照...』ボタン（または Ctrl+O）を押し、管理表Excelを選択します。\n", "body")
    txt1.insert(tk.END, "・右端の『📋 管理表を開く』ボタンを押すと、現在指定されている管理表ファイルを直接Excelで開いて確認できます。\n", "body")
    txt1.insert(tk.END, "★【自動抽出＆列構造の自動認識】管理表内の全シートから「管理No（8桁）」を自動抽出し、A列（現品確認日）、AB列（購入数）、AH列（納期）、AI列（請求日）などを自動認識します。\n", "point")

    txt1.insert(tk.END, "\n■ ② 注文書フォルダパスの指定（高速枝刈りスキャン）\n", "h2")
    txt1.insert(tk.END, "・『② 注文書フォルダパス』行にある『フォルダを参照...』ボタン（または Ctrl+D）を押し、注文書が保管されている親フォルダを選択します。\n", "body")
    txt1.insert(tk.END, "★【枝刈りスキャン】親フォルダ配下の多数のサブフォルダのうち、管理表にない管理Noフォルダはフォルダごと丸ごとスキップするため、何千フォルダある環境でも数秒で探索が完了します。\n", "point")

    txt1.insert(tk.END, "\n  STEP 2 : 判定基準日・日数・月数除外・設定の保存  \n", "h1")
    txt1.insert(tk.END, "・『③ 判定基準日 (今日)』には判定基準日を入力します（『今日に戻す』ボタンあり）。\n", "body")
    txt1.insert(tk.END, "・『納期未記載時: 注文日から [ 7 ] 日経過で超過』により、納期空欄明細も自動でみなし納期を算出します。\n", "body")
    txt1.insert(tk.END, "・『経過除外: 注文日から [ 12 ] ヶ月経過で除外』により、古い過去データを自動的に除外します（0を入力すると全期間表示）。\n", "body")
    txt1.insert(tk.END, "・『💾 設定を保存』ボタンを押すと、パス・超過日数・除外月数・指示書フィルタ・除外管理No・表示列設定が config.json に保存されます。\n", "point")

    txt1.insert(tk.END, "\n  STEP 3 : 除外する管理Noの設定 ＆ 指示書列(J列)フィルタ  \n", "h1")
    txt1.insert(tk.END, "・『🚫 除外する管理Noを設定・編集...』から、テスト案件や除外したい管理Noを改行区切りで登録できます（Excel複数セルの直接貼り付けに対応）。\n", "body")
    txt1.insert(tk.END, "・『📋 管理表 指示書列(J列)フィルタ』から、「「有り」のみ（初期値）」「有り または 空白 のみ」「「空白」のみ」「すべて表示」をワンクリックで切り替えられます。\n", "body")

    txt1.insert(tk.END, "\n  STEP 4 : 照合・遅延チェックの実行と中止  \n", "h1")
    txt1.insert(tk.END, "・青い『▶ 照合・遅延チェック実行 (F5)』ボタンで照合を開始します。実行中は赤い『⏹ 中止 (Esc)』ボタンまたは Esc キーでいつでも安全に中断できます。\n", "point")
    txt1.configure(state="disabled")

    # タブ2: 照合ルール・タブ分離の仕組み
    t2 = ttk.Frame(notebook, padding=8)
    notebook.add(t2, text="  ② 照合ルール・タブ分離の仕組み  ")
    txt2 = _create_scrollable_text(t2)

    txt2.insert(tk.END, "  1. 2系統タブによる明確な分離表示  \n", "h1")
    txt2.insert(tk.END, "・注文書が存在する案件はすべて『タブ①（納期超過・照合結果一覧）』に集約されます。\n", "body")
    txt2.insert(tk.END, "・管理表にのみ記載があって注文書が見当たらない未発注案件は『タブ②（未発注・注文書なし一覧）』へ完全に分離されます。\n", "body")
    txt2.insert(tk.END, "★【管理表AB列（購入数0）の自動除外】管理表AB列の購入数が0の明細は在庫充当品としてタブ②から自動除外されます。\n", "point")

    txt2.insert(tk.END, "\n  2. 注文品管理表の納期（AH列）優先参照と日付型判定ルール  \n", "h1")
    txt2.insert(tk.END, "★【管理表AH列納期の最優先適用】注文品管理表のAH列（34列目）に納期が記入されている場合は、注文書に書かれている納期ではなく、注文品管理表の納期を最優先で参照して超過判定を行います。\n", "point")
    txt2.insert(tk.END, "★【日付型判定による誤判定防止】管理表の納期セルの型が日付型（Date/Datetime）ではない場合（文字列で「未定」「要相談」等のメモが書かれている場合や空欄など）は自動的に無視され、従来通り注文書の納期（または注文日+猶予日数）を参照します。\n", "point")

    txt2.insert(tk.END, "\n  3. 高精度マッチングアルゴリズム  \n", "h1")
    txt2.insert(tk.END, "★【単価・金額の照合除外】誤判定を防ぐため、単価・金額は照合計算から完全に除外しています。\n", "point")
    txt2.insert(tk.END, "・管理No、図面番号、品番、品名、仕様、業者/メーカー、数量で1対1の全体最適マッチングを行います。\n", "body")
    txt2.configure(state="disabled")

    # タブ3: 便利機能・ショートカットキー
    t3 = ttk.Frame(notebook, padding=8)
    notebook.add(t3, text="  ③ 便利機能・ショートカットキー  ")
    txt3 = _create_scrollable_text(t3)

    txt3.insert(tk.END, "  1. 表示列の設定機能（自由なカスタマイズ ＆ 設定保存）  \n", "h1")
    txt3.insert(tk.END, "・一覧上部の『👁 表示列の設定...』ボタン（またはメニュー『設定(S)』）から、表示する列を自由に選択・保存できます。\n", "point")
    txt3.insert(tk.END, "★【デフォルト15列表示】業務で最も確認頻度の高い15列（総合判定、不一致理由、依頼区分、管理No、指示書、出荷先、品名/ライン名、製品詳細、発注先、メーカー、型番/規格、数量、注文書納期、基準日超過、照合元情報）が初期表示されます。\n", "point")

    txt3.insert(tk.END, "\n  2. 詳細プレビュー ＆ 原本誤編集完全ブロック  \n", "h1")
    txt3.insert(tk.END, "・明細行をダブルクリックすると、比較診断レポートやアプリ内シートプレビュー（注文書・管理表）が開きます。\n", "body")
    txt3.insert(tk.END, "・注文書をExcelで開く際は、読み取り専用強制やシート編集ロック等の5重の保護により、誤上書き事故を完全に防止します。\n", "point")

    txt3.insert(tk.END, "\n  3. Excel新規ブックへの直接展開（Ctrl + S）  \n", "h1")
    txt3.insert(tk.END, "・『📊 Excelで直接開く』ボタンを押すと、一時ファイルなしで新規ブックに色分け・フィルター・A4横印刷設定済みの表を直接展開します。\n", "point")
    txt3.configure(state="disabled")

    notebook.select(initial_tab)

    btn_frame = tk.Frame(win, bg="#f1f5f9", padx=16, pady=8)
    btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
    ttk.Button(btn_frame, text="閉じる", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)


def show_about(parent: tk.Tk | tk.Toplevel, icon_path: str | None = None):
    """バージョン情報ダイアログを表示"""
    about_text = (
        f"{APP_NAME}\n"
        "注文書 ⇔ 管理表 双方向照合・現品/請求超過システム\n\n"
        "・Excel直接起動: 一時ファイルなしで未保存新規ブック展開（A4横・1ページ収まり印刷設定済）\n"
        "・原本誤編集ブロック: 注文書参照時の読み取り専用・メモリ上保護ロック\n"
        "・超高速スキャン: 管理No単位の枝刈りスキップ＆インメモリキャッシュ\n"
        "・指示書(J列)フィルタ: 「有り」のみ等のワンクリック絞り込み\n"
        "・経過除外: 注文日から指定月数経過の古いデータを自動除外\n"
        "・表示列設定: 業務に合わせた自由な列の表示・非表示カスタマイズ\n"
        "・納期AH列優先: 管理表の確定納期を最優先参照（日付型自動判定）\n"
    )
    messagebox.showinfo("バージョン情報", about_text, parent=parent)
