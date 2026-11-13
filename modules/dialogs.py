"""
StockMatch_Miya ダイアログ画面モジュール
表示列設定、除外No設定、保存先設定、照合詳細・シートプレビュー、Excel風列フィルター、ヘルプ、バージョン情報を集約
"""

import os
import re
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
try:
    import openpyxl
except ImportError:
    openpyxl = None

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

    win.bind("<MouseWheel>", _on_mousewheel)

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
        win.destroy()
        if on_apply:
            on_apply(selected, save_to_config)

    def _on_cancel(_event=None):
        win.destroy()
        return "break"

    win.protocol("WM_DELETE_WINDOW", _on_cancel)
    win.bind("<Escape>", _on_cancel)

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
    win.bind("<Escape>", lambda _e: (win.destroy(), "break")[1])


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
    win.bind("<Escape>", lambda _e: (win.destroy(), "break")[1])

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
            filetypes=[("JSON 設定ファイル (*.json)", "*.json")],
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
        text="移行先に設定ファイルがまだない場合のみ現在の設定を引き継ぐ（既存ファイルは上書き保護）",
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
    win.resizable(True, True)
    win.minsize(340, 420)

    win_w = 360
    win_h = 480
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    pos_x = max(10, min(x, screen_w - win_w - 20))
    pos_y = max(10, min(y, screen_h - win_h - 60))
    win.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    # --- 最下部フッター（フィルター解除・キャンセル・OK）を先に BOTTOM で配置 ---
    footer = ttk.Frame(win, padding=(10, 8))
    footer.pack(side=tk.BOTTOM, fill=tk.X)

    # --- フッター上の区切り線 ---
    ttk.Separator(win, orient="horizontal").pack(side=tk.BOTTOM, fill=tk.X, pady=(2, 0))

    # --- 一括選択操作バー（すべて選択・すべて解除） ---
    btn_row = ttk.Frame(win, padding=(10, 4))
    btn_row.pack(side=tk.BOTTOM, fill=tk.X)

    # --- 最上部: ソートフレーム ---
    sort_frame = ttk.Frame(win, padding=(10, 8, 10, 4))
    sort_frame.pack(side=tk.TOP, fill=tk.X)

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

    ttk.Separator(win, orient="horizontal").pack(side=tk.TOP, fill=tk.X, pady=4)

    # --- 検索フレーム ---
    search_frame = ttk.Frame(win, padding=(10, 4))
    search_frame.pack(side=tk.TOP, fill=tk.X)
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

    # --- 中央: スクロール可能な値リスト（残りの全領域を占有） ---
    list_container = ttk.Frame(win, padding=(10, 4))
    list_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(list_container, bg="#ffffff", highlightthickness=1, highlightbackground="#cbd5e1")
    scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg="#ffffff")

    scroll_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    c_win = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(c_win, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind("<MouseWheel>", _on_mousewheel)

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
        cb.bind("<MouseWheel>", _on_mousewheel)
        c_frame.bind("<MouseWheel>", _on_mousewheel)
        cb_widgets.append((v, c_frame))

    def _filter_cb_list(*_args):
        query = search_var.get().strip().lower()
        for v, frame in cb_widgets:
            if not query or query in v.lower():
                frame.pack(fill=tk.X, pady=1)
            else:
                frame.pack_forget()

    search_var.trace_add("write", _filter_cb_list)

    def _select_all(state):
        for v, frame in cb_widgets:
            if frame.winfo_ismapped():
                check_vars[v].set(state)

    ttk.Button(btn_row, text="すべて選択", command=lambda: _select_all(True)).pack(side=tk.LEFT, padx=(0, 4))
    ttk.Button(btn_row, text="すべて解除", command=lambda: _select_all(False)).pack(side=tk.LEFT)

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

    win.bind("<Escape>", lambda _e: (win.destroy(), "break")[1])
    win.bind("<Return>", lambda _e: _apply_filter())


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
    win.bind("<Escape>", lambda _e: (win.destroy(), "break")[1])

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

    cmp_tree_frame = ttk.Frame(tab_cmp)
    cmp_tree_frame.pack(fill=tk.BOTH, expand=True)

    cmp_cols = ("field", "order_val", "mgmt_val", "judge")
    cmp_tree = ttk.Treeview(cmp_tree_frame, columns=cmp_cols, show="headings", height=10)
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

    cmp_vsb = ttk.Scrollbar(cmp_tree_frame, orient="vertical", command=cmp_tree.yview)
    cmp_hsb = ttk.Scrollbar(cmp_tree_frame, orient="horizontal", command=cmp_tree.xview)
    cmp_tree.configure(yscrollcommand=cmp_vsb.set, xscrollcommand=cmp_hsb.set)
    cmp_vsb.pack(side=tk.RIGHT, fill=tk.Y)
    cmp_hsb.pack(side=tk.BOTTOM, fill=tk.X)
    cmp_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

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
    tree.column("row_no", width=65, minwidth=55, anchor="center", stretch=False)

    for i in range(1, max_preview_cols + 1):
        col_letter = chr(64 + i)
        tree.heading(f"c{i}", text=f"{col_letter}列")
        tree.column(f"c{i}", width=125, minwidth=80, anchor="w", stretch=False)

    tree.tag_configure("TARGET_ROW", background="#fef08a", foreground="#713f12")
    tree.tag_configure("HEADER_ROW", background="#e2e8f0", foreground="#0f172a")

    vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.bind("<Shift-MouseWheel>", lambda e: (tree.xview_scroll(-4 if e.delta > 0 else 4, "units"), "break")[1])

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tbl_frame.grid_rowconfigure(0, weight=1)
    tbl_frame.grid_columnconfigure(0, weight=1)

    wb = None
    tmp_dir = None
    try:
        import shutil
        from .matcher import _load_workbook_safely

        wb, tmp_dir = _load_workbook_safely(filepath, data_only=True, read_only=True)
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
    except Exception as e:
        ttk.Label(parent, text=f"シート読込エラー: {e}", foreground="#dc2626").pack(anchor="w")
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if tmp_dir:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass


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
        tree.column(cid, width=width, minwidth=65, anchor=anchor, stretch=False)

    tree.tag_configure("TARGET_ROW", background="#fef08a", foreground="#713f12")
    tree.tag_configure("NORMAL_ROW", background="#ffffff", foreground="#334155")

    vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.bind("<Shift-MouseWheel>", lambda e: (tree.xview_scroll(-4 if e.delta > 0 else 4, "units"), "break")[1])

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
    """取扱説明書・操作ガイドをリッチかつ分かりやすく表示するダイアログ"""
    win = tk.Toplevel(parent)
    win.title(f"{APP_NAME} - ヘルプ＆操作ガイド（取扱説明書）")
    apply_window_icon(win, icon_path)
    win.geometry("1060x780")
    win.minsize(900, 640)

    header = tk.Frame(win, bg="#1e293b", padx=20, pady=12)
    header.pack(fill=tk.X)

    tk.Label(
        header,
        text=f"📖 {APP_NAME}  取扱説明書・総合操作ガイド",
        bg="#1e293b",
        fg="#ffffff",
        font=("Meiryo UI", 13, "bold"),
    ).pack(anchor="w")

    tk.Label(
        header,
        text="注文品管理表（1ファイル） ⇔ 注文書フォルダ（全階層） 双方向自動照合・納期超過チェックシステム 総合マニュアル",
        bg="#1e293b",
        fg="#94a3b8",
        font=("Meiryo UI", 9),
    ).pack(anchor="w", pady=(2, 0))

    notebook = ttk.Notebook(win, padding=6)
    notebook.pack(fill=tk.BOTH, expand=True)

    def _create_scrollable_text(p):
        txt_frame = ttk.Frame(p)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(
            txt_frame,
            wrap="word",
            font=("Meiryo UI", 9),
            padx=16,
            pady=12,
            bg="#ffffff",
            fg="#0f172a",
            relief="flat",
            highlightthickness=0,
        )
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        txt.tag_configure("h1", font=("Meiryo UI", 11, "bold"), foreground="#1e3a8a", background="#eff6ff", spacing1=12, spacing3=4)
        txt.tag_configure("h2", font=("Meiryo UI", 10, "bold"), foreground="#1e293b", spacing1=10, spacing3=2)
        txt.tag_configure("body", font=("Meiryo UI", 9), foreground="#334155", spacing1=2, spacing3=2)
        txt.tag_configure("body_lead", font=("Meiryo UI", 9, "bold"), foreground="#1e293b", spacing1=4, spacing3=2)
        txt.tag_configure("point", font=("Meiryo UI", 9, "bold"), foreground="#0f172a", background="#fef08a")
        txt.tag_configure("warn", font=("Meiryo UI", 9, "bold"), foreground="#991b1b", background="#fee2e2")
        txt.tag_configure("feature", font=("Meiryo UI", 9, "bold"), foreground="#166534", background="#dcfce7")
        txt.tag_configure("key", font=("Meiryo UI", 9, "bold"), foreground="#0f172a", background="#e2e8f0")
        txt.tag_configure("bullet", font=("Meiryo UI", 9, "bold"), foreground="#2563eb")
        txt.tag_configure("qa_q", font=("Meiryo UI", 10, "bold"), foreground="#1d4ed8", spacing1=10, spacing3=2)
        txt.tag_configure("qa_a", font=("Meiryo UI", 9), foreground="#1e293b", spacing1=2, spacing3=6)
        txt.tag_configure("sub_note", font=("Meiryo UI", 8), foreground="#64748b")
        return txt

    # -------------------------------------------------------------
    # タブ1: 使い方・基本操作手順
    # -------------------------------------------------------------
    t1 = ttk.Frame(notebook, padding=6)
    notebook.add(t1, text="  ① 📖 使い方・基本操作手順  ")
    txt1 = _create_scrollable_text(t1)

    txt1.insert(tk.END, " 【システムの概要と目的】 \n", "h1")
    txt1.insert(tk.END, "本システムは、社内の「注文品管理表（1ファイル）」と「注文書フォルダ（配下のサブフォルダを含む全注文書Excel）」を双方向で高精度に自動照合し、以下の案件を瞬時に抽出・可視化する業務支援ツールです。\n", "body")
    txt1.insert(tk.END, "・納期を過ぎているのに「現品確認日（A列）」や「請求日（AI列）」が未記入となっている納期超過・遅延案件\n", "body")
    txt1.insert(tk.END, "・管理表に記載があるのに注文書Excelが存在しない未発注・手配漏れ案件（別タブに完全分離表示）\n", "body")
    txt1.insert(tk.END, "・注文書と管理表で品名・品番・図番・数量などが食い違っている不一致案件\n", "body")

    txt1.insert(tk.END, "\n 基本操作の 6 ステップ（日常の業務フロー） \n", "h1")

    txt1.insert(tk.END, "■ STEP 1 : 注文品管理表ファイル（Excel）の指定\n", "h2")
    txt1.insert(tk.END, "1. 画面上部『① 注文品管理表ファイル』の右側にある「ファイルを参照...」ボタン（または Ctrl+O）をクリックし、管理表Excel（例: 注文品管理表.xlsx）を選択します。\n", "body")
    txt1.insert(tk.END, "2. 右端の「📋 管理表を開く」ボタンを押すと、現在指定されている管理表をいつでも直接Excelで開いて閲覧できます。\n", "body")
    txt1.insert(tk.END, "★【全シート自動解析 ＆ Excelを開いたままでOK】\n", "feature")
    txt1.insert(tk.END, "・ブック内の全シート（購入品リスト、外注加工リスト等）を自動で網羅探索し、8桁の管理Noを瞬時に抽出します。\n", "body")
    txt1.insert(tk.END, "・Excelで管理表を開いたまま作業していても、アプリが一時的に安全に読み込むため、Excelを閉じる必要はありません。\n", "body")

    txt1.insert(tk.END, "\n■ STEP 2 : 注文書フォルダの指定（超高速枝刈りスキャン）\n", "h2")
    txt1.insert(tk.END, "1. 『② 注文書フォルダパス』の右側にある「フォルダを参照...」ボタン（または Ctrl+D）をクリックし、注文書が保管されている親フォルダを選択します。\n", "body")
    txt1.insert(tk.END, "2. 右端の「📂 フォルダを開く」ボタンを押すと、エクスプローラーで保管フォルダを直接開けます。\n", "body")
    txt1.insert(tk.END, "★【管理Noフォルダ単位の枝刈りスキップ】\n", "feature")
    txt1.insert(tk.END, "・親フォルダの中に1,000個以上のサブフォルダがあっても、管理表に記載のある管理Noフォルダだけを自動判別して探索します。\n", "body")
    txt1.insert(tk.END, "・関係のないフォルダや見積専用ファイルは一切開かずに丸ごとスキップするため、Googleドライブ等のクラウド同期フォルダや社内サーバー上でも数秒で高速に探索が完了します。\n", "body")

    txt1.insert(tk.END, "\n■ STEP 3 : 判定基準日・日数・月数除外の設定と保存\n", "h2")
    txt1.insert(tk.END, "・判定基準日 (今日): 超過を判定する基準日を入力します（初期値は当日。「今日に戻す」ボタンでいつでも当日に復帰）。\n", "body")
    txt1.insert(tk.END, "・納期未記載時 [ 7 ] 日経過で超過: 注文書に納期が書かれていない場合、注文日（注文書4行目）から指定日数経過でみなし超過判定します。\n", "body")
    txt1.insert(tk.END, "・経過除外 [ 12 ] ヶ月経過で除外: 注文日から指定ヶ月数以上が経過した古い完了データを自動的に一覧や集計から除外します（0で全期間表示）。\n", "body")
    txt1.insert(tk.END, "★「💾 設定を保存」ボタンを押すと、ファイルパス・超過日数・除外月数・除外管理No・指示書フィルタ・表示列設定が設定ファイル（config.json）に保存され、次回起動時に自動復元されます。\n", "point")

    txt1.insert(tk.END, "\n■ STEP 4 : 除外管理No・指示書列(J列)フィルタの確認\n", "h2")
    txt1.insert(tk.END, "・『④ 一覧から除外する管理No』の「🚫 除外する管理Noを設定・編集...」から、照合対象外にしたい管理Noを登録・編集できます。\n", "body")
    txt1.insert(tk.END, "・右側の「📋 管理表 指示書列(J列)フィルタ」から、注文品管理表の指示書列（初期値: 「有り」のみ）による絞り込みをワンクリックで選択できます。\n", "body")

    txt1.insert(tk.END, "\n■ STEP 5 : 照合・遅延チェックの実行と中止\n", "h2")
    txt1.insert(tk.END, "・右端の大きなロイヤルブルーのボタン「▶ 照合・遅延チェック実行 (F5)」を押すと照合が始まります。\n", "body")
    txt1.insert(tk.END, "・プログレスバーに進捗状況（読み込み済みファイル数／全件数）がリアルタイム表示されます。\n", "body")
    txt1.insert(tk.END, "★【安心の中止機能】処理を中断したいときは、左隣の赤い「⏹ 中止 (Esc)」ボタンを押すか、キーボードの「Esc」キーを押すだけで、いつでも安全かつ即座に処理を中断できます。\n", "point")

    txt1.insert(tk.END, "\n■ STEP 6 : 件数サマリーカードと2つのタブによる確認\n", "h2")
    txt1.insert(tk.END, "・照合が完了すると、中段の8枚のサマリーカードに件数が集計されます。\n", "body")
    txt1.insert(tk.END, "・見たいカード（例:「【警告】現品・請求の両方超過」など）をクリックするだけで、該当する明細が下部一覧へワンクリックで絞り込まれます。\n", "body")
    txt1.insert(tk.END, "・下部の一覧エリアは2つの独立したタブで構成されています:\n", "body")
    txt1.insert(tk.END, "  【タブ① 納期超過・照合結果 一覧】: 注文書が存在する案件の納期超過や管理表との不一致を確認。\n", "body")
    txt1.insert(tk.END, "  【タブ② 【未発注・注文書なし】管理表のみ記載 一覧】: 管理表に載っているのに注文書が見当たらない未発注案件を確認。\n", "body")
    txt1.configure(state="disabled")

    # -------------------------------------------------------------
    # タブ2: 判定ロジック・色分け・2系統タブ
    # -------------------------------------------------------------
    t2 = ttk.Frame(notebook, padding=6)
    notebook.add(t2, text="  ② 🎯 判定ロジック・色分け・2系統タブ  ")
    txt2 = _create_scrollable_text(t2)

    txt2.insert(tk.END, " 8 枚の件数サマリーカードと警告色の意味 \n", "h1")
    txt2.insert(tk.END, "画面中段のカードをクリックすると、対象の明細へ瞬時に切り替わります（初期選択: 最重要の【警告】現品・請求の両方超過）。\n\n", "body")

    cards_info = [
        ("① 注文書明細 総数", "白", "注文書から検出された全明細の総件数を表示します。"),
        ("② すべての異常(超過・不一致)", "薄赤枠", "納期超過・伝票未着・不一致・管理表未登録など、何らかの対応・確認が必要な全異常件数を一括表示します。"),
        ("③ 【警告】現品・請求の両方超過", "薄赤背景・赤文字", "【最重要確認】納期を過ぎても「現品確認日(A列)」と「請求日(AI列)」の両方が未記入となっている遅延案件を表示します。"),
        ("④ 【現品超過】現品確認日の超過", "薄橙背景・橙文字", "請求日は記入済みですが、現品確認日(A列)が未記入のまま納期を過ぎている案件を表示します。"),
        ("⑤ 【請求超過・未着】請求日(伝票)超過", "薄黄背景・茶文字", "現品は確認済みですが、請求書・伝票（請求日: AI列）が未記入のまま納期を過ぎている案件を表示します。"),
        ("⑥ 【不一致】管理表との不一致", "薄紫背景・紫文字", "注文書と管理表で品名・品番・図番・数量などが食い違っている案件、または注文書はあるが管理表に記載がない案件を表示します。"),
        ("⑦ 正常・納期前・確認済", "白背景・緑枠", "現品・請求ともに確認済みの案件、またはまだ納期前の正常な案件を表示します。"),
        ("⑧ 【別タブ】注文書未検出", "琥珀色背景", "管理表に記載があるのに注文書フォルダ内に対応する注文書Excelが存在しない案件数を表示します（クリックで自動的にタブ②へ移動）。"),
    ]
    for c_title, c_color, c_desc in cards_info:
        txt2.insert(tk.END, f"■ {c_title}  [{c_color}]\n", "h2")
        txt2.insert(tk.END, f"  {c_desc}\n", "body")

    txt2.insert(tk.END, "\n 注文品管理表の納期（AH列）優先参照 ＆ 日付型判定ルール \n", "h1")
    txt2.insert(tk.END, "★【管理表納期の最優先適用】\n", "feature")
    txt2.insert(tk.END, "・注文品管理表のAH列（34列目: 納期列）に納期が記載されている場合、注文書側に書かれている日付ではなく、注文品管理表の納期を最優先で参照して超過判定を行います。\n", "body")
    txt2.insert(tk.END, "★【日付型判定による誤判定の完全防止】\n", "point")
    txt2.insert(tk.END, "・注文品管理表の納期セルが「日付型（Excelの表示形式が日付）」である場合のみ納期として有効判定します。\n", "body")
    txt2.insert(tk.END, "・管理表の納期欄に「未定」「要相談」「都度」「4月中旬」などの文字列メモが書かれている場合や空欄の場合は、誤判定を防ぐため自動的に無視され、注文書に記載された本来の納期を参照します。\n", "body")
    txt2.insert(tk.END, "・文字列として入力された日付形式（例: '2026-07-30 00:00:00' や '2026/07/30' 等）も自動で正規の日付として解析・認識します。\n", "body")

    txt2.insert(tk.END, "\n 納期の柔軟な期間自動変換ルール \n", "h1")
    txt2.insert(tk.END, "注文書や管理表に書かれた柔軟な納期表記を、自動的に正確な判定期日へ変換します:\n", "body")
    txt2.insert(tk.END, "・「上旬・初旬・頭」: その月の 10 日までとして判定（例: 11/上 → 11月10日）\n", "body")
    txt2.insert(tk.END, "・「中旬・半ば・中頃」: その月の 20 日までとして判定（例: 11/中 → 11月20日）\n", "body")
    txt2.insert(tk.END, "・「下旬・月末・いっぱい」: その月の最終日までとして判定（例: 11/下 → 11月30日、2月 → 28日/29日）\n", "body")
    txt2.insert(tk.END, "・「11/10〜15」等の範囲表記: 後ろの期日（15日）を基準に判定\n", "body")
    txt2.insert(tk.END, "・和暦表記（令和・R表記）や年省略表記にも自動対応\n", "body")
    txt2.insert(tk.END, "・納期未記載（空欄）時: 注文書4行目の「注文日」から指定猶予日数（初期値: 7日）を加算した日を期日として自動計算\n", "body")

    txt2.insert(tk.END, "\n 2 系統タブによる完全分離 ＆ AB列「購入数0」自動除外 \n", "h1")
    txt2.insert(tk.END, "・タブ①『納期超過・照合結果 一覧』: 注文書が存在する案件を一覧表示します。\n", "body")
    txt2.insert(tk.END, "・タブ②『【未発注・注文書なし】管理表のみ記載 一覧』: 管理表にあるのに注文書が見当たらない未発注案件のみを完全分離して表示します。\n", "body")
    txt2.insert(tk.END, "★【管理表AB列「購入数0」の自動除外】\n", "feature")
    txt2.insert(tk.END, "・注文品管理表のAB列（28列目）「購入数」が「0」（または「0個」「0台」等）となっている在庫充当品や発注不要項目は、未発注一覧から自動的に除外されます。\n", "body")

    txt2.insert(tk.END, "\n 単価・金額の照合除外と高精度 1 対 1 マッチング \n", "h1")
    txt2.insert(tk.END, "・単価未入力や金額0円、消費税率表記の違いによる誤判定を防ぐため、単価・金額は照合計算から完全に除外しています。\n", "body")
    txt2.insert(tk.END, "・管理No（8桁）、図面番号、品番、品名、仕様、業者/メーカー、購入数による多次元スコアリングを行い、全体最適1対1マッチングを実行します。\n", "body")
    txt2.configure(state="disabled")

    # -------------------------------------------------------------
    # タブ3: 除外設定・フィルタ・設定管理
    # -------------------------------------------------------------
    t3 = ttk.Frame(notebook, padding=6)
    notebook.add(t3, text="  ③ ⚙ 除外設定・フィルタ・設定管理  ")
    txt3 = _create_scrollable_text(t3)

    txt3.insert(tk.END, " 専用画面での除外管理No設定 ＆ Excel直接貼り付け \n", "h1")
    txt3.insert(tk.END, "1. 画面上部『④ 一覧から除外する管理No』の「🚫 除外する管理Noを設定・編集...」ボタンをクリックすると専用ウィンドウが開きます。\n", "body")
    txt3.insert(tk.END, "2. 1行に1件ずつ管理Noが並ぶ見やすいエディタ形式で、除外したい管理Noを管理できます。\n", "body")
    txt3.insert(tk.END, "★【Excelからのコピー＆ペーストに対応】\n", "feature")
    txt3.insert(tk.END, "・Excelの管理No列（縦一列や複数セル）をまとめてコピー（Ctrl+C）し、設定画面で「📋 クリップボードから貼り付け」ボタンを押す（または Ctrl+V）だけで、何十件・何百件でも一括登録できます。\n", "body")
    txt3.insert(tk.END, "・「✨ 重複・空白を整理」ボタンで、重複した管理Noや余分な改行・空白を瞬時にクリーンアップできます。\n", "body")
    txt3.insert(tk.END, "3. 「✔ 適用して設定を保存」ボタンを押すと、再スキャン不要で現在表示中の一覧・サマリーから即座に対象案件が除外され、設定ファイルへ自動保存されます。\n", "point")
    txt3.insert(tk.END, "4. メイン画面の「全解除」ボタンを押すと、ワンクリックで除外指定をすべて解除できます。\n", "body")

    txt3.insert(tk.END, "\n 注文品管理表「指示書列(J列)」フィルタ（設定保存対応） \n", "h1")
    txt3.insert(tk.END, "画面上部『④』の右側にある「📋 管理表 指示書列(J列)フィルタ」から、ワンクリックで表示対象を切り替えられます:\n", "body")
    txt3.insert(tk.END, "1. 「有り」のみ (初期値): J列に「有り（有・〇・済・完等）」と書かれた指示書発行済み案件のみを表示。\n", "body")
    txt3.insert(tk.END, "2. 有り または 空白 のみ: J列が「有り」または「空白」の案件を表示し、「無し（不要・×・ー等）」の対象外案件を除外。\n", "body")
    txt3.insert(tk.END, "3. 「空白」のみ: J列が未記入の案件のみを表示。\n", "body")
    txt3.insert(tk.END, "4. すべて表示 (フィルタなし): J列の値に関係なく全件を表示。\n", "body")
    txt3.insert(tk.END, "※コンボボックスを切り替えると、再照合を行わなくても一覧・サマリーカード・Excel出力へ即座に反映されます。\n", "point")

    txt3.insert(tk.END, "\n 注文日からの経過月数自動除外（初期値: 12ヶ月） \n", "h1")
    txt3.insert(tk.END, "・『経過除外: 注文日から [ 12 ] ヶ月経過で除外』欄により、完了済みの古い過去データを自動的に集計・一覧から除外します。\n", "body")
    txt3.insert(tk.END, "・数値を変更して Enter を押すだけで即座に反映されます（「0」を入力すると経過月数除外を行わず全期間を表示）。\n", "body")

    txt3.insert(tk.END, "\n 設定ファイル (config.json) の保存先変更 ＆ 初期値復帰 \n", "h1")
    txt3.insert(tk.END, "・「⚙ 保存先設定...」ボタン（またはメニュー『設定(S)』）から、設定ファイルの保存先フォルダを自由に変更できます。\n", "body")
    txt3.insert(tk.END, "・社内共有サーバー（NAS）やGoogleドライブ等を指定することで、複数端末で同一の除外設定やパスを共有できます。\n", "body")
    txt3.insert(tk.END, "★【初期値へのワンクリック復帰】「🏠 アプリと同じ場所 (初期値) に戻す」ボタンを押せば、いつでも初期状態に戻せます。\n", "feature")
    txt3.insert(tk.END, "★【設定の自動引き継ぎ】保存先変更時に既存の設定内容が自動コピーされるため、入力済みのパスや除外設定が消える心配はありません。\n", "body")
    txt3.configure(state="disabled")

    # -------------------------------------------------------------
    # タブ4: 便利機能・ショートカット・マウス操作
    # -------------------------------------------------------------
    t4 = ttk.Frame(notebook, padding=6)
    notebook.add(t4, text="  ④ 💡 便利機能・ショートカット・マウス操作  ")
    txt4 = _create_scrollable_text(t4)

    txt4.insert(tk.END, " 全 18 列の表示列設定（管理表順準拠 ＆ 設定保存） \n", "h1")
    txt4.insert(tk.END, "・一覧上部『👁 表示列の設定...』ボタンをクリックすると専用設定画面が開きます。\n", "body")
    txt4.insert(tk.END, "・チェックボックスのON/OFFで表示列を自由にカスタマイズでき、「💾 適用して保存」で次回起動時にも復元されます。\n", "body")
    txt4.insert(tk.END, "【全18列の構成】\n", "body_lead")
    txt4.insert(tk.END, " 1.総合判定  2.不一致理由・照合診断  3.納期  4.プロジェクト名  5.管理No  6.出荷先\n", "body")
    txt4.insert(tk.END, " 7.品名、ライン名等  8.製品詳細  9.製作数  10.図面番号  11.品番  12.品名\n", "body")
    txt4.insert(tk.END, " 13.メーカー  14.単数  15.仕様  16.仕入先  17.在庫数  18.注文日\n", "body")

    txt4.insert(tk.END, "\n 列幅の自動調整 ＆ 標準推奨幅へのリセット \n", "h1")
    txt4.insert(tk.END, "・「📐 列幅自動調整」ボタン（一覧下部またはメニュー設定）: 表示中の文字長に合わせて全列の幅を一括拡大フィット（文字数の少ない列も標準幅を下限保証して潰れを防止）。\n", "body")
    txt4.insert(tk.END, "・「↺ 標準列幅に戻す」ボタン: いつでもワンクリックで見やすい標準初期幅にリセット。\n", "body")
    txt4.insert(tk.END, "・列ヘッダー境界（セパレータ）のダブルクリックでも個別自動調整が可能です。\n", "body")

    txt4.insert(tk.END, "\n 快適な横スクロール操作（4つの操作方法） \n", "h1")
    txt4.insert(tk.END, "① 下部スクロールバーの左右移動ボタン: 「⏮ 左端」「◀ 左へ」「右へ ▶」「右端 ⏭」でワンクリック高速移動\n", "body")
    txt4.insert(tk.END, "② Shift + マウスホイール: 表の上で Shift（または Ctrl）を押しながらホイールを回すと左右にスムーズにスクロール\n", "body")
    txt4.insert(tk.END, "③ 横スクロールバー上のホイール回転: 下部横スクロールバーの上にマウスを乗せてホイールを回すだけでも横スクロール\n", "body")
    txt4.insert(tk.END, "④ キーボード操作: 表を選択した状態で「←」「→」キー（Shift + ← / → で高速移動）\n", "body")

    txt4.insert(tk.END, "\n セル・ツールチップ表示（マウスを乗せるだけで全文表示） \n", "h1")
    txt4.insert(tk.END, "・一覧のセルにマウスカーソルを合わせるだけで、見切れている文字を含めた完全な内容が吹き出し（ツールチップ）でポップアップ表示されます。\n", "body")

    txt4.insert(tk.END, "\n 注文書の原本誤保存・誤編集完全ブロック（5重の安全保護） \n", "h1")
    txt4.insert(tk.END, "「📄 注文書を参照 (読取専用)」ボタンまたは行の右クリックから開くと、Excelで該当シート・該当行がハイライトされた状態で開きます。\n", "body")
    txt4.insert(tk.END, "① 強制読取専用オープン\n", "body")
    txt4.insert(tk.END, "② 既存通常オープン検知時の強制読取専用切り替え\n", "body")
    txt4.insert(tk.END, "③ メモリ上の全シート保護ロック（セル書き換え・削除禁止）\n", "body")
    txt4.insert(tk.END, "④ 終了時の「保存しますか？」ポップアップ抑止\n", "body")
    txt4.insert(tk.END, "⑤ 通常モード起動の完全遮断ガード\n", "body")
    txt4.insert(tk.END, "※この5重ガードにより、原本の注文書を誤って編集・上書き保存してしまう事故を100%防止します。\n", "point")

    txt4.insert(tk.END, "\n 一時ファイルなしの「📊 Excelで直接開く」（Ctrl + S） \n", "h1")
    txt4.insert(tk.END, "・一時ファイルを作成せず、未保存のExcel新規ブックに直接結果を展開（シート1: 納期超過・照合結果、シート2: 注文書未検出）。\n", "body")
    txt4.insert(tk.END, "・アプリと同じ色分け、見出し固定、オートフィルター、最適列幅（フィルター▼が欠けない最小幅保証＆広すぎ防止）、A4横向き1ページ収まり印刷設定を自動適用。\n", "body")
    txt4.insert(tk.END, "・展開完了後は即座にアプリ側が待機状態に戻るため、Excelを閉じても安全です。\n", "body")

    txt4.insert(tk.END, "\n キーボードショートカット ＆ マウス操作一覧 \n", "h1")
    shortcuts = [
        ("F1", "取扱説明書・操作ガイドを開く"),
        ("F5", "▶ 照合・遅延チェックを実行する"),
        ("Esc", "実行中の処理（照合 または Excel展開）を安全に中止する"),
        ("Ctrl + O", "① 注文品管理表ファイルを選択する"),
        ("Ctrl + D", "② 注文書フォルダを選択する"),
        ("Ctrl + S", "照合結果を「📊 Excelで直接開く」（新規ブック展開）"),
        ("Alt + F4", "アプリケーションを終了する"),
        ("行ダブルクリック", "照合詳細・アプリ内シートプレビュー画面を開く"),
        ("行右クリック", "コンテキストメニュー（注文書参照／管理表参照／フォルダを開く／列幅調整）"),
        ("列ヘッダー左クリック", "Excel風 列フィルター＆並べ替え（昇順・降順）メニューを開く"),
        ("列ヘッダー境界ダブルクリック", "該当列の幅を表示文字長に合わせて自動フィットする"),
        ("Shift + ホイール", "一覧表を左右にスムーズに横スクロールする"),
        ("← / → キー", "一覧表を左右に横スクロールする（Shiftキー併用で高速移動）"),
        ("セルにマウスホバー", "見切れている文字列の全文を吹き出しで確認する"),
    ]
    for key_name, key_desc in shortcuts:
        txt4.insert(tk.END, f"・{key_name: <22} : {key_desc}\n", "body")
    txt4.configure(state="disabled")

    # -------------------------------------------------------------
    # タブ5: よくある質問 (Q&A)・トラブルシューティング
    # -------------------------------------------------------------
    t5 = ttk.Frame(notebook, padding=6)
    notebook.add(t5, text="  ⑤ ❓ よくある質問 (Q&A)  ")
    txt5 = _create_scrollable_text(t5)

    txt5.insert(tk.END, " よくある質問 (Q&A) と トラブルシューティング \n", "h1")

    qa_list = [
        (
            "Q1. 注文品管理表をExcelで開いたままでも照合を実行できますか？",
            "A. はい、まったく問題ありません。\n"
            "本アプリはExcelファイルの一時的なメモリ読み込みを行うため、Excelで管理表を開いたまま作業していてもエラーにならず、そのまま照合を実行できます。"
        ),
        (
            "Q2. 納期が注文書に書いてある日付と違っている行があります。なぜですか？",
            "A. 注文品管理表のAH列（納期列）に確定納期が入力されているためです。\n"
            "本アプリの仕様として、管理表の納期欄に有効な日付が記入されている場合は、管理表の納期を最優先で適用します。管理表側が空欄または「未定」等の文字列メモの場合は、注文書の納期が正しく参照されます。"
        ),
        (
            "Q3. 注文書に納期が書かれていない（空欄）案件はどう判定されますか？",
            "A. 注文書4行目に記載されている「注文日」から、設定された猶予日数（初期値: 7日）を加算した日付をみなし納期として自動計算します。\n"
            "一覧表の納期欄にも「未記載 (2026-07-10注文+7日)」のように計算根拠が明記されます。"
        ),
        (
            "Q4. Excelへの直接展開を実行した後、Excelを閉じるとどうなりますか？",
            "A. 展開完了と同時にアプリは通常の待機状態に戻っていますので、Excelをいつ閉じてもアプリの動作には何の影響もありません。\n"
            "保存が必要な場合は、Excel側で「名前を付けて保存」を行ってください。"
        ),
        (
            "Q5. 特定のテスト案件や終了済み管理Noを一覧から非表示にしたいです。",
            "A. 画面上部『④』の「🚫 除外する管理Noを設定・編集...」ボタンを開き、除外したい管理Noを改行区切りで入力（またはExcelからコピー＆ペースト）して適用してください。\n"
            "再スキャンを行わなくても、現在表示中の一覧や集計カードから即座に除外されます。"
        ),
        (
            "Q6. 表示される列が多すぎるので、必要な列だけに絞り込みたいです。",
            "A. 一覧上部『👁 表示列の設定...』ボタンをクリックし、表示したい列にのみチェックを入れて「💾 適用して保存」を押してください。\n"
            "次回アプリを起動した際も、保存した列構成が自動的に復元されます。"
        ),
        (
            "Q7. 列の文字が見切れていて読めないときはどうすればいいですか？",
            "A. 以下の方法で快適に確認できます:\n"
            "1. セルにマウスカーソルを合わせると、吹き出し（ツールチップ）で全文が表示されます。\n"
            "2. 下部の「📐 列幅自動調整」ボタンを押すと、表示中の文字長に合わせて列幅が一括で広がります。\n"
            "3. 元の見やすい幅に戻したいときは、隣の「↺ 標準列幅に戻す」ボタンを押してください。"
        ),
        (
            "Q8. 初回起動時にWindowsの青い警告画面（SmartScreen）が出ました。",
            "A. 故障やウイルスではありません。以下の手順でそのまま実行してください:\n"
            "1. 画面中央の文章下にある「詳細情報」という文字をクリックします。\n"
            "2. 右下に現れる「実行」ボタンをクリックすると、アプリが正常に起動します（2回目以降は表示されません）。"
        ),
    ]

    for q_text, a_text in qa_list:
        txt5.insert(tk.END, f"■ {q_text}\n", "qa_q")
        txt5.insert(tk.END, f"  {a_text}\n\n", "qa_a")

    txt5.configure(state="disabled")

    # 指定されたタブを選択
    try:
        notebook.select(min(max(0, initial_tab), 4))
    except Exception:
        notebook.select(0)

    btn_frame = tk.Frame(win, bg="#f1f5f9", padx=16, pady=8)
    btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
    ttk.Button(btn_frame, text="閉じる (Esc)", style="Primary.TButton", command=win.destroy).pack(side=tk.RIGHT)
    win.bind("<Escape>", lambda _e: (win.destroy(), "break")[1])


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


class TreeviewCellTooltip:
    """
    Treeviewの各セルにマウスを合わせた際に、見切れている文字を含めた完全な内容を
    吹き出し（ツールチップ）でポップアップ表示するクラス
    """

    def __init__(self, tree: ttk.Treeview, col_titles: dict[str, str], delay_ms: int = 350):
        self.tree = tree
        self.col_titles = col_titles
        self.delay_ms = delay_ms
        self.tip_window = None
        self.after_id = None
        self.current_cell = None

        self.tree.bind("<Motion>", self._on_motion, add="+")
        self.tree.bind("<Leave>", self._on_leave, add="+")
        self.tree.bind("<ButtonPress>", self._on_button_press, add="+")
        self.tree.bind("<MouseWheel>", self._on_leave, add="+")

    def _on_motion(self, event):
        item_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not item_id or not col_id:
            self._cancel_timer()
            self._hide_tooltip()
            self.current_cell = None
            return

        cell_key = (item_id, col_id)
        if cell_key == self.current_cell:
            return

        self._cancel_timer()
        self._hide_tooltip()
        self.current_cell = cell_key

        x_root = event.x_root
        y_root = event.y_root
        self.after_id = self.tree.after(
            self.delay_ms, lambda: self._show_tooltip(item_id, col_id, x_root, y_root)
        )

    def _on_leave(self, event=None):
        self._cancel_timer()
        self._hide_tooltip()
        self.current_cell = None

    def _on_button_press(self, event=None):
        self._cancel_timer()
        self._hide_tooltip()

    def _cancel_timer(self):
        if self.after_id:
            try:
                self.tree.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None

    def _hide_tooltip(self):
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None

    def _show_tooltip(self, item_id: str, col_id: str, x: int, y: int):
        self.after_id = None
        if not self.tree.exists(item_id):
            return

        try:
            col_index = int(col_id.replace("#", "")) - 1
        except Exception:
            return

        all_columns = list(self.tree.cget("columns"))
        columns = self.tree.cget("displaycolumns")
        if not columns or columns == ("#all",) or columns == "#all":
            columns = all_columns
        elif isinstance(columns, str):
            columns = columns.split()
        else:
            columns = list(columns)

        if col_index < 0 or col_index >= len(columns):
            return

        target_col_key = columns[col_index]
        raw_idx = all_columns.index(target_col_key) if target_col_key in all_columns else col_index
        item_values = self.tree.item(item_id, "values")
        if not item_values or raw_idx >= len(item_values):
            return

        cell_val = str(item_values[raw_idx] or "").strip()
        if not cell_val or cell_val in ("-", "(空欄)", "(空白)", "なし"):
            return

        col_name = self.col_titles.get(target_col_key, target_col_key)

        self._hide_tooltip()
        self.tip_window = tw = tk.Toplevel(self.tree)
        tw.wm_overrideredirect(True)
        try:
            tw.attributes("-topmost", True)
        except Exception:
            pass

        frame = tk.Frame(
            tw,
            bg="#1e293b",
            relief=tk.SOLID,
            bd=1,
            highlightbackground="#475569",
            highlightthickness=1,
        )
        frame.pack(fill=tk.BOTH, expand=True)

        header_lbl = tk.Label(
            frame,
            text=f"【{col_name}】",
            font=("Meiryo UI", 8, "bold"),
            bg="#1e293b",
            fg="#94a3b8",
            anchor="w",
        )
        header_lbl.pack(fill=tk.X, padx=8, pady=(5, 1))

        body_lbl = tk.Label(
            frame,
            text=cell_val,
            font=("Meiryo UI", 9),
            bg="#1e293b",
            fg="#f8fafc",
            justify=tk.LEFT,
            wraplength=520,
            anchor="w",
        )
        body_lbl.pack(fill=tk.BOTH, padx=8, pady=(0, 6))

        tw.update_idletasks()
        req_w = tw.winfo_reqwidth()
        req_h = tw.winfo_reqheight()

        screen_w = tw.winfo_screenwidth()
        screen_h = tw.winfo_screenheight()

        pos_x = x + 16
        pos_y = y + 20
        if pos_x + req_w > screen_w - 10:
            pos_x = screen_w - req_w - 10
        if pos_y + req_h > screen_h - 10:
            pos_y = y - req_h - 10

        tw.wm_geometry(f"+{max(0, pos_x)}+{max(0, pos_y)}")


def autofit_treeview_columns(
    tree: ttk.Treeview,
    col_configs: list,
    col_titles: dict[str, str],
    max_sample_rows: int = 200,
    min_width: int = 85,
    max_width: int = 450,
):
    """
    Treeviewに表示されているテキストの最長文字幅、および各列の推奨幅・ヘッダー文字幅に合わせて
    すべての列の幅を自動的に最適フィット（オートフィット）させる。
    列が小さくなりすぎないよう、各列の初期推奨幅を下限基準として保証し、stretch=Falseを維持する。
    """
    try:
        font_name = "Meiryo UI"
        font_size = 9
        data_font = tkfont.Font(family=font_name, size=font_size)
        heading_font = tkfont.Font(family=font_name, size=font_size, weight="bold")
    except Exception:
        data_font = None
        heading_font = None

    all_items = tree.get_children()
    sample_items = all_items[:max_sample_rows]

    all_cols = list(tree.cget("columns"))
    display_cols = tree.cget("displaycolumns")
    if not display_cols or display_cols == ("#all",) or display_cols == "#all":
        display_cols = all_cols
    elif isinstance(display_cols, str):
        display_cols = display_cols.split()
    else:
        display_cols = list(display_cols)

    col_index_map = {col_id: idx for idx, col_id in enumerate(all_cols)}
    default_width_map = {c[0]: c[2] for c in col_configs}
    anchor_map = {c[0]: c[3] for c in col_configs}

    for col_id in display_cols:
        col_title = col_titles.get(col_id, col_id)
        default_w = default_width_map.get(col_id, min_width)
        anchor = anchor_map.get(col_id, "w")

        if heading_font:
            title_w = heading_font.measure(col_title) + 44
        else:
            title_w = len(col_title) * 14 + 44

        max_w = max(title_w, default_w)
        raw_idx = col_index_map.get(col_id)
        if raw_idx is not None:
            for item in sample_items:
                vals = tree.item(item, "values")
                if vals and raw_idx < len(vals):
                    val_str = str(vals[raw_idx] or "").strip()
                    if val_str:
                        for line in val_str.split("\n"):
                            line_str = line.strip()
                            if data_font:
                                w = data_font.measure(line_str) + 30
                            else:
                                w = len(line_str) * 11 + 30
                            if w > max_w:
                                max_w = w

        optimal_w = max(min_width, default_w, min(max_w, max_width))
        tree.column(col_id, width=optimal_w, minwidth=80, anchor=anchor, stretch=False)
