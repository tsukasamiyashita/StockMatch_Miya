# -*- coding: utf-8 -*-
"""
設定ファイル (config.json) の保存先管理および安全な読み書きを担うモジュール。

配布用 .exe 環境（PyInstaller）でも設定の引継ぎや共有フォルダ上の config.json 参照、
デフォルト位置への確実なリセットが可能となるよう設計している。
また、管理表・注文書などの Excel ファイルを誤って設定ファイルとして上書きすることを防ぐ
多重安全ガードを備える。
"""

import json
import os
import sys
from typing import Iterable


DEFAULT_CONFIG_FILENAME = "config.json"
POINTER_FILENAME = "config_location.txt"
APPDATA_SUBDIR = "StockMatch_Miya"

FORBIDDEN_DATA_EXTENSIONS = (
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
    ".xltx",
    ".xltm",
    ".csv",
    ".tsv",
)


def resolve_app_dir() -> str:
    """実行形態に応じたアプリケーション基準ディレクトリを返す。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    # modules/config_manager.py の親ディレクトリ（プロジェクトルート）
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _norm_win_path(path: str) -> str:
    """Windows パス比較用に正規化する。"""
    if not path:
        return ""
    try:
        return os.path.normcase(os.path.abspath(os.path.normpath(path)))
    except Exception:
        return os.path.normcase(path.strip())


class ConfigManager:
    """設定ファイル保存先の解決・切り替え・安全なI/Oを担当するマネージャ。"""

    def __init__(self, app_dir: str | None = None):
        self.app_dir = os.path.abspath(app_dir) if app_dir else resolve_app_dir()
        self.base_dir = self.app_dir

    def get_default_config_path(self) -> str:
        """デフォルトの設定ファイルパス (<app_dir>/config.json) を返す。"""
        return os.path.join(self.app_dir, DEFAULT_CONFIG_FILENAME)

    def get_pointer_file_path(self) -> str:
        """優先ポインタファイル (%APPDATA%/StockMatch_Miya/config_location.txt) のパスを返す。"""
        appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(appdata, APPDATA_SUBDIR, POINTER_FILENAME)

    def get_local_pointer_file_path(self) -> str:
        """アプリ同梱のローカルポインタファイル (<app_dir>/config_location.txt) のパスを返す。"""
        return os.path.join(self.app_dir, POINTER_FILENAME)

    @staticmethod
    def _read_pointer_file(pointer_path: str) -> str:
        """指定ポインタファイルから有効な設定パス文字列を1行読み取る。"""
        if not pointer_path or not os.path.isfile(pointer_path):
            return ""
        try:
            with open(pointer_path, "r", encoding="utf-8-sig") as f:
                for line in f:
                    cand = line.strip().strip('"').strip("'")
                    if cand:
                        return os.path.abspath(cand)
        except Exception:
            pass
        return ""

    @staticmethod
    def is_forbidden_config_path(
        target_path: str,
        extra_forbidden_paths: Iterable[str] | None = None,
    ) -> tuple[bool, str]:
        """設定ファイル先として禁止されたパス（Excelファイル等・管理表と同一パス）かを判定する。"""
        if not target_path or not str(target_path).strip():
            return True, "設定ファイルのパスが空です。"

        cleaned = str(target_path).strip()
        _, ext = os.path.splitext(cleaned)
        if ext.lower() in FORBIDDEN_DATA_EXTENSIONS:
            return (
                True,
                f"データファイル（{ext}）を設定ファイル保存先として指定・上書きすることは禁止されています。",
            )

        norm_target = _norm_win_path(cleaned)
        if extra_forbidden_paths:
            for fp in extra_forbidden_paths:
                if not fp:
                    continue
                if norm_target == _norm_win_path(str(fp)):
                    return (
                        True,
                        "指定されたファイルは注文品管理表などのデータファイルと同一のため、設定ファイルとして上書きできません。",
                    )

        return False, ""

    @staticmethod
    def resolve_target_config_path(raw_path: str) -> tuple[bool, str]:
        """
        ユーザー指定パスを設定ファイル (.json) の絶対パスへ正規化する。
        既存フォルダ・末尾区切り・拡張子なしのパスはフォルダ指定とみなして config.json を補完する。
        """
        cleaned = (raw_path or "").strip().strip('"').strip("'")
        if not cleaned:
            return False, "保存先のパスを入力してください。"

        forbidden, reason = ConfigManager.is_forbidden_config_path(cleaned)
        if forbidden:
            return False, reason

        ends_with_sep = cleaned.endswith(("/", "\\"))
        abs_input = os.path.abspath(cleaned)

        if os.path.isdir(abs_input) or ends_with_sep:
            return True, os.path.join(abs_input, DEFAULT_CONFIG_FILENAME)

        _, ext = os.path.splitext(abs_input)
        if not ext:
            # 拡張子なし & 既存ファイルでない場合のみ新規フォルダ指定とみなし config.json を補完
            if not os.path.isfile(abs_input):
                return True, os.path.join(abs_input, DEFAULT_CONFIG_FILENAME)
            return False, "拡張子のない既存ファイルは設定ファイルとして利用できません。.json ファイルを指定してください。"

        if ext.lower() != ".json":
            return False, f"設定ファイルの拡張子は .json を指定してください（指定: {ext}）。"

        return True, abs_input

    def load_pointer_path(self) -> str:
        """
        現在有効なポインタファイルからカスタム設定ファイルパスを取得する。
        %APPDATA% 側を優先し、未設定・空の場合はアプリ直下 config_location.txt を参照する。
        """
        for ptr_path in (self.get_pointer_file_path(), self.get_local_pointer_file_path()):
            cand = self._read_pointer_file(ptr_path)
            if not cand:
                continue
            ok, resolved = self.resolve_target_config_path(cand)
            if ok:
                return resolved
        return ""

    def get_active_config_path(self) -> str:
        """実際に読み書き対象となる設定ファイルパスを返す。"""
        custom = self.load_pointer_path()
        if custom:
            return custom
        return self.get_default_config_path()

    def get_config_file_path(self) -> str:
        """現在有効な設定ファイルパスの別名（ダイアログ・互換呼び出し用）。"""
        return self.get_active_config_path()

    def _write_pointer_files(self, target_config_path: str) -> None:
        """ポインタファイルへ保存先を記録する（%APPDATA% 必須、ローカル側が存在時は同期）。"""
        ptr_path = self.get_pointer_file_path()
        ptr_dir = os.path.dirname(ptr_path)
        os.makedirs(ptr_dir, exist_ok=True)
        with open(ptr_path, "w", encoding="utf-8") as f:
            f.write(target_config_path)

        local_ptr = self.get_local_pointer_file_path()
        if os.path.isfile(local_ptr):
            try:
                with open(local_ptr, "w", encoding="utf-8") as f:
                    f.write(target_config_path)
            except Exception:
                pass

    def _delete_pointer_files(self) -> None:
        """すべてのポインタファイル（%APPDATA% およびアプリ直下）を削除する。"""
        for ptr_path in (self.get_pointer_file_path(), self.get_local_pointer_file_path()):
            if os.path.isfile(ptr_path):
                try:
                    os.remove(ptr_path)
                except OSError:
                    # 削除できない場合でも空にして無効化を試みる
                    try:
                        with open(ptr_path, "w", encoding="utf-8") as f:
                            f.write("")
                    except OSError:
                        pass

    def set_config_file_path(
        self,
        new_path: str,
        copy_existing: bool = True,
        extra_forbidden_paths: Iterable[str] | None = None,
    ) -> tuple[bool, str]:
        """
        設定ファイルの保存先を変更し、ポインタファイルへ記録する。
        移行先に既存の .json ファイルが存在する場合は共有設定保護のため上書きせずそのまま引き継ぐ。
        """
        ok, target_or_err = self.resolve_target_config_path(new_path)
        if not ok:
            return False, target_or_err

        target_abs = target_or_err
        forbidden, reason = self.is_forbidden_config_path(target_abs, extra_forbidden_paths)
        if forbidden:
            return False, reason

        old_abs = os.path.abspath(self.get_active_config_path())
        try:
            target_dir = os.path.dirname(target_abs)
            if target_dir:
                os.makedirs(target_dir, exist_ok=True)

            # 移行先にまだ設定ファイルが存在しない場合のみ現在設定をコピーする
            if (
                copy_existing
                and not os.path.isfile(target_abs)
                and os.path.isfile(old_abs)
                and _norm_win_path(old_abs) != _norm_win_path(target_abs)
            ):
                data, load_err = self.load_config_from_path(old_abs)
                if load_err is None:
                    saved, save_err = self.save_config_to_path(
                        target_abs,
                        data,
                        extra_forbidden_paths=extra_forbidden_paths,
                    )
                    if not saved:
                        return False, f"移行先への設定コピーに失敗しました:\n{save_err}"

            self._write_pointer_files(target_abs)
            return True, target_abs
        except Exception as e:
            return False, f"保存先の変更中にエラーが発生しました:\n{e}"

    def reset_to_default(
        self,
        copy_existing: bool = False,
        extra_forbidden_paths: Iterable[str] | None = None,
    ) -> tuple[bool, str]:
        """設定ファイルの保存先をデフォルト位置 (<app_dir>/config.json) へ戻す。"""
        default_abs = os.path.abspath(self.get_default_config_path())
        old_abs = os.path.abspath(self.get_active_config_path())

        try:
            if (
                copy_existing
                and not os.path.isfile(default_abs)
                and os.path.isfile(old_abs)
                and _norm_win_path(old_abs) != _norm_win_path(default_abs)
            ):
                data, load_err = self.load_config_from_path(old_abs)
                if load_err is None:
                    saved, save_err = self.save_config_to_path(
                        default_abs,
                        data,
                        extra_forbidden_paths=extra_forbidden_paths,
                    )
                    if not saved:
                        return False, f"デフォルト設定先へのコピーに失敗しました:\n{save_err}"

            self._delete_pointer_files()
            return True, default_abs
        except Exception as e:
            return False, f"デフォルト位置へのリセットに失敗しました:\n{e}"

    def load_config_from_path(self, filepath: str) -> tuple[dict, str | None]:
        """指定パスの設定ファイルを安全に読み込む。戻り値は (設定辞書, エラーメッセージまたはNone)。"""
        if not filepath:
            return {}, "設定ファイルパスが未指定です。"

        forbidden, reason = self.is_forbidden_config_path(filepath)
        if forbidden:
            return {}, reason

        if not os.path.isfile(filepath):
            return {}, None

        try:
            with open(filepath, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {}, "設定ファイルの形式が不正です（JSONオブジェクトではありません）。"
            return data, None
        except Exception as e:
            return {}, f"設定ファイルの読み込みに失敗しました:\n{e}"

    def load_config(self) -> tuple[dict, str, str | None]:
        """現在有効な設定ファイルを読み込む。(data, active_path, error_or_none) を返す。"""
        active_path = self.get_active_config_path()
        data, err = self.load_config_from_path(active_path)
        return data, active_path, err

    def save_config_to_path(
        self,
        filepath: str,
        data: dict,
        extra_forbidden_paths: Iterable[str] | None = None,
    ) -> tuple[bool, str]:
        """指定パスへ設定辞書を安全に保存する（禁止拡張子ガード＋原子的置換）。"""
        ok, target_or_err = self.resolve_target_config_path(filepath)
        if not ok:
            return False, target_or_err

        target_abs = target_or_err
        forbidden, reason = self.is_forbidden_config_path(target_abs, extra_forbidden_paths)
        if forbidden:
            return False, reason

        try:
            target_dir = os.path.dirname(target_abs)
            if target_dir:
                os.makedirs(target_dir, exist_ok=True)

            tmp_path = f"{target_abs}.tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, target_abs)
            return True, target_abs
        except Exception as e:
            return False, f"設定ファイルの保存に失敗しました:\n{e}"

    def save_config(
        self,
        data: dict,
        extra_forbidden_paths: Iterable[str] | None = None,
    ) -> tuple[bool, str]:
        """現在有効な保存先へ設定辞書を保存する。"""
        active_path = self.get_active_config_path()
        return self.save_config_to_path(
            active_path,
            data,
            extra_forbidden_paths=extra_forbidden_paths,
        )
