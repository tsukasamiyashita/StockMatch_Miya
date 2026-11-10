"""
StockMatch_Miya 設定管理モジュール
config.json の保存・読み込み、保存先フォルダの切り替え、セキュリティ制限チェックを集約
"""

import os
import sys
import json
import shutil
import datetime


class ConfigManager:
    """設定ファイルの保存先管理および読み書きを行うマネージャークラス"""

    def __init__(self, app_dir: str):
        self.app_dir = app_dir
        self.default_config_path = os.path.join(app_dir, "config.json")

        appdata_dir = os.environ.get("APPDATA")
        if appdata_dir:
            user_config_dir = os.path.join(appdata_dir, "StockMatch_Miya")
            try:
                os.makedirs(user_config_dir, exist_ok=True)
            except Exception:
                pass
            self.pointer_file_path = os.path.join(user_config_dir, "config_location.txt")
        else:
            self.pointer_file_path = os.path.join(app_dir, "config_location.txt")

        self.current_config_path = self._resolve_initial_config_path()

    def _resolve_initial_config_path(self) -> str:
        """ポインタファイルまたはデフォルトパスから初期設定パスを解決"""
        try:
            if os.path.isfile(self.pointer_file_path):
                with open(self.pointer_file_path, "r", encoding="utf-8") as f:
                    stored_path = f.read().strip()
                if stored_path:
                    stored_norm = os.path.normpath(stored_path)
                    dir_name = os.path.dirname(stored_norm)
                    if dir_name and os.path.isdir(dir_name) and self.is_strictly_allowed_path(stored_norm):
                        return stored_norm
        except Exception:
            pass

        # 旧仕様: app_dir 内のポインタファイルも互換性確認
        legacy_pointer = os.path.join(self.app_dir, "config_location.txt")
        try:
            if os.path.isfile(legacy_pointer):
                with open(legacy_pointer, "r", encoding="utf-8") as f:
                    stored_path = f.read().strip()
                if stored_path:
                    stored_norm = os.path.normpath(stored_path)
                    dir_name = os.path.dirname(stored_norm)
                    if dir_name and os.path.isdir(dir_name) and self.is_strictly_allowed_path(stored_norm):
                        return stored_norm
        except Exception:
            pass

        return os.path.normpath(self.default_config_path)

    def is_strictly_allowed_path(self, target_path: str) -> bool:
        """Windowsのシステム重要領域への不正書き込みを防止するセキュリティ判定"""
        try:
            target_norm = os.path.normcase(os.path.abspath(os.path.normpath(target_path)))
            windir = os.environ.get("WINDIR") or os.environ.get("SYSTEMROOT") or "C:\\Windows"
            system_dirs = [os.path.normcase(os.path.abspath(windir))]
            for env_var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
                p = os.environ.get(env_var)
                if p:
                    system_dirs.append(os.path.normcase(os.path.abspath(p)))

            for sdir in system_dirs:
                if target_norm == sdir or target_norm.startswith(sdir + os.sep):
                    return False

            base_name = os.path.basename(target_norm)
            root_name, _ = os.path.splitext(base_name)
            if root_name.upper() in (
                "CON", "PRN", "AUX", "NUL",
                "COM1", "COM2", "COM3", "COM4",
                "LPT1", "LPT2", "LPT3", "LPT4",
            ):
                return False

            return True
        except Exception:
            return False

    def get_config_file_path(self) -> str:
        """現在の有効な設定ファイルパスを返す"""
        return self.current_config_path

    def set_config_file_path(self, new_dir_or_file: str, copy_existing: bool = True) -> tuple[bool, str]:
        """設定ファイルの保存先パスを変更する"""
        new_dir_or_file = new_dir_or_file.strip().strip('"')
        if not new_dir_or_file:
            return False, "パスが指定されていません。"

        if os.path.isdir(new_dir_or_file):
            target_path = os.path.join(new_dir_or_file, "config.json")
        else:
            if not new_dir_or_file.lower().endswith(".json"):
                target_path = new_dir_or_file + ".json"
            else:
                target_path = new_dir_or_file

        target_norm = os.path.normpath(os.path.abspath(target_path))
        if not self.is_strictly_allowed_path(target_norm):
            return False, f"指定されたパスはセキュリティ上の理由により使用できません。\n対象: {target_norm}"

        old_path = self.get_config_file_path()

        # 既存設定の自動引き継ぎコピー
        if copy_existing and os.path.isfile(old_path) and os.path.normcase(old_path) != os.path.normcase(target_norm):
            try:
                os.makedirs(os.path.dirname(target_norm), exist_ok=True)
                shutil.copy2(old_path, target_norm)
            except Exception as e:
                return False, f"既存設定ファイルのコピーに失敗しました:\n{e}"

        # ポインタファイルへ記録
        try:
            os.makedirs(os.path.dirname(self.pointer_file_path), exist_ok=True)
            with open(self.pointer_file_path, "w", encoding="utf-8") as f:
                f.write(target_norm)
        except Exception as e:
            return False, f"設定先ポインタファイルの保存に失敗しました:\n{e}"

        self.current_config_path = target_norm
        return True, target_norm

    def reset_to_default(self, copy_existing: bool = False) -> tuple[bool, str]:
        """設定保存先を初期値（アプリ配置フォルダ）へリセット"""
        default_norm = os.path.normpath(os.path.abspath(self.default_config_path))
        old_path = self.get_config_file_path()

        if copy_existing and os.path.isfile(old_path) and os.path.normcase(old_path) != os.path.normcase(default_norm):
            try:
                os.makedirs(os.path.dirname(default_norm), exist_ok=True)
                shutil.copy2(old_path, default_norm)
            except Exception:
                pass

        try:
            if os.path.isfile(self.pointer_file_path):
                os.remove(self.pointer_file_path)
        except Exception:
            pass

        self.current_config_path = default_norm
        return True, default_norm

    def save_config(self, data: dict) -> tuple[bool, str]:
        """設定辞書を現在指定されている config.json に書き込む"""
        target_file = self.get_config_file_path()
        if not self.is_strictly_allowed_path(target_file):
            return False, f"保存エラー: 許可されていない保存先への書き込み要求です。\n対象: {target_file}"

        data["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            target_dir = os.path.dirname(target_file)
            if target_dir:
                os.makedirs(target_dir, exist_ok=True)
            temp_file = target_file + ".tmp"
            with open(temp_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if os.path.exists(target_file):
                os.replace(temp_file, target_file)
            else:
                os.rename(temp_file, target_file)
            return True, target_file
        except Exception as e:
            return False, str(e)

    def load_config(self) -> dict:
        """現在指定されている config.json から設定辞書を読み出す"""
        cfg_path = self.get_config_file_path()
        if os.path.exists(cfg_path) and self.is_strictly_allowed_path(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}
