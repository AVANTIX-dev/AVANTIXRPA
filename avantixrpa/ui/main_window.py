
from __future__ import annotations

import threading
from pathlib import Path
from typing import List, Dict, Any, Optional
import json
import shutil
import urllib.request
import urllib.error
from urllib.parse import urlparse
import html as html_lib
import re
import unicodedata


import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# --- Drag & Drop 用 (あれば使う / なければ無効化) ---
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES  # type: ignore
    DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None  # type: ignore
    DND_FILES = None   # type: ignore
    DND_AVAILABLE = False

import yaml  # YAML から name を読む＆書く

from avantixrpa.core.flow_loader import load_flow
from avantixrpa.core.engine import Engine
from avantixrpa.config.paths import FLOWS_DIR, CONFIG_DIR, RESOURCES_FILE
from avantixrpa.actions.builtins import BUILTIN_ACTIONS

# パス定義（config.paths と共有）
TRASH_DIR = FLOWS_DIR / ".trash"

DEFAULT_RESOURCES = {
    "sites": {
        "google": {
            "label": "Google",
            "url": "https://www.google.com",
        },
    },
    "files": {},
}


class StepEditor(tk.Toplevel):
    """
    1ステップ分（action + params）の編集ダイアログ。
    画面では日本語だけ見せて、内部で action_id / params を組み立てる。
    """

    def __init__(
        self,
        master: tk.Tk,
        action_ids: List[str],
        initial_step: Optional[Dict[str, Any]] = None,
        resources: Optional[Dict[str, Any]] = None,
    ) -> None:
        super().__init__(master)
        self.title("ステップ編集")
        self.resizable(False, False)
        self.grab_set()  # モーダルっぽく

        self._result: Optional[Dict[str, Any]] = None

        # 座標フィールド用
        self._current_action_id: str = ""
        self._x_var: Optional[tk.StringVar] = None
        self._y_var: Optional[tk.StringVar] = None

        # リソース情報（サイト / ファイルのキー一覧に使う）
        if resources is None:
            resources = {}
        self.resources: Dict[str, Any] = {
            "sites": resources.get("sites") or {},
            "files": resources.get("files") or {},
        }

        # アクション定義
        self.action_defs: List[Dict[str, Any]] = [
            {
                "id": "print",
                "label": "メッセージを表示する",
                "help": "ログにメッセージを出します（画面の右側に出るログ）。",
                "fields": [
                    {"name": "prefix", "label": "先頭につける文字（任意）", "type": "str", "default": "[AVANTIXRPA]"},
                    {"name": "message", "label": "メッセージ本体", "type": "str", "default": "ここに表示したい文章"},
                ],
            },
            {
                "id": "wait",
                "label": "指定秒数だけ待つ",
                "help": "次のステップに進む前に、指定した秒数だけ待機します。",
                "fields": [
                    {"name": "seconds", "label": "待機秒数（秒）", "type": "float", "default": 1.0},
                ],
            },
            {
                "id": "browser.open",
                "label": "ブラウザでURLを開く",
                "help": "既定のブラウザでURLを開きます。",
                "fields": [
                    {"name": "url", "label": "URL", "type": "str", "default": "https://www.google.com"},
                ],
            },
            {
                "id": "resource.open_site",
                "label": "登録済みサイトを開く",
                "help": "リソース管理タブで登録した「サイト名（キー）」のURLを開きます。",
                "fields": [
                    {"name": "key", "label": "サイト名（キー）", "type": "str", "default": "google"},
                ],
            },
            {
                "id": "resource.open_file",
                "label": "登録済みファイルを開く",
                "help": "リソース管理タブで登録した「ファイル名（キー）」のファイルを開きます。",
                "fields": [
                    {"name": "key", "label": "ファイル名（キー）", "type": "str", "default": "sample_excel"},
                ],
            },
            {
                "id": "run.program",
                "label": "プログラムを起動する",
                "help": "指定したプログラム（EXEなど）を起動します。",
                "fields": [
                    {"name": "program", "label": "プログラム名 or パス", "type": "str", "default": "notepad.exe"},
                    {"name": "args", "label": "引数（必要な場合のみ）", "type": "str", "default": "", "optional": True},
                ],
            },
            {
                "id": "ui.type",
                "label": "文字を入力する（キーボード）",
                "help": "アクティブなウィンドウに文字列をタイプします。",
                "fields": [
                    {"name": "text", "label": "入力する文字列", "type": "str", "default": "これはAVANTIXRPAのテストです。"},
                ],
            },
            {
                "id": "ui.hotkey",
                "label": "キー操作を送る（Enter / Ctrl+Sなど）",
                "help": "Enter や Ctrl+S などのキー操作を送ります。",
                "fields": [
                    {
                        "name": "keys",
                        "label": "キー（カンマ区切り） 例: ctrl,s / enter",
                        "type": "list_str",
                        "default": "enter",
                    },
                ],
            },
            {
                "id": "ui.move",
                "label": "マウスを座標へ移動する",
                "help": "画面上の座標（x, y）へマウスカーソルを移動します。",
                "fields": [
                    {"name": "x", "label": "X座標", "type": "int", "default": 500},
                    {"name": "y", "label": "Y座標", "type": "int", "default": 300},
                    {"name": "duration", "label": "移動時間（秒）", "type": "float", "default": 0.3},
                ],
            },
                        {
                "id": "ui.click",
                "label": "マウスクリックする",
                "help": "マウスクリックをします。座標を空欄にすると現在位置でクリックします。",
                "fields": [
                    {"name": "button", "label": "ボタン（left/right/middle）", "type": "str", "default": "left"},
                    {"name": "clicks", "label": "クリック回数", "type": "int", "default": 1},
                    {"name": "x", "label": "X座標（任意）", "type": "int", "default": None, "optional": True},
                    {"name": "y", "label": "Y座標（任意）", "type": "int", "default": None, "optional": True},
                ],
            },
            {
                "id": "ui.scroll",
                "label": "画面をスクロールする",
                "help": "マウスホイールで画面をスクロールします。プラスで上、マイナスで下にスクロールします。",
                "fields": [
                    {"name": "amount", "label": "スクロール量（+で上 / -で下）", "type": "int", "default": -500},
                    {"name": "x", "label": "X座標（任意）", "type": "int", "default": None, "optional": True},
                    {"name": "y", "label": "Y座標（任意）", "type": "int", "default": None, "optional": True},
                ],
            },
            {
                "id": "file.copy",
                "label": "ファイルをコピーする",
                "help": "ファイルを別の場所にコピーします。",
                "fields": [
                    {"name": "src", "label": "コピー元ファイルパス", "type": "str", "default": "src.txt"},
                    {"name": "dst", "label": "コピー先ファイルパス", "type": "str", "default": "dst.txt"},
                ],
            },
            {
                "id": "file.move",
                "label": "ファイルを移動する",
                "help": "ファイルを別の場所に移動します。",
                "fields": [
                    {"name": "src", "label": "移動元ファイルパス", "type": "str", "default": "old.txt"},
                    {"name": "dst", "label": "移動先ファイルパス", "type": "str", "default": "new.txt"},
                ],
            },
        ]

        self._label_to_def = {d["label"]: d for d in self.action_defs}
        self._id_to_def = {d["id"]: d for d in self.action_defs}

        self.action_label_var = tk.StringVar()
        self.on_error_var = tk.StringVar()
        self.help_text_var = tk.StringVar()

        self.field_vars: Dict[str, tuple[tk.StringVar, Dict[str, Any]]] = {}

        self._create_widgets()

        self.action_label_var.trace_add("write", lambda *args: self._on_action_changed())

        if initial_step:
            action_id = initial_step.get("action", "")
            params = initial_step.get("params") or {}
            action_def = self._id_to_def.get(action_id)
            if action_def:
                self.action_label_var.set(action_def["label"])
            else:
                self.action_label_var.set(self.action_defs[0]["label"])
            if "on_error" in initial_step:
                self.on_error_var.set(str(initial_step["on_error"]))
            self._initial_params = params
        else:
            self.action_label_var.set(self.action_defs[0]["label"])
            self._initial_params = {}

        self._on_action_changed()

    def _create_widgets(self) -> None:
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="やりたいこと").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        action_combo = ttk.Combobox(
            self,
            textvariable=self.action_label_var,
            state="readonly",
            values=[d["label"] for d in self.action_defs],
            width=40,
        )
        action_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        help_label = ttk.Label(
            self,
            textvariable=self.help_text_var,
            foreground="gray",
            wraplength=420,
            justify="left",
        )
        help_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=4, pady=(0, 4))

        ttk.Label(self, text="エラー時の動き").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        on_error_combo = ttk.Combobox(
            self,
            textvariable=self.on_error_var,
            state="readonly",
            values=["", "stop", "continue"],
            width=10,
        )
        on_error_combo.grid(row=2, column=1, sticky="w", padx=4, pady=4)
        on_error_combo.set("")

        params_frame = ttk.LabelFrame(self, text="このステップの設定")
        params_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=4, pady=(4, 4))
        params_frame.columnconfigure(1, weight=1)
        self.params_frame = params_frame

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=4, column=0, columnspan=2, sticky="e", padx=4, pady=4)

        ttk.Button(btn_frame, text="OK", command=self._on_ok).grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="キャンセル", command=self._on_cancel).grid(row=0, column=1, padx=4)

    def _on_action_changed(self) -> None:
        label = self.action_label_var.get().strip()
        action_def = self._label_to_def.get(label)
        if not action_def:
            self.help_text_var.set("")
            return

        self._current_action_id = action_def["id"]
        self.help_text_var.set(action_def.get("help", ""))

        # パラメータ欄リセット
        for child in self.params_frame.winfo_children():
            child.destroy()
        self.field_vars.clear()
        self._x_var = None
        self._y_var = None

        # resources.json から取得済みのサイト/ファイル情報
        sites = self.resources.get("sites") or {}
        files = self.resources.get("files") or {}

        # 各フィールドの入力欄を作る
        for row, field in enumerate(action_def.get("fields", [])):
            fname = field["name"]
            flabel = field.get("label", fname)
            default = field.get("default", "")

            ttk.Label(self.params_frame, text=flabel).grid(
                row=row, column=0, sticky="e", padx=4, pady=2
            )

            # --- resource.open_site: key はサイト一覧から選択 ---
            if self._current_action_id == "resource.open_site" and fname == "key":
                site_keys = sorted(sites.keys())
                var = tk.StringVar()
                if self._initial_params and fname in self._initial_params:
                    var.set(str(self._initial_params[fname]))
                elif default:
                    var.set(str(default))
                elif site_keys:
                    var.set(site_keys[0])

                entry = ttk.Combobox(
                    self.params_frame,
                    textvariable=var,
                    values=site_keys,
                    state="readonly",
                    width=30,
                )
                entry.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
                self.field_vars[fname] = (var, field)
                continue

            # --- resource.open_file: key はファイル一覧から選択 ---
            if self._current_action_id == "resource.open_file" and fname == "key":
                file_keys = sorted(files.keys())
                var = tk.StringVar()
                if self._initial_params and fname in self._initial_params:
                    var.set(str(self._initial_params[fname]))
                elif default:
                    var.set(str(default))
                elif file_keys:
                    var.set(file_keys[0])

                entry = ttk.Combobox(
                    self.params_frame,
                    textvariable=var,
                    values=file_keys,
                    state="readonly",
                    width=30,
                )
                entry.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
                self.field_vars[fname] = (var, field)
                continue

            # --- run.program: program だけ「参照...」ボタン付き & ドラッグ＆ドロップ対応 ---
            if self._current_action_id == "run.program" and fname == "program":
                var = tk.StringVar()
                if self._initial_params and fname in self._initial_params:
                    var.set(str(self._initial_params[fname]))
                elif default is not None:
                    var.set(str(default))
                else:
                    var.set("")

                container = ttk.Frame(self.params_frame)
                container.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
                container.columnconfigure(0, weight=1)

                entry = ttk.Entry(container, textvariable=var)
                entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))

                # --- ここから D&D 対応 ---
                if DND_AVAILABLE:
                    def _on_drop(event, target_var=var) -> None:
                        # Explorer からの D&D は {C:\Program Files\foo.exe} みたいな形式で来る
                        data = event.data
                        if data.startswith("{") and data.endswith("}"):
                            data = data[1:-1]
                        target_var.set(data)

                    try:
                        entry.drop_target_register(DND_FILES)
                        entry.dnd_bind("<<Drop>>", _on_drop)
                    except Exception:
                        # 失敗しても GUI 全体が死なないように握りつぶす
                        pass
                # --- ここまで D&D ---

                # 参照ボタン（ファイルダイアログ）
                def _browse_program(target_var=var) -> None:
                    path = filedialog.askopenfilename(
                        title="起動するプログラムを選択",
                        filetypes=[
                            ("実行ファイル", "*.exe *.bat *.cmd *.lnk"),
                            ("すべてのファイル", "*.*"),
                        ],
                    )
                    if path:
                        target_var.set(path)

                ttk.Button(container, text="参照...", command=_browse_program).grid(
                    row=0, column=1, sticky="w"
                )

                self.field_vars[fname] = (var, field)
                continue

            # --- デフォルト: 単純なテキスト入力 ---
            var = tk.StringVar()
            if self._initial_params and fname in self._initial_params:
                var.set(str(self._initial_params[fname]))
            elif default is not None:
                var.set(str(default))
            else:
                var.set("")
            entry = ttk.Entry(self.params_frame, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
            self.field_vars[fname] = (var, field)

            # ui.move / ui.click 用の座標キャプチャ
            if fname == "x":
                self._x_var = var
            if fname == "y":
                self._y_var = var

            if self._current_action_id in ("ui.move", "ui.click", "ui.scroll") and fname == "x":
                ttk.Button(
                    self.params_frame,
                    text="画面から取得",
                    command=self._capture_xy,
                ).grid(row=row, column=2, padx=4, pady=2)

        # 初期パラメータは使い終わったのでリセット
        self._initial_params = {}

    def _capture_xy(self) -> None:
        if self._x_var is None or self._y_var is None:
            messagebox.showerror("エラー", "X座標 / Y座標フィールドが見つかりません。", parent=self)
            return

        parent = self

        class InlineCapture(tk.Toplevel):
            def __init__(self, owner: StepEditor) -> None:
                super().__init__(owner)
                self.owner = owner
                self.title("画面から座標を取得")
                self.resizable(False, False)

                msg = (
                    "1. 押したい場所にマウスカーソルを動かしてください。\n"
                    "2. このウィンドウをアクティブにして Enter を押すと、\n"
                    "   その位置の座標を X/Y にセットします。"
                )
                ttk.Label(self, text=msg, justify="left").pack(padx=8, pady=(8, 4))
                self.pos_label = ttk.Label(self, text="現在の座標: x=--, y=--")
                self.pos_label.pack(padx=8, pady=(0, 8))

                ttk.Button(self, text="今の座標を反映して閉じる", command=self._finish).pack(
                    padx=8, pady=(0, 8)
                )

                self.bind("<Return>", lambda e: self._finish())
                self.bind("<space>", lambda e: self._finish())

                self._update_position()
                self.grab_set()
                self.focus_set()

            def _update_position(self) -> None:
                try:
                    x = self.winfo_pointerx()
                    y = self.winfo_pointery()
                    self.pos_label.config(text=f"現在の座標: x={x}, y={y}")
                except Exception:
                    pass
                self.after(100, self._update_position)

            def _finish(self) -> None:
                x = self.winfo_pointerx()
                y = self.winfo_pointery()
                if parent._x_var is not None:
                    parent._x_var.set(str(x))
                if parent._y_var is not None:
                    parent._y_var.set(str(y))
                self.destroy()
                parent.deiconify()
                parent.lift()
                parent.focus_force()

        self.withdraw()
        InlineCapture(self)

    def _on_ok(self) -> None:
        label = self.action_label_var.get().strip()
        action_def = self._label_to_def.get(label)
        if not action_def:
            messagebox.showerror("エラー", "アクションの選択が不正です。", parent=self)
            return

        action_id = action_def["id"]
        params: Dict[str, Any] = {}

        for fname, (var, field) in self.field_vars.items():
            raw = var.get().strip()
            ftype = field.get("type", "str")
            optional = field.get("optional", False)

            if raw == "":
                if optional:
                    continue
                messagebox.showwarning("入力不足", f"「{field.get('label', fname)}」を入力してください。", parent=self)
                return

            try:
                if ftype == "int":
                    value: Any = int(raw)
                elif ftype == "float":
                    value = float(raw)
                elif ftype == "list_str":
                    value = [x.strip() for x in raw.split(",") if x.strip()]
                else:
                    value = raw
            except ValueError:
                messagebox.showerror(
                    "形式エラー",
                    f"「{field.get('label', fname)}」の値が不正です。",
                    parent=self,
                )
                return

            params[fname] = value

        step: Dict[str, Any] = {"action": action_id, "params": params}
        oe = self.on_error_var.get().strip()
        if oe:
            step["on_error"] = oe

        self._result = step
        self.destroy()

    def _on_cancel(self) -> None:
        self._result = None
        self.destroy()

    def get_result(self) -> Optional[Dict[str, Any]]:
        return self._result


class CoordinateCapture(tk.Toplevel):
    """
    画面上でマウスを動かして、Enterキーを押した時点の座標を取得する。
    """

    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master)
        self.title("マウス座標キャプチャ")
        self.resizable(False, False)

        msg = (
            "1. 押したい場所にマウスカーソルを動かしてください。\n"
            "2. このウィンドウをアクティブにして Enter を押すと、\n"
            "   その位置の座標をクリップボードにコピーします。"
        )
        ttk.Label(self, text=msg, justify="left").pack(padx=8, pady=(8, 4))

        self.pos_label = ttk.Label(self, text="現在の座標: x=--, y=--")
        self.pos_label.pack(padx=8, pady=(0, 8))

        ttk.Button(self, text="今の座標をコピーして閉じる", command=self._finish).pack(
            padx=8, pady=(0, 8)
        )

        self.bind("<Return>", lambda e: self._finish())
        self.bind("<space>", lambda e: self._finish())

        self._update_position()
        self.grab_set()
        self.focus_set()

    def _update_position(self) -> None:
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            self.pos_label.config(text=f"現在の座標: x={x}, y={y}")
        except Exception:
            pass
        self.after(100, self._update_position)

    def _finish(self) -> None:
        x = self.winfo_pointerx()
        y = self.winfo_pointery()
        try:
            self.clipboard_clear()
            self.clipboard_append(f"{x},{y}")
        except Exception:
            pass
        self.destroy()

# D&D が使える環境なら TkinterDnD.Tk を継承、それ以外は普通の tk.Tk
class MainWindow(TkinterDnD.Tk if DND_AVAILABLE else tk.Tk):

    def __init__(self) -> None:
        super().__init__()
        self.title("AVANTIXRPA Launcher")
        self.geometry("980x560")

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.option_add("*Font", "{Meiryo UI} 9")

        self.engine = Engine(BUILTIN_ACTIONS)
        self._running_thread: Optional[threading.Thread] = None

        self.resources: Dict[str, Any] = self._load_resources()
        self._flow_entries: List[Dict[str, Any]] = []

        self.edit_flow_name_var = tk.StringVar()
        self.edit_on_error_var = tk.StringVar()
        self.edit_steps: List[Dict[str, Any]] = []

        # ★ 追加：今編集中のフロー(YAML)のパス（新規のときは None）
        self.current_edit_flow_path: Optional[Path] = None

        self._create_widgets()
        self._load_flows_list()

    def _load_resources(self) -> Dict[str, Any]:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        if not RESOURCES_FILE.exists():
            with RESOURCES_FILE.open("w", encoding="utf-8") as f:
                json.dump(DEFAULT_RESOURCES, f, ensure_ascii=False, indent=2)
            return DEFAULT_RESOURCES.copy()
        try:
            with RESOURCES_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("Invalid resources.json format")

            # ここで古い形式(string)も吸収しておく
            sites_raw = data.get("sites", {})
            files_raw = data.get("files", {})

            norm_sites: Dict[str, Dict[str, str]] = {}
            if isinstance(sites_raw, dict):
                for key, v in sites_raw.items():
                    if isinstance(v, str):
                        norm_sites[key] = {"label": key, "url": v}
                    elif isinstance(v, dict):
                        url = v.get("url") or ""
                        label = v.get("label") or key
                        norm_sites[key] = {"label": label, "url": url}

            norm_files: Dict[str, Dict[str, str]] = {}
            if isinstance(files_raw, dict):
                for key, v in files_raw.items():
                    if isinstance(v, str):
                        norm_files[key] = {"label": key, "path": v}
                    elif isinstance(v, dict):
                        path = v.get("path") or ""
                        label = v.get("label") or key
                        norm_files[key] = {"label": label, "path": path}

            data["sites"] = norm_sites
            data["files"] = norm_files

            return data
        except Exception as exc:
            messagebox.showerror("リソース読み込みエラー", f"resources.json の読み込みに失敗しました。\n{exc}")
            return DEFAULT_RESOURCES.copy()

    def _generate_resource_key(self, label: str, prefix: str, existing: Dict[str, Any]) -> str:
        """表示名から内部キーを自動生成する.

        - 日本語などは落ちるので、全部 ASCII にできなかった場合は prefix ベースで作る
        - 既存のキーと被る場合は _2, _3... を付けてずらす
        """
        text = unicodedata.normalize("NFKC", label)
        ascii_text = text.encode("ascii", "ignore").decode("ascii").lower()
        ascii_text = re.sub(r"[^a-z0-9]+", "_", ascii_text).strip("_")

        base = ascii_text or prefix  # ぜんぶ消えたら prefix を使う（site, file など）
        key = base
        i = 2
        while key in existing:
            key = f"{base}_{i}"
            i += 1
        return key

    def _fetch_title_from_url(self, url: str) -> str | None:
        """URL から <title> を引っこ抜いて返す。失敗したら None。"""
        if not url:
            return None

        try:
            with urllib.request.urlopen(url, timeout=5) as resp:
                # サーバから charset が来てたらそれ優先
                charset = resp.headers.get_content_charset() or "utf-8"
                data = resp.read()
        except Exception as e:
            print(f"[RPA] タイトル取得失敗: {e}")
            return None

        try:
            text = data.decode(charset, errors="ignore")
        except Exception:
            text = data.decode("utf-8", errors="ignore")

        m = re.search(r"<title[^>]*>(.*?)</title>", text, re.IGNORECASE | re.DOTALL)
        if not m:
            return None

        title = m.group(1)
        # 改行・連続スペースを1個に
        title = re.sub(r"\s+", " ", title).strip()
        title = html_lib.unescape(title)
        return title or None

    def _save_resources(self) -> None:
        try:
            with RESOURCES_FILE.open("w", encoding="utf-8") as f:
                json.dump(self.resources, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            messagebox.showerror("リソース保存エラー", f"resources.json の保存に失敗しました。\n{exc}")

    def _create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # ★ Notebook とタブをインスタンス変数で保持しておく
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.flow_tab = ttk.Frame(self.notebook)
        self.resource_tab = ttk.Frame(self.notebook)
        self.editor_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.flow_tab, text="フロー実行")
        self.notebook.add(self.resource_tab, text="リソース管理")
        self.notebook.add(self.editor_tab, text="フローを作成する（β）")

        self._create_flow_tab(self.flow_tab)
        self._create_resource_tab(self.resource_tab)
        self._create_flow_editor_tab(self.editor_tab)

        status_frame = ttk.Frame(self)
        status_frame.grid(row=1, column=0, sticky="ew")
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(status_frame, text="準備完了")
        self.status_label.grid(row=0, column=0, sticky="w", padx=8)

        bottom = ttk.Frame(self)
        bottom.grid(row=2, column=0, sticky="ew", pady=(2, 4))
        bottom.columnconfigure(0, weight=1)

        coord_btn = ttk.Button(bottom, text="マウス座標キャプチャ", command=self._open_coord_capture)
        coord_btn.grid(row=0, column=0, sticky="w", padx=(8, 0))

        self.run_button = ttk.Button(bottom, text="選択フロー実行", command=self._on_run_clicked)
        self.run_button.grid(row=0, column=1, padx=(8, 0))

        self.reload_button = ttk.Button(bottom, text="フロー再読み込み", command=self._load_flows_list)
        self.reload_button.grid(row=0, column=2, padx=(8, 0))

    def _create_flow_tab(self, tab: ttk.Frame) -> None:
        left_frame = ttk.Frame(tab, padding=(0, 0, 8, 0))
        left_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        left_frame.rowconfigure(1, weight=1)
        left_frame.rowconfigure(2, weight=0)
        left_frame.rowconfigure(3, weight=0)
        left_frame.rowconfigure(4, weight=0)
        left_frame.columnconfigure(0, weight=1)

        lbl_flows = ttk.Label(left_frame, text="フロー一覧（RPA名）")
        lbl_flows.grid(row=0, column=0, sticky="w")

        self.flows_listbox = tk.Listbox(left_frame, height=18)
        self.flows_listbox.grid(row=1, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.flows_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.flows_listbox.config(yscrollcommand=scrollbar.set)

        # ダブルクリックで実行
        self.flows_listbox.bind("<Double-Button-1>", self._on_flow_double_click)

        # ★ 右クリック用コンテキストメニュー
        self.flow_list_menu = tk.Menu(self, tearoff=0)
        self.flow_list_menu.add_command(label="フローを実行", command=self._on_run_clicked)
        self.flow_list_menu.add_command(label="編集（フローエディタで開く）", command=self._on_edit_flow_from_list)
        self.flow_list_menu.add_separator()
        self.flow_list_menu.add_command(label="削除", command=self._on_delete_flow)
        self.flow_list_menu.add_separator()
        self.flow_list_menu.add_command(label="名前変更...", command=self._on_rename_flow)
        self.flow_list_menu.add_command(label="複製して新規フローを作成", command=self._on_duplicate_flow)

        # ★ 右クリックでコンテキストメニューを表示
        self.flows_listbox.bind("<Button-3>", self._on_flows_listbox_right_click)

        delete_btn = ttk.Button(left_frame, text="選択フロー削除", command=self._on_delete_flow)
        delete_btn.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        restore_btn = ttk.Button(left_frame, text="削除したフローを復元...", command=self._open_trash_manager)
        restore_btn.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        # ★ おまけ：ボタンでも編集できるようにしておく
        edit_btn = ttk.Button(left_frame, text="選択フローを編集（エディタ）", command=self._on_edit_flow_from_list)
        edit_btn.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        right_frame = ttk.Frame(tab)
        right_frame.grid(row=0, column=1, sticky="nsew")
        right_frame.rowconfigure(1, weight=1)
        right_frame.columnconfigure(0, weight=1)

        lbl_log = ttk.Label(right_frame, text="実行ログ（セッション内）")
        lbl_log.grid(row=0, column=0, sticky="w")

        self.log_text = tk.Text(right_frame, height=18, state="disabled")
        self.log_text.grid(row=1, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=1, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=log_scroll.set)

    def _create_resource_tab(self, tab: ttk.Frame) -> None:
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(1, weight=1)

        site_frame = ttk.LabelFrame(tab, text="サイト（URL）")
        site_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=8, pady=8)
        site_frame.columnconfigure(1, weight=1)
        site_frame.rowconfigure(3, weight=1)

        ttk.Label(site_frame, text="表示名").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        ttk.Label(site_frame, text="URL").grid(row=1, column=0, sticky="e", padx=4, pady=2)

        # キーは内部用（入力欄は出さない）
        self.site_key_var = tk.StringVar()
        self.site_label_var = tk.StringVar()
        self.site_url_var = tk.StringVar()
        self._site_title_after_id = None  # URL変更時の after() 用

        ttk.Entry(site_frame, textvariable=self.site_label_var).grid(
            row=0, column=1, sticky="ew", padx=4, pady=2
        )

        url_entry = ttk.Entry(site_frame, textvariable=self.site_url_var)
        url_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        # URL が変更されたらタイトル自動取得をスケジュール
        self.site_url_var.trace_add("write", self._on_site_url_changed)

        btn_frame_site = ttk.Frame(site_frame)
        btn_frame_site.grid(row=0, column=2, rowspan=3, sticky="ns", padx=4)

        ttk.Button(btn_frame_site, text="新規", command=self._on_site_new).grid(row=0, column=0, pady=2)
        ttk.Button(btn_frame_site, text="保存", command=self._on_site_save).grid(row=1, column=0, pady=2)
        ttk.Button(btn_frame_site, text="削除", command=self._on_site_delete).grid(row=2, column=0, pady=2)

        self.site_listbox = tk.Listbox(site_frame, height=10)
        self.site_listbox.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=4, pady=(4, 4))

        site_scroll = ttk.Scrollbar(site_frame, orient="vertical", command=self.site_listbox.yview)
        site_scroll.grid(row=3, column=3, sticky="ns")
        self.site_listbox.config(yscrollcommand=site_scroll.set)
        self.site_listbox.bind("<<ListboxSelect>>", self._on_site_selected)

        file_frame = ttk.LabelFrame(tab, text="ファイル（Excel / ショートカットなど）")
        file_frame.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=8, pady=8)
        file_frame.columnconfigure(1, weight=1)
        file_frame.rowconfigure(3, weight=1)

        ttk.Label(file_frame, text="表示名").grid(row=1, column=0, sticky="e", padx=4, pady=2)
        ttk.Label(file_frame, text="ファイルパス").grid(row=2, column=0, sticky="e", padx=4, pady=2)

        # キーは内部管理用。入力欄は出さない。
        self.file_key_var = tk.StringVar()
        self.file_label_var = tk.StringVar()
        self.file_path_var = tk.StringVar()
        self._file_title_after_id = None  # パス変更時 after() 用

        ttk.Entry(file_frame, textvariable=self.file_label_var).grid(
            row=1, column=1, sticky="ew", padx=4, pady=2
        )

        # ファイルパス入力欄（ここに D&D も仕込む）
        file_path_entry = ttk.Entry(file_frame, textvariable=self.file_path_var)
        file_path_entry.grid(row=2, column=1, sticky="ew", padx=4, pady=2)

        # パスが変更されたら、少し待ってから表示名を自動補完
        self.file_path_var.trace_add("write", self._on_file_path_changed)

        # ★ D&D 対応（tkinterdnd2 が使える環境だけ）
        if DND_AVAILABLE:
            def _on_drop_file(event, target_var=self.file_path_var):
                data = event.data
                # {C:\foo bar\baz.txt} みたいな形式の場合は括弧を剥がす
                if data.startswith("{") and data.endswith("}"):
                    data = data[1:-1]
                target_var.set(data)

            try:
                file_path_entry.drop_target_register(DND_FILES)
                file_path_entry.dnd_bind("<<Drop>>", _on_drop_file)
            except Exception:
                # D&D がうまく初期化できなくてもアプリ全体は落とさない
                pass

        # ボタン類（新規 / 保存 / 削除 / 参照）
        btn_frame_file = ttk.Frame(file_frame)
        btn_frame_file.grid(row=0, column=2, rowspan=3, sticky="ns", padx=4)

        ttk.Button(btn_frame_file, text="新規", command=self._on_file_new).grid(row=0, column=0, pady=2)
        ttk.Button(btn_frame_file, text="保存", command=self._on_file_save).grid(row=1, column=0, pady=2)
        ttk.Button(btn_frame_file, text="削除", command=self._on_file_delete).grid(row=2, column=0, pady=2)
        ttk.Button(btn_frame_file, text="参照...", command=self._on_file_browse).grid(row=3, column=0, pady=2)

        self.file_listbox = tk.Listbox(file_frame, height=10)
        self.file_listbox.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=4, pady=(4, 4))

        file_scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.file_listbox.yview)
        file_scroll.grid(row=3, column=3, sticky="ns")
        self.file_listbox.config(yscrollcommand=file_scroll.set)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_selected)

        self._refresh_site_list()
        self._refresh_file_list()

    def _create_flow_editor_tab(self, tab: ttk.Frame) -> None:
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)

        top_frame = ttk.Frame(tab)
        top_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        top_frame.columnconfigure(1, weight=1)

        ttk.Label(top_frame, text="フロー名（RPA名）").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(top_frame, textvariable=self.edit_flow_name_var).grid(
            row=0, column=1, sticky="ew", padx=4, pady=2
        )

        ttk.Label(top_frame, text="エラー時の動き（全体）").grid(row=1, column=0, sticky="e", padx=4, pady=2)
        on_error_combo = ttk.Combobox(
            top_frame,
            textvariable=self.edit_on_error_var,
            state="readonly",
            values=["", "stop", "continue"],
            width=10,
        )
        on_error_combo.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        on_error_combo.set("stop")

        middle_frame = ttk.LabelFrame(tab, text="ステップ一覧")
        middle_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.rowconfigure(0, weight=1)

        self.edit_steps_list = tk.Listbox(middle_frame, height=10)
        self.edit_steps_list.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)

        steps_scroll = ttk.Scrollbar(middle_frame, orient="vertical", command=self.edit_steps_list.yview)
        steps_scroll.grid(row=0, column=1, sticky="ns", pady=4)
        self.edit_steps_list.config(yscrollcommand=steps_scroll.set)

        btn_frame_steps = ttk.Frame(middle_frame)
        btn_frame_steps.grid(row=0, column=2, sticky="ns", padx=4, pady=4)

        ttk.Button(btn_frame_steps, text="ステップ追加", command=self._editor_add_step).grid(row=0, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="選択ステップ編集", command=self._editor_edit_step).grid(row=1, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="選択ステップ削除", command=self._editor_delete_step).grid(row=2, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="上へ", command=lambda: self._editor_move_step(-1)).grid(row=3, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="下へ", command=lambda: self._editor_move_step(1)).grid(row=4, column=0, pady=2)

        bottom_frame = ttk.Frame(tab)
        bottom_frame.grid(row=3, column=0, sticky="ew", padx=8, pady=(4, 8))
        bottom_frame.columnconfigure(0, weight=0)
        bottom_frame.columnconfigure(1, weight=0)
        bottom_frame.columnconfigure(2, weight=0)
        bottom_frame.columnconfigure(3, weight=1)  # 右側を余白で伸ばす

        # ★ 新しいフロー作成
        ttk.Button(bottom_frame, text="新しいフロー", command=self._editor_new_flow).grid(
            row=0, column=0, sticky="w", padx=4
        )

        # ★ 既存フローを読み込む
        ttk.Button(bottom_frame, text="既存フローを読み込む...", command=self._editor_load_flow).grid(
            row=0, column=1, sticky="w", padx=4
        )

        # ★ フローを保存
        ttk.Button(bottom_frame, text="フローを保存", command=self._editor_save_flow).grid(
            row=0, column=2, sticky="w", padx=4
        )

        # ★ 今開いているフローを実行
        ttk.Button(bottom_frame, text="このフローを実行", command=self._editor_run_flow).grid(
            row=0, column=3, sticky="w", padx=4
        )

        ttk.Label(bottom_frame, text="※ flows フォルダに YAML として保存されます").grid(
            row=1, column=0, columnspan=4, sticky="w", padx=4, pady=(2, 0)
        )

    def _load_flows_list(self) -> None:
        self.flows_listbox.delete(0, tk.END)
        self._flow_entries.clear()

        FLOWS_DIR.mkdir(parents=True, exist_ok=True)

        yaml_files = sorted(FLOWS_DIR.glob("*.yaml"))
        for p in yaml_files:
            try:
                with p.open("r", encoding="utf-8") as f:
                    data = yaml.safe_load(f) or {}
                if not isinstance(data, dict):
                    raise ValueError("root is not mapping")
                name = data.get("name") or p.stem
                enabled = data.get("enabled", True)
            except Exception:
                name = p.stem
                enabled = True

            self._flow_entries.append(
                {
                    "name": name,
                    "file": p,
                    "enabled": enabled,
                }
            )
            flow_name = f"{name} ({p.name})" if enabled else f"[無効] {name} ({p.name})"
            self.flows_listbox.insert(tk.END, flow_name)

        self._append_log(f"[INFO] フロー一覧を読み込みました ({len(self._flow_entries)} 件)")
        self.status_label.config(text="フロー一覧を更新しました")

    def _append_log(self, message: str) -> None:
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _on_flow_double_click(self, event) -> None:
        self._on_run_clicked()

    def _on_edit_flow_from_list(self) -> None:
        """フロー一覧で選択中のフローを、フローエディタタブで開く。"""
        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showwarning("フロー未選択", "編集するフローを一覧から選択してください。")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            messagebox.showerror("エラー", "内部のフロー一覧と表示がずれています。")
            return

        entry = self._flow_entries[idx]
        flow_path: Path = entry["file"]

        if not flow_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません:\n{flow_path}")
            return

        # 実際の読み込みロジックに委譲
        self._editor_load_from_path(flow_path)

        # エディタタブに切り替え
        try:
            self.notebook.select(self.editor_tab)
        except Exception:
            pass

    def _on_run_clicked(self) -> None:
        if self._running_thread and self._running_thread.is_alive():
            messagebox.showinfo("実行中", "現在フロー実行中です。完了をお待ちください。")
            return

        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showwarning("フロー未選択", "実行するフローを一覧から選択してください。")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            messagebox.showerror("エラー", "内部データと表示がずれています。")
            return

        entry = self._flow_entries[idx]
        flow_path: Path = entry["file"]
        flow_name: str = entry["name"]

        if not flow_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません: {flow_path}")
            return

        self.status_label.config(text=f"フロー実行中: {flow_name}")
        self._append_log(f"[RUN] {flow_name} ({flow_path.name})")

        self.run_button.config(state="disabled")
        self.reload_button.config(state="disabled")

        t = threading.Thread(
            target=self._run_flow_thread,
            args=(flow_path, flow_name),
            daemon=True,
        )
        self._running_thread = t
        t.start()

    def _on_flows_listbox_right_click(self, event) -> None:
        """フロー一覧の右クリックでコンテキストメニューを出す。"""
        if self.flows_listbox.size() == 0:
            return

        # マウス位置に最も近い行インデックスを取得
        index = self.flows_listbox.nearest(event.y)
        if index < 0:
            return

        # その行を選択状態にする
        self.flows_listbox.selection_clear(0, tk.END)
        self.flows_listbox.selection_set(index)
        self.flows_listbox.activate(index)

        try:
            self.flow_list_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.flow_list_menu.grab_release()

    def _on_edit_flow_from_list(self) -> None:
        """フロー一覧で選択中のフローを、エディタタブで編集する。"""
        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showwarning("フロー未選択", "編集するフローを一覧から選択してください。")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            messagebox.showerror("エラー", "内部データと表示がずれています。")
            return

        entry = self._flow_entries[idx]
        flow_path: Path = entry["file"]

        if not flow_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません: {flow_path}")
            return

        # 実際の読み込みロジックに委譲
        self._editor_load_from_path(flow_path)

        # エディタタブに切り替え
        try:
            self.notebook.select(self.editor_tab)
        except Exception:
            # notebook がまだ無いとかは普通起きないけど、一応握りつぶす
            pass

    def _on_rename_flow(self) -> None:
        """選択中のフローの name とファイル名をまとめて変更する。"""
        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showinfo("フロー未選択", "名前を変更するフローを一覧から選択してください。")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            messagebox.showerror("エラー", "内部データと表示がずれています。")
            return

        entry = self._flow_entries[idx]
        old_name: str = entry["name"]
        old_path: Path = entry["file"]

        if not old_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません:\n{old_path}")
            return

        # 新しいフロー名を聞く
        new_name = simpledialog.askstring(
            "フロー名の変更",
            f"現在のフロー名:\n  {old_name}\n\n新しいフロー名を入力してください。",
            initialvalue=old_name,
            parent=self,
        )
        if new_name is None:
            # キャンセル
            return

        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("名前が空です", "フロー名を入力してください。")
            return

        # フロー名からファイル名を生成
        safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in new_name).strip()
        safe_name = safe_name.replace(" ", "_")
        if not safe_name:
            safe_name = "flow"

        new_path = FLOWS_DIR / f"{safe_name}.yaml"

        # 既に別のファイルがある場合は拒否
        if new_path != old_path and new_path.exists():
            messagebox.showerror(
                "既に存在します",
                f"別のフローが同じファイル名を使用しています:\n{new_path.name}\n\n別の名前を指定してください。",
            )
            return

        # YAML を読み込んで name だけ差し替えつつ、新しいパスに保存
        try:
            with old_path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            if not isinstance(data, dict):
                data = {}

            data["name"] = new_name

            with new_path.open("w", encoding="utf-8") as f:
                yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)

            # パスが変わっているなら元ファイルを削除（実質 rename）
            if new_path != old_path and old_path.exists():
                old_path.unlink()

        except Exception as exc:
            messagebox.showerror("名前変更エラー", f"フロー名の変更に失敗しました。\n{exc}")
            return

        # 一覧を再読み込み
        self._load_flows_list()
        self.status_label.config(text=f"フロー名を変更しました: {new_name}")

    def _on_duplicate_flow(self) -> None:
        """選択中のフローを複製して、新しいフローとして保存＆エディタで開く。"""
        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showinfo("フロー未選択", "複製するフローを一覧から選択してください。")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            messagebox.showerror("エラー", "内部データと表示がずれています。")
            return

        entry = self._flow_entries[idx]
        old_name: str = entry["name"]
        old_path: Path = entry["file"]

        if not old_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません:\n{old_path}")
            return

        # 新しいフロー名の候補（デフォルトは「〇〇（コピー）」）
        default_new_name = f"{old_name}（コピー）" if old_name else "新しいフロー"

        new_name = simpledialog.askstring(
            "フローを複製",
            f"元のフロー名:\n  {old_name}\n\n複製後のフロー名を入力してください。",
            initialvalue=default_new_name,
            parent=self,
        )
        if new_name is None:
            # キャンセル
            return

        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("名前が空です", "フロー名を入力してください。")
            return

        # フロー名からベースとなるファイル名を生成
        base_safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in new_name).strip()
        base_safe_name = base_safe_name.replace(" ", "_")
        if not base_safe_name:
            base_safe_name = "flow"

        # 同名ファイルがすでにある場合は _2, _3… とずらす
        candidate = base_safe_name
        i = 2
        while True:
            candidate_path = FLOWS_DIR / f"{candidate}.yaml"
            if not candidate_path.exists():
                new_path = candidate_path
                break
            candidate = f"{base_safe_name}_{i}"
            i += 1

        # 元の YAML を読み込んで、name だけ差し替えて新パスに保存
        try:
            with old_path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            if not isinstance(data, dict):
                data = {}

            data["name"] = new_name

            with new_path.open("w", encoding="utf-8") as f:
                yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)

        except Exception as exc:
            messagebox.showerror("複製エラー", f"フローの複製に失敗しました。\n{exc}")
            return

        # 一覧を更新
        self._load_flows_list()

        # せっかくなので、複製したフローをエディタで即開く
        try:
            self._editor_load_from_path(new_path)
            self.notebook.select(self.editor_tab)
        except Exception:
            # エディタ側で何か死んでもアプリ全体が落ちないようにする
            pass

        self.status_label.config(text=f"フローを複製しました: {new_name}")

    def _run_flow_thread(self, flow_path: Path, flow_name: str) -> None:
        try:
            flow_def = load_flow(flow_path)
            self.engine.run_flow(flow_def)
        except Exception as exc:
            self._append_log(f"[ERROR] フロー実行中にエラーが発生しました: {exc}")
            self.status_label.config(text=f"フロー実行エラー: {flow_name}")
        else:
            self._append_log(f"[DONE] フロー実行完了: {flow_name}")
            self.status_label.config(text=f"フロー実行完了: {flow_name}")
        finally:
            self.run_button.config(state="normal")
            self.reload_button.config(state="normal")

    def _on_delete_flow(self) -> None:
        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showinfo("フロー未選択", "削除するフローを一覧から選択してください。")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            messagebox.showerror("エラー", "内部データと表示がずれています。")
            return

        entry = self._flow_entries[idx]
        flow_name = entry["name"]
        flow_path: Path = entry["file"]

        if not flow_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません: {flow_path}")
            return

        if not messagebox.askyesno(
            "削除確認",
            f"フロー '{flow_name}' を削除しますか？\n"
            f"ファイルは AVANTIXRPA のゴミ箱 (.trash) に移動されます。",
        ):
            return

        try:
            TRASH_DIR.mkdir(parents=True, exist_ok=True)

            target = TRASH_DIR / flow_path.name
            if target.exists():
                from datetime import datetime

                stem = flow_path.stem
                suffix = flow_path.suffix
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                target = TRASH_DIR / f"{stem}_{ts}{suffix}"

            shutil.move(str(flow_path), str(target))
        except OSError as exc:
            messagebox.showerror("削除失敗", f"フローファイルの移動に失敗しました。\n{exc}")
            return

        self._append_log(f"[DELETE] フロー '{flow_name}' をゴミ箱に移動しました。 ({flow_path.name})")
        self.status_label.config(text=f"フロー '{flow_name}' を削除しました。（ゴミ箱に移動）")
        self._load_flows_list()

    def _open_trash_manager(self) -> None:
        if not TRASH_DIR.exists():
            messagebox.showinfo("ゴミ箱なし", "削除されたフローはまだありません。")
            return
        TrashManager(self, TRASH_DIR, FLOWS_DIR, on_restored=self._load_flows_list)

    def _refresh_site_list(self) -> None:
        self.site_listbox.delete(0, tk.END)
        sites = self.resources.get("sites", {})
        for key, site in sites.items():
            label = site.get("label") or key
            # 画面には表示名だけ出す
            self.site_listbox.insert(tk.END, label)

    def _on_site_selected(self, event) -> None:
        selection = self.site_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        sites = self.resources.get("sites", {})
        if idx >= len(sites):
            return
        key = list(sites.keys())[idx]
        site = sites[key]
        # キーは裏で保持、画面には出さない
        self.site_key_var.set(key)
        self.site_label_var.set(site.get("label", ""))
        self.site_url_var.set(site.get("url", ""))

    def _on_site_new(self) -> None:
        # 新規はキーを空にしておく（保存時に自動生成）
        self.site_key_var.set("")
        self.site_label_var.set("")
        self.site_url_var.set("")

    def _on_site_save(self) -> None:
        label = self.site_label_var.get().strip()
        url = self.site_url_var.get().strip()
        if not label or not url:
            messagebox.showwarning("入力不足", "表示名とURLは必須です。")
            return

        sites = self.resources.setdefault("sites", {})

        key = self.site_key_var.get().strip()
        if not key:
            # 新規登録 → 表示名からキー自動生成
            key = self._generate_resource_key(label, "site", sites)

        sites[key] = {"label": label, "url": url}
        self.site_key_var.set(key)  # 裏で保持
        self._save_resources()
        self._refresh_site_list()
        self.status_label.config(text=f"サイトリソースを保存しました: {label}")

    def _on_site_url_changed(self, *args) -> None:
        """URL欄が変更されたときに呼ばれる（即取得せず、少し待ってから実行）。"""
        # すでにスケジュールがあればキャンセル
        if getattr(self, "_site_title_after_id", None) is not None:
            try:
                self.after_cancel(self._site_title_after_id)
            except Exception:
                pass
            self._site_title_after_id = None

        # 0.8秒後に実行（タイプ中に連打しないように）
        self._site_title_after_id = self.after(800, self._auto_fill_site_title_from_url)

    def _auto_fill_site_title_from_url(self) -> None:
        self._site_title_after_id = None

        url = self.site_url_var.get().strip()
        if not url:
            return

        # すでに表示名が入っているなら何もしない
        if self.site_label_var.get().strip():
            return

        # 入力途中の「h」とかで取りに行かない
        if "://" not in url and "." not in url:
            return

        title = self._fetch_title_from_url(url)

        if title:
            # 正常にタイトル取れたケース
            self.site_label_var.set(title)
            self.status_label.config(text="URL からタイトルを自動取得しました")
            return

        # ★ タイトル取れなかったときの fallback
        guess = self._guess_label_from_url(url)
        if guess:
            self.site_label_var.set(guess)
            self.status_label.config(
                text="ページタイトルは取得できなかったため、URLから簡易な表示名を設定しました"
            )

    def _fetch_title_from_url(self, url: str) -> str | None:
        """URL から <title> を引っこ抜いて返す。失敗したら None。"""
        if not url:
            return None

        url = url.strip()

        # スキームが無い場合は https を補完（chatgpt.com だけ貼ったとき用）
        if "://" not in url:
            url = "https://" + url

        try:
            # ブラウザっぽい User-Agent を名乗る
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/122.0 Safari/537.36"
                    )
                },
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                charset = resp.headers.get_content_charset() or "utf-8"
                data = resp.read()
        except Exception as e:
            print(f"[RPA] タイトル取得失敗: {e}")
            return None

        try:
            text = data.decode(charset, errors="ignore")
        except Exception:
            text = data.decode("utf-8", errors="ignore")

        m = re.search(r"<title[^>]*>(.*?)</title>", text, re.IGNORECASE | re.DOTALL)
        if not m:
            return None

        title = m.group(1)
        title = re.sub(r"\s+", " ", title).strip()
        title = html_lib.unescape(title)
        return title or None
    
    def _guess_label_from_url(self, url: str) -> str:
        """タイトルが取れなかったとき用に、URLからそれっぽい表示名を作る。"""
        if not url:
            return ""

        # scheme 無しなら https を補完
        if "://" not in url:
            url = "https://" + url

        parsed = urlparse(url)
        host = parsed.hostname or ""
        path = (parsed.path or "").strip("/")

        # ホスト部分からベースの名前を作る
        if host:
            parts = host.split(".")
            # outlook.office.com → outlook
            base = parts[0].capitalize()
        else:
            base = url

        if path:
            # /mail/ → Mail
            first = path.split("/")[0]
            base = f"{base} {first.capitalize()}"

        return base
    
    def _guess_label_from_path(self, path: str) -> str:
        """ファイルパスから表示名候補を作る。

        例:
          C:\\foo\\bar\\report_2025-12.xlsx → "report_2025-12"
        """
        if not path:
            return ""

        p = Path(path)
        name = p.name or str(path)

        # 拡張子を落とした名前
        stem = p.stem or name
        return stem
    
    def _on_site_fetch_title(self) -> None:
        url = self.site_url_var.get().strip()
        if not url:
            messagebox.showwarning("URLが未入力です", "先に URL を入力してください。")
            return

        self.status_label.config(text="URL からタイトルを取得しています...")
        self.update_idletasks()

        title = self._fetch_title_from_url(url)
        if not title:
            messagebox.showinfo(
                "取得できませんでした",
                "ページタイトルを取得できませんでした。\nログインが必要なページや、特殊なサイトの可能性があります。",
            )
            self.status_label.config(text="タイトル取得に失敗しました")
            return

        current = self.site_label_var.get().strip()
        if not current:
            # まだ何も入っていないならそのままセット
            self.site_label_var.set(title)
            self.status_label.config(text="ページタイトルを表示名に設定しました")
        else:
            # 既に表示名があるなら上書き確認
            if messagebox.askyesno(
                "表示名を上書きしますか？",
                f"現在の表示名:\n  {current}\n\n取得したタイトル:\n  {title}\n\n上書きしてもよいですか？",
            ):
                self.site_label_var.set(title)
                self.status_label.config(text="ページタイトルで表示名を更新しました")
            else:
                self.status_label.config(text="タイトル取得は行いました（表示名は変更していません）")

    def _on_site_delete(self) -> None:
        key = self.site_key_var.get().strip()
        if not key:
            messagebox.showwarning("選択なし", "削除するサイトを一覧から選択してください。")
            return
        sites = self.resources.get("sites", {})
        if key not in sites:
            messagebox.showwarning("存在しません", "選択されたサイトは登録されていません。")
            return

        label = sites[key].get("label") or key
        if not messagebox.askyesno("確認", f"サイト '{label}' を削除しますか？"):
            return

        del sites[key]
        self._save_resources()
        self._refresh_site_list()
        self._on_site_new()
        self.status_label.config(text=f"サイトリソースを削除しました: {label}")

    def _refresh_file_list(self) -> None:
        self.file_listbox.delete(0, tk.END)
        files = self.resources.get("files", {})
        for key, item in files.items():
            label = item.get("label") or key
            # 画面には表示名だけ
            self.file_listbox.insert(tk.END, label)

    def _on_file_selected(self, event) -> None:
        selection = self.file_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        files = self.resources.get("files", {})
        if idx >= len(files):
            return
        key = list(files.keys())[idx]
        item = files[key]
        self.file_key_var.set(key)
        self.file_label_var.set(item.get("label", ""))
        self.file_path_var.set(item.get("path", ""))

    def _on_file_new(self) -> None:
        self.file_key_var.set("")
        self.file_label_var.set("")
        self.file_path_var.set("")

    def _on_file_save(self) -> None:
        label = self.file_label_var.get().strip()
        path = self.file_path_var.get().strip()
        if not path:
            messagebox.showwarning("入力不足", "表示名とファイルパスは必須です。")
            return

        # 表示名が空なら、パスから推測して補完
        if not label:
            guess = self._guess_label_from_path(path)
            if guess:
                label = guess
                self.file_label_var.set(guess)
                self.status_label.config(text="ファイルパスから表示名を自動設定しました")
            else:
                messagebox.showwarning(
                    "表示名がありません",
                    "表示名が空で、ファイルパスから名前を推測できませんでした。\n手動で表示名を入力してください。",
                )
                return

        files = self.resources.setdefault("files", {})

        key = self.file_key_var.get().strip()
        if not key:
            key = self._generate_resource_key(label, "file", files)

        files[key] = {"label": label, "path": path}
        self.file_key_var.set(key)
        self._save_resources()
        self._refresh_file_list()
        self.status_label.config(text=f"ファイルリソースを保存しました: {label}")

    def _on_file_delete(self) -> None:
        key = self.file_key_var.get().strip()
        if not key:
            messagebox.showwarning("選択なし", "削除するファイルを一覧から選択してください。")
            return
        files = self.resources.get("files", {})
        if key not in files:
            messagebox.showwarning("存在しません", "選択されたファイルは登録されていません。")
            return

        label = files[key].get("label") or key
        if not messagebox.askyesno("確認", f"ファイル '{label}' を削除しますか？"):
            return

        del files[key]
        self._save_resources()
        self._refresh_file_list()
        self._on_file_new()
        self.status_label.config(text=f"ファイルリソースを削除しました: {label}")

    def _on_file_path_changed(self, *args) -> None:
        """ファイルパス欄が変更されたときに呼ばれる（少し待ってから実行）。"""
        if getattr(self, "_file_title_after_id", None) is not None:
            try:
                self.after_cancel(self._file_title_after_id)
            except Exception:
                pass
            self._file_title_after_id = None

        # 0.5秒後に実行（タイプ中に連打しないように）
        self._file_title_after_id = self.after(500, self._auto_fill_file_label_from_path)

    def _auto_fill_file_label_from_path(self) -> None:
        """ファイルパスから表示名を自動セットする（表示名が空のときだけ）。"""
        self._file_title_after_id = None

        path = self.file_path_var.get().strip()
        if not path:
            return

        # すでに表示名が入っていたら何もしない（手入力を優先）
        if self.file_label_var.get().strip():
            return

        guess = self._guess_label_from_path(path)
        if not guess:
            return

        self.file_label_var.set(guess)
        self.status_label.config(text="ファイルパスから表示名を自動設定しました")

    def _on_file_browse(self) -> None:
        path = filedialog.askopenfilename(title="ファイルを選択")
        if path:
            self.file_path_var.set(path)

    def _editor_add_step(self) -> None:
        actions = list(BUILTIN_ACTIONS.keys())
        dialog = StepEditor(self, actions, resources=self.resources)
        self.wait_window(dialog)
        result = dialog.get_result()
        if result is None:
            return
        self.edit_steps.append(result)
        self._refresh_edit_steps_list()

    def _editor_edit_step(self) -> None:
        if not self.edit_steps_list:
            return
        sel = self.edit_steps_list.curselection()
        if not sel:
            messagebox.showinfo("選択なし", "編集するステップを選択してください。")
            return
        idx = sel[0]
        if idx < 0 or idx >= len(self.edit_steps):
            return
        current = self.edit_steps[idx]
        actions = list(BUILTIN_ACTIONS.keys())
        dialog = StepEditor(self, actions, initial_step=current, resources=self.resources)
        self.wait_window(dialog)
        result = dialog.get_result()
        if result is None:
            return
        self.edit_steps[idx] = result
        self._refresh_edit_steps_list()

    def _editor_delete_step(self) -> None:
        sel = self.edit_steps_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx < 0 or idx >= len(self.edit_steps):
            return
        del self.edit_steps[idx]
        self._refresh_edit_steps_list()

    def _editor_move_step(self, direction: int) -> None:
        sel = self.edit_steps_list.curselection()
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.edit_steps):
            return
        self.edit_steps[idx], self.edit_steps[new_idx] = self.edit_steps[new_idx], self.edit_steps[idx]
        self._refresh_edit_steps_list()
        self.edit_steps_list.selection_set(new_idx)

    def _editor_load_from_path(self, path: Path) -> None:
        """指定された YAML フローを読み込み、フローエディタに反映する。"""
        try:
            with path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception as exc:
            messagebox.showerror("読み込み失敗", f"フローの読み込みに失敗しました。\n{exc}")
            return

        if not isinstance(data, dict):
            messagebox.showerror("形式エラー", "フローファイルの形式が不正です。")
            return

        name = data.get("name", "") or ""
        on_error = data.get("on_error", "stop") or "stop"
        steps_raw = data.get("steps") or []

        if not isinstance(steps_raw, list):
            messagebox.showerror("形式エラー", "steps が配列ではありません。このフローは編集できません。")
            return

        # 編集状態にセット
        self.edit_flow_name_var.set(name)
        self.edit_on_error_var.set(on_error)

        self.edit_steps = []
        for step in steps_raw:
            if not isinstance(step, dict):
                continue
            action = step.get("action")
            params = step.get("params") or {}
            on_err = step.get("on_error")

            step_data: Dict[str, Any] = {
                "action": action,
                "params": params,
            }
            if on_err is not None:
                step_data["on_error"] = on_err

            self.edit_steps.append(step_data)

        self._refresh_edit_steps_list()

        # 「今編集中のファイル」として記録
        self.current_edit_flow_path = path

        self.status_label.config(text=f"フローを読み込みました: {path.name}")

    def _editor_load_from_path(self, path: Path) -> None:
        """指定された YAML フローを読み込み、フローエディタに反映する。"""
        try:
            with path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception as exc:
            messagebox.showerror("読み込み失敗", f"フローの読み込みに失敗しました。\n{exc}")
            return

        if not isinstance(data, dict):
            messagebox.showerror("形式エラー", "フローファイルの形式が不正です。")
            return

        name = data.get("name", "") or ""
        on_error = data.get("on_error", "stop") or "stop"
        steps_raw = data.get("steps") or []

        if not isinstance(steps_raw, list):
            messagebox.showerror("形式エラー", "steps が配列ではありません。このフローは編集できません。")
            return

        # 不正な要素を落として、辞書だけにしておく
        steps: List[Dict[str, Any]] = [s for s in steps_raw if isinstance(s, dict)]

        self.edit_flow_name_var.set(name)
        self.edit_on_error_var.set(on_error)
        self.edit_steps = steps
        self._refresh_edit_steps_list()

        # 以後「保存」したときはこのファイルに上書き
        self.current_edit_flow_path = path

        self.status_label.config(text=f"フローを読み込みました: {path.name}")

    def _refresh_edit_steps_list(self) -> None:
        self.edit_steps_list.delete(0, tk.END)
        for i, step in enumerate(self.edit_steps, start=1):
            action = step.get("action", "?")
            params = step.get("params", {})
            on_error = step.get("on_error")
            label = f"{i}. {action} {params}"
            if on_error:
                label += f" (on_error={on_error})"
            self.edit_steps_list.insert(tk.END, label)

    def _editor_new_flow(self) -> None:
        """フローエディタをリセットして、新規作成モードにする。"""
        self.edit_flow_name_var.set("")
        self.edit_on_error_var.set("stop")
        self.edit_steps.clear()
        self.edit_steps_list.delete(0, tk.END)
        self.current_edit_flow_path = None
        self.status_label.config(text="新しいフローの作成を開始しました")

    def _editor_load_flow(self) -> None:
        """flows ディレクトリから既存フロー(YAML)を読み込んでエディタに反映する。"""
        FLOWS_DIR.mkdir(parents=True, exist_ok=True)

        path_str = filedialog.askopenfilename(
            title="編集するフローを選択",
            initialdir=FLOWS_DIR,
            filetypes=[("フローファイル", "*.yaml *.yml"), ("すべてのファイル", "*.*")],
        )
        if not path_str:
            return

        path = Path(path_str)

        try:
            with path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception as exc:
            messagebox.showerror("読み込み失敗", f"フローの読み込みに失敗しました。\n{exc}")
            return

        if not isinstance(data, dict):
            messagebox.showerror("形式エラー", "フローファイルの形式が不正です。")
            return

        name = data.get("name", "") or ""
        on_error = data.get("on_error", "stop") or "stop"
        steps_raw = data.get("steps") or []

        if not isinstance(steps_raw, list):
            messagebox.showerror("形式エラー", "steps が配列ではありません。このフローは編集できません。")
            return

        # 編集状態に反映
        self.edit_flow_name_var.set(name)
        self.edit_on_error_var.set(on_error)

        self.edit_steps = []
        for step in steps_raw:
            if not isinstance(step, dict):
                continue
            action = step.get("action")
            params = step.get("params") or {}
            on_err = step.get("on_error")

            step_data: Dict[str, Any] = {
                "action": action,
                "params": params,
            }
            if on_err is not None:
                step_data["on_error"] = on_err

            self.edit_steps.append(step_data)

        self._refresh_edit_steps_list()

        # 以後の保存はこのファイルに上書き
        self.current_edit_flow_path = path

        self.status_label.config(text=f"フローを読み込みました: {path.name}")

    def _editor_save_flow(self) -> None:
        name = self.edit_flow_name_var.get().strip()
        if not name:
            messagebox.showwarning("フロー名不足", "フロー名（RPA名）を入力してください。")
            return
        if not self.edit_steps:
            messagebox.showwarning("ステップなし", "少なくとも1つ以上のステップを追加してください。")
            return

        on_error = self.edit_on_error_var.get().strip() or "stop"

        data = {
            "name": name,
            "on_error": on_error,
            "steps": self.edit_steps,
        }

        FLOWS_DIR.mkdir(parents=True, exist_ok=True)

        # ★ 新規作成か、既存フローの上書きかを判定
        if self.current_edit_flow_path is not None:
            # 既存フロー編集 → そのファイルに上書き
            path = self.current_edit_flow_path
        else:
            # 新規フロー → フロー名からファイル名を生成
            safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name).strip()
            safe_name = safe_name.replace(" ", "_")
            if not safe_name:
                safe_name = "flow"

            path = FLOWS_DIR / f"{safe_name}.yaml"

            if path.exists():
                if not messagebox.askyesno(
                    "上書き確認",
                    f"{path.name} は既に存在します。上書きしてもよろしいですか？",
                ):
                    return

        try:
            with path.open("w", encoding="utf-8") as f:
                yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)
        except Exception as exc:
            messagebox.showerror("保存エラー", f"フローの保存に失敗しました。\n{exc}")
            return

        # 新規保存だった場合も、以後はこのファイルを「編集中」とみなす
        self.current_edit_flow_path = path

        messagebox.showinfo("保存完了", f"フローを保存しました。\n{path}")
        self.status_label.config(text=f"フローを保存しました: {path.name}")
        self._load_flows_list()

    def _editor_run_flow(self) -> None:
        """フローエディタで開いているフローを保存してから実行する。"""
        # すでに実行中なら弾く（フロー一覧の実行ボタンと同じルール）
        if self._running_thread and self._running_thread.is_alive():
            messagebox.showinfo("実行中", "現在フロー実行中です。完了をお待ちください。")
            return

        # まず保存されているかチェック
        if self.current_edit_flow_path is None:
            # まだ一度も保存していないフロー
            if not messagebox.askyesno(
                "保存されていません",
                "このフローはまだファイルに保存されていません。\n"
                "保存してから実行しますか？",
            ):
                return

            # 保存実行（失敗したりユーザーがキャンセルしたら current_edit_flow_path は None のまま）
            self._editor_save_flow()
            if self.current_edit_flow_path is None:
                # 保存失敗 or キャンセル
                return

        flow_path = self.current_edit_flow_path
        assert flow_path is not None  # 型的なおまじない

        if not flow_path.exists():
            messagebox.showerror("ファイルなし", f"フローファイルが見つかりません: {flow_path}")
            return

        flow_name = self.edit_flow_name_var.get().strip() or flow_path.stem

        # ステータス＆ログ出力
        self.status_label.config(text=f"フロー実行中: {flow_name}")
        self._append_log(f"[RUN] {flow_name} ({flow_path.name})")

        # 実行中はメイン画面側の実行ボタンをロック
        try:
            self.run_button.config(state="disabled")
            self.reload_button.config(state="disabled")
        except Exception:
            # 念のため。エディタからだけ使うケースとかでも落ちないように。
            pass

        # いつもの実行スレッドに丸投げ
        t = threading.Thread(
            target=self._run_flow_thread,
            args=(flow_path, flow_name),
            daemon=True,
        )
        self._running_thread = t
        t.start()

    def _open_coord_capture(self) -> None:
        CoordinateCapture(self)

    def _open_trash_manager(self) -> None:
        if not TRASH_DIR.exists():
            messagebox.showinfo("ゴミ箱なし", "削除されたフローはまだありません。")
            return
        TrashManager(self, TRASH_DIR, FLOWS_DIR, on_restored=self._load_flows_list)


class TrashManager(tk.Toplevel):
    """flows/.trash にある削除済みフローの一覧と復元/完全削除を行うダイアログ。"""

    def __init__(
        self,
        master: tk.Tk,
        trash_dir: Path,
        flows_dir: Path,
        on_restored: Optional[callable] = None,
    ) -> None:
        super().__init__(master)
        self.title("削除したフローの管理")
        self.resizable(False, False)

        self.trash_dir = trash_dir
        self.flows_dir = flows_dir
        self.on_restored = on_restored

        self._files: list[Path] = []

        self._create_widgets()
        self._load_trash_list()

        self.grab_set()
        self.focus_set()

    def _create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ttk.Label(self, text="ゴミ箱にあるフロー（.trash）").grid(
            row=0, column=0, sticky="w", padx=8, pady=(8, 4)
        )

        frame = ttk.Frame(self)
        frame.grid(row=1, column=0, sticky="nsew", padx=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(frame, height=12, width=60)
        self.listbox.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.listbox.config(yscrollcommand=scroll.set)

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=2, column=0, sticky="e", padx=8, pady=(4, 8))
        ttk.Button(btn_frame, text="復元", command=self._restore_selected).grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="完全に削除", command=self._delete_selected).grid(row=0, column=1, padx=4)
        ttk.Button(btn_frame, text="閉じる", command=self.destroy).grid(row=0, column=2, padx=4)

    def _load_trash_list(self) -> None:
        self.listbox.delete(0, tk.END)
        self._files.clear()

        if not self.trash_dir.exists():
            return

        yaml_files = sorted(self.trash_dir.glob("*.yaml"))
        for p in yaml_files:
            display = p.name
            try:
                with p.open("r", encoding="utf-8") as f:
                    data = yaml.safe_load(f) or {}
                if isinstance(data, dict) and data.get("name"):
                    display = f"{data['name']} ({p.name})"
            except Exception:
                pass

            self._files.append(p)
            self.listbox.insert(tk.END, display)

        if not self._files:
            self.listbox.insert(tk.END, "[ゴミ箱は空です]")

    def _get_selected_path(self) -> Optional[Path]:
        if not self._files:
            messagebox.showinfo("空", "削除されたフローはありません。")
            return None
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo("選択なし", "対象のフローを選択してください。")
            return None
        idx = sel[0]
        if idx < 0 or idx >= len(self._files):
            return None
        return self._files[idx]

    def _restore_selected(self) -> None:
        p = self._get_selected_path()
        if not p:
            return

        target = self.flows_dir / p.name
        if target.exists():
            if not messagebox.askyesno(
                "上書き確認",
                f"{target.name} は既に flows に存在します。\n"
                "上書きして復元しますか？",
                parent=self,
            ):
                return

        try:
            self.flows_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(p), str(target))
        except OSError as exc:
            messagebox.showerror("復元失敗", f"フローの復元に失敗しました。\n{exc}", parent=self)
            return

        messagebox.showinfo("復元完了", f"フローを復元しました。\n{target.name}", parent=self)
        if self.on_restored:
            self.on_restored()
        self._load_trash_list()

    def _delete_selected(self) -> None:
        p = self._get_selected_path()
        if not p:
            return

        if not messagebox.askyesno(
            "完全削除確認",
            f"'{p.name}' をゴミ箱から完全に削除しますか？\nこの操作は元に戻せません。",
            parent=self,
        ):
            return

        try:
            p.unlink()
        except OSError as exc:
            messagebox.showerror("削除失敗", f"ファイルの削除に失敗しました。\n{exc}", parent=self)
            return

        self._load_trash_list()


def main() -> None:
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()