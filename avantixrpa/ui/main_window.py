from __future__ import annotations

import threading
from pathlib import Path
from typing import List, Dict, Any, Optional
import json
import shutil
import zipfile
import urllib.request
import urllib.error
from urllib.parse import urlparse
import html as html_lib
import re
import unicodedata
from datetime import datetime

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
from avantixrpa.core.engine import Engine, FlowStoppedException
from avantixrpa.config.paths import FLOWS_DIR, CONFIG_DIR, RESOURCES_FILE
from avantixrpa.actions.builtins import BUILTIN_ACTIONS

# パス定義（config.paths と共有）
TRASH_DIR = FLOWS_DIR / ".trash"

# ロゴ画像（config/avantix_logo.png に置く想定）
LOGO_FILE = CONFIG_DIR / "avantix_logo.png"
LOGO_FILE_DARK = CONFIG_DIR / "avantix_logo_dark.png"

# ★ 設定ファイル
SETTINGS_FILE = CONFIG_DIR / "settings.json"

APP_COPYRIGHT = "© 2025 Toshiki Azuma. All rights reserved."

DEFAULT_RESOURCES = {
    "sites": {
        "google": {
            "label": "Google",
            "url": "https://www.google.com",
        },
    },
    "files": {},
}


class DraggableStepList(tk.Frame):
    """
    ドラッグ&ドロップで並び替え可能なステップリスト。
    Canvasベースで各アイテムがヌルヌル動く。
    フローチャート風の表示。
    """
    
    ITEM_HEIGHT = 32  # 各アイテムの高さ（ボタン部分）
    ARROW_HEIGHT = 20  # 矢印部分の高さ
    ITEM_PADDING = 2   # アイテム間の余白
    
    def __init__(self, master, dark_mode: bool = False, **kwargs):
        super().__init__(master, **kwargs)
        
        self._dark_mode = dark_mode
        self._items: List[str] = []  # 表示テキストのリスト
        self._selected_index: Optional[int] = None
        self._item_widgets: List[dict] = []  # Canvasアイテムの情報
        self._last_canvas_width = 0  # 前回のCanvas幅
        
        # ドラッグ状態
        self._drag_data = {
            "active": False,
            "index": None,
            "start_y": 0,
            "current_y": 0,
        }
        
        # 色設定
        self._update_colors()
        
        # Canvas + Scrollbar
        self.canvas = tk.Canvas(
            self,
            bg=self._bg,
            highlightthickness=1,
            highlightbackground=self._border,
        )
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # イベントバインド
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_drop)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_resize)  # ★リサイズ監視
        
        # コールバック
        self._on_select_callback = None
        self._on_double_click_callback = None
        self._on_right_click_callback = None
        self._on_reorder_callback = None
    
    def _on_canvas_resize(self, event) -> None:
        """Canvasがリサイズされたら再描画"""
        new_width = event.width
        if new_width != self._last_canvas_width and new_width > 1:
            self._last_canvas_width = new_width
            self._render_items()
    
    def _update_colors(self) -> None:
        """ダークモード対応の色設定"""
        if self._dark_mode:
            self._bg = "#505050"
            self._fg = "#f0f0f0"
            self._item_bg = "#606060"        # ボタン背景
            self._item_border = "#707070"    # ボタン枠線
            self._item_hover = "#707070"
            self._item_selected = "#0078d7"
            self._item_dragging = "#707070"
            self._border = "#404040"
            self._arrow_color = "#888888"    # 矢印の色
        else:
            self._bg = "#ffffff"
            self._fg = "#333333"
            self._item_bg = "#f8f8f8"        # ボタン背景
            self._item_border = "#dddddd"    # ボタン枠線
            self._item_hover = "#f0f0f0"
            self._item_selected = "#0078d7"
            self._item_dragging = "#ffffcc"
            self._border = "#cccccc"
            self._arrow_color = "#999999"    # 矢印の色
    
    def set_dark_mode(self, dark_mode: bool) -> None:
        """ダークモードを切り替え"""
        self._dark_mode = dark_mode
        self._update_colors()
        self.canvas.configure(bg=self._bg, highlightbackground=self._border)
        self._render_items()
    
    def insert(self, index: int, text: str) -> None:
        """アイテムを挿入"""
        if index == tk.END or index >= len(self._items):
            self._items.append(text)
        else:
            self._items.insert(index, text)
        self._render_items()
    
    def delete(self, first, last=None) -> None:
        """アイテムを削除"""
        if first == 0 and last == tk.END:
            self._items.clear()
            self._selected_index = None
        elif last is None:
            if 0 <= first < len(self._items):
                del self._items[first]
                if self._selected_index == first:
                    self._selected_index = None
        self._render_items()
    
    def get(self, index) -> str:
        """アイテムを取得"""
        if 0 <= index < len(self._items):
            return self._items[index]
        return ""
    
    def size(self) -> int:
        """アイテム数を返す"""
        return len(self._items)
    
    def curselection(self) -> tuple:
        """選択中のインデックスを返す"""
        if self._selected_index is not None:
            return (self._selected_index,)
        return ()
    
    def selection_clear(self, first, last=None) -> None:
        """選択を解除"""
        self._selected_index = None
        self._render_items()
    
    def selection_set(self, index) -> None:
        """選択を設定"""
        if 0 <= index < len(self._items):
            self._selected_index = index
            self._render_items()
            self._ensure_visible(index)
    
    def _ensure_visible(self, index: int) -> None:
        """指定インデックスが見えるようにスクロール"""
        if not self._items:
            return
        slot_height = self.ITEM_HEIGHT + self.ARROW_HEIGHT + self.ITEM_PADDING
        total_height = len(self._items) * slot_height
        item_top = index * slot_height
        item_bottom = item_top + self.ITEM_HEIGHT
        
        canvas_height = self.canvas.winfo_height()
        if canvas_height <= 1:
            return
        
        # 現在の表示範囲
        view_top = self.canvas.canvasy(0)
        view_bottom = view_top + canvas_height
        
        if item_top < view_top:
            self.canvas.yview_moveto(item_top / total_height)
        elif item_bottom > view_bottom:
            self.canvas.yview_moveto((item_bottom - canvas_height) / total_height)
    
    def _render_items(self) -> None:
        """全アイテムを描画（フローチャート風ボタン＋矢印）"""
        self.canvas.delete("all")
        self._item_widgets.clear()
        
        canvas_width = self.canvas.winfo_width()
        if canvas_width <= 1:
            canvas_width = 400  # デフォルト幅
        
        drag_active = self._drag_data.get("active", False)
        drag_idx = self._drag_data.get("index")
        drag_y = self._drag_data.get("current_y", 0)
        
        # 1スロットの高さ（ボタン + 矢印）
        slot_height = self.ITEM_HEIGHT + self.ARROW_HEIGHT + self.ITEM_PADDING
        
        # ドラッグ中のアイテムが入る予定の位置（スロット）を計算
        if drag_active and drag_idx is not None:
            target_slot = int((drag_y + self.ITEM_HEIGHT // 2) // slot_height)
            target_slot = max(0, min(target_slot, len(self._items) - 1))
        else:
            target_slot = None
        
        # 描画用のマージン
        margin_x = 8
        button_left = margin_x
        button_right = canvas_width - margin_x
        
        # 各アイテムを描画
        slot = 0  # 描画するスロット位置
        total_slots = len(self._items)
        
        for i, text in enumerate(self._items):
            # ドラッグ中のアイテムはスキップ（後で描画）
            if drag_active and i == drag_idx:
                total_slots -= 1  # ドラッグ中のものは数えない
                continue
            
            # ドラッグ中で、現在のスロットがターゲット位置なら、1つずらす（隙間を作る）
            if drag_active and target_slot is not None and slot == target_slot:
                slot += 1
            
            y = slot * slot_height
            
            # 背景色を決定
            if i == self._selected_index and not drag_active:
                bg = self._item_selected
                border_color = "#005a9e"
                fg = "#ffffff"
            else:
                bg = self._item_bg
                border_color = self._item_border
                fg = self._fg
            
            # テキストを整形
            clean_text = self._strip_number(text)
            icon = self._get_step_icon(clean_text)
            formatted_text = self._format_step_text(clean_text)
            
            # ボタン風の矩形を描画（角丸風に見せるため枠線付き）
            rect = self.canvas.create_rectangle(
                button_left, y + 2,
                button_right, y + self.ITEM_HEIGHT,
                fill=bg,
                outline=border_color,
                width=1,
                tags=f"item_{i}",
            )
            
            # アイコンを描画（固定位置）
            self.canvas.create_text(
                button_left + 12, y + self.ITEM_HEIGHT // 2 + 1,
                text=icon,
                anchor="w",
                fill=fg,
                font=("Meiryo UI", 9),
                tags=f"item_{i}",
            )
            
            # テキストを描画（アイコンの後の固定位置から）
            txt = self.canvas.create_text(
                button_left + 32, y + self.ITEM_HEIGHT // 2 + 1,
                text=formatted_text,
                anchor="w",
                fill=fg,
                font=("Meiryo UI", 9),
                tags=f"item_{i}",
            )
            
            # 矢印を描画（最後のアイテム以外）
            actual_remaining = total_slots - slot - 1
            if actual_remaining > 0 or (drag_active and slot < len(self._items) - 1):
                arrow_y = y + self.ITEM_HEIGHT + self.ARROW_HEIGHT // 2 + 2
                arrow_x = canvas_width // 2
                
                # 矢印の線
                self.canvas.create_line(
                    arrow_x, y + self.ITEM_HEIGHT + 2,
                    arrow_x, y + self.ITEM_HEIGHT + self.ARROW_HEIGHT - 2,
                    fill=self._arrow_color,
                    width=2,
                    tags=f"arrow_{i}",
                )
                
                # 矢印の先端（三角形）
                self.canvas.create_polygon(
                    arrow_x - 5, y + self.ITEM_HEIGHT + self.ARROW_HEIGHT - 8,
                    arrow_x + 5, y + self.ITEM_HEIGHT + self.ARROW_HEIGHT - 8,
                    arrow_x, y + self.ITEM_HEIGHT + self.ARROW_HEIGHT - 2,
                    fill=self._arrow_color,
                    outline="",
                    tags=f"arrow_{i}",
                )
            
            self._item_widgets.append({"rect": rect, "text": txt, "index": i})
            slot += 1
        
        # ドラッグ中のアイテムを最前面に描画
        if drag_active and drag_idx is not None and 0 <= drag_idx < len(self._items):
            text = self._items[drag_idx]
            clean_text = self._strip_number(text)
            icon = self._get_step_icon(clean_text)
            formatted_text = self._format_step_text(clean_text)
            
            # ドラッグ中アイテムの背景（影付き風）
            shadow_offset = 4
            self.canvas.create_rectangle(
                button_left + shadow_offset, drag_y + 2 + shadow_offset,
                button_right + shadow_offset, drag_y + self.ITEM_HEIGHT + shadow_offset,
                fill="#00000022",
                outline="",
                tags="dragging_shadow",
            )
            
            # ドラッグ中アイテム本体
            self.canvas.create_rectangle(
                button_left, drag_y + 2,
                button_right, drag_y + self.ITEM_HEIGHT,
                fill=self._item_selected,
                outline="#ffffff",
                width=2,
                tags="dragging",
            )
            
            # ドラッグ中アイコン
            self.canvas.create_text(
                button_left + 12, drag_y + self.ITEM_HEIGHT // 2 + 1,
                text=icon,
                anchor="w",
                fill="#ffffff",
                font=("Meiryo UI", 9, "bold"),
                tags="dragging",
            )
            
            # ドラッグ中テキスト
            self.canvas.create_text(
                button_left + 32, drag_y + self.ITEM_HEIGHT // 2 + 1,
                text=formatted_text,
                anchor="w",
                fill="#ffffff",
                font=("Meiryo UI", 9, "bold"),
                tags="dragging",
            )
        
        # スクロール領域を更新（コンテンツがCanvas高さより小さい場合はスクロール無効）
        total_height = len(self._items) * slot_height + 10
        canvas_height = self.canvas.winfo_height()
        if canvas_height > 1 and total_height <= canvas_height:
            # コンテンツが表示領域に収まる場合はスクロール不要
            self.canvas.configure(scrollregion=(0, 0, canvas_width, canvas_height))
            self.canvas.yview_moveto(0)  # 先頭に戻す
        else:
            self.canvas.configure(scrollregion=(0, 0, canvas_width, total_height))
    
    def _strip_number(self, text: str) -> str:
        """テキストから先頭の番号部分を削除する"""
        import re
        # [1] や [12] などのパターンを削除
        text = re.sub(r'^\[\d+\]\s*', '', text)
        # [↕] パターンも削除
        text = re.sub(r'^\[↕\]\s*', '', text)
        # 1. や 12. などのパターンも削除
        text = re.sub(r'^\d+\.\s*', '', text)
        return text.strip()
    
    def _get_step_icon(self, text: str) -> str:
        """ステップの種類に応じたアイコンを返す"""
        text_lower = text.lower()
        if "プログラム" in text or "起動" in text:
            return "🚀"
        elif "一時停止" in text or "pause" in text_lower:
            return "⏸️"
        elif "待" in text or "wait" in text_lower:
            return "⏱️"
        elif "クリック" in text or "click" in text_lower:
            return "👆"
        elif "マウス" in text and "移動" in text:
            return "🖱️"
        elif "入力" in text or "type" in text_lower or "キーボード" in text:
            return "⌨️"
        elif "キー" in text or "hotkey" in text_lower:
            return "⌨️"
        elif "ブラウザ" in text or "url" in text_lower:
            return "🌐"
        elif "サイト" in text:
            return "🌐"
        elif "ファイル" in text:
            return "📁"
        elif "メッセージ" in text or "print" in text_lower:
            return "💬"
        else:
            return "▶️"
    
    def _format_step_text(self, text: str) -> str:
        """ステップのテキストをユーザーフレンドリーに整形"""
        # パスを短くする
        import re
        
        # C:/Program Files/.../xxx.exe → xxx.exe または フォルダ名
        def shorten_path(match):
            path = match.group(0)
            # ファイル名だけ取り出す
            parts = path.replace("\\", "/").split("/")
            filename = parts[-1] if parts else path
            # 拡張子を除いた名前
            name = filename.rsplit(".", 1)[0] if "." in filename else filename
            return name
        
        # Windowsパスのパターン
        text = re.sub(r'[A-Za-z]:[/\\][^\s\[\]]+', shorten_path, text)
        
        # [エラー時:stop] や [エラー時:continue] を削除（一旦非表示）
        text = re.sub(r'\s*\[エラー時:[^\]]+\]', '', text)
        
        # ユーザーフレンドリーな表現に変換
        text = text.replace("マウスを座標へ移動する", "マウスを移動")
        text = text.replace("マウスクリックする", "クリック")
        text = text.replace("指定秒数だけ待つ", "待機")
        text = text.replace("プログラムを起動する", "プログラム起動")
        
        # 余分なスペースを整理
        text = re.sub(r'\s+', ' ', text).strip()
        
        # " - " の前後を整理
        text = re.sub(r'\s*-\s*', ': ', text, count=1)
        
        return text
    
    def _get_index_at_y(self, y: int) -> int:
        """Y座標からインデックスを取得"""
        canvas_y = self.canvas.canvasy(y)
        slot_height = self.ITEM_HEIGHT + self.ARROW_HEIGHT + self.ITEM_PADDING
        index = int(canvas_y // slot_height)
        return max(0, min(index, len(self._items) - 1))
    
    def _on_click(self, event) -> None:
        """クリック処理"""
        if not self._items:
            return
        
        index = self._get_index_at_y(event.y)
        self._selected_index = index
        
        # ドラッグ開始準備
        slot_height = self.ITEM_HEIGHT + self.ARROW_HEIGHT + self.ITEM_PADDING
        self._drag_data = {
            "active": False,
            "index": index,
            "start_y": event.y,
            "start_canvas_y": self.canvas.canvasy(event.y),
            "current_y": index * slot_height,
        }
        
        self._render_items()
        
        if self._on_select_callback:
            self._on_select_callback(index)
    
    def _on_double_click(self, event) -> None:
        """ダブルクリック処理"""
        if self._on_double_click_callback and self._selected_index is not None:
            self._on_double_click_callback(self._selected_index)
    
    def _on_right_click(self, event) -> None:
        """右クリック処理"""
        if not self._items:
            return
        
        index = self._get_index_at_y(event.y)
        self._selected_index = index
        self._render_items()
        
        if self._on_right_click_callback:
            self._on_right_click_callback(event, index)
    
    def _on_drag(self, event) -> None:
        """ドラッグ処理"""
        if self._drag_data["index"] is None:
            return
        
        # ある程度動いたらドラッグ開始
        if not self._drag_data["active"]:
            if abs(event.y - self._drag_data["start_y"]) > 5:
                self._drag_data["active"] = True
            else:
                return
        
        # ドラッグ中の位置を更新
        slot_height = self.ITEM_HEIGHT + self.ARROW_HEIGHT + self.ITEM_PADDING
        canvas_y = self.canvas.canvasy(event.y)
        offset = canvas_y - self._drag_data["start_canvas_y"]
        original_y = self._drag_data["index"] * slot_height
        self._drag_data["current_y"] = original_y + offset
        
        self._render_items()
    
    def _on_drop(self, event) -> None:
        """ドロップ処理"""
        if not self._drag_data["active"]:
            self._drag_data = {"active": False, "index": None, "start_y": 0, "current_y": 0}
            return
        
        from_index = self._drag_data["index"]
        to_index = self._get_index_at_y(event.y)
        
        self._drag_data = {"active": False, "index": None, "start_y": 0, "current_y": 0}
        
        if from_index != to_index and from_index is not None:
            # アイテムを移動
            item = self._items.pop(from_index)
            self._items.insert(to_index, item)
            self._selected_index = to_index
            
            if self._on_reorder_callback:
                self._on_reorder_callback(from_index, to_index)
        
        self._render_items()
    
    def _on_mousewheel(self, event) -> None:
        """マウスホイールでスクロール"""
        self.canvas.yview_scroll(-1 * (event.delta // 120), "units")
    
    # コールバック設定
    def set_on_select(self, callback) -> None:
        self._on_select_callback = callback
    
    def set_on_double_click(self, callback) -> None:
        self._on_double_click_callback = callback
    
    def set_on_right_click(self, callback) -> None:
        self._on_right_click_callback = callback
    
    def set_on_reorder(self, callback) -> None:
        self._on_reorder_callback = callback
    
    # Listbox互換メソッド
    def bind(self, sequence, func):
        """Listbox互換: bind"""
        # 一部のイベントは内部で処理するのでスキップ
        if sequence in ("<Button-1>", "<B1-Motion>", "<ButtonRelease-1>", "<Double-Button-1>", "<Button-3>"):
            return
        self.canvas.bind(sequence, func)
    
    def config(self, **kwargs):
        """Listbox互換: config"""
        if "yscrollcommand" in kwargs:
            self.canvas.configure(yscrollcommand=kwargs["yscrollcommand"])
    
    def yview(self, *args):
        """Listbox互換: yview"""
        return self.canvas.yview(*args)


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
        dark_mode: bool = False,
    ) -> None:
        super().__init__(master)
        self.title("ステップ編集")
        self.resizable(False, False)
        self.grab_set()  # モーダルっぽく

        self._dark_mode = dark_mode
        self._result: Optional[Dict[str, Any]] = None

        # 座標フィールド用
        self._current_action_id: str = ""
        self._x_var: Optional[tk.StringVar] = None
        self._y_var: Optional[tk.StringVar] = None

        # リソース情報（サイト / ファイル）
        if resources is None:
            resources = {}
        self.resources: Dict[str, Any] = {
            "sites": resources.get("sites") or {},
            "files": resources.get("files") or {},
        }

        # ★ ダークモード時の色設定
        self._apply_dialog_theme()

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
                "help": "リソース管理タブで登録した「サイト」を開きます。",
                "fields": [
                    {"name": "key", "label": "サイト（表示名）", "type": "str", "default": "google"},
                ],
            },
            {
                "id": "resource.open_file",
                "label": "登録済みファイルを開く",
                "help": "リソース管理タブで登録した「ファイル」を開きます。",
                "fields": [
                    {"name": "key", "label": "ファイル（表示名）", "type": "str", "default": "sample_excel"},
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
                    {"name": "delay", "label": "実行前の待機（秒）", "type": "float", "default": None, "optional": True},
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
                    {"name": "delay", "label": "実行前の待機（秒）", "type": "float", "default": None, "optional": True},
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
                    {"name": "delay", "label": "実行前の待機（秒）", "type": "float", "default": None, "optional": True},
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
            {
                "id": "pause",
                "label": "一時停止（手動で再開）",
                "help": "ダイアログが表示され、「OK」を押すまでフローが一時停止します。手動作業を挟みたい時に使います。",
                "fields": [
                    {"name": "message", "label": "表示するメッセージ", "type": "str", "default": "準備ができたら「OK」を押してください"},
                ],
            },
        ]

        self._label_to_def = {d["label"]: d for d in self.action_defs}
        self._id_to_def = {d["id"]: d for d in self.action_defs}

        self.action_label_var = tk.StringVar()
        self.on_error_var = tk.StringVar()
        self.help_text_var = tk.StringVar()

        # name -> (tk.StringVar, field_dict)
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

        # ダークモード用の色
        if self._dark_mode:
            help_fg = "#aaaaaa"
        else:
            help_fg = "gray"

        ttk.Label(self, text="やりたいこと", style="Dialog.TLabel").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        action_combo = ttk.Combobox(
            self,
            textvariable=self.action_label_var,
            state="readonly",
            values=[d["label"] for d in self.action_defs],
            width=40,
            style="Dialog.TCombobox",
        )
        action_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        help_label = ttk.Label(
            self,
            textvariable=self.help_text_var,
            foreground=help_fg,
            wraplength=420,
            justify="left",
            style="Dialog.TLabel",
        )
        help_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=4, pady=(0, 4))

        ttk.Label(self, text="エラー時の動き", style="Dialog.TLabel").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        on_error_combo = ttk.Combobox(
            self,
            textvariable=self.on_error_var,
            state="readonly",
            values=["", "stop", "continue"],
            width=10,
            style="Dialog.TCombobox",
        )
        on_error_combo.grid(row=2, column=1, sticky="w", padx=4, pady=4)
        on_error_combo.set("")

        params_frame = ttk.LabelFrame(self, text="このステップの設定", style="Dialog.TLabelframe")
        params_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=4, pady=(4, 4))
        params_frame.columnconfigure(1, weight=1)
        self.params_frame = params_frame

        btn_frame = ttk.Frame(self, style="Dialog.TFrame")
        btn_frame.grid(row=4, column=0, columnspan=2, sticky="e", padx=4, pady=4)

        ttk.Button(btn_frame, text="OK", command=self._on_ok, style="Dialog.TButton").grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="キャンセル", command=self._on_cancel, style="Dialog.TButton").grid(row=0, column=1, padx=4)

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

        sites = self.resources.get("sites") or {}
        files = self.resources.get("files") or {}

        for row, field in enumerate(action_def.get("fields", [])):
            fname = field["name"]
            flabel = field.get("label", fname)
            default = field.get("default", "")

            ttk.Label(self.params_frame, text=flabel).grid(
                row=row, column=0, sticky="e", padx=4, pady=2
            )

            # --- resource.open_site: 表示名だけ見せるコンボ + 新規/編集 ---
            if self._current_action_id == "resource.open_site" and fname == "key":
                # keys -> displays (表示名 or key)
                site_keys = sorted(sites.keys())
                display_values = []
                for k in site_keys:
                    item = sites.get(k) or {}
                    display_values.append(item.get("label") or k)

                var = tk.StringVar()

                # initial_params に key が入っているので、表示名に変換
                if self._initial_params and fname in self._initial_params:
                    key = str(self._initial_params[fname])
                    item = sites.get(key) or {}
                    disp = item.get("label") or key
                    var.set(disp)
                elif default:
                    key = str(default)
                    item = sites.get(key) or {}
                    disp = item.get("label") or key
                    if display_values:
                        # defaultがリストにない場合もあるので、一応セット
                        var.set(disp)
                    else:
                        var.set("")
                elif display_values:
                    var.set(display_values[0])
                else:
                    var.set("")

                container = ttk.Frame(self.params_frame)
                container.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
                container.columnconfigure(0, weight=1)

                combo = ttk.Combobox(
                    container,
                    textvariable=var,
                    values=display_values,
                    state="readonly",
                    width=30,
                )
                combo.grid(row=0, column=0, sticky="ew", padx=(0, 4))

                # field情報にマッピングと種別を埋め込む
                fcopy = dict(field)
                fcopy["resource_type"] = "site"
                fcopy["keys"] = site_keys
                fcopy["display_values"] = display_values

                ttk.Button(
                    container,
                    text="新規",
                    command=lambda v=var, c=combo, fn=fname: self._open_site_resource_editor(
                        v, c, fn, is_new=True
                    ),
                ).grid(row=0, column=1, padx=(0, 2))

                ttk.Button(
                    container,
                    text="編集",
                    command=lambda v=var, c=combo, fn=fname: self._open_site_resource_editor(
                        v, c, fn, is_new=False
                    ),
                ).grid(row=0, column=2)

                self.field_vars[fname] = (var, fcopy)
                continue

            # --- resource.open_file: 表示名だけ見せるコンボ + 新規/編集 ---
            if self._current_action_id == "resource.open_file" and fname == "key":
                file_keys = sorted(files.keys())
                display_values = []
                for k in file_keys:
                    item = files.get(k) or {}
                    display_values.append(item.get("label") or k)

                var = tk.StringVar()

                if self._initial_params and fname in self._initial_params:
                    key = str(self._initial_params[fname])
                    item = files.get(key) or {}
                    disp = item.get("label") or key
                    var.set(disp)
                elif default:
                    key = str(default)
                    item = files.get(key) or {}
                    disp = item.get("label") or key
                    if display_values:
                        var.set(disp)
                    else:
                        var.set("")
                elif display_values:
                    var.set(display_values[0])
                else:
                    var.set("")

                container = ttk.Frame(self.params_frame)
                container.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
                container.columnconfigure(0, weight=1)

                combo = ttk.Combobox(
                    container,
                    textvariable=var,
                    values=display_values,
                    state="readonly",
                    width=30,
                )
                combo.grid(row=0, column=0, sticky="ew", padx=(0, 4))

                fcopy = dict(field)
                fcopy["resource_type"] = "file"
                fcopy["keys"] = file_keys
                fcopy["display_values"] = display_values

                ttk.Button(
                    container,
                    text="新規",
                    command=lambda v=var, c=combo, fn=fname: self._open_file_resource_editor(
                        v, c, fn, is_new=True
                    ),
                ).grid(row=0, column=1, padx=(0, 2))

                ttk.Button(
                    container,
                    text="編集",
                    command=lambda v=var, c=combo, fn=fname: self._open_file_resource_editor(
                        v, c, fn, is_new=False
                    ),
                ).grid(row=0, column=2)

                self.field_vars[fname] = (var, fcopy)
                continue

            # --- run.program: program だけ「参照...」ボタン付き & D&D ---
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

                if DND_AVAILABLE:
                    def _on_drop(event, target_var=var) -> None:
                        data = event.data
                        if data.startswith("{") and data.endswith("}"):
                            data = data[1:-1]
                        target_var.set(data)

                    try:
                        entry.drop_target_register(DND_FILES)
                        entry.dnd_bind("<<Drop>>", _on_drop)
                    except Exception:
                        pass

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

        self._initial_params = {}

    # ---- resources 保存ヘルパー ----
    def _save_resources_from_editor(self) -> None:
        master = self.master
        try:
            if hasattr(master, "resources"):
                master.resources = self.resources
            if hasattr(master, "_save_resources"):
                master._save_resources()
        except Exception as exc:
            print(f"[RPA] resources 保存失敗: {exc}")

    # ---- サイト用クイック編集（表示名だけ見せる版） ----
    def _open_site_resource_editor(
        self,
        target_var: tk.StringVar,
        combo: ttk.Combobox,
        field_name: str,
        is_new: bool,
    ) -> None:
        sites = self.resources.setdefault("sites", {})

        # 現在のフィールド情報（keys / display_values）を取る
        var, fdict = self.field_vars.get(field_name, (target_var, {}))
        keys: List[str] = list(fdict.get("keys") or [])
        displays: List[str] = list(fdict.get("display_values") or [])

        # 編集モードなら、現在選択中の表示名から key を逆引き
        current_key: Optional[str] = None
        if not is_new:
            current_disp = target_var.get().strip()
            if current_disp and displays and keys and len(displays) == len(keys):
                try:
                    idx = displays.index(current_disp)
                    current_key = keys[idx]
                except ValueError:
                    current_key = None

        initial_label = ""
        initial_url = ""
        if current_key and current_key in sites:
            item = sites[current_key]
            initial_label = item.get("label", "")
            initial_url = item.get("url", "")

        top = tk.Toplevel(self)
        top.title("サイトリソースの編集")
        top.resizable(False, False)
        top.transient(self)  # StepEditor を親にする
        top.grab_set()

        # ★ ダークモード対応
        if self._dark_mode:
            top.configure(bg="#505050")

        frame = ttk.Frame(top, padding=8, style="Dialog.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="表示名", style="Dialog.TLabel").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        label_var = tk.StringVar(value=initial_label)
        ttk.Entry(frame, textvariable=label_var, width=40, style="Dialog.TEntry").grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=4, pady=4
        )

        ttk.Label(frame, text="URL", style="Dialog.TLabel").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        url_var = tk.StringVar(value=initial_url)
        url_entry = ttk.Entry(frame, textvariable=url_var, width=40, style="Dialog.TEntry")
        url_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=4, pady=4)

        fetch_after_id = {"id": None}

        def _schedule_auto_fill(*_args: object) -> None:
            if fetch_after_id["id"] is not None:
                try:
                    top.after_cancel(fetch_after_id["id"])
                except Exception:
                    pass
                fetch_after_id["id"] = None
            fetch_after_id["id"] = top.after(800, _auto_fill_label_from_url)

        def _auto_fill_label_from_url() -> None:
            fetch_after_id["id"] = None

            if label_var.get().strip():
                return

            url = url_var.get().strip()
            if not url:
                return
            if "://" not in url and "." not in url:
                return

            master = self.master
            fetch_title = getattr(master, "_fetch_title_from_url", None)
            guess_label = getattr(master, "_guess_label_from_url", None)

            title = fetch_title(url) if callable(fetch_title) else None
            if title:
                label_var.set(title)
                return

            guess = guess_label(url) if callable(guess_label) else None
            if guess:
                label_var.set(guess)

        url_var.trace_add("write", _schedule_auto_fill)

        btn_frame = ttk.Frame(frame, style="Dialog.TFrame")
        btn_frame.grid(row=2, column=0, columnspan=3, sticky="e", pady=(4, 0))

        def _on_ok() -> None:
            label = label_var.get().strip()
            url = url_var.get().strip()
            if not label or not url:
                messagebox.showwarning("入力不足", "表示名とURLは必須です。", parent=top)
                return

            key = current_key or ""
            if not key:
                master = self.master
                gen_key = getattr(master, "_generate_resource_key", None)
                if callable(gen_key):
                    key = gen_key(label, "site", sites)
                else:
                    key = label

            sites[key] = {"label": label, "url": url}

            # keys / displays を更新（表示名モード）
            new_keys = sorted(sites.keys())
            new_displays = []
            for k in new_keys:
                item = sites.get(k) or {}
                new_displays.append(item.get("label") or k)

            # 対応フィールドのメタデータを更新
            if field_name in self.field_vars:
                v2, f2 = self.field_vars[field_name]
                f2["keys"] = new_keys
                f2["display_values"] = new_displays

            combo["values"] = new_displays
            # 今追加/更新したものの表示名を選択
            disp_new = sites[key].get("label") or key
            target_var.set(disp_new)

            self._save_resources_from_editor()
            top.destroy()

            # StepEditor を前面に戻す
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass

        def _on_cancel() -> None:
            top.destroy()
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass

        ttk.Button(btn_frame, text="OK", command=_on_ok, style="Dialog.TButton").grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="キャンセル", command=_on_cancel, style="Dialog.TButton").grid(row=0, column=1, padx=4)

        url_entry.focus_set()

    # ---- ファイル用クイック編集（表示名だけ見せる版） ----
    def _open_file_resource_editor(
        self,
        target_var: tk.StringVar,
        combo: ttk.Combobox,
        field_name: str,
        is_new: bool,
    ) -> None:
        files = self.resources.setdefault("files", {})

        var, fdict = self.field_vars.get(field_name, (target_var, {}))
        keys: List[str] = list(fdict.get("keys") or [])
        displays: List[str] = list(fdict.get("display_values") or [])

        current_key: Optional[str] = None
        if not is_new:
            current_disp = target_var.get().strip()
            if current_disp and displays and keys and len(displays) == len(keys):
                try:
                    idx = displays.index(current_disp)
                    current_key = keys[idx]
                except ValueError:
                    current_key = None

        initial_label = ""
        initial_path = ""
        if current_key and current_key in files:
            item = files[current_key]
            initial_label = item.get("label", "")
            initial_path = item.get("path", "")

        top = tk.Toplevel(self)
        top.title("ファイルリソースの編集")
        top.resizable(False, False)
        top.transient(self)
        top.grab_set()

        # ★ ダークモード対応
        if self._dark_mode:
            top.configure(bg="#505050")

        frame = ttk.Frame(top, padding=8, style="Dialog.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="表示名", style="Dialog.TLabel").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        label_var = tk.StringVar(value=initial_label)
        ttk.Entry(frame, textvariable=label_var, width=40, style="Dialog.TEntry").grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=4, pady=4
        )

        ttk.Label(frame, text="ファイルパス", style="Dialog.TLabel").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        path_var = tk.StringVar(value=initial_path)
        path_entry = ttk.Entry(frame, textvariable=path_var, width=40, style="Dialog.TEntry")
        path_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=4)

        def _on_browse() -> None:
            path = filedialog.askopenfilename(parent=top, title="ファイルを選択")
            if path:
                path_var.set(path)

        ttk.Button(frame, text="参照...", command=_on_browse, style="Dialog.TButton").grid(
            row=1, column=2, sticky="w", padx=(0, 4), pady=4
        )

        guess_after_id = {"id": None}

        def _schedule_auto_fill(*_args: object) -> None:
            if guess_after_id["id"] is not None:
                try:
                    top.after_cancel(guess_after_id["id"])
                except Exception:
                    pass
                guess_after_id["id"] = None
            guess_after_id["id"] = top.after(500, _auto_fill_label_from_path)

        def _auto_fill_label_from_path() -> None:
            guess_after_id["id"] = None
            if label_var.get().strip():
                return

            path = path_var.get().strip()
            if not path:
                return

            master = self.master
            guess_fn = getattr(master, "_guess_label_from_path", None)
            if callable(guess_fn):
                guess = guess_fn(path)
            else:
                p = Path(path)
                guess = p.stem or p.name

            if guess:
                label_var.set(guess)

        path_var.trace_add("write", _schedule_auto_fill)

        btn_frame = ttk.Frame(frame, style="Dialog.TFrame")
        btn_frame.grid(row=2, column=0, columnspan=3, sticky="e", pady=(4, 0))

        def _on_ok() -> None:
            label = label_var.get().strip()
            path = path_var.get().strip()
            if not label or not path:
                messagebox.showwarning("入力不足", "表示名とファイルパスは必須です。", parent=top)
                return

            key = current_key or ""
            if not key:
                master = self.master
                gen_key = getattr(master, "_generate_resource_key", None)
                if callable(gen_key):
                    key = gen_key(label, "file", files)
                else:
                    key = label

            files[key] = {"label": label, "path": path}

            new_keys = sorted(files.keys())
            new_displays = []
            for k in new_keys:
                item = files.get(k) or {}
                new_displays.append(item.get("label") or k)

            if field_name in self.field_vars:
                v2, f2 = self.field_vars[field_name]
                f2["keys"] = new_keys
                f2["display_values"] = new_displays

            combo["values"] = new_displays
            disp_new = files[key].get("label") or key
            target_var.set(disp_new)

            self._save_resources_from_editor()
            top.destroy()
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass

        def _on_cancel() -> None:
            top.destroy()
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass

        ttk.Button(btn_frame, text="OK", command=_on_ok, style="Dialog.TButton").grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="キャンセル", command=_on_cancel, style="Dialog.TButton").grid(row=0, column=1, padx=4)

        path_entry.focus_set()

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

                # ★ ダークモード対応
                if parent._dark_mode:
                    bg = "#505050"
                else:
                    bg = "#e1e1e1"
                self.configure(bg=bg)

                msg = (
                    "1. 押したい場所にマウスカーソルを動かしてください。\n"
                    "2. このウィンドウをアクティブにして Enter を押すと、\n"
                    "   その位置の座標を X/Y にセットします。"
                )
                ttk.Label(self, text=msg, justify="left", style="Dialog.TLabel").pack(padx=8, pady=(8, 4))
                self.pos_label = ttk.Label(self, text="現在の座標: x=--, y=--", style="Dialog.TLabel")
                self.pos_label.pack(padx=8, pady=(0, 8))

                ttk.Button(self, text="今の座標を反映して閉じる", command=self._finish, style="Dialog.TButton").pack(
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

            # ★ リソース系だけ、表示名→キーへの変換を挟む
            rtype = field.get("resource_type")
            if rtype in ("site", "file"):
                displays = field.get("display_values") or []
                keys = field.get("keys") or []
                value: Any = raw
                if displays and keys and len(displays) == len(keys):
                    try:
                        idx = displays.index(raw)
                        value = keys[idx]  # ← paramsには key を入れる
                    except ValueError:
                        # 万が一見つからなければそのまま raw を使う
                        value = raw
                params[fname] = value
                continue

            try:
                if ftype == "int":
                    value = int(raw)
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

    def _apply_dialog_theme(self) -> None:
        """ダークモード時にダイアログの色を設定する。"""
        if self._dark_mode:
            bg = "#505050"
            fg = "#f0f0f0"
            entry_bg = "#606060"
        else:
            bg = "#e1e1e1"
            fg = "#000000"
            entry_bg = "#ffffff"

        self.configure(bg=bg)

        # ttkスタイルをこのダイアログ用に設定
        style = ttk.Style()
        style.configure("Dialog.TFrame", background=bg)
        style.configure("Dialog.TLabel", background=bg, foreground=fg)
        style.configure("Dialog.TLabelframe", background=bg)
        style.configure("Dialog.TLabelframe.Label", background=bg, foreground=fg)
        style.configure("Dialog.TButton", background=entry_bg, foreground=fg)
        style.configure("Dialog.TEntry", fieldbackground=entry_bg, foreground=fg)
        style.configure("Dialog.TCombobox", fieldbackground=entry_bg, foreground=fg)


class CoordinateCapture(tk.Toplevel):
    """
    画面上でマウスを動かして、Enterキーを押した時点の座標を取得する。
    """

    def __init__(self, master: tk.Tk, dark_mode: bool = False) -> None:
        super().__init__(master)
        self.title("マウス座標キャプチャ")
        self.resizable(False, False)

        # ★ ダークモード対応
        if dark_mode:
            bg = "#505050"
            fg = "#f0f0f0"
        else:
            bg = "#e1e1e1"
            fg = "#000000"
        self.configure(bg=bg)

        msg = (
            "1. 押したい場所にマウスカーソルを動かしてください。\n"
            "2. このウィンドウをアクティブにして Enter を押すと、\n"
            "   その位置の座標をクリップボードにコピーします。"
        )
        ttk.Label(self, text=msg, justify="left", style="Dialog.TLabel").pack(padx=8, pady=(8, 4))

        self.pos_label = ttk.Label(self, text="現在の座標: x=--, y=--", style="Dialog.TLabel")
        self.pos_label.pack(padx=8, pady=(0, 8))

        ttk.Button(self, text="今の座標をコピーして閉じる", command=self._finish, style="Dialog.TButton").pack(
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
        self.geometry("1200x750")  # 縦幅を少し拡大（700→750）

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.option_add("*Font", "{Meiryo UI} 9")

        # ★ 設定を読み込み（ダークモードなど）
        self._settings = self._load_settings()
        self._dark_mode = self._settings.get("dark_mode", False)

        # ★ ロゴ画像の読み込み（ライト / ダーク 両方）
        self.logo_image: Optional[tk.PhotoImage] = None
        self.logo_image_dark: Optional[tk.PhotoImage] = None
        self._logo_label: Optional[ttk.Label] = None  # ロゴ表示用ラベルへの参照

        def _load_logo(logo_path: Path) -> Optional[tk.PhotoImage]:
            """ロゴ画像を読み込んで適切なサイズに縮小して返す。"""
            if not logo_path.exists():
                return None
            try:
                original = tk.PhotoImage(file=str(logo_path))
                max_width = 300
                if original.width() > max_width:
                    scale = int(original.width() / max_width)
                    if scale < 1:
                        scale = 1
                    return original.subsample(scale)
                return original
            except Exception as exc:
                print(f"[RPA] ロゴ画像の読み込みに失敗しました: {exc}")
                return None

        self.logo_image = _load_logo(LOGO_FILE)
        self.logo_image_dark = _load_logo(LOGO_FILE_DARK)

        if self.logo_image is None:
            print(f"[RPA] ロゴ画像が見つかりません: {LOGO_FILE}")

        # ===== テーマ適用 =====
        self._apply_theme()


        # ★ ウィンドウ全体のグリッド設定
        self.columnconfigure(0, weight=1)
        # row=0 にヘッダー、row=1 に Notebook を置く想定
        self.rowconfigure(1, weight=1)

        # ★ ウィンドウ上部のヘッダー（ロゴ + タイトル）
        self.header_frame = ttk.Frame(self, padding=(8, 8, 8, 4), style="AppHeader.TFrame")
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.columnconfigure(1, weight=1)

        if self.logo_image is not None:
            self._logo_label = ttk.Label(self.header_frame, image=self.logo_image, style="AppHeader.TLabel")
            self._logo_label.grid(row=0, column=0, sticky="w")
            ttk.Label(
                self.header_frame,
                text="AVANTIXRPA Launcher",
                style="AppHeader.TLabel",
            ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        else:
            ttk.Label(
                self.header_frame,
                text="AVANTIXRPA Launcher",
                style="AppHeader.TLabel",
            ).grid(row=0, column=0, sticky="w")

        self.engine = Engine(BUILTIN_ACTIONS)
        self._running_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()  # ★ 中断用イベント
        
        # 工程プレビュー表示用：アクションID → 日本語ラベル
        self._action_id_to_label = {
            "print": "メッセージを表示する",
            "wait": "指定秒数だけ待つ",
            "browser.open": "ブラウザでURLを開く",
            "resource.open_site": "登録済みサイトを開く",
            "resource.open_file": "登録済みファイルを開く",
            "run.program": "プログラムを起動する",
            "ui.type": "文字を入力する（キーボード）",
            "ui.hotkey": "キー操作を送る（Enter / Ctrl+Sなど）",
            "ui.move": "マウスを座標へ移動する",
            "ui.click": "マウスクリックする",
            "ui.scroll": "画面をスクロールする",
            "file.copy": "ファイルをコピーする",
            "file.move": "ファイルを移動する",
        }
        
        self.resources: Dict[str, Any] = self._load_resources()
        self._flow_entries: List[Dict[str, Any]] = []

        # フロー編集用
        self.edit_flow_name_var = tk.StringVar()
        self.edit_on_error_var = tk.StringVar()
        self.edit_flow_description_var = tk.StringVar()  # ★ フロー説明（1行）用
        self.edit_steps: List[Dict[str, Any]] = []

        # ★ 追加：今編集中のフロー(YAML)のパス（新規のときは None）
        self.current_edit_flow_path: Optional[Path] = None

        # フロー実行タブの詳細表示（説明＋工程プレビュー）用
        self.flow_detail_var = tk.StringVar()          # 互換用（念のため残す）
        self.flow_detail_text: Optional[tk.Text] = None  # 説明＋工程の表示用 Text

        self._create_widgets()
        self._load_flows_list()

    def _load_settings(self) -> Dict[str, Any]:
        """設定ファイルを読み込む。"""
        if not SETTINGS_FILE.exists():
            return {}
        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception as exc:
            print(f"[RPA] 設定ファイルの読み込みに失敗: {exc}")
            return {}

    def _save_settings(self) -> None:
        """設定ファイルに保存する。"""
        self._settings["dark_mode"] = self._dark_mode
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with SETTINGS_FILE.open("w", encoding="utf-8") as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            print(f"[RPA] 設定ファイルの保存に失敗: {exc}")

    def _apply_theme(self) -> None:
        """現在のダークモード状態に応じてテーマを適用する。"""
        if self._dark_mode:
            base_bg = "#505050"       # 背景（グレー寄り）
            panel_bg = "#606060"      # パネル
            fg_color = "#f0f0f0"      # 文字色
            fg_muted = "#aaaaaa"      # 薄い文字
            select_bg = "#0078d7"     # 選択色
            button_bg = "#686868"     # ボタン背景
            button_active = "#787878" # ボタンhover
            tab_bg = "#585858"        # タブ背景
            tab_selected = "#686868"  # タブ選択時
            scrollbar_bg = "#707070"  # スクロールバー
            scrollbar_trough = "#505050"
            entry_bg = "#606060"      # 入力欄背景
        else:
            base_bg = "#e1e1e1"
            panel_bg = "#ffffff"
            fg_color = "#000000"
            fg_muted = "#888888"
            select_bg = "#0078d7"
            button_bg = "#e1e1e1"
            button_active = "#c9c9c9"
            tab_bg = "#e1e1e1"
            tab_selected = "#ffffff"
            scrollbar_bg = "#c1c1c1"
            scrollbar_trough = "#e1e1e1"
            entry_bg = "#ffffff"

        self.configure(bg=base_bg)

        self.style.configure("TFrame", background=base_bg)
        self.style.configure("Main.TFrame", background=base_bg)

        # ヘッダーはダークモードのときだけpanel_bgに揃える
        header_bg = panel_bg if self._dark_mode else base_bg
        self.style.configure("AppHeader.TFrame", background=header_bg)
        self.style.configure(
            "AppHeader.TLabel",
            background=header_bg,
            foreground=fg_color,
            font=("{Meiryo UI}", 11, "bold"),
        )

        self.style.configure(
            "Card.TFrame",
            relief="groove",
            borderwidth=1,
            background=panel_bg,
        )

        self.style.configure(
            "Footer.TLabel",
            font=("{Meiryo UI}", 8),
            foreground=fg_muted,
            background=base_bg,
        )

        self.style.configure(
            "FlowDetailHeader.TLabel",
            background=base_bg,
            foreground=fg_color,
            font=("{Meiryo UI}", 9, "bold"),
        )

        self.style.configure("TLabel", background=base_bg, foreground=fg_color)
        self.style.configure("TLabelframe", background=base_bg)
        self.style.configure("TLabelframe.Label", background=base_bg, foreground=fg_color)

        # ★ ボタン
        self.style.configure(
            "TButton",
            background=button_bg,
            foreground=fg_color,
        )
        self.style.map(
            "TButton",
            background=[("active", button_active), ("pressed", button_active)],
            foreground=[("active", fg_color), ("pressed", fg_color)],
        )

        # ★ Notebook（タブ）- ダークモード時はヘッダーと同じ色に
        notebook_bg = header_bg
        self.style.configure("TNotebook", background=notebook_bg, borderwidth=0)
        self.style.configure(
            "TNotebook.Tab",
            background=tab_bg,
            foreground=fg_color,
            padding=(8, 4),
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", tab_selected), ("active", button_active)],
            foreground=[("selected", fg_color), ("active", fg_color)],
        )

        # ★ スクロールバー
        self.style.configure(
            "TScrollbar",
            background=scrollbar_bg,
            troughcolor=scrollbar_trough,
            borderwidth=0,
        )
        self.style.map(
            "TScrollbar",
            background=[("active", button_active), ("pressed", button_active)],
        )

        # ★ Entry（テキスト入力欄）
        self.style.configure(
            "TEntry",
            fieldbackground=entry_bg,
            foreground=fg_color,
            insertcolor=fg_color,
        )

        # ★ Combobox
        self.style.configure(
            "TCombobox",
            fieldbackground=entry_bg,
            background=button_bg,
            foreground=fg_color,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", entry_bg)],
            foreground=[("readonly", fg_color)],
        )

        # Listbox / Text は ttk じゃないので直接設定
        for widget in [
            getattr(self, "flows_listbox", None),
            getattr(self, "site_listbox", None),
            getattr(self, "file_listbox", None),
        ]:
            if widget:
                try:
                    widget.config(bg=panel_bg, fg=fg_color, selectbackground=select_bg)
                except Exception:
                    pass

        # ★ DraggableStepList のダークモード切り替え
        if hasattr(self, "edit_steps_list") and self.edit_steps_list:
            try:
                self.edit_steps_list.set_dark_mode(self._dark_mode)
            except Exception:
                pass

        for widget in [
            getattr(self, "log_text", None),
            getattr(self, "flow_detail_text", None),
        ]:
            if widget:
                try:
                    widget.config(bg=panel_bg, fg=fg_color)
                except Exception:
                    pass

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

    def _save_resources(self) -> None:
        try:
            with RESOURCES_FILE.open("w", encoding="utf-8") as f:
                json.dump(self.resources, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            messagebox.showerror("リソース保存エラー", f"resources.json の保存に失敗しました。\n{exc}")

    def _create_menubar(self) -> None:
        """メニューバーを作成する。"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="エクスポート...", command=self._on_export_data)
        file_menu.add_command(label="インポート...", command=self._on_import_data)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.destroy)

        # ★ 表示メニュー（新規追加）
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="表示", menu=view_menu)
        self._dark_mode_var = tk.BooleanVar(value=self._dark_mode)
        view_menu.add_checkbutton(
            label="ダークモード",
            variable=self._dark_mode_var,
            command=self._toggle_dark_mode,
        )

        # ツールメニュー
        tool_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ツール", menu=tool_menu)
        tool_menu.add_command(label="マウス座標キャプチャ", command=self._open_coord_capture)
        tool_menu.add_command(label="削除したフローを復元...", command=self._open_trash_manager)

    def _toggle_dark_mode(self) -> None:
        """ダークモードの切り替え。"""
        self._dark_mode = self._dark_mode_var.get()
        self._apply_theme()
        self._update_logo()
        self._save_settings()  # ★ 設定を保存

    def _update_logo(self) -> None:
        """ダークモード状態に応じてロゴ画像を切り替える。"""
        if self._logo_label is None:
            return

        if self._dark_mode and self.logo_image_dark is not None:
            self._logo_label.config(image=self.logo_image_dark)
        elif self.logo_image is not None:
            self._logo_label.config(image=self.logo_image)

    def _create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        # 行の weight は __init__ で設定

        # ★ メニューバー追加
        self._create_menubar()

        # ★ Notebook とタブをインスタンス変数で保持しておく
        self.notebook = ttk.Notebook(self)
        # ヘッダーの下（row=1）に配置
        self.notebook.grid(row=1, column=0, sticky="nsew")

        self.flow_tab = ttk.Frame(self.notebook, style="Main.TFrame")
        self.resource_tab = ttk.Frame(self.notebook, style="Main.TFrame")
        self.editor_tab = ttk.Frame(self.notebook, style="Main.TFrame")

        self.notebook.add(self.flow_tab, text="フロー実行")
        self.notebook.add(self.resource_tab, text="リソース管理")
        self.notebook.add(self.editor_tab, text="フローを作成・編集")

        self._create_flow_tab(self.flow_tab)
        self._create_resource_tab(self.resource_tab)
        self._create_flow_editor_tab(self.editor_tab)

        status_frame = ttk.Frame(self, padding=(8, 2), style="Main.TFrame")
        status_frame.grid(row=2, column=0, sticky="ew")
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(status_frame, text="準備完了")
        self.status_label.grid(row=0, column=0, sticky="w", padx=8)

        bottom = ttk.Frame(self, style="Main.TFrame")
        bottom.grid(row=3, column=0, sticky="ew", pady=(2, 4))
        bottom.columnconfigure(0, weight=1)

        self.run_button = ttk.Button(bottom, text="▶ フローを実行", command=self._on_run_clicked)
        self.run_button.grid(row=0, column=0, sticky="w", padx=(8, 0))

        # ★ 中断ボタン
        self.stop_button = ttk.Button(bottom, text="■ 中断", command=self._on_stop_clicked, state="disabled")
        self.stop_button.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.reload_button = ttk.Button(bottom, text="フロー再読み込み", command=self._load_flows_list)
        self.reload_button.grid(row=0, column=2, padx=(8, 0))

        # ★ フッター（コピーライト表示）
        footer = ttk.Frame(self)
        footer.grid(row=4, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        footer_label = ttk.Label(
            footer,
            text=APP_COPYRIGHT,
            anchor="center",      # ← ここを center に
            style="Footer.TLabel",
        )
        footer_label.grid(row=0, column=0, sticky="ew", padx=8)  # ← sticky を "ew" に

        # ★ キーボードショートカット設定
        self._setup_keyboard_shortcuts()

        # ★ 起動時にダークモードが有効なら適用
        if self._dark_mode:
            self._apply_theme()
            self._update_logo()

    def _setup_keyboard_shortcuts(self) -> None:
        """キーボードショートカットを設定する。"""
        # ファイル操作系
        self.bind_all("<Control-s>", lambda e: self._shortcut_save())
        self.bind_all("<Control-n>", lambda e: self._shortcut_new_flow())
        self.bind_all("<Control-o>", lambda e: self._shortcut_load_flow())

        # 実行系
        self.bind_all("<F5>", lambda e: self._on_run_clicked())
        self.bind_all("<Control-r>", lambda e: self._load_flows_list())

        # ステップ操作系（エディタタブ用）
        self.bind_all("<Delete>", lambda e: self._shortcut_delete_step())
        self.bind_all("<Control-Up>", lambda e: self._editor_move_step(-1))
        self.bind_all("<Control-Down>", lambda e: self._editor_move_step(1))

    def _shortcut_save(self) -> None:
        """Ctrl+S: 現在のタブに応じて保存処理。"""
        # エディタタブがアクティブなら保存
        try:
            current = self.notebook.index(self.notebook.select())
            if current == 2:  # フローを作成・編集タブ
                self._editor_save_flow()
        except Exception:
            pass

    def _shortcut_new_flow(self) -> None:
        """Ctrl+N: 新しいフロー作成。"""
        self._editor_new_flow()
        # エディタタブに切り替え
        try:
            self.notebook.select(self.editor_tab)
        except Exception:
            pass

    def _shortcut_load_flow(self) -> None:
        """Ctrl+O: 既存フロー読み込み。"""
        self._editor_load_flow()
        try:
            self.notebook.select(self.editor_tab)
        except Exception:
            pass

    def _shortcut_delete_step(self) -> None:
        """Delete: エディタタブでステップ削除。"""
        try:
            current = self.notebook.index(self.notebook.select())
            if current == 2:  # フローを作成・編集タブ
                self._editor_delete_step()
        except Exception:
            pass

    def _create_flow_tab(self, tab: ttk.Frame) -> None:
        # タブ全体のグリッド設定（ヘッダーはウィンドウ共通なのでここには置かない）
        tab.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)

        # --------------------------------------------------
        # 左側：フロー一覧
        # --------------------------------------------------
        left_frame = ttk.Frame(tab, padding=8, style="Card.TFrame")
        left_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(8, 4), pady=8)
        left_frame.rowconfigure(1, weight=1)
        left_frame.rowconfigure(2, weight=0)
        left_frame.rowconfigure(3, weight=0)
        left_frame.rowconfigure(4, weight=0)
        left_frame.columnconfigure(0, weight=1)

        lbl_flows = ttk.Label(left_frame, text="フロー一覧（RPA名）")
        lbl_flows.grid(row=0, column=0, sticky="w")

        # ★ selectmode="extended" で複数選択対応（Shift/Ctrl+クリック）
        self.flows_listbox = tk.Listbox(left_frame, height=18, selectmode="extended")
        self.flows_listbox.grid(row=1, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.flows_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.flows_listbox.config(yscrollcommand=scrollbar.set)

        # ダブルクリックで実行
        self.flows_listbox.bind("<Double-Button-1>", self._on_flow_double_click)

        # 選択変更で詳細表示を更新
        self.flows_listbox.bind("<<ListboxSelect>>", self._on_flow_selection_changed)

        # ★ 右クリック用コンテキストメニュー
        self.flow_list_menu = tk.Menu(self, tearoff=0)
        self.flow_list_menu.add_command(label="フローを実行", command=self._on_run_clicked)
        self.flow_list_menu.add_command(label="編集（フローエディタで開く）", command=self._on_edit_flow_from_list)
        self.flow_list_menu.add_separator()
        self.flow_list_menu.add_command(label="削除", command=self._on_delete_flow)
        self.flow_list_menu.add_command(label="削除したフローを復元...", command=self._open_trash_manager)
        self.flow_list_menu.add_separator()
        self.flow_list_menu.add_command(label="名前変更...", command=self._on_rename_flow)
        self.flow_list_menu.add_command(label="複製して新規フローを作成", command=self._on_duplicate_flow)

        # ★ 右クリックでコンテキストメニューを表示
        self.flows_listbox.bind("<Button-3>", self._on_flows_listbox_right_click)

        # ★ Deleteキーでフロー削除
        self.flows_listbox.bind("<Delete>", lambda e: self._on_delete_flow())
        self.flows_listbox.bind("<BackSpace>", lambda e: self._on_delete_flow())

        # ★ 編集ボタンだけ残す（削除・復元は右クリックに統一）
        edit_btn = ttk.Button(left_frame, text="選択フローを編集（エディタ）", command=self._on_edit_flow_from_list)
        edit_btn.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        # フロー概要 / 工程プレビュー（フロー一覧や実行ログと同じノリのラベルにする）
        detail_label = ttk.Label(
            left_frame,
            text="フロー概要 / 工程プレビュー",
        )
        detail_label.grid(row=3, column=0, sticky="w", pady=(6, 0))

        self.flow_detail_text = tk.Text(
            left_frame,
            height=5,      # 高さを3→5行に拡大
            wrap="word",
            state="disabled",
            # relief / border はデフォルトのままにして、
            # フロー一覧の Listbox や 実行ログの Text と同じ枠にする
        )
        self.flow_detail_text.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(2, 0))

        # --------------------------------------------------
        # 右側：ログエリア
        # --------------------------------------------------
        right_frame = ttk.Frame(tab, padding=8, style="Card.TFrame")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)
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
        tab.rowconfigure(1, weight=1)  # ステップ一覧が伸縮するように

        top_frame = ttk.Frame(tab)
        top_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        top_frame.columnconfigure(1, weight=1)

        ttk.Label(top_frame, text="フロー名（RPA名）").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(top_frame, textvariable=self.edit_flow_name_var).grid(
            row=0, column=1, sticky="ew", padx=4, pady=2
        )

        ttk.Label(top_frame, text="エラー時の動き（フロー全体）").grid(row=1, column=0, sticky="e", padx=4, pady=2)
        on_error_combo = ttk.Combobox(
            top_frame,
            textvariable=self.edit_on_error_var,
            state="readonly",
            values=["", "stop", "continue"],
            width=10,
        )
        on_error_combo.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        on_error_combo.set("stop")

        ttk.Label(top_frame, text="説明（任意）").grid(row=2, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(top_frame, textvariable=self.edit_flow_description_var).grid(
            row=2, column=1, sticky="ew", padx=4, pady=2
        )

        middle_frame = ttk.LabelFrame(tab, text="ステップ一覧")
        middle_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.rowconfigure(0, weight=1)

        # ★ Canvas版ドラッグ&ドロップ対応ステップリスト
        self.edit_steps_list = DraggableStepList(middle_frame, dark_mode=self._dark_mode)
        self.edit_steps_list.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)

        # ★ ダブルクリックでステップ編集
        self.edit_steps_list.set_on_double_click(lambda idx: self._editor_edit_step())

        # ★ 右クリックでコンテキストメニュー表示
        def _show_step_context(event, index):
            self.step_context_menu.tk_popup(event.x_root, event.y_root)
        self.edit_steps_list.set_on_right_click(_show_step_context)

        # ★ 並び替え時のコールバック
        def _on_reorder(from_idx, to_idx):
            if from_idx < len(self.edit_steps) and to_idx <= len(self.edit_steps):
                step = self.edit_steps.pop(from_idx)
                self.edit_steps.insert(to_idx, step)
        self.edit_steps_list.set_on_reorder(_on_reorder)

        # ★ ステップ用の右クリックメニュー
        self.step_context_menu = tk.Menu(self, tearoff=0)
        self.step_context_menu.add_command(label="編集", command=self._editor_edit_step)
        self.step_context_menu.add_command(label="複製", command=self._editor_duplicate_step)
        self.step_context_menu.add_command(label="削除", command=self._editor_delete_step)
        self.step_context_menu.add_separator()
        self.step_context_menu.add_command(label="上へ移動", command=lambda: self._editor_move_step(-1))
        self.step_context_menu.add_command(label="下へ移動", command=lambda: self._editor_move_step(1))

        # スクロールバーはDraggableStepList内部で管理するので不要
        # steps_scroll = ttk.Scrollbar(...)

        btn_frame_steps = ttk.Frame(middle_frame)
        btn_frame_steps.grid(row=0, column=1, sticky="ns", padx=4, pady=4)

        ttk.Button(btn_frame_steps, text="ステップを追加", command=self._editor_add_step).grid(row=0, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="選択したステップを編集", command=self._editor_edit_step).grid(row=1, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="選択したステップを複製", command=self._editor_duplicate_step).grid(row=2, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="選択したステップを削除", command=self._editor_delete_step).grid(row=3, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="上へ移動", command=lambda: self._editor_move_step(-1)).grid(row=4, column=0, pady=2)
        ttk.Button(btn_frame_steps, text="下へ移動", command=lambda: self._editor_move_step(1)).grid(row=5, column=0, pady=2)
        ttk.Separator(btn_frame_steps, orient="horizontal").grid(row=6, column=0, sticky="ew", pady=6)
        ttk.Button(btn_frame_steps, text="座標キャプチャ", command=self._open_coord_capture).grid(row=7, column=0, pady=2)

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
        self.editor_run_button = ttk.Button(bottom_frame, text="このフローを実行", command=self._editor_run_flow)
        self.editor_run_button.grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(bottom_frame, text="※ flows フォルダに YAML として保存されます").grid(
            row=1, column=0, columnspan=4, sticky="w", padx=4, pady=(2, 0)
        )

        # ★ 初回起動時にプレースホルダーを表示
        self._refresh_edit_steps_list()

    def _load_flows_list(self) -> None:
        self.flows_listbox.delete(0, tk.END)
        self._flow_entries.clear()

        FLOWS_DIR.mkdir(parents=True, exist_ok=True)

        yaml_files = sorted(FLOWS_DIR.glob("*.yaml"))
        for p in yaml_files:
            description = ""
            steps_raw: List[Dict[str, Any]] = []
            try:
                with p.open("r", encoding="utf-8") as f:
                    data = yaml.safe_load(f) or {}
                if not isinstance(data, dict):
                    raise ValueError("root is not mapping")
                name = data.get("name") or p.stem
                enabled = data.get("enabled", True)
                description = data.get("description") or ""
                steps_raw = data.get("steps") or []
            except Exception:
                name = p.stem
                enabled = True
                description = ""
                steps_raw = []

            self._flow_entries.append(
                {
                    "name": name,
                    "file": p,
                    "enabled": enabled,
                    "description": description,
                    "steps": steps_raw,
                }
            )

            # 表示はフロー名だけにする（ファイル名 *.yaml は隠す）
            flow_name = name if enabled else f"[無効] {name}"
            self.flows_listbox.insert(tk.END, flow_name)

        self._append_log(f"[INFO] フロー一覧を読み込みました ({len(self._flow_entries)} 件)")
        self.status_label.config(text="フロー一覧を更新しました")
        
        # 先頭のフローがあれば、その詳細を表示
        if self._flow_entries:
            self.flows_listbox.selection_clear(0, tk.END)
            self.flows_listbox.selection_set(0)
            self._on_flow_selection_changed()


    def _append_log(self, message: str) -> None:
        """
        実行ログをテキストエリアに追記する。
        先頭に [INFO] などのタグがあれば色分けし、時刻も付ける。
        """
        # ログ種別を判定（[INFO] / [RUN] / [ERROR] / [DONE] / [DELETE]）
        level_tag: Optional[str] = None
        if message.startswith("[") and "]" in message:
            level = message[1 : message.index("]")]
            if level in ("INFO", "RUN", "ERROR", "DONE", "DELETE"):
                level_tag = level

        # 時刻を付ける
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"{ts} {message}"

        self.log_text.config(state="normal")
        if level_tag:
            # レベルタグがあれば、そのタグで色分け
            self.log_text.insert(tk.END, line + "\n", (level_tag,))
        else:
            self.log_text.insert(tk.END, line + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _on_flow_double_click(self, event) -> None:
        self._on_run_clicked()

    def _on_flow_selection_changed(self, event=None) -> None:
        """フロー一覧の選択が変わったとき、説明と工程プレビューを更新する。"""
        selection = self.flows_listbox.curselection()
        if not selection:
            self.flow_detail_var.set("")
            if self.flow_detail_text is not None:
                self.flow_detail_text.configure(state="normal")
                self.flow_detail_text.delete("1.0", tk.END)
                self.flow_detail_text.configure(state="disabled")
            return

        idx = selection[0]
        if idx >= len(self._flow_entries):
            self.flow_detail_var.set("")
            if self.flow_detail_text is not None:
                self.flow_detail_text.configure(state="normal")
                self.flow_detail_text.delete("1.0", tk.END)
                self.flow_detail_text.configure(state="disabled")
            return

        entry = self._flow_entries[idx]
        description: str = entry.get("description") or ""
        steps = entry.get("steps") or []

        # 工程プレビュー（アクション名の簡易列挙）
        actions: list[str] = []
        for step in steps:
            if not isinstance(step, dict):
                continue
            action_id = step.get("action")
            if not action_id:
                continue
            # MainWindow 側で持っている ID → 日本語ラベルの表を使う
            label = self._action_id_to_label.get(str(action_id), str(action_id))
            actions.append(label)

        preview = ""
        if actions:
            # 長すぎるとウザいので先頭数件だけ表示
            preview = " → ".join(actions[:6])
            if len(actions) > 6:
                preview += " → …"

        parts: list[str] = []
        if description:
            parts.append(description)
        if preview:
            parts.append(f"[工程] {preview}")

        text = "\n".join(parts)

        if self.flow_detail_text is not None:
            self.flow_detail_text.configure(state="normal")
            self.flow_detail_text.delete("1.0", tk.END)
            if text:
                self.flow_detail_text.insert("1.0", text)
            self.flow_detail_text.configure(state="disabled")
        else:
            # 万一 Text がまだ無い場合の保険（古い UI でも落ちないように）
            self.flow_detail_var.set(text)

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

    def _on_stop_clicked(self) -> None:
        """中断ボタンが押されたとき。"""
        if not self._running_thread or not self._running_thread.is_alive():
            return

        self._stop_event.set()
        self._append_log("[INFO] 中断リクエストを送信しました（次のステップ終了時に停止します）")
        self.status_label.config(text="中断リクエスト中...")
        self.stop_button.config(state="disabled", text="中断中...")

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

        # ★ 実行中の視覚フィードバック
        self.run_button.config(state="disabled", text="⏳ 実行中...")
        self.stop_button.config(state="normal")  # ★ 中断ボタン有効化
        self.reload_button.config(state="disabled")

        # ★ 中断フラグをリセット
        self._stop_event.clear()

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

        # ★ 既に選択されている項目の上で右クリックした場合は選択を維持
        #    そうでなければ、クリックした項目だけを選択
        current_selection = self.flows_listbox.curselection()
        if index not in current_selection:
            self.flows_listbox.selection_clear(0, tk.END)
            self.flows_listbox.selection_set(index)
        self.flows_listbox.activate(index)

        try:
            self.flow_list_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.flow_list_menu.grab_release()

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
        success = True
        error_msg = ""
        stopped = False
        try:
            flow_def = load_flow(flow_path)
            # ★ 中断フラグをエンジンに渡す
            self.engine.stop_event = self._stop_event
            self.engine.run_flow(flow_def)
        except FlowStoppedException:
            # ★ ユーザーによる中断
            stopped = True
        except Exception as exc:
            success = False
            error_msg = str(exc)
        finally:
            # ★ メインスレッドにUI更新を投げる（スレッドセーフ）
            self.after(0, lambda: self._on_flow_finished(flow_name, success, error_msg, stopped))

    def _on_flow_finished(self, flow_name: str, success: bool, error_msg: str, stopped: bool = False) -> None:
        """フロー実行完了後のUI更新（メインスレッドで実行される）。"""
        # ★ 中断された場合
        if stopped:
            self._append_log(f"[STOP] フロー実行を中断しました: {flow_name}")
            self.status_label.config(text=f"フロー実行を中断しました: {flow_name}")
        elif success:
            self._append_log(f"[DONE] フロー実行完了: {flow_name}")
            self.status_label.config(text=f"フロー実行完了: {flow_name}")
        else:
            self._append_log(f"[ERROR] フロー実行中にエラーが発生しました: {error_msg}")
            self.status_label.config(text=f"フロー実行エラー: {flow_name}")

        # ボタンを元に戻す
        self.run_button.config(state="normal", text="▶ フローを実行")
        self.stop_button.config(state="disabled", text="■ 中断")  # ★ 中断ボタン無効化&テキスト戻す
        self.reload_button.config(state="normal")
        try:
            self.editor_run_button.config(state="normal", text="このフローを実行")
        except Exception:
            pass

    def _on_delete_flow(self) -> None:
        selection = self.flows_listbox.curselection()
        if not selection:
            messagebox.showinfo("フロー未選択", "削除するフローを一覧から選択してください。")
            return

        # 複数選択対応
        entries_to_delete = []
        for idx in selection:
            if idx < len(self._flow_entries):
                entries_to_delete.append(self._flow_entries[idx])

        if not entries_to_delete:
            messagebox.showerror("エラー", "内部データと表示がずれています。")
            return

        # 確認メッセージ
        if len(entries_to_delete) == 1:
            flow_name = entries_to_delete[0]["name"]
            confirm_msg = f"フロー '{flow_name}' を削除しますか？\nファイルは AVANTIXRPA のゴミ箱 (.trash) に移動されます。"
        else:
            names = [e["name"] for e in entries_to_delete]
            names_preview = "\n".join(f"  • {n}" for n in names[:5])
            if len(names) > 5:
                names_preview += f"\n  ...他 {len(names) - 5} 件"
            confirm_msg = f"{len(entries_to_delete)} 件のフローを削除しますか？\n\n{names_preview}\n\nファイルは AVANTIXRPA のゴミ箱 (.trash) に移動されます。"

        if not messagebox.askyesno("削除確認", confirm_msg):
            return

        # 削除実行
        deleted_count = 0
        for entry in entries_to_delete:
            flow_name = entry["name"]
            flow_path: Path = entry["file"]

            if not flow_path.exists():
                continue

            try:
                TRASH_DIR.mkdir(parents=True, exist_ok=True)

                target = TRASH_DIR / flow_path.name
                if target.exists():
                    stem = flow_path.stem
                    suffix = flow_path.suffix
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    target = TRASH_DIR / f"{stem}_{ts}{suffix}"

                shutil.move(str(flow_path), str(target))
                deleted_count += 1
                self._append_log(f"[DELETE] フロー '{flow_name}' をゴミ箱に移動しました。 ({flow_path.name})")
            except OSError as exc:
                self._append_log(f"[ERROR] フロー '{flow_name}' の削除に失敗: {exc}")

        self.status_label.config(text=f"{deleted_count} 件のフローを削除しました。（ゴミ箱に移動）")
        self._load_flows_list()

    def _open_trash_manager(self) -> None:
        if not TRASH_DIR.exists():
            messagebox.showinfo("ゴミ箱なし", "削除されたフローはまだありません。")
            return
        TrashManager(self, TRASH_DIR, FLOWS_DIR, on_restored=self._load_flows_list)

    def _on_export_data(self) -> None:
        """flows/*.yaml と resources.json を ZIP にエクスポートする。"""
        default_name = datetime.now().strftime("avantixrpa_export_%Y%m%d_%H%M%S.zip")
        path = filedialog.asksaveasfilename(
            title="フローとリソースをエクスポート",
            defaultextension=".zip",
            filetypes=[("ZIP ファイル", "*.zip")],
            initialfile=default_name,
        )
        if not path:
            return

        zip_path = Path(path)
        try:
            FLOWS_DIR.mkdir(parents=True, exist_ok=True)

            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                # flows/*.yaml （.trash は除外）
                for flow_path in sorted(FLOWS_DIR.glob("*.yaml")):
                    zf.write(flow_path, arcname=f"flows/{flow_path.name}")

                # config/resources.json
                if RESOURCES_FILE.exists():
                    zf.write(RESOURCES_FILE, arcname="config/resources.json")

            self.status_label.config(text=f"エクスポートしました: {zip_path.name}")
            self._append_log(f"[INFO] エクスポート: {zip_path}")
            messagebox.showinfo("エクスポート完了", f"フローとリソースをエクスポートしました:\n{zip_path}")
        except Exception as e:
            messagebox.showerror("エクスポート失敗", f"エクスポート中にエラーが発生しました:\n{e}")

    def _on_import_data(self) -> None:
        """ZIP から flows/*.yaml と config/resources.json をインポートする。"""
        path = filedialog.askopenfilename(
            title="フローとリソースをインポート",
            filetypes=[("ZIP ファイル", "*.zip"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return

        zip_path = Path(path)
        if not zip_path.exists():
            messagebox.showerror("ファイルなし", f"ZIP ファイルが見つかりません:\n{zip_path}")
            return

        try:
            FLOWS_DIR.mkdir(parents=True, exist_ok=True)

            imported_flows = 0

            with zipfile.ZipFile(zip_path, "r") as zf:
                names = zf.namelist()

                # --- flows/*.yaml をインポート ---
                for name in names:
                    if not name.endswith(".yaml"):
                        continue
                    # flows/xxx.yaml だけ対象（他のパスは無視）
                    if not (name.startswith("flows/") or "/" not in name):
                        continue

                    src_name = name
                    filename = Path(name).name
                    target = FLOWS_DIR / filename

                    # 既に同名がある場合は xxx_importN.yaml にリネーム
                    if target.exists():
                        base = target.stem
                        suffix = target.suffix
                        i = 1
                        while True:
                            candidate = FLOWS_DIR / f"{base}_import{i}{suffix}"
                            if not candidate.exists():
                                target = candidate
                                break
                            i += 1

                    with zf.open(src_name) as src, target.open("wb") as dst:
                        dst.write(src.read())
                    imported_flows += 1

                # --- resources.json をマージ ---
                if "config/resources.json" in names:
                    try:
                        with zf.open("config/resources.json") as f:
                            imported_res = json.load(f)
                    except Exception:
                        imported_res = None

                    if imported_res is not None:
                        RESOURCES_FILE.parent.mkdir(parents=True, exist_ok=True)
                        if RESOURCES_FILE.exists():
                            try:
                                with RESOURCES_FILE.open("r", encoding="utf-8") as f:
                                    current_res = json.load(f)
                            except Exception:
                                current_res = {}
                        else:
                            current_res = {}

                        # 既存優先で、無いキーだけ追加するゆるいマージ
                        def merge_dict(dst: Dict[str, Any], src: Dict[str, Any]) -> None:
                            for key, value in src.items():
                                if isinstance(value, dict) and isinstance(dst.get(key), dict):
                                    for k2, v2 in value.items():
                                        if k2 not in dst[key]:
                                            dst[key][k2] = v2
                                else:
                                    if key not in dst:
                                        dst[key] = value

                        if isinstance(current_res, dict) and isinstance(imported_res, dict):
                            merge_dict(current_res, imported_res)
                            with RESOURCES_FILE.open("w", encoding="utf-8") as f:
                                json.dump(current_res, f, ensure_ascii=False, indent=2)

            # フロー一覧を更新
            self._load_flows_list()
            self.status_label.config(text=f"インポートしました: {zip_path.name}")
            self._append_log(f"[INFO] インポート: {zip_path} （{imported_flows} 件）")
            messagebox.showinfo("インポート完了", f"{imported_flows} 件のフローをインポートしました。")
        except Exception as e:
            messagebox.showerror("インポート失敗", f"インポート中にエラーが発生しました:\n{e}")

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
        dialog = StepEditor(self, actions, resources=self.resources, dark_mode=self._dark_mode)
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
            messagebox.showinfo("ステップ未選択", "編集するステップを一覧から選択してください。")
            return
        idx = sel[0]
        if idx < 0 or idx >= len(self.edit_steps):
            return
        current = self.edit_steps[idx]
        actions = list(BUILTIN_ACTIONS.keys())
        dialog = StepEditor(self, actions, initial_step=current, resources=self.resources, dark_mode=self._dark_mode)
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

    def _editor_duplicate_step(self) -> None:
        """選択中のステップを複製して直下に挿入する。"""
        sel = self.edit_steps_list.curselection()
        if not sel:
            messagebox.showinfo("ステップ未選択", "複製するステップを選択してください。")
            return
        idx = sel[0]
        if idx < 0 or idx >= len(self.edit_steps):
            return
        
        import copy
        original = self.edit_steps[idx]
        duplicated = copy.deepcopy(original)
        
        # 複製したステップを直下に挿入
        self.edit_steps.insert(idx + 1, duplicated)
        self._refresh_edit_steps_list()
        
        # 複製したステップを選択状態にする
        self.edit_steps_list.selection_set(idx + 1)

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
        description = data.get("description", "") or ""
        steps_raw = data.get("steps") or []

        if not isinstance(steps_raw, list):
            messagebox.showerror("形式エラー", "steps が配列ではありません。このフローは編集できません。")
            return

        # 不正な要素を落として、辞書だけにしておく
        steps: List[Dict[str, Any]] = [s for s in steps_raw if isinstance(s, dict)]

        self.edit_flow_name_var.set(name)
        self.edit_on_error_var.set(on_error)
        self.edit_flow_description_var.set(description)
        self.edit_steps = steps
        self._refresh_edit_steps_list()

        # 以後「保存」したときはこのファイルに上書き
        self.current_edit_flow_path = path

        self.status_label.config(text=f"フローを読み込みました: {path.name}")

    def _refresh_edit_steps_list(self) -> None:
        """ステップ一覧の表示を、人間が読める日本語ベースに整える。"""
        self.edit_steps_list.delete(0, tk.END)

        sites = (self.resources or {}).get("sites", {})
        files = (self.resources or {}).get("files", {})

        for i, step in enumerate(self.edit_steps, start=1):
            action = step.get("action", "?")
            params = step.get("params") or {}
            on_error = step.get("on_error")

            base_label = self._action_id_to_label.get(action, action)

            # ざっくり内容の要約を作る
            summary = ""

            if action == "print":
                msg = str(params.get("message", "")).strip()
                if msg:
                    short = msg[:30]
                    if len(msg) > 30:
                        short += "…"
                    summary = f"「{short}」"

            elif action == "wait":
                sec = params.get("seconds")
                if sec is not None:
                    summary = f"{sec} 秒待つ"

            elif action == "browser.open":
                url = str(params.get("url", "")).strip()
                if url:
                    summary = url

            elif action == "resource.open_site":
                key = params.get("key")
                item = sites.get(key, {}) if key else {}
                label = item.get("label") or str(key or "")
                if label:
                    summary = f"{label}（サイト）"

            elif action == "resource.open_file":
                key = params.get("key")
                item = files.get(key, {}) if key else {}
                label = item.get("label") or str(key or "")
                if label:
                    summary = f"{label}（ファイル）"

            elif action == "run.program":
                prog = str(params.get("program", "")).strip()
                if prog:
                    summary = prog

            elif action == "ui.type":
                txt = str(params.get("text", "")).strip()
                if txt:
                    short = txt[:20]
                    if len(txt) > 20:
                        short += "…"
                    summary = f"「{short}」を入力"

            elif action == "ui.hotkey":
                keys = params.get("keys") or []
                if isinstance(keys, list) and keys:
                    summary = "+".join(keys)

            elif action in ("ui.move", "ui.click", "ui.scroll"):
                x = params.get("x")
                y = params.get("y")
                pos = ""
                if x is not None and y is not None:
                    pos = f"({x}, {y})"
                if action == "ui.scroll":
                    amount = params.get("amount")
                    if amount is not None:
                        summary = f"{pos} amount={amount}" if pos else f"amount={amount}"
                else:
                    if pos:
                        summary = pos

            elif action in ("file.copy", "file.move"):
                src = params.get("src")
                dst = params.get("dst")
                if src and dst:
                    summary = f"{src} → {dst}"

            # 最終的な表示文字列を組み立てる
            text = f"{i}. {base_label}"
            if summary:
                text += f" - {summary}"
            if on_error:
                text += f"  [エラー時: {on_error}]"

            self.edit_steps_list.insert(tk.END, text)

        # ★ ステップが空の時はプレースホルダーを表示
        if not self.edit_steps:
            self.edit_steps_list.insert(tk.END, "（ステップがありません。「ステップを追加」で追加してください）")

    def _editor_new_flow(self) -> None:
        """フローエディタをリセットして、新規作成モードにする。"""
        self.edit_flow_name_var.set("")
        self.edit_flow_description_var.set("")  # 説明もクリア
        self.edit_on_error_var.set("stop")
        self.edit_steps.clear()
        self._refresh_edit_steps_list()  # プレースホルダー表示のため
        self.current_edit_flow_path = None
        self.status_label.config(text="新しいフローの作成を開始しました")

    # ------------------------------------------------------------
    # フロー作成タブから使う「既存フロー選択」用の小さなダイアログ
    # ------------------------------------------------------------
    def _choose_flow_for_edit(self) -> Optional[Path]:
        """
        フロー実行タブと同じ一覧（self._flow_entries）から、
        エディタで編集するフローを選ばせるダイアログを出す。
        選ばれたフローの Path を返し、キャンセル時は None を返す。
        """
        # 一覧を最新状態にしておく
        self._load_flows_list()

        if not self._flow_entries:
            messagebox.showinfo(
                "フローがありません",
                "flows フォルダにフローがありません。\n先にフローを作成してください。",
                parent=self,
            )
            return None

        dlg = tk.Toplevel(self)
        dlg.title("編集するフローを選択")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=8)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="編集するフローを選択してください").grid(
            row=0, column=0, sticky="w", padx=4, pady=(0, 4)
        )

        lb = tk.Listbox(frame, height=12)
        lb.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))

        scroll = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
        scroll.grid(row=1, column=1, sticky="ns", pady=(0, 4))
        lb.config(yscrollcommand=scroll.set)

        # 表示は「フロー名だけ」 or 「[無効] フロー名」
        for entry in self._flow_entries:
            name = entry.get("name", "")
            enabled = entry.get("enabled", True)
            label = name if enabled else f"[無効] {name}"
            lb.insert(tk.END, label)

        # すでに何か編集中なら、そのフローを初期選択にする
        if self.current_edit_flow_path is not None:
            for idx, entry in enumerate(self._flow_entries):
                if entry.get("file") == self.current_edit_flow_path:
                    lb.selection_set(idx)
                    lb.see(idx)
                    break
        else:
            lb.selection_set(0)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="e")

        selected: list[Optional[Path]] = [None]

        def _on_ok() -> None:
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("フロー未選択", "編集するフローを一覧から選択してください。", parent=dlg)
                return
            idx = sel[0]
            if idx >= len(self._flow_entries):
                messagebox.showerror("エラー", "内部データと一覧がずれています。", parent=dlg)
                return
            selected[0] = self._flow_entries[idx]["file"]
            dlg.destroy()

        def _on_cancel() -> None:
            dlg.destroy()

        ttk.Button(btn_frame, text="OK", command=_on_ok).grid(row=0, column=0, padx=4, pady=(4, 0))
        ttk.Button(btn_frame, text="キャンセル", command=_on_cancel).grid(
            row=0, column=1, padx=4, pady=(4, 0)
        )

        # ダブルクリックでも OK
        lb.bind("<Double-Button-1>", lambda e: _on_ok())

        self.wait_window(dlg)
        return selected[0]

    def _editor_load_flow(self) -> None:
        """既存フロー一覧から1つ選んで、エディタに読み込む。"""
        FLOWS_DIR.mkdir(parents=True, exist_ok=True)

        path = self._choose_flow_for_edit()
        if path is None:
            return

        try:
            with path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception as exc:
            messagebox.showerror("読み込み失敗", f"フローの読み込みに失敗しました。\n{exc}")
            return

        if not isinstance(data, dict):
            messagebox.showerror("形式エラー", "フローファイルの形式が不正です。")
            return

        name = data.get("name") or ""
        on_error = data.get("on_error") or "stop"
        description = data.get("description") or ""
        steps_raw = data.get("steps") or []

        self.edit_flow_name_var.set(name)
        self.edit_on_error_var.set(on_error)
        self.edit_flow_description_var.set(description)

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

        self.current_edit_flow_path = path
        self._refresh_edit_steps_list()
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
        description = self.edit_flow_description_var.get().strip()

        data = {
            "name": name,
            "on_error": on_error,
            "steps": self.edit_steps,
        }
        if description:
            data["description"] = description

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

        # ★ 中断フラグをリセット
        self._stop_event.clear()

        # ★ 実行中はボタンをロック＆フィードバック表示
        try:
            self.run_button.config(state="disabled", text="⏳ 実行中...")
            self.stop_button.config(state="normal")  # ★ 中断ボタン有効化
            self.reload_button.config(state="disabled")
            self.editor_run_button.config(state="disabled", text="⏳ 実行中...")
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
        CoordinateCapture(self, dark_mode=self._dark_mode)

    def _open_trash_manager(self) -> None:
        if not TRASH_DIR.exists():
            messagebox.showinfo("ゴミ箱なし", "削除されたフローはまだありません。")
            return
        TrashManager(self, TRASH_DIR, FLOWS_DIR, on_restored=self._load_flows_list, dark_mode=self._dark_mode)


class TrashManager(tk.Toplevel):
    """flows/.trash にある削除済みフローの一覧と復元/完全削除を行うダイアログ。"""

    def __init__(
        self,
        master: tk.Tk,
        trash_dir: Path,
        flows_dir: Path,
        on_restored: Optional[callable] = None,
        dark_mode: bool = False,
    ) -> None:
        super().__init__(master)
        self.title("削除したフローの管理")
        self.resizable(False, False)

        self.trash_dir = trash_dir
        self.flows_dir = flows_dir
        self.on_restored = on_restored
        self._dark_mode = dark_mode

        # ★ ダークモード対応
        if dark_mode:
            self._bg = "#505050"
            self._fg = "#f0f0f0"
            self._panel_bg = "#606060"
        else:
            self._bg = "#e1e1e1"
            self._fg = "#000000"
            self._panel_bg = "#ffffff"
        self.configure(bg=self._bg)

        self._files: list[Path] = []

        self._create_widgets()
        self._load_trash_list()

        self.grab_set()
        self.focus_set()

    def _create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ttk.Label(self, text="ゴミ箱にあるフロー（.trash）", style="Dialog.TLabel").grid(
            row=0, column=0, sticky="w", padx=8, pady=(8, 4)
        )

        frame = ttk.Frame(self, style="Dialog.TFrame")
        frame.grid(row=1, column=0, sticky="nsew", padx=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(frame, height=12, width=60, bg=self._panel_bg, fg=self._fg, selectbackground="#0078d7")
        self.listbox.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.listbox.config(yscrollcommand=scroll.set)

        btn_frame = ttk.Frame(self, style="Dialog.TFrame")
        btn_frame.grid(row=2, column=0, sticky="e", padx=8, pady=(4, 8))
        ttk.Button(btn_frame, text="復元", command=self._restore_selected, style="Dialog.TButton").grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="完全に削除", command=self._delete_selected, style="Dialog.TButton").grid(row=0, column=1, padx=4)
        ttk.Button(btn_frame, text="閉じる", command=self.destroy, style="Dialog.TButton").grid(row=0, column=2, padx=4)

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