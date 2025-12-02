from typing import Dict, Any

try:
    import pyautogui
except ImportError as e:
    pyautogui = None
    _import_error = e
else:
    _import_error = None

from avantixrpa.actions.base import ActionBase


def _ensure_pyautogui():
    if pyautogui is None:
        raise RuntimeError(
            "pyautogui がインストールされていません。"
            " 'py -m pip install pyautogui' を実行してください。"
        )

class UiScrollAction(ActionBase):
    """マウスホイールで画面をスクロールするアクション。

    例:
      - action: ui.scroll
        params:
          amount: -500   # 下方向にスクロール（プラスで上 / マイナスで下）
          x: 960         # 任意。省略時は現在のマウス位置
          y: 540
    """

    id = "ui.scroll"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        _ensure_pyautogui()

        amount = params.get("amount")
        if amount is None:
            raise ValueError("ui.scroll には 'amount' パラメータが必要です（+で上 / -で下）")

        try:
            clicks = int(amount)
        except (TypeError, ValueError):
            raise ValueError(f"ui.scroll: amount は整数で指定してください（指定値: {amount!r}）")

        x = params.get("x")
        y = params.get("y")

        if x is not None and y is not None:
            print(f"[RPA] スクロール: amount={clicks}, x={x}, y={y}")
            pyautogui.scroll(clicks, x=int(x), y=int(y))
        else:
            print(f"[RPA] スクロール: amount={clicks}（現在のマウス位置）")
            pyautogui.scroll(clicks)


class UiMoveAction(ActionBase):
    """マウスカーソルを指定座標に移動するアクション。

    例:
      - action: ui.move
        params:
          x: 100
          y: 200
          duration: 0.2
    """

    id = "ui.move"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        _ensure_pyautogui()
        x = params.get("x")
        y = params.get("y")
        duration = float(params.get("duration", 0.0))

        if x is None or y is None:
            raise ValueError("ui.move には 'x', 'y' パラメータが必要です")

        print(f"[RPA] マウス移動: x={x}, y={y}, duration={duration}")
        pyautogui.moveTo(int(x), int(y), duration=duration)


class UiClickAction(ActionBase):
    """マウスクリックを行うアクション。

    例:
      - action: ui.click
        params:
          x: 100       # 省略すると現在位置
          y: 200
          button: "left"
          clicks: 1
    """

    id = "ui.click"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        _ensure_pyautogui()
        x = params.get("x")
        y = params.get("y")
        button = params.get("button", "left")
        clicks = int(params.get("clicks", 1))
        interval = float(params.get("interval", 0.0))

        print(f"[RPA] クリック: x={x}, y={y}, button={button}, clicks={clicks}")

        if x is not None and y is not None:
            pyautogui.click(int(x), int(y), clicks=clicks, interval=interval, button=button)
        else:
            pyautogui.click(clicks=clicks, interval=interval, button=button)


class UiTypeAction(ActionBase):
    """キーボード入力を行うアクション。

    例:
      - action: ui.type
        params:
          text: "これはテストです"
          interval: 0.05
    """

    id = "ui.type"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        _ensure_pyautogui()
        text = params.get("text")
        interval = float(params.get("interval", 0.0))

        if not text:
            raise ValueError("ui.type には 'text' パラメータが必要です")

        print(f"[RPA] キーボード入力: {text!r}")
        pyautogui.write(text, interval=interval)


class UiHotkeyAction(ActionBase):
    """ショートカットキー（ホットキー）を送信するアクション。

    例:
      - action: ui.hotkey
        params:
          keys: ["ctrl", "s"]
    """

    id = "ui.hotkey"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        _ensure_pyautogui()
        keys = params.get("keys")
        if not keys or not isinstance(keys, (list, tuple)):
            raise ValueError("ui.hotkey には 'keys' リストが必要です (例: ['ctrl', 's'])")

        print(f"[RPA] ホットキー送信: {keys}")
        pyautogui.hotkey(*keys)
