from __future__ import annotations

from .print_action import PrintAction
from .wait_action import WaitAction
from .file_actions import FileCopyAction, FileMoveAction
from .run_program_action import RunProgramAction
from .browser_actions import BrowserOpenAction
from .ui_actions import UiTypeAction, UiHotkeyAction, UiMoveAction, UiClickAction, UiScrollAction
from .resource_actions import ResourceOpenSiteAction, ResourceOpenFileAction


# NOTE:
# Engine は Dict[str, Type[ActionBase]]（アクションクラス）を前提としているので
# ここでは **インスタンスではなくクラス** を登録すること。
BUILTIN_ACTIONS = {
    "print": PrintAction,
    "wait": WaitAction,
    "file.copy": FileCopyAction,
    "file.move": FileMoveAction,
    "run.program": RunProgramAction,
    "browser.open": BrowserOpenAction,
    "ui.type": UiTypeAction,
    "ui.hotkey": UiHotkeyAction,
    "ui.move": UiMoveAction,
    "ui.click": UiClickAction,
    "ui.scroll": UiScrollAction,
    "resource.open_site": ResourceOpenSiteAction,
    "resource.open_file": ResourceOpenFileAction,
}
