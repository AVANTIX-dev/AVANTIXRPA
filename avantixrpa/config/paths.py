from __future__ import annotations

from pathlib import Path
import os
import json
from typing import Dict, Any

# プロジェクトルート（AVANTIXRPA のルートフォルダ）
PROJECT_ROOT = Path(__file__).resolve().parents[2]

CONFIG_DIR = PROJECT_ROOT / "config"
FLOWS_DIR = PROJECT_ROOT / "flows"
LOGS_DIR = PROJECT_ROOT / "logs"
RUNS_DIR = PROJECT_ROOT / "runs"
RESOURCES_FILE = CONFIG_DIR / "resources.json"

# env.local.json キャッシュ
_ENV_CACHE: Dict[str, Any] | None = None


def _load_env() -> Dict[str, Any]:
    """config/env.local.json を読み込む（あれば）。"""
    global _ENV_CACHE
    if _ENV_CACHE is not None:
        return _ENV_CACHE

    env_file = CONFIG_DIR / "env.local.json"
    data: Dict[str, Any] = {}
    if env_file.exists():
        try:
            with env_file.open("r", encoding="utf-8") as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    data = loaded
        except Exception:
            # env.local.json が壊れてても、とりあえず空として扱う
            data = {}

    _ENV_CACHE = data
    return data


def _get_default_placeholders() -> Dict[str, str]:
    """OS依存の標準プレースホルダを返す。"""
    home = Path.home()

    # Windows 前提だけど、一応他OSでもそれっぽく動くように書いておく
    desktop = home / "Desktop"
    documents = home / "Documents"

    return {
        "PROJECT_ROOT": str(PROJECT_ROOT),
        "AVANTIX_ROOT": str(PROJECT_ROOT),
        "DESKTOP": str(desktop),
        "DOCUMENTS": str(documents),
    }


def expand_path(template: str) -> Path:
    """
    パス文字列に含まれる {PLACEHOLDER} を実行環境に合わせて展開して Path を返す。

    例:
      "{DESKTOP}/AVANTIXRPA/log.xlsx"
      "{AVANTIX_ROOT}/test_data/source/sample_log.txt"

    さらに env.local.json の "placeholders" で任意のプレースホルダを追加可能:

      {
        "placeholders": {
          "DATA_ROOT": "D:/shared/data"
        }
      }

    フロー側では:

      src: "{DATA_ROOT}/input.csv"
    """
    if not template:
        raise ValueError("expand_path に空のパス文字列が渡されました")

    env = _load_env()
    placeholders = _get_default_placeholders()

    # env.local.json の placeholders を上書き/追加
    user_ph = env.get("placeholders") if isinstance(env, dict) else None
    if isinstance(user_ph, dict):
        for k, v in user_ph.items():
            placeholders[str(k)] = str(v)

    result = str(template)

    for key, value in placeholders.items():
        token = "{" + key + "}"
        if token in result:
            result = result.replace(token, value)

    # チルダ展開（~）も一応対応
    return Path(os.path.expanduser(result)).resolve()