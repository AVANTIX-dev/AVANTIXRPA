
from __future__ import annotations

import json
import os
import webbrowser
from typing import Any, Dict
from pathlib import Path

from ..base import ActionBase

from avantixrpa.config.paths import CONFIG_DIR, RESOURCES_FILE
# リソースファイル（config.paths と共有）


def _load_resources() -> Dict[str, Any]:
    """UI側と同じ resources.json を読み込む。古い形式(string)と新形式(dict)の両方に対応。"""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not RESOURCES_FILE.exists():
        return {"sites": {}, "files": {}}

    with RESOURCES_FILE.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        return {"sites": {}, "files": {}}

    sites_raw = data.get("sites", {})
    files_raw = data.get("files", {})

    sites: Dict[str, Dict[str, str]] = {}
    if isinstance(sites_raw, dict):
        for key, v in sites_raw.items():
            if isinstance(v, str):
                sites[key] = {"label": key, "url": v}
            elif isinstance(v, dict):
                url = v.get("url") or ""
                label = v.get("label") or key
                sites[key] = {"label": label, "url": url}

    files: Dict[str, Dict[str, str]] = {}
    if isinstance(files_raw, dict):
        for key, v in files_raw.items():
            if isinstance(v, str):
                files[key] = {"label": key, "path": v}
            elif isinstance(v, dict):
                path = v.get("path") or ""
                label = v.get("label") or key
                files[key] = {"label": label, "path": path}

    return {"sites": sites, "files": files}


class ResourceOpenSiteAction(ActionBase):
    action_id = "resource.open_site"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        key = params.get("key")
        if not key:
            raise ValueError("resource.open_site: 'key' パラメータが指定されていません。")

        resources = _load_resources()
        sites = resources.get("sites", {})

        entry = sites.get(key)
        if entry is None:
            raise KeyError(f"resources.json の sites にキー '{key}' が見つかりません。")

        if isinstance(entry, dict):
            url = entry.get("url")
        else:
            url = str(entry)

        if not url:
            raise ValueError(f"サイト '{key}' の URL が空です。")

        print(f"[RPA] 登録サイトを開きます: key={key}, url={url}")
        opened = webbrowser.open(url, new=2)
        if not opened:
            raise RuntimeError(f"ブラウザで URL を開けませんでした: {url}")


class ResourceOpenFileAction(ActionBase):
    action_id = "resource.open_file"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        key = params.get("key")
        if not key:
            raise ValueError("resource.open_file: 'key' パラメータが指定されていません。")

        resources = _load_resources()
        files = resources.get("files", {})

        entry = files.get(key)
        if entry is None:
            raise KeyError(f"resources.json の files にキー '{key}' が見つかりません。")

        if isinstance(entry, dict):
            path = entry.get("path")
        else:
            path = str(entry)

        if not path:
            raise ValueError(f"ファイル '{key}' のパスが空です。")

        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"登録ファイルが存在しません: {p}")

        print(f"[RPA] 登録ファイルを開きます: key={key}, path={p}")
        os.startfile(str(p))