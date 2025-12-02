from typing import Dict, Any
import webbrowser

from avantixrpa.actions.base import ActionBase


class BrowserOpenAction(ActionBase):
    """デフォルトブラウザでURLを開くアクション。

    例:
      - action: browser.open
        params:
          url: "https://www.google.com"
    """

    id = "browser.open"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        url = params.get("url")
        if not url:
            raise ValueError("browser.open には 'url' パラメータが必要です")

        # 将来ブラウザ指定したくなったとき用の予約パラメータ
        browser_name = params.get("browser")  # 例: "chrome"（今は未使用）

        print(f"[RPA] ブラウザでURLを開きます: {url}")

        # シンプルにデフォルトブラウザで開く
        opened = webbrowser.open(url, new=2)  # new=2: 新しいタブで開く（対応ブラウザなら）

        if not opened:
            # 一応失敗検知できる可能性はあるが、環境依存なので軽めにエラー扱い
            raise RuntimeError(f"URLを開けませんでした: {url}")
