from typing import Dict, Any
import shutil

from avantixrpa.actions.base import ActionBase
from avantixrpa.config.paths import expand_path


class FileCopyAction(ActionBase):
    """ファイルコピーアクション。

    例:
      - action: file.copy
        params:
          src: "{AVANTIX_ROOT}/test_data/source/sample_log.txt"
          dst: "{DESKTOP}/AVANTIX_BACKUP/sample_log_backup.txt"
    """

    id = "file.copy"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        src = params.get("src")
        dst = params.get("dst")
        if not src or not dst:
            raise ValueError("file.copy には 'src' と 'dst' が必要です")

        src_path = expand_path(str(src))
        dst_path = expand_path(str(dst))

        print(f"[RPA] Copying {src_path} -> {dst_path}")
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dst_path)


class FileMoveAction(ActionBase):
    """ファイル移動アクション。"""

    id = "file.move"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        src = params.get("src")
        dst = params.get("dst")
        if not src or not dst:
            raise ValueError("file.move には 'src' と 'dst' が必要です")

        src_path = expand_path(str(src))
        dst_path = expand_path(str(dst))

        print(f"[RPA] Moving {src_path} -> {dst_path}")
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(src_path, dst_path)
