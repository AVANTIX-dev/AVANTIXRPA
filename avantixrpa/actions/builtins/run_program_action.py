from typing import Dict, Any
from pathlib import Path
import os
import subprocess
import shlex

from avantixrpa.actions.base import ActionBase


class RunProgramAction(ActionBase):
    """外部プログラムを起動するアクション。

    例:
      - action: run.program
        params:
          program: "notepad.exe"
          args: ""
    """

    id = "run.program"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        # 新仕様: program + args
        program = params.get("program")
        args_str = params.get("args") or ""

        # 旧仕様: command だけ（互換用）
        legacy_command = params.get("command")

        if not program and not legacy_command:
            raise ValueError("run.program には 'program' か 'command' パラメータのどちらかが必要です")

        # 新しい形式が優先
        if program:
            program_str = str(program)

            if args_str:
                try:
                    args_list = shlex.split(str(args_str))
                except ValueError:
                    args_list = [str(args_str)]
            else:
                args_list = []

            p = Path(program_str)

            # 「パスとして書いてるくせに存在しない」場合は分かりやすく怒る
            if any(sep in program_str for sep in ("\\", "/")) and not p.exists():
                raise FileNotFoundError(f"run.program: 指定されたパスが存在しません: {program_str}")

            # ★ .lnk (ショートカット) は os.startfile で開く
            if p.suffix.lower() == ".lnk" and p.exists():
                print(f"[RPA] ショートカットを起動します: {program_str}")
                os.startfile(program_str)
                return

            cmd_list = [program_str] + args_list
            print(f"[RPA] プログラム起動: {' '.join(cmd_list)}")
            subprocess.Popen(cmd_list)
            return

