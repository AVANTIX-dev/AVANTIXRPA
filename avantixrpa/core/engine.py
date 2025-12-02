from typing import Dict, Type, Optional, Any

from avantixrpa.actions.base import ActionBase
from avantixrpa.logging.logger import get_logger


class Engine:
    """Very small RPA engine that runs a flow definition dict.

    flow_def expected structure:

    {
      "name": "Flow name",
      "on_error": "stop" | "continue",   # optional, default: "stop"
      "steps": [
        {
          "action": "print",
          "params": {...},
          "continue_on_error": true      # optional, step-level override
        },
        ...
      ]
    }
    """

    def __init__(self, actions: Dict[str, Type[ActionBase]]):
        self.actions = actions
        self.logger = get_logger("avantixrpa.engine")

    def run_flow(self, flow_def: dict, context: Optional[dict] = None) -> None:
        if context is None:
            context = {}

        name = flow_def.get("name", "Unnamed Flow")
        steps = flow_def.get("steps", [])

        # フロー全体のエラー方針
        flow_on_error = (flow_def.get("on_error") or "stop").lower()
        if flow_on_error not in ("stop", "continue"):
            raise ValueError("on_error must be 'stop' or 'continue'")

        msg = f"Start flow: {name} (on_error={flow_on_error})"
        print(f"[ENGINE] {msg}")
        self.logger.info(msg)

        if not isinstance(steps, list):
            raise ValueError("flow_def['steps'] must be a list")

        for idx, step in enumerate(steps, start=1):
            action_id = step.get("action")
            params: dict = step.get("params") or {}

            # ステップ単位のエラー方針
            step_continue_flag = step.get("continue_on_error", False)
            # True ならこのステップはエラーでも続行、それ以外はフローの on_error に従う
            step_policy = "continue" if step_continue_flag else flow_on_error

            step_msg = f"Step {idx}: {action_id} (on_error={step_policy})"
            print(f"[ENGINE] {step_msg}")
            self.logger.info(step_msg)

            if action_id not in self.actions:
                err = f"Unknown action: {action_id!r}"
                self.logger.error(err)
                raise ValueError(err)

            action_cls = self.actions[action_id]
            action = action_cls()

            try:
                action.execute(context, params)
                self.logger.info(f"Step {idx} completed: {action_id}")
            except Exception as exc:  # noqa: BLE001
                err = f"Step {idx} failed: {exc}"
                print(f"[ENGINE] {err}")
                # stack trace付き
                self.logger.exception(err)

                if step_policy == "continue":
                    # ログだけ吐いて次のステップへ
                    continue

                # デフォルト: 停止
                raise

        end_msg = "Flow finished"
        print(f"[ENGINE] {end_msg}")
        self.logger.info(end_msg)
