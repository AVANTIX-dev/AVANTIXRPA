from typing import Dict, Any
import time

from avantixrpa.actions.base import ActionBase


class WaitAction(ActionBase):
    """Sleeps for a given number of seconds."""

    id = "wait"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        seconds = float(params.get("seconds", 1))
        print(f"[RPA] Waiting {seconds} seconds...")
        time.sleep(seconds)
