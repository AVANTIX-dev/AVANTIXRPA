from typing import Dict, Any
from avantixrpa.actions.base import ActionBase


class PrintAction(ActionBase):
    """Prints a message to stdout."""

    id = "print"

    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        prefix = params.get("prefix", "[RPA]")
        message = params.get("message", "")
        print(f"{prefix} {message}")
