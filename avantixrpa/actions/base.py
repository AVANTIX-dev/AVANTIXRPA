from abc import ABC, abstractmethod
from typing import Dict, Any


class ActionBase(ABC):
    """Base class for all RPA actions.

    Each concrete action must provide:
      - id: a unique string identifier (e.g. "print", "file.copy")
      - execute(context, params)
    """

    # Override in subclasses
    id: str = "base"

    @abstractmethod
    def execute(self, context: Dict[str, Any], params: Dict[str, Any]) -> None:
        raise NotImplementedError
