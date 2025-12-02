from pathlib import Path
from typing import Union
import yaml  # Requires: pip install pyyaml

from avantixrpa.config.paths import FLOWS_DIR


def load_flow(path_or_name: Union[str, Path]) -> dict:
    """Load a YAML flow definition.

    If a relative path or name is given, it is resolved under FLOWS_DIR.
    """
    p = Path(path_or_name)
    if not p.is_absolute():
        p = FLOWS_DIR / p

    if not p.exists():
        raise FileNotFoundError(f"Flow file not found: {p}")

    with p.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    if not isinstance(data, dict):
        raise ValueError("Flow definition root must be an object (mapping)")

    return data
