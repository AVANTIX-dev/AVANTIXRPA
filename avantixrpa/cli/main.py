import argparse
from typing import Dict, Type

from avantixrpa.core.engine import Engine
from avantixrpa.core.flow_loader import load_flow
from avantixrpa.actions.base import ActionBase
from avantixrpa.actions.builtins import BUILTIN_ACTIONS


def build_actions_registry() -> Dict[str, Type[ActionBase]]:
    """Collect all available actions.

    Later this can be extended to scan plugins/ dynamically.
    """
    return dict(BUILTIN_ACTIONS)


def main() -> int:
    parser = argparse.ArgumentParser(description="AVANTIXRPA minimal CLI")
    parser.add_argument(
        "flow",
        help="Flow YAML path or name under flows/ (e.g. sample_flow.yaml)",
    )
    args = parser.parse_args()

    actions = build_actions_registry()
    engine = Engine(actions)

    flow_def = load_flow(args.flow)
    engine.run_flow(flow_def)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
