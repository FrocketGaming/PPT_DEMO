"""Command line: `deckgen configs/demo.yaml [-o out/demo.pptx]`."""

from __future__ import annotations

import argparse
import sys

from .builder import ConfigError, build


def main() -> None:
    parser = argparse.ArgumentParser(prog="deckgen", description="Build a .pptx from a YAML config.")
    parser.add_argument("config", help="path to the deck YAML")
    parser.add_argument("-o", "--output", help="output .pptx (default: out/<config name>.pptx)")
    args = parser.parse_args()
    try:
        path = build(args.config, args.output)
    except (ConfigError, FileNotFoundError) as exc:
        sys.exit(f"error: {exc}")
    print(f"wrote {path}")


if __name__ == "__main__":
    main()
