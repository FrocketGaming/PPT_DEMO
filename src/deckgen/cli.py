"""Command line.

    deckgen configs/demo.yaml [-o out/demo.pptx]     build a deck
    deckgen layouts "Brand Template"                   list a template's layouts/placeholders
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .builder import ConfigError, build
from .template import describe, find_template


def main() -> None:
    if sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) > 1 and sys.argv[1] == "layouts":
        parser = argparse.ArgumentParser(prog="deckgen layouts",
                                         description="List a template's layouts and placeholders.")
        parser.add_argument("template", help="template name in templates/ or a path")
        args = parser.parse_args(sys.argv[2:])
        try:
            path = find_template(args.template, [Path.cwd(), Path.cwd() / "templates"])
        except FileNotFoundError as exc:
            sys.exit(f"error: {exc}")
        print(describe(path))
        return

    parser = argparse.ArgumentParser(prog="deckgen", description="Build a .pptx from a YAML config.")
    parser.add_argument("config", help="path to the deck YAML")
    parser.add_argument("-o", "--output", help="output .pptx (default: ./out/<config name>.pptx)")
    args = parser.parse_args()
    try:
        path = build(args.config, args.output)
    except (ConfigError, FileNotFoundError) as exc:
        sys.exit(f"error: {exc}")
    print(f"wrote {path}")


if __name__ == "__main__":
    main()
