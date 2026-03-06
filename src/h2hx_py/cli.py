from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .converter import convert_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert a .hwp file to .hwpx.")
    parser.add_argument("source", type=Path, help="Source .hwp file")
    parser.add_argument("-o", "--output", type=Path, help="Destination .hwpx file")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    result = convert_file(args.source, args.output)
    print(result.output_path)
    for warning in result.warnings:
        print(f"warning: {warning}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
