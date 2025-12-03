from __future__ import annotations

import argparse
from pathlib import Path

from exstruct import process_excel


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Dev-only CLI stub for ExStruct extraction."
    )
    parser.add_argument("input", type=Path, help="Excel file (.xlsx/.xlsm/.xls)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output path (defaults to <input>.json)",
    )
    parser.add_argument(
        "-f",
        "--format",
        default="json",
        choices=["json"],
        help="Export format (dev stub only supports json)",
    )
    parser.add_argument(
        "--image",
        action="store_true",
        help="(placeholder) Render PNG alongside JSON",
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="(placeholder) Render PDF alongside JSON",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=144,
        help="DPI for image rendering (placeholder)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path: Path = args.input
    output_path: Path = args.output or input_path.with_suffix(".json")

    process_excel(
        file_path=input_path,
        output_path=output_path,
        out_fmt=args.format,
        image=args.image,
        pdf=args.pdf,
        dpi=args.dpi,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
