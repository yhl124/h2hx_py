from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .parser import parse_hwp
from .writer import write_hwpx


@dataclass(slots=True)
class ConversionResult:
    output_path: Path
    warnings: list[str]


def convert_file(source: str | Path, destination: str | Path | None = None) -> ConversionResult:
    source = Path(source)
    destination = Path(destination) if destination else source.with_suffix(".hwpx")

    document = parse_hwp(source)
    output_path = write_hwpx(document, destination)
    return ConversionResult(output_path=output_path, warnings=document.warnings)
