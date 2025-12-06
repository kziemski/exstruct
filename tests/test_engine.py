from pathlib import Path

import pytest

from exstruct.engine import ExStructEngine, OutputOptions, StructOptions
from exstruct.models import Chart, ChartSeries, SheetData, Shape, WorkbookData


def test_engine_extract_uses_mode(monkeypatch, tmp_path: Path) -> None:
    called = {}

    def fake_extract(path: Path, mode: str):
        called["mode"] = mode
        return WorkbookData(book_name=path.name, sheets={})

    monkeypatch.setattr("exstruct.engine.extract_workbook", fake_extract)
    engine = ExStructEngine(options=StructOptions(mode="standard"))
    engine.extract(tmp_path / "book.xlsx", mode="verbose")
    assert called["mode"] == "verbose"


def _sample_workbook() -> WorkbookData:
    shape = Shape(text="x", l=0, t=0, w=10, h=10, type="Rect")
    chart = Chart(
        name="c1",
        chart_type="Line",
        title=None,
        y_axis_title="",
        y_axis_range=[],
        series=[ChartSeries(name="s1")],
        l=0,
        t=0,
        error=None,
    )
    sheet = SheetData(rows=[], shapes=[shape], charts=[chart], table_candidates=["A1:B2"])
    return WorkbookData(book_name="book.xlsx", sheets={"Sheet1": sheet})


def test_engine_serialize_filters_shapes(tmp_path: Path) -> None:
    wb = _sample_workbook()
    engine = ExStructEngine(output=OutputOptions(include_shapes=False))
    text = engine.serialize(wb, fmt="json")
    assert '"shapes"' not in text


def test_engine_serialize_filters_tables(tmp_path: Path) -> None:
    wb = _sample_workbook()
    engine = ExStructEngine(output=OutputOptions(include_tables=False))
    text = engine.serialize(wb, fmt="json")
    assert "table_candidates" not in text


def test_engine_export_respects_sheets_dir(tmp_path: Path) -> None:
    wb = _sample_workbook()
    sheets_dir = tmp_path / "sheets"
    engine = ExStructEngine(output=OutputOptions(sheets_dir=sheets_dir))
    out = tmp_path / "out.json"
    engine.export(wb, output_path=out)
    assert out.exists()
    assert sheets_dir.exists()
    files = list(sheets_dir.glob("*.json"))
    assert len(files) == 1
