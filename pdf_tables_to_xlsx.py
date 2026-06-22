#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract main-text tables from the W-isotope PDF folder and write one XLSX per PDF.

This script is designed to reproduce the table-cleaning workflow used in the
current Codex run: no CSV, no metadata columns, one worksheet per table, with
caption/title rows, headers, grouped rows, notes, and readable Excel formatting.

Requirements:
  - Python 3.9+ standard library only.
  - `pdftotext` must be available on PATH. On Windows, MiKTeX/Xpdf/Poppler can
    provide it.

Usage:
  python pdf_tables_to_xlsx.py "C:\Users\Spica\Desktop\W同位素 表生"

Optional:
  python pdf_tables_to_xlsx.py INPUT_FOLDER --output OUTPUT_FOLDER --keep-text
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional
from xml.sax.saxutils import escape as xml_escape_base


PDFS = {
    "wasy": "Differential behavior of tungsten stable isotopes during sorption to Fe versus.pdf",
    "kurz2021": "kurzweil-et-al-2021-redox-control-on-the-tungsten-isotope-composition-of-seawater.pdf",
    "alam": "Palaeoredox reconstruction in the eastern Arabian Sea since.pdf",
    "yang": "Stable tungsten isotope systematics on the Earth’s surface.pdf",
    "kurz2022": "The stable tungsten isotope composition.pdf",
}


@dataclass
class Row:
    cells: List[str]
    kind: str = "body"


@dataclass
class Sheet:
    name: str
    rows: List[Row]


@dataclass
class WorkbookSpec:
    pdf_name: str
    sheets: List[Sheet]


def clean_text(value: object) -> str:
    text = str(value or "")
    return (
        text.replace("\x03", "×")
        .replace("\x04", "~")
        .replace("\x01", "-")
        .replace("ﬁ", "fi")
        .replace("ﬂ", "fl")
        .replace("\f", "")
        .strip()
    )


def split_columns(line: str) -> List[str]:
    clean = clean_text(line)
    if not clean:
        return []
    parts = [p for p in re.split(r"\s{2,}", clean) if p]
    if len(parts) <= 1:
        parts = [p for p in re.split(r"\s+", clean) if p]
    return [clean_text(p) for p in parts]


def xml_escape(value: object) -> str:
    return xml_escape_base(clean_text(value), {"\"": "&quot;"})


def run_pdftotext(pdf: Path, output: Path, *, first: Optional[int] = None, last: Optional[int] = None) -> None:
    cmd = ["pdftotext", "-layout", "-enc", "UTF-8"]
    if first is not None:
        cmd.extend(["-f", str(first)])
    if last is not None:
        cmd.extend(["-l", str(last)])
    cmd.extend([str(pdf), str(output)])
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    except FileNotFoundError as exc:
        raise RuntimeError("未找到 pdftotext。请先安装 Poppler/Xpdf/MiKTeX，并确保 pdftotext 在 PATH 中。") from exc
    except subprocess.CalledProcessError as exc:
        msg = exc.stderr.strip() or exc.stdout.strip() or str(exc)
        raise RuntimeError(f"pdftotext 抽取失败：{pdf.name}\n{msg}") from exc


def read_lines(path: Path) -> List[str]:
    return path.read_text(encoding="utf-8", errors="replace").splitlines()


def find_line(lines: List[str], pattern: str, start: int = 0) -> int:
    regex = re.compile(pattern)
    for i in range(start, len(lines)):
        if regex.search(clean_text(lines[i])):
            return i
    raise ValueError(f"Cannot find line matching: {pattern}")


def row_from_raw(raw: str, caption: Optional[str] = None) -> Row:
    raw = clean_text(raw)
    if not raw:
        return Row([], "blank")
    if re.match(r"^Table\s", raw, re.I):
        return Row([raw], "title")
    if caption and raw == caption:
        return Row([raw], "caption")
    if re.match(r"^(A dash indicates|nm:|Computed as|The larger of two|The instrumental settings|Procedures of|Elemental concentration)", raw, re.I):
        kind = "caption" if re.match(r"^(Elemental|Procedures|The instrumental)", raw) else "note"
        return Row([raw], kind)
    if re.match(r"^[ab]$", raw):
        return Row([raw], "note")
    if re.match(r"^(Ferrihydrite|Birnessite|Expt\.)", raw, re.I):
        return Row([raw], "group")
    cells = split_columns(raw)
    first = cells[0] if cells else ""
    if re.match(r"^(Experiment|Parameters|Core Depth|\(m\)|Step|RF power|Selection of cones|Inlet system|Nebulizer|Cooling gas|Sample gas|Auxiliary gas|Aridus|Sensitivity|Number of cycles|Sample uptake time|Integration time|Washout time|Cup configuration)$", first, re.I):
        return Row(cells, "header")
    if re.match(r"^[A-Za-z]+\d|^\d|^Mn|^Fe|^Column|^Condition|^Load|^Rinse|^Collect", first):
        return Row(cells, "data")
    return Row(cells, "body")


def range_rows(lines: List[str], start: int, end: int, *, caption: Optional[str] = None, left_width: Optional[int] = None, keep=None) -> List[Row]:
    rows: List[Row] = []
    for i in range(start, min(end, len(lines))):
        raw = lines[i]
        if left_width is not None:
            raw = raw[:left_width]
        raw = clean_text(raw)
        if not raw:
            continue
        if keep and not keep(raw):
            continue
        rows.append(row_from_raw(raw, caption))
    return rows


def table_range(lines: List[str], table_id: str, next_table: Optional[str] = None) -> tuple[int, int]:
    start = find_line(lines, rf"^\s*{re.escape(table_id)}\b")
    if next_table:
        end = find_line(lines, rf"^\s*{re.escape(next_table)}\b", start + 1)
    else:
        end = len(lines)
    return start, end


def make_wasy_tables(lines: List[str]) -> List[Sheet]:
    captions = {
        "Table 1": "Parameters and results from fixed-duration tungsten adsorption experiments with ferrihydrite.",
        "Table 2": "Parameters and results from fixed-duration tungsten adsorption experiments with birnessite.",
        "Table 3": "Extents of sorption from time-series experiments.",
        "Table 4": "Isotope results for birnessite pH 5 time-series experiments (δ183/182Wstock = −0.27 ± 0.16‰, other parameters in Table 3).",
    }

    t1_start, t2_start = table_range(lines, "Table 1", "Table 2")
    t2_start, t3_start = table_range(lines, "Table 2", "Table 3")
    t3_start, t4_start = table_range(lines, "Table 3", "Table 4")
    t4_start = find_line(lines, r"^\s*Table 4\b")

    def stop_after_note_b(start: int) -> int:
        seen_b = False
        for i in range(start, min(len(lines), start + 80)):
            raw = clean_text(lines[i])
            if raw == "b":
                seen_b = True
            elif seen_b and raw.startswith("Computed as"):
                return i + 1
        return min(len(lines), start + 80)

    def keep_table3(raw: str) -> bool:
        if re.match(r"^\d+$", raw):
            return False
        noise = r"^(sorption continued|more slowly|shows a continuous|components|preferentially|solved component|For ferrihydrite|where|average fractionation|Fractionation|182|Wdissolved|removed lighter|from \+|greatest|even a larger|experiments with|of Δ|nations|To confirm|isotopic mass|δ183|to the product|3\.2\.|3\.3\.)"
        return not re.match(noise, raw, re.I)

    t1_end = stop_after_note_b(t1_start)
    t2_end = stop_after_note_b(t2_start)
    t3_end = t4_start
    for i in range(t3_start, t4_start):
        if clean_text(lines[i]).startswith("Mn5-TS-504"):
            t3_end = i + 1
            break

    return [
        Sheet("Table 1", range_rows(lines, t1_start, t1_end, caption=captions["Table 1"])),
        Sheet("Table 2", range_rows(lines, t2_start, t2_end, caption=captions["Table 2"])),
        Sheet("Table 3", range_rows(lines, t3_start, t3_end, caption=captions["Table 3"], left_width=100, keep=keep_table3)),
        Sheet("Table 4", range_rows(lines, t4_start, stop_after_note_b(t4_start), caption=captions["Table 4"])),
    ]


def make_alam_table(layout_lines: List[str], fallback_lines: List[str]) -> List[Sheet]:
    source = layout_lines or fallback_lines
    table_idx = find_line(source, r"^\s*Table 1\s*$")
    header_idx = find_line(source, r"Core Depth\s+Age\s+Al", table_idx)
    unit_idx = find_line(source, r"\(m\)\s+\(Ma\)", header_idx)
    rows = [
        Row(["Table 1"], "title"),
        Row(["Elemental concentration and isotopic data of sediment samples from IODP Site U1457 in the eastern Arabian Sea."], "caption"),
        Row(split_columns(source[header_idx]), "header"),
        Row(split_columns(source[unit_idx]), "header"),
    ]
    for i in range(unit_idx + 1, min(len(source), unit_idx + 120)):
        raw = clean_text(source[i])
        if "continued on next page" in raw:
            break
        cells = split_columns(raw)
        if cells and re.match(r"^-?\d+(\.\d+)?$", cells[0]) and len(cells) >= 15:
            rows.append(Row(cells, "data"))
    rows.append(Row([""], "spacer"))
    rows.append(Row(["nm: not measured."], "note"))
    return [Sheet("Table 1", rows)]


def make_yang_tables(lines: List[str]) -> List[Sheet]:
    # Table 1 is embedded in the right column of a two-column page. The slice
    # position mirrors the current manually verified extraction.
    t1_idx = find_line(lines, r"Table 1")
    table1_rows = [
        Row(["Table 1"], "title"),
        Row(["Procedures of column chemistry for the separation of W."], "caption"),
        Row(["Step", "Volumes", "Reagent"], "header"),
    ]
    for i in range(t1_idx + 3, min(len(lines), t1_idx + 24)):
        right = clean_text(lines[i][60:]).lstrip("- ")
        cells = split_columns(right)
        if cells and re.match(r"^(Column clean|Condition|Load sample|Rinse matrix|Collect W)", cells[0]):
            table1_rows.append(Row(cells, "data"))

    t2_idx = find_line(lines, r"^\s*Table 2\s*$")
    table2_rows = [
        Row(["Table 2"], "title"),
        Row(["The instrumental settings, data acquisition parameters, and cup configuration for MC-ICPMS analysis."], "caption"),
    ]
    for i in range(t2_idx + 2, min(len(lines), t2_idx + 24)):
        raw = clean_text(lines[i])
        if not raw:
            continue
        table2_rows.append(row_from_raw(raw))
        if raw.endswith("Os"):
            break

    return [Sheet("Table 1", table1_rows), Sheet("Table 2", table2_rows)]


def col_name(n: int) -> str:
    name = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(value: str, row_idx: int, col_idx: int, style: int) -> str:
    ref = f"{col_name(col_idx)}{row_idx}"
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{xml_escape(value)}</t></is></c>'


def sheet_xml(sheet: Sheet) -> str:
    max_cols = max(1, *(len(row.cells) for row in sheet.rows))
    col_xml = []
    for i in range(max_cols):
        width = max(10, min(28, max((len(clean_text(row.cells[i])) + 2 if i < len(row.cells) else 10) for row in sheet.rows)))
        col_xml.append(f'<col min="{i + 1}" max="{i + 1}" width="{width}" customWidth="1"/>')

    merges = []
    row_xml = []
    for r, row in enumerate(sheet.rows, 1):
        if row.kind == "header":
            style = 2
        elif row.kind in {"title", "caption"}:
            style = 1
        elif row.kind in {"group", "note"}:
            style = 3
        else:
            style = 0
        if row.kind in {"title", "caption", "group", "note"} and len(row.cells) == 1 and max_cols > 1:
            merges.append(f'<mergeCell ref="A{r}:{col_name(max_cols)}{r}"/>')
        cells = "".join(cell_xml(v, r, c + 1, style) for c, v in enumerate(row.cells))
        row_xml.append(f'<row r="{r}">{cells}</row>')

    merge_xml = f'<mergeCells count="{len(merges)}">{"".join(merges)}</mergeCells>' if merges else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/>'
        f'<cols>{"".join(col_xml)}</cols><sheetData>{"".join(row_xml)}</sheetData>{merge_xml}</worksheet>'
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="3"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><i/><sz val="10"/><name val="Calibri"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFEFEF"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    )


def write_xlsx(spec: WorkbookSpec, output_file: Path) -> None:
    sheets = spec.sheets
    workbook_sheets = "".join(
        f'<sheet name="{xml_escape(sheet.name)[:31]}" sheetId="{i}" r:id="rId{i}"/>' for i, sheet in enumerate(sheets, 1)
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{workbook_sheets}</sheets></workbook>'
    )
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(sheets) + 1)
    )
    rels += f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    workbook_rels = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>'
    sheet_types = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, len(sheets) + 1)
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f'{sheet_types}</Types>'
    )

    with zipfile.ZipFile(output_file, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles_xml())
        for i, sheet in enumerate(sheets, 1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(sheet))


def build_workbooks(text_dir: Path, input_dir: Path) -> List[WorkbookSpec]:
    text = {key: read_lines(text_dir / PDFS[key].replace(".pdf", ".txt")) for key in PDFS}
    alam_layout_file = text_dir / "Palaeoredox reconstruction in the eastern Arabian Sea since.pages_5_9_layout.txt"
    alam_layout = read_lines(alam_layout_file) if alam_layout_file.exists() else []
    return [
        WorkbookSpec(PDFS["wasy"], make_wasy_tables(text["wasy"])),
        WorkbookSpec(PDFS["kurz2021"], [Sheet("No tables", [Row(["No main-text tables detected."], "note")])]),
        WorkbookSpec(PDFS["alam"], make_alam_table(alam_layout, text["alam"])),
        WorkbookSpec(PDFS["yang"], make_yang_tables(text["yang"])),
        WorkbookSpec(PDFS["kurz2022"], [Sheet("No tables", [Row(["No main-text tables detected."], "note")])]),
    ]


def extract_all_text(input_dir: Path, text_dir: Path) -> None:
    for pdf_name in PDFS.values():
        pdf = input_dir / pdf_name
        if not pdf.exists():
            raise FileNotFoundError(f"缺少 PDF：{pdf}")
        run_pdftotext(pdf, text_dir / pdf_name.replace(".pdf", ".txt"))
    alam_pdf = input_dir / PDFS["alam"]
    run_pdftotext(alam_pdf, text_dir / "Palaeoredox reconstruction in the eastern Arabian Sea since.pages_5_9_layout.txt", first=5, last=9)


def write_manifest(output_dir: Path, specs: Iterable[WorkbookSpec]) -> None:
    manifest = []
    for spec in specs:
        manifest.append({
            "pdf": spec.pdf_name,
            "xlsx": spec.pdf_name.replace(".pdf", ".xlsx"),
            "sheets": [
                {
                    "name": sheet.name,
                    "rows": len(sheet.rows),
                    "cols": max(1, *(len(row.cells) for row in sheet.rows)),
                }
                for sheet in spec.sheets
            ],
        })
    (output_dir / "extraction_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract W-isotope PDF tables into clean XLSX workbooks.")
    parser.add_argument("input_folder", type=Path, help="Folder containing the five PDF files.")
    parser.add_argument("--output", type=Path, default=None, help="Output folder. Defaults to input_folder.")
    parser.add_argument("--text-dir", type=Path, default=None, help="Optional folder for extracted text cache.")
    parser.add_argument("--skip-extract", action="store_true", help="Use an existing --text-dir instead of running pdftotext.")
    parser.add_argument("--keep-text", action="store_true", help="Keep temporary extracted text files when --text-dir is not set.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_dir = args.input_folder.expanduser().resolve()
    output_dir = (args.output or input_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    temp_obj = None
    if args.text_dir:
        text_dir = args.text_dir.expanduser().resolve()
        text_dir.mkdir(parents=True, exist_ok=True)
    else:
        temp_obj = tempfile.TemporaryDirectory(prefix="w_isotope_text_")
        text_dir = Path(temp_obj.name)

    try:
        if not args.skip_extract:
            extract_all_text(input_dir, text_dir)
        specs = build_workbooks(text_dir, input_dir)
        for spec in specs:
            write_xlsx(spec, output_dir / spec.pdf_name.replace(".pdf", ".xlsx"))
        write_manifest(output_dir, specs)
        if args.keep_text and temp_obj:
            saved = output_dir / "_extracted_text_cache"
            if saved.exists():
                shutil.rmtree(saved)
            shutil.copytree(text_dir, saved)
        print(f"Done. XLSX files written to: {output_dir}")
        for spec in specs:
            print(f"- {spec.pdf_name}: {', '.join(sheet.name for sheet in spec.sheets)}")
        return 0
    finally:
        if temp_obj:
            temp_obj.cleanup()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
