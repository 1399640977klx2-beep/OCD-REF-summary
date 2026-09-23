"""Build Excel-native Match charts from a Match worksheet.

The output is a new workbook containing:
    * a copy of the selected source sheet (values, with formula fallback)
    * a "Match Charts" sheet with the standard 3-chart-per-parameter set
    * an optional "Custom Charts" sheet with user-selected charts

This avoids modifying the source workbook because openpyxl does not preserve
existing Excel charts when an existing workbook is loaded and saved.
"""
from dataclasses import dataclass, field
from copy import copy
from pathlib import Path
import re
import warnings

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, ScatterChart, Series, Reference
from openpyxl.chart.data_source import AxDataSource, MultiLevelStrRef, StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.chart.trendline import Trendline
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties, Font as DrawingFont, Paragraph,
    ParagraphProperties, RichTextProperties,
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string

DEFAULT_END_ROW = 3000
DEFAULT_CATEGORY_COLS = (10, 11, 12, 13)
SERIES_COLORS = ("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47")


@dataclass
class ParamInfo:
    name: str
    raw_col: int = None
    pmish_col: int = None
    ref_col: int = None
    bias_cols: dict = field(default_factory=dict)


@dataclass
class MatchAnalysis:
    input_path: str
    sheet_name: str
    header_row: int
    data_start_row: int
    end_row: int
    category_cols: tuple
    headers: dict
    parameters: list
    warnings: list


def _norm(value):
    return str(value).strip().lower()


def _compact(value):
    return "".join(ch for ch in _norm(value) if ch.isalnum())


def _is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def _load_workbooks(path):
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Unable to read chart.*")
        values_wb = load_workbook(path, data_only=True)
        formula_wb = load_workbook(path, data_only=False)
    return values_wb, formula_wb


def list_sheets(input_path):
    values_wb, _ = _load_workbooks(input_path)
    return values_wb.sheetnames


def _find_header_row(ws):
    limit = min(ws.max_row, 100)
    for row in range(1, limit + 1):
        values = {_norm(ws.cell(row, col).value)
                  for col in range(1, min(ws.max_column, 200) + 1)}
        has_wafer = "wafer id" in values or "waferid" in values
        has_tool = "tool" in values
        has_product = "product" in values or "productname" in values
        if has_wafer and has_tool and has_product:
            return row
    raise ValueError("未找到 Match 表头行：需要包含 Wafer ID、Tool、Product")


def _resolve_header_formula(formula, header_map):
    if not _is_formula(formula):
        return None
    compact = formula.replace(" ", "")
    match = re.fullmatch(r"=([A-Z]{1,3})(\d+)&\"([^\"]+)\"", compact)
    if match:
        col = column_index_from_string(match.group(1))
        suffix = match.group(3)
        base = header_map.get(col)
        if base:
            return f"{base}{suffix}"
    match = re.fullmatch(r"=([A-Z]{1,3})(\d+)", compact)
    if match:
        return header_map.get(column_index_from_string(match.group(1)))
    return None


def _build_headers(values_ws, formula_ws, header_row):
    max_col = max(values_ws.max_column, formula_ws.max_column)
    headers = {col: None for col in range(1, max_col + 1)}
    for col in range(1, max_col + 1):
        value = values_ws.cell(header_row, col).value
        if value is not None and str(value).strip():
            headers[col] = str(value).strip()
    for col in range(1, max_col + 1):
        if headers[col] is None:
            headers[col] = _resolve_header_formula(
                formula_ws.cell(header_row, col).value, headers)
    return headers


def _find_category_cols(headers, formula_ws, header_row):
    tool_cols = [
        col for col, name in headers.items() if name and _norm(name) == "tool"
    ]
    product_cols = [
        col for col, name in headers.items()
        if name and _norm(name) in ("product", "productname")
    ]
    anchor = None
    if tool_cols:
        anchor = min(tool_cols)
    elif product_cols:
        anchor = min(product_cols)

    wafer_cols = [
        col for col, name in headers.items()
        if name and _norm(name) in ("wafer id", "waferid")
        and (anchor is None or col < anchor)
    ]
    if wafer_cols:
        wafer_col = max(wafer_cols)
        if wafer_col - 4 >= 1:
            candidate = tuple(range(wafer_col - 4, wafer_col))
            if len(candidate) == 4:
                return candidate

    max_col = formula_ws.max_column
    for start in range(1, max_col - 3):
        formulas = [
            formula_ws.cell(header_row + 1, start + i).value
            for i in range(4)
        ]
        if not all(_is_formula(value) for value in formulas):
            continue
        joined = " ".join(str(value).lower() for value in formulas)
        if "if(" in joined and "left(" in joined and "right(" in joined:
            return tuple(range(start, start + 4))
    return DEFAULT_CATEGORY_COLS


def _bias_variant(header):
    lower = header.lower()
    idx = lower.find("bias")
    tail = header[idx + 4:].strip("_ ") if idx >= 0 else header
    compact = _compact(tail)
    if "addoffset" in compact:
        return "Addoffset"
    if "raw" in compact:
        return "Raw"
    if "old" in compact:
        return "Old"
    if not compact:
        return "Bias"
    return tail or "Bias"


def _find_parameters(headers):
    by_key = {}
    order = []
    for col in sorted(headers):
        header = headers[col]
        if not header:
            continue
        norm = _norm(header)
        if norm.endswith("_pmish"):
            name = header[:-6].strip()
            key = _norm(name)
            if key not in by_key:
                by_key[key] = ParamInfo(name=name)
                order.append(key)
            by_key[key].pmish_col = col

    for col in sorted(headers):
        header = headers[col]
        if not header:
            continue
        norm = _norm(header)
        key = _norm(header)
        if key in by_key:
            pmish_col = by_key[key].pmish_col
            if pmish_col is not None and col < pmish_col:
                current = by_key[key].raw_col
                if current is None or col > current:
                    by_key[key].raw_col = col
        if norm.endswith("_ref"):
            name = header[:-4].strip()
            key = _norm(name)
            if key in by_key:
                pmish_col = by_key[key].pmish_col
                if pmish_col is None or col > pmish_col:
                    current = by_key[key].ref_col
                    if current is None or col < current:
                        by_key[key].ref_col = col
        if "bias" in norm:
            name = re.split(r"bias", header, maxsplit=1, flags=re.I)[0]
            key = _norm(name.strip("_ "))
            if key in by_key:
                variant = _bias_variant(header)
                by_key[key].bias_cols[variant] = col

    parameters = [by_key[key] for key in order]
    return parameters


def analyze_match_workbook(input_path, sheet_name=None, end_row=DEFAULT_END_ROW):
    input_path = str(input_path)
    values_wb, formula_wb = _load_workbooks(input_path)
    if sheet_name is None:
        sheet_name = values_wb.sheetnames[0]
    if sheet_name not in values_wb.sheetnames:
        raise ValueError(f"工作簿中不存在 sheet: {sheet_name}")

    values_ws = values_wb[sheet_name]
    formula_ws = formula_wb[sheet_name]
    header_row = _find_header_row(values_ws)
    data_start_row = header_row + 1
    end_row = int(end_row)
    if end_row < data_start_row:
        raise ValueError(
            f"绘图结束行 {end_row} 小于数据起始行 {data_start_row}")

    headers = _build_headers(values_ws, formula_ws, header_row)
    category_cols = _find_category_cols(headers, formula_ws, header_row)
    parameters = _find_parameters(headers)
    if not parameters:
        raise ValueError("未识别到任何带 _PMISH 后缀的参数列")

    warnings_list = []
    for param in parameters:
        if param.pmish_col is None:
            warnings_list.append(f"{param.name}: 缺少 _PMISH 列")
        if param.ref_col is None:
            warnings_list.append(f"{param.name}: 缺少 _REF 列")
        if param.raw_col is None:
            warnings_list.append(
                f"{param.name}: 缺少原始 PMISH 列，散点图将使用 PMISH 列代替")

    all_variants = []
    for param in parameters:
        for variant in param.bias_cols:
            if variant not in all_variants:
                all_variants.append(variant)
    for variant in all_variants:
        missing = [p.name for p in parameters if variant not in p.bias_cols]
        if missing:
            warnings_list.append(
                f"Bias 类型 {variant} 在以下参数中缺失: {', '.join(missing)}")

    return MatchAnalysis(
        input_path=input_path,
        sheet_name=sheet_name,
        header_row=header_row,
        data_start_row=data_start_row,
        end_row=end_row,
        category_cols=category_cols,
        headers=headers,
        parameters=parameters,
        warnings=warnings_list,
    )


def _multi_level_category(data_ws, category_cols, start_row, end_row):
    ref = Reference(
        data_ws,
        min_col=min(category_cols),
        max_col=max(category_cols),
        min_row=start_row,
        max_row=end_row,
    )
    return AxDataSource(multiLvlStrRef=MultiLevelStrRef(f=str(ref)))


def _single_category(data_ws, col, start_row, end_row):
    ref = Reference(
        data_ws, min_col=col, min_row=start_row, max_row=end_row)
    return AxDataSource(strRef=StrRef(f=str(ref)))


def _style_line_series(series, color="4472C4"):
    series.graphicalProperties.line.width = 28575
    series.graphicalProperties.line.solidFill = color
    series.marker.symbol = "circle"
    series.marker.size = 5
    series.marker.graphicalProperties.solidFill = color
    series.marker.graphicalProperties.line.solidFill = color


def _text_properties(size, bold=False):
    font = DrawingFont(typeface="Times New Roman")
    return CharacterProperties(
        sz=size, b=bold, latin=font, ea=font, cs=font)


def _make_text_properties(size):
    props = _text_properties(size)
    return RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(
            pPr=ParagraphProperties(defRPr=props),
            endParaRPr=props,
        )],
    )


def _set_title_font(title, size=1400):
    if title is None:
        return
    props = _text_properties(size, bold=False)
    try:
        for paragraph in title.tx.rich.p:
            if paragraph.pPr is None:
                paragraph.pPr = ParagraphProperties()
            paragraph.pPr.defRPr = props
            paragraph.endParaRPr = props
            for run in paragraph.r or []:
                run.rPr = props
    except Exception:
        pass


def _apply_chart_fonts(chart):
    _set_title_font(chart.title, 1400)
    for axis in (getattr(chart, "x_axis", None),
                 getattr(chart, "y_axis", None)):
        if axis is None:
            continue
        _set_title_font(axis.title, 1200)
        axis.txPr = _make_text_properties(1200)
    if getattr(chart, "legend", None) is not None:
        chart.legend.txPr = _make_text_properties(1200)


def _create_scatter_chart(data_ws, start_row, end_row,
                          x_col, y_col, title):
    chart = ScatterChart()
    chart.title = title
    chart.style = 2
    chart.x_axis.title = "PMISH"
    chart.y_axis.title = "REF"
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend = None
    chart.height = 9
    chart.width = 15

    series = Series(
        Reference(
            data_ws, min_col=y_col, min_row=start_row, max_row=end_row),
        xvalues=Reference(
            data_ws, min_col=x_col, min_row=start_row, max_row=end_row),
    )
    color = SERIES_COLORS[0]
    series.marker.symbol = "circle"
    series.marker.size = 5
    series.marker.graphicalProperties.solidFill = color
    series.marker.graphicalProperties.line.solidFill = color
    series.graphicalProperties.line.noFill = True
    series.trendline = Trendline(
        trendlineType="linear", dispRSqr=True, dispEq=True,
        spPr=GraphicalProperties(
            ln=LineProperties(
                w=19050, solidFill=color, prstDash="sysDot")),
    )
    chart.series.append(series)
    _apply_chart_fonts(chart)
    return chart


def _create_compare_line_chart(data_ws, start_row, end_row,
                               category_cols, pmish_col, ref_col,
                               param_name, headers):
    chart = LineChart()
    chart.title = param_name
    chart.style = 2
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.height = 9
    chart.width = 15

    categories = _multi_level_category(
        data_ws, category_cols, start_row, end_row)
    for idx, col in enumerate((pmish_col, ref_col)):
        series = Series(
            Reference(data_ws, min_col=col, min_row=start_row, max_row=end_row),
            title=str(headers.get(col) or get_column_letter(col)),
        )
        series.cat = categories
        _style_line_series(series, SERIES_COLORS[idx % len(SERIES_COLORS)])
        chart.series.append(series)
    _apply_chart_fonts(chart)
    return chart


def _create_bias_line_chart(data_ws, start_row, end_row,
                            category_cols, bias_columns, param_name, headers):
    chart = LineChart()
    chart.title = f"{param_name}_Bias"
    chart.style = 2
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.height = 9
    chart.width = 15

    categories = _multi_level_category(
        data_ws, category_cols, start_row, end_row)
    for idx, col in enumerate(bias_columns):
        series = Series(
            Reference(data_ws, min_col=col, min_row=start_row, max_row=end_row),
            title=str(headers.get(col) or get_column_letter(col)),
        )
        series.cat = categories
        _style_line_series(series, SERIES_COLORS[idx % len(SERIES_COLORS)])
        chart.series.append(series)
    _apply_chart_fonts(chart)
    return chart


def _create_custom_scatter(data_ws, start_row, end_row,
                           x_col, y_col, x_name, y_name):
    chart = ScatterChart()
    chart.title = f"{x_name}_vs_{y_name}"
    chart.style = 2
    chart.x_axis.title = x_name
    chart.y_axis.title = y_name
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend = None
    chart.height = 9
    chart.width = 15

    series = Series(
        Reference(
            data_ws, min_col=y_col, min_row=start_row, max_row=end_row),
        xvalues=Reference(
            data_ws, min_col=x_col, min_row=start_row, max_row=end_row),
    )
    color = SERIES_COLORS[0]
    series.marker.symbol = "circle"
    series.marker.size = 5
    series.marker.graphicalProperties.solidFill = color
    series.marker.graphicalProperties.line.solidFill = color
    series.graphicalProperties.line.noFill = True
    chart.series.append(series)
    _apply_chart_fonts(chart)
    return chart


def _create_custom_line(data_ws, start_row, end_row,
                        columns, label_col, category_cols, title, headers):
    chart = LineChart()
    chart.title = title
    chart.style = 2
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.height = 9
    chart.width = 15

    if label_col is None:
        categories = _multi_level_category(
            data_ws, category_cols, start_row, end_row)
    else:
        categories = _single_category(
            data_ws, label_col, start_row, end_row)
    for idx, col in enumerate(columns):
        series = Series(
            Reference(data_ws, min_col=col, min_row=start_row, max_row=end_row),
            title=str(headers.get(col) or get_column_letter(col)),
        )
        series.cat = categories
        _style_line_series(series, SERIES_COLORS[idx % len(SERIES_COLORS)])
        chart.series.append(series)
    _apply_chart_fonts(chart)
    return chart


def _safe_sheet_title(title):
    clean = re.sub(r"[\[\]:*?/\\]", "_", str(title)).strip()
    return (clean or "Data")[:31]


def _copy_sheet(src_ws, out_wb):
    dst_ws = out_wb.create_sheet(_safe_sheet_title(src_ws.title))
    style_cache = {}
    for src_row in src_ws.iter_rows():
        for src_cell in src_row:
            if src_cell.value is None:
                continue
            dst_cell = dst_ws.cell(
                row=src_cell.row, column=src_cell.column,
                value=src_cell.value)
            if src_cell.has_style:
                dst_cell.number_format = src_cell.number_format
                key = id(src_cell._style)
                cached = style_cache.get(key)
                if cached is None:
                    cached = (
                        copy(src_cell.font),
                        copy(src_cell.fill),
                        copy(src_cell.border),
                        copy(src_cell.alignment),
                        copy(src_cell.protection),
                    )
                    style_cache[key] = cached
                (dst_cell.font, dst_cell.fill, dst_cell.border,
                 dst_cell.alignment, dst_cell.protection) = cached

    for key, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            dst_ws.column_dimensions[key].width = dim.width
        if dim.hidden:
            dst_ws.column_dimensions[key].hidden = True
    for idx, dim in src_ws.row_dimensions.items():
        if dim.height is not None:
            dst_ws.row_dimensions[idx].height = dim.height
        if dim.hidden:
            dst_ws.row_dimensions[idx].hidden = True
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes
    if src_ws.auto_filter.ref:
        dst_ws.auto_filter.ref = src_ws.auto_filter.ref
    return dst_ws


def _copy_all_sheets(formula_wb, out_wb, selected_sheet):
    out_wb.remove(out_wb.active)
    copied = {}
    for src_ws in formula_wb.worksheets:
        copied[src_ws.title] = _copy_sheet(src_ws, out_wb)
    return copied[selected_sheet]


def _add_standard_charts(data_ws, chart_ws, analysis, bias_selection):
    generated = 0
    start_row = analysis.data_start_row
    end_row = analysis.end_row
    headers = analysis.headers

    for idx, param in enumerate(analysis.parameters):
        column_anchor = 1 + idx * 16
        chart_col = get_column_letter(column_anchor)

        if param.ref_col and (param.raw_col or param.pmish_col):
            chart = _create_scatter_chart(
                data_ws, start_row, end_row,
                param.raw_col or param.pmish_col, param.ref_col, param.name)
            chart_ws.add_chart(chart, f"{chart_col}1")
            generated += 1

        if param.pmish_col and param.ref_col:
            chart = _create_compare_line_chart(
                data_ws, start_row, end_row, analysis.category_cols,
                param.pmish_col, param.ref_col, param.name, headers)
            chart_ws.add_chart(chart, f"{chart_col}18")
            generated += 1

        selected = bias_selection.get(param.name, [])
        selected_cols = [
            param.bias_cols[v] for v in selected if v in param.bias_cols
        ]
        if selected_cols:
            chart = _create_bias_line_chart(
                data_ws, start_row, end_row, analysis.category_cols,
                selected_cols, param.name, headers)
            chart_ws.add_chart(chart, f"{chart_col}35")
            generated += 1
    return generated


def _add_custom_charts(data_ws, chart_ws, analysis, custom_charts):
    generated = 0
    start_row = analysis.data_start_row
    end_row = analysis.end_row
    headers = analysis.headers

    for idx, item in enumerate(custom_charts.get("scatter", [])):
        x_col = item["x_col"]
        y_col = item["y_col"]
        x_name = str(headers.get(x_col) or get_column_letter(x_col))
        y_name = str(headers.get(y_col) or get_column_letter(y_col))
        chart = _create_custom_scatter(
            data_ws, start_row, end_row, x_col, y_col, x_name, y_name)
        chart_ws.add_chart(
            chart, f"{get_column_letter(1 + (idx % 2) * 16)}"
                   f"{1 + (idx // 2) * 18}")
        generated += 1

    scatter_count = len(custom_charts.get("scatter", []))
    for idx, item in enumerate(custom_charts.get("line", [])):
        columns = item.get("columns", [])
        if not columns:
            continue
        chart = _create_custom_line(
            data_ws, start_row, end_row, columns,
            item.get("label_col"), analysis.category_cols,
            item.get("title", "Custom_Line"), headers)
        chart_ws.add_chart(
            chart, f"{get_column_letter(1 + ((scatter_count + idx) % 2) * 16)}"
                   f"{1 + ((scatter_count + idx) // 2) * 18}")
        generated += 1
    return generated


def generate_match_charts(input_path, output_path, sheet_name=None,
                          end_row=DEFAULT_END_ROW, bias_selection=None,
                          custom_charts=None):
    """Generate a new workbook containing data and native Excel charts."""
    analysis = analyze_match_workbook(input_path, sheet_name, end_row)
    warnings_list = list(analysis.warnings)
    _, formula_wb = _load_workbooks(input_path)

    if bias_selection is None:
        bias_selection = {
            p.name: list(p.bias_cols.keys()) for p in analysis.parameters
        }
    for param in analysis.parameters:
        for variant in bias_selection.get(param.name, []):
            if variant not in param.bias_cols:
                warnings_list.append(
                    f"{param.name}: 选择的 Bias 类型 {variant} 不存在，已跳过")
    custom_charts = custom_charts or {}

    out_wb = Workbook()
    out_wb.calculation.fullCalcOnLoad = True
    data_ws = _copy_all_sheets(
        formula_wb, out_wb, analysis.sheet_name)
    chart_ws = out_wb.create_sheet("Match Charts")
    standard_count = _add_standard_charts(
        data_ws, chart_ws, analysis, bias_selection)

    custom_ws = None
    custom_count = 0
    if custom_charts.get("scatter") or custom_charts.get("line"):
        custom_ws = out_wb.create_sheet("Custom Charts")
        custom_count = _add_custom_charts(
            data_ws, custom_ws, analysis, custom_charts)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(output_path)
    return {
        "output_path": str(output_path),
        "analysis": analysis,
        "warnings": warnings_list,
        "standard_charts": standard_count,
        "custom_charts": custom_count,
    }
