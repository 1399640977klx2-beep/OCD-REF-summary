"""Per-Tool/Product/Wafer pivot summary for OCD pasted Excel data.

Standalone script; it does not import other modules in this project.

Run:
    python utils/pivot_summary.py                          # 使用默认 Test1.xlsx
    python utils/pivot_summary.py <input.xlsx> <output.xlsx>
    python utils/pivot_summary.py <input.xlsx> <output.xlsx> --sheet Sheet1

The input may contain extra rows above the real header. The script finds the
header row by looking for the cell names ``Die Seq`` and ``Wafer ID``
(case-insensitive) and ignores everything above it.

Output layout (one row per unique Tool / Product / Wafer combination):
    Tool_label, Product_label, Wafer_label,   # 同上一行时留空，用于图表轴标签
    Tool, Product, Wafer,
    <param>_PMISH_Mean, <param>_REF_Mean,
    <param>_Bias_Mean, <param>_oldBias_Mean,
    <param>_PMISH_STD, <param>_REF_STD,
    <空列>, ... repeated for every parameter
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

DEFAULT_PARAMETERS = ["CD_Bot", "CD_Top", "HIGH", "SPA", "THK"]

DEFAULT_INPUT_PATH = (
    Path(__file__).resolve().parents[1] / "Data" / "绘制数据透视表" / "Test1.xlsx"
)

OUTPUT_COLUMN_GROUPS = (
    ("PMISH_Mean", "REF_Mean"),
    ("Bias_Mean", "oldBias_Mean"),
    ("PMISH_STD", "REF_STD"),
)


def _norm(value):
    return str(value).strip().lower()


def _find_header_row(raw):
    """Return the 0-based row index of the real data header."""
    limit = min(len(raw), 200)
    for idx in range(limit):
        values = {_norm(v) for v in raw.iloc[idx].tolist()}
        if "die seq" in values and "wafer id" in values:
            return idx
    for idx in range(limit):
        values = {_norm(v) for v in raw.iloc[idx].tolist()}
        if "wafer id" in values:
            return idx
    raise ValueError("未找到数据表头行：需要包含 Die Seq 和 Wafer ID")


def _read_raw(path, sheet=None):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, header=None, dtype=str)
    return pd.read_excel(
        path, sheet_name=sheet if sheet is not None else 0,
        header=None, dtype=str,
    )


def load_dataframe(path, sheet=None):
    """Read pasted data and return a DataFrame with the real header applied."""
    raw = _read_raw(path, sheet)
    header_idx = _find_header_row(raw)
    header = [
        str(v).strip() if not pd.isna(v) else f"Unnamed_{i}"
        for i, v in enumerate(raw.iloc[header_idx])
    ]
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = header
    df = df.reset_index(drop=True)
    if df.empty:
        raise ValueError("表头下方没有数据")
    return df


def _find_group_columns(df):
    norm_to_orig = {}
    for col in df.columns:
        norm_to_orig.setdefault(_norm(col), col)

    wafer_col = norm_to_orig.get("wafer id") or norm_to_orig.get("waferid")
    tool_col = norm_to_orig.get("tool")
    product_col = norm_to_orig.get("product")

    missing = []
    if tool_col is None:
        missing.append("Tool")
    if product_col is None:
        missing.append("Product")
    if wafer_col is None:
        missing.append("Wafer ID")
    if missing:
        raise ValueError("缺少分组列: " + ", ".join(missing))
    return tool_col, product_col, wafer_col


def _classify_suffix(suffix):
    compact = "".join(ch for ch in _norm(suffix) if ch.isalnum())
    if compact == "pmish":
        return "pmish"
    if compact == "ref":
        return "ref"
    if compact == "bias":
        return "bias"
    if "addoffset" in compact:
        return "bias_adoffset"
    if "old" in compact:
        return "old_bias"
    return None


def _collect_param_columns(df, parameters):
    found = {
        p: {"pmish": None, "ref": None, "bias": None,
            "bias_adoffset": None, "old_bias": None}
        for p in parameters
    }
    for col in df.columns:
        norm = _norm(col)
        for p in parameters:
            prefix = _norm(p) + "_"
            if norm.startswith(prefix):
                kind = _classify_suffix(norm[len(prefix):])
                if kind:
                    found[p][kind] = col
    return found


def _resolve_param_columns(found, parameters):
    warnings = []
    resolved = {}
    for p in parameters:
        m = found[p]
        if m["pmish"] is None:
            warnings.append(f"缺少 {p}_PMISH 后缀")
        if m["ref"] is None:
            warnings.append(f"缺少 {p}_REF 后缀")

        bias_col = m["bias_adoffset"]
        if bias_col is None and m["bias"] is not None:
            bias_col = m["bias"]
            warnings.append(f"缺少 {p}_bias_Addoffset 后缀，改用 {m['bias']}")
        if bias_col is None:
            warnings.append(f"缺少 {p} 的 Bias 后缀（bias_Addoffset 或 Bias）")

        if m["old_bias"] is None:
            warnings.append(f"缺少 {p} 的 old Bias 后缀")

        resolved[p] = {
            "pmish": m["pmish"],
            "ref": m["ref"],
            "bias": bias_col,
            "old_bias": m["old_bias"],
        }
    return resolved, warnings


def _to_numeric(series):
    return pd.to_numeric(series, errors="coerce")


def _group_mean(group_df, col):
    if col is None:
        return ""
    try:
        return _empty_if_unavailable(_to_numeric(group_df[col]).mean())
    except Exception:
        return ""


def _group_std(group_df, col):
    if col is None:
        return ""
    try:
        return _empty_if_unavailable(_to_numeric(group_df[col]).std())
    except Exception:
        return ""


def _empty_if_unavailable(value):
    try:
        if value is None or pd.isna(value):
            return ""
    except (TypeError, ValueError):
        return ""
    return value


def _same_key(a, b):
    try:
        if pd.isna(a) and pd.isna(b):
            return True
    except (TypeError, ValueError):
        pass
    return a == b


def _build_output_columns(parameters):
    cols = [
        "Tool_label", "Product_label", "Wafer_label",
        "Tool", "Product", "Wafer",
    ]
    for idx, p in enumerate(parameters):
        for col1, col2 in OUTPUT_COLUMN_GROUPS:
            cols.append(f"{p}_{col1}")
            cols.append(f"{p}_{col2}")
        if idx < len(parameters) - 1:
            cols.append("")
    return cols


def compute_pivot_summary(df, parameters=None):
    if parameters is None:
        parameters = DEFAULT_PARAMETERS

    tool_col, product_col, wafer_col = _find_group_columns(df)
    found = _collect_param_columns(df, parameters)
    resolved, warnings = _resolve_param_columns(found, parameters)

    group_cols = [tool_col, product_col, wafer_col]
    output_cols = _build_output_columns(parameters)
    rows = []
    prev_tool = prev_product = prev_wafer = object()
    grouped = df.groupby(group_cols, dropna=False, sort=True)
    for keys, group_df in grouped:
        tool_key, product_key, wafer_key = keys
        row = {
            "Tool_label": "" if _same_key(tool_key, prev_tool) else tool_key,
            "Product_label": "" if _same_key(product_key, prev_product) else product_key,
            "Wafer_label": "" if _same_key(wafer_key, prev_wafer) else wafer_key,
            "Tool": tool_key,
            "Product": product_key,
            "Wafer": wafer_key,
        }
        for p in parameters:
            r = resolved[p]
            row[f"{p}_PMISH_Mean"] = _group_mean(group_df, r["pmish"])
            row[f"{p}_REF_Mean"] = _group_mean(group_df, r["ref"])
            row[f"{p}_Bias_Mean"] = _group_mean(group_df, r["bias"])
            row[f"{p}_oldBias_Mean"] = _group_mean(group_df, r["old_bias"])
            row[f"{p}_PMISH_STD"] = _group_std(group_df, r["pmish"])
            row[f"{p}_REF_STD"] = _group_std(group_df, r["ref"])
        rows.append([row.get(col, "") for col in output_cols])
        prev_tool, prev_product, prev_wafer = tool_key, product_key, wafer_key

    summary = pd.DataFrame(rows, columns=output_cols)
    return summary, warnings


def save_summary(summary_df, output_path, warnings=None):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        if warnings:
            pd.DataFrame({"Warning": warnings}).to_excel(
                writer, sheet_name="Warnings", index=False)
    return output_path


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "input", nargs="?", default=None,
        help="Excel/CSV 输入文件；不填时使用 Data\\绘制数据透视表\\Test1.xlsx",
    )
    parser.add_argument(
        "output", nargs="?", default=None,
        help="输出 xlsx 文件；不填时自动生成 <输入名>_pivot_summary.xlsx",
    )
    parser.add_argument("--sheet", default=None, help="Excel sheet 名称，默认取第一个 sheet")
    args = parser.parse_args(argv)

    input_path = args.input
    if input_path is None:
        input_path = str(DEFAULT_INPUT_PATH)
        if not Path(input_path).exists():
            parser.error(f"未提供输入文件，且默认文件不存在: {input_path}")
        print(f"[信息] 未提供输入文件，使用默认: {input_path}")

    df = load_dataframe(input_path, args.sheet)
    summary, warnings = compute_pivot_summary(df)
    for w in warnings:
        print(f"[警告] {w}", file=sys.stderr)

    output = args.output or str(
        Path(input_path).with_name(Path(input_path).stem + "_pivot_summary.xlsx"))
    save_summary(summary, output, warnings)
    print(f"已生成: {Path(output).resolve()}")
    print(f"分组数: {len(summary)}（Tool x Product x Wafer）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
