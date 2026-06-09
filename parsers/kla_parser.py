import re
import pandas as pd


def parse_kla_file(filepath):
    """
    Parse a KLA-format CSV file into a unified DataFrame.

    KLA format has metadata lines followed by a data table.
    The data table starts after the line beginning with 'Site #'.
    """
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Extract metadata from header
    metadata = {}
    for line in lines:
        line_stripped = line.strip()
        if line_stripped.startswith('WAFER ID,'):
            metadata['WaferID'] = line_stripped.split(',')[1].strip()
        elif line_stripped.startswith('LOT ID,'):
            metadata['LotID'] = line_stripped.split(',')[1].strip()
        elif line_stripped.startswith('DATE/TIME,'):
            metadata['Date'] = line_stripped.split(',')[1].strip()
        elif line_stripped.startswith('RECIPE,'):
            metadata['Recipe'] = line_stripped.split(',')[1].strip()
        elif line_stripped.startswith('TOOL ID,'):
            metadata['ToolID'] = line_stripped.split(',')[1].strip()

    # Find the data header line (starts with 'Site #')
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('Site #'):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError(f"Cannot find data header in KLA file: {filepath}")

    # Read data table
    df = pd.read_csv(filepath, skiprows=header_idx, encoding='utf-8', on_bad_lines='skip')

    # Add metadata columns
    for key, value in metadata.items():
        df[key] = value


    return df
