"""Parser dispatcher."""
from .kla_parser import parse_kla_file
from .nova_parser import parse_nova_file
from .pmish_parser import parse_pmish_file

PARSER_MAP = {
    "KLA": parse_kla_file,
    "NOVA": parse_nova_file,
    "PMISH": parse_pmish_file,
}

def parse_files(filepaths, vendor, error_handler=None):
    parser = PARSER_MAP.get(vendor)
    if parser is None:
        raise ValueError(f"Unknown vendor: {vendor}")

    dfs = []
    pad_combined = {}

    for fp in filepaths:
        try:
            result = parser(fp)
            if isinstance(result, dict):
                for pad, df in result.items():
                    if pad not in pad_combined:
                        pad_combined[pad] = []
                    pad_combined[pad].append(df)
            elif result is not None and len(result) > 0:
                dfs.append(result)
            else:
                msg = f"{fp.split(chr(92))[-1]}: empty result"
                if error_handler: error_handler(msg)
        except Exception as e:
            msg = f"{fp.split(chr(92))[-1]}: {e}"
            if error_handler: error_handler(msg)

    if pad_combined:
        import pandas as pd
        result = {}
        for pad, pad_dfs in pad_combined.items():
            result[pad] = pd.concat(pad_dfs, ignore_index=True)
        return result

    return dfs if dfs else None
