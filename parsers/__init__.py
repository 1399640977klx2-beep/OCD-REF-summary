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
    for fp in filepaths:
        try:
            df = parser(fp)
            if df is not None and len(df) > 0:
                dfs.append(df)
            else:
                msg = f"{str(fp.split(chr(92))[-1])}: empty result"
                if error_handler: error_handler(msg)
        except Exception as e:
            msg = f"{str(fp.split(chr(92))[-1])}: {e}"
            if error_handler: error_handler(msg)
    return dfs if dfs else None
