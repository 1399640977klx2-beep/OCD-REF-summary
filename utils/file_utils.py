"""
File scanning and path utility functions.
"""
import os

def scan_data_files(directory, extensions=None):
    """
    Recursively scan a directory for data files.
    Returns list of full file paths.
    """
    if extensions is None:
        extensions = {'.csv', '.txt'}
    files = []
    for root, _dirs, filenames in os.walk(directory):
        for fn in filenames:
            ext = os.path.splitext(fn)[1].lower()
            if ext in extensions:
                files.append(os.path.join(root, fn))
    return sorted(files)

def extract_wafer_id_from_sme_path(sme_path, depth=-3):
    """
    Extract wafer ID from the SME file path.
    Default: take the directory at depth=-3 (relative to the file),
    then take the prefix before '_'.
    Example path: .../X1L086#05_20260527_123023/P228/xxx.sme
    -> folder = 'X1L086#05_20260527_123023'
    -> wafer ID = 'X1L086#05'
    """
    import re
    parts = sme_path.replace('\\', '/').split('/')
    if len(parts) >= abs(depth):
        folder = parts[depth]
        prefix = folder.split('_')[0]
        return prefix
    return ''
