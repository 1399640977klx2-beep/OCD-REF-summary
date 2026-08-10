"""Organize spectral measurement folders by machine / product / wafer.

The organizer scans a source tree generically instead of depending on a
specific layout such as ``OCD_*.rcp`` folders:

* Machine, product and wafer names are found with configurable regexes.
* Wafer directories are detected by name and copied with their internal
  tree preserved.
* Product is taken from any ancestor directory; if it cannot be found the
  wafer is still kept directly under the machine folder.
* CSV files outside wafer folders are copied to the product level, under
  ``ResultData`` when the source path contains a ``ResultData`` folder.
* Existing targets are skipped unless ``overwrite=True``.

CLI usage:
    python organize_spectra.py -s Raw -d Organized
    python organize_spectra.py -s Raw -d Organized -n
    python organize_spectra.py   # interactive dialog
"""

import argparse
import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Pattern

DEFAULT_PRODUCT_PATTERN = r'[A-Za-z0-9]{4}'
DEFAULT_MACHINE_PATTERN = r'[A-Za-z]{3}\d{2}'
DEFAULT_WAFER_PATTERNS = [
    r'^X[A-Z0-9]{2}\d{3}#\d{2}(_\d{8}_\d{6})?$',
    r'^X[A-Z0-9]{2}\d{3}\.\d{2}$',
    r'^\d{8}_\d{6}$',
]

_DATE_DIR_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')
_LOG_FILE = 'reorganize_log.txt'


def _match_group(m):
    return m.group(1) if m.groups() else m.group()


def _is_wafer_name(name, wafer_res):
    return any(r.match(name) for r in wafer_res)


def _is_date_dir(name):
    return bool(_DATE_DIR_RE.match(name))


def _is_within(path, root):
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def _extract_machine(name, machine_res):
    for r in machine_res:
        m = r.search(name)
        if m:
            return _match_group(m)
    return None


def _extract_product(name, product_re):
    """Extract product from a folder/file name.

    Tries a full match first, then the common ``OCD_<code>_OFFSET``
    convention, then a tokenized match. Pure-digit tokens are ignored to
    avoid treating dates or lot numbers as product codes.
    """
    m = product_re.fullmatch(name)
    if m:
        return _match_group(m)
    m = re.search(r'OCD_([A-Za-z0-9]+)_OFFSET', name)
    if m:
        return m.group(1)
    for token in re.split(r'[^A-Za-z0-9]+', name):
        if not token:
            continue
        m = product_re.fullmatch(token)
        if m:
            val = _match_group(m)
            if not val.isdigit():
                return val
    return None


def _ancestor_parts(item_path, source_root):
    """Yield (name, path) for ancestors of item_path up to source_root."""
    current = item_path.parent
    while current is not None and current != current.parent:
        yield current.name, current
        if current == source_root:
            break
        current = current.parent


def _find_machine(item_path, source_root, machine_res):
    for name, _path in _ancestor_parts(item_path, source_root):
        machine = _extract_machine(name, machine_res)
        if machine:
            return machine
    return None


def _find_product(item_path, source_root, product_re, wafer_res, machine_res):
    for name, _path in _ancestor_parts(item_path, source_root):
        # Product lookup stops at the machine folder; never scan above it.
        if _extract_machine(name, machine_res):
            break
        if _is_date_dir(name) or _is_wafer_name(name, wafer_res):
            continue
        product = _extract_product(name, product_re)
        if product:
            return product
    return None


def _walk_dirs(root_dir):
    """Return root_dir plus every subdirectory below it."""
    result = [root_dir]
    for root, dirs, files in os.walk(root_dir):
        for d in dirs:
            result.append(Path(root) / d)
    return result


def _has_spectrum_data(wafer_dir):
    for root, dirs, files in os.walk(wafer_dir):
        for f in files:
            low = f.lower()
            if low.endswith('.sme') or low.endswith('.csv'):
                return True
    return False


def _is_grouping_dir(candidate, candidates):
    """Lot/date folders contain wafer folders; a real wafer does not."""
    for other in candidates:
        if other != candidate and _is_within(other, candidate):
            return True
    return False


def _collect_wafer_dirs(source_dir, wafer_res, exclude_path=None):
    candidates = []
    for root, dirs, files in os.walk(source_dir):
        root_path = Path(root)
        if exclude_path is not None and _is_within(root_path, exclude_path):
            dirs[:] = []
            continue
        dirs[:] = [d for d in dirs
                   if not (exclude_path is not None
                           and _is_within(root_path / d, exclude_path))]
        for d in dirs:
            if _is_wafer_name(d, wafer_res):
                candidates.append(root_path / d)

    real = []
    for cand in candidates:
        if _is_grouping_dir(cand, candidates):
            continue
        if _has_spectrum_data(cand):
            real.append(cand)
    return real


def _copy_tree(src_dir, dst_dir, dry_run, overwrite, log):
    if dst_dir.exists():
        if overwrite:
            if not dry_run:
                shutil.rmtree(dst_dir)
            log("  [OVERWRITE] %s" % dst_dir)
        else:
            log("  [SKIP] target exists: %s" % dst_dir)
            return False
    if dry_run:
        log("  [DRY-RUN] %s  ->  %s" % (src_dir, dst_dir))
        return True
    dst_dir.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(src_dir, dst_dir)
    log("  [OK] %s  ->  %s" % (src_dir, dst_dir))
    return True


def _copy_file(src_file, dst_file, dry_run, overwrite, log):
    if dst_file.exists():
        if overwrite:
            if not dry_run:
                dst_file.unlink()
            log("  [OVERWRITE] %s" % dst_file)
        else:
            log("  [SKIP] target exists: %s" % dst_file)
            return False
    if dry_run:
        log("  [DRY-RUN] %s  ->  %s" % (src_file, dst_file))
        return True
    dst_file.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src_file, dst_file)
    log("  [OK] %s  ->  %s" % (src_file, dst_file))
    return True


def organize(source_dir, dest_dir, product_re, machine_filters,
             wafer_patterns=None, dry_run=False, overwrite=False,
             log_file=None):
    source_dir = Path(source_dir).resolve()
    dest_dir = Path(dest_dir).resolve()
    if wafer_patterns is None:
        wafer_patterns = [re.compile(p, re.IGNORECASE) for p in DEFAULT_WAFER_PATTERNS]

    if dest_dir == source_dir:
        print("[ERROR] destination cannot be the same as source")
        return 0, 0

    exclude_path = dest_dir if _is_within(dest_dir, source_dir) else None

    log_lines = []

    def log(msg):
        print(msg)
        log_lines.append(msg)

    log("=== Spectra folder organize start: %s ==="
        % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log("Source: %s" % source_dir)
    log("Dest:   %s" % dest_dir)
    log("Product pattern: %s" % product_re.pattern)
    log("Machine patterns: %s" % [p.pattern for p in machine_filters])
    log("Wafer patterns: %s" % [p.pattern for p in wafer_patterns])
    log("Mode: %s" % ("dry-run" if dry_run else "execute"))
    log("-" * 70)

    wafer_dirs = _collect_wafer_dirs(source_dir, wafer_patterns, exclude_path)
    log("Found %d wafer directory(ies)" % len(wafer_dirs))

    covered_dirs = set()
    for w in wafer_dirs:
        covered_dirs.update(_walk_dirs(w))

    ok = 0
    skipped = 0
    by_machine = {}
    for w in wafer_dirs:
        machine = _find_machine(w, source_dir, machine_filters)
        if machine:
            by_machine.setdefault(machine, []).append(w)
        else:
            log("[SKIP] no machine found for: %s" % w)
            skipped += 1

    for machine in sorted(by_machine):
        dest_machine = dest_dir / machine
        log("")
        log("-- Machine: %s" % machine)
        for w in by_machine[machine]:
            product = _find_product(w, source_dir, product_re,
                                    wafer_patterns, machine_filters)
            dest_product = dest_machine / product if product else dest_machine
            if not product:
                log("  [NOTE] product not found, wafer kept under machine: %s"
                    % w.name)
            dst_wafer = dest_product / w.name
            if _copy_tree(w, dst_wafer, dry_run, overwrite, log):
                ok += 1
            else:
                skipped += 1

    for root, dirs, files in os.walk(source_dir):
        root_path = Path(root)
        if exclude_path is not None and _is_within(root_path, exclude_path):
            continue
        for f in files:
            if not f.lower().endswith('.csv'):
                continue
            csv_path = root_path / f
            if csv_path.parent in covered_dirs:
                continue
            machine = _find_machine(csv_path, source_dir, machine_filters)
            if not machine:
                log("[SKIP] no machine found for CSV: %s" % csv_path)
                skipped += 1
                continue
            product = _find_product(csv_path, source_dir, product_re,
                                    wafer_patterns, machine_filters)
            if not product:
                product = _extract_product(csv_path.stem, product_re)
            dest_product = dest_dir / machine / product if product else dest_dir / machine
            rel = csv_path.parent.relative_to(source_dir).parts
            use_result_dir = any(p.lower() == 'resultdata' for p in rel)
            dst_dir_csv = dest_product / 'ResultData' if use_result_dir else dest_product
            dst_csv = dst_dir_csv / f
            if _copy_file(csv_path, dst_csv, dry_run, overwrite, log):
                ok += 1
            else:
                skipped += 1

    log("")
    log("-" * 70)
    log("Done: ok=%d skipped=%d" % (ok, skipped))

    if not dry_run:
        dest_dir.mkdir(parents=True, exist_ok=True)
        log_path = dest_dir / (log_file or _LOG_FILE)
        try:
            log_path.write_text("\n".join(log_lines), encoding="utf-8")
            log("Log saved to: %s" % log_path)
        except OSError as e:
            log("Failed to write log: %s" % e)

    return ok, skipped


def _interactive():
    print("=" * 60)
    print("  Spectra Folder Organizer")
    print("=" * 60)
    print()

    while True:
        src = input("Source directory [Raw]: ").strip()
        if not src:
            src = "Raw"
        src_path = Path(src).resolve()
        if src_path.is_dir():
            break
        print("  [ERROR] directory not found: %s" % src_path)
        print()

    while True:
        dst = input("Destination directory [Organized]: ").strip()
        if not dst:
            dst = "Organized"
        dst_path = Path(dst).resolve()
        if dst_path == src_path:
            print("  [ERROR] destination cannot be the same as source")
            print()
            continue
        break

    print()

    inp = input("Product regex (Enter={}): ".format(DEFAULT_PRODUCT_PATTERN)).strip()
    product_re = re.compile(inp) if inp else re.compile(DEFAULT_PRODUCT_PATTERN)

    machine_filters = []
    print("Machine regexes (one per line, empty = done)")
    print("  Default: %s" % DEFAULT_MACHINE_PATTERN)
    while True:
        inp = input("  > ").strip()
        if not inp:
            if not machine_filters:
                machine_filters.append(re.compile(DEFAULT_MACHINE_PATTERN))
            break
        try:
            machine_filters.append(re.compile(inp))
        except re.error as e:
            print("  [ERROR] %s" % e)

    wafer_patterns = []
    print("Wafer regexes (one per line, empty = use defaults)")
    for p in DEFAULT_WAFER_PATTERNS:
        print("  default: %s" % p)
    while True:
        inp = input("  > ").strip()
        if not inp:
            if not wafer_patterns:
                wafer_patterns = [re.compile(p, re.IGNORECASE)
                                  for p in DEFAULT_WAFER_PATTERNS]
            break
        try:
            wafer_patterns.append(re.compile(inp, re.IGNORECASE))
        except re.error as e:
            print("  [ERROR] %s" % e)

    print()

    dry_run = False
    inp = input("Mode (Enter=execute, p=preview): ").strip().lower()
    if inp in ("p", "preview"):
        dry_run = True

    overwrite = False
    inp = input("Overwrite existing files? (Enter=no, y=yes): ").strip().lower()
    if inp in ("y", "yes"):
        overwrite = True

    print()
    print("-" * 40)
    print("  Source:    %s" % src_path)
    print("  Dest:      %s" % dst_path)
    print("  Product:   %s" % product_re.pattern)
    print("  Machine:   %s" % [p.pattern for p in machine_filters])
    print("  Wafer:     %s" % [p.pattern for p in wafer_patterns])
    print("  Mode:      %s" % ("Preview" if dry_run else "Execute"))
    print("  Overwrite: %s" % overwrite)
    print("-" * 40)
    inp = input("Continue? (Enter=yes, n=cancel): ").strip().lower()
    if inp in ("n", "no"):
        print("Cancelled")
        return

    organize(src_path, dst_path, product_re, machine_filters,
             wafer_patterns=wafer_patterns, dry_run=dry_run,
             overwrite=overwrite)
    print()
    print("Done." if not dry_run else "Preview complete")
    input("Press Enter to exit...")


def main():
    if len(sys.argv) == 1:
        _interactive()
        return

    parser = argparse.ArgumentParser(
        description='Organize spectral measurement folders into unified hierarchy.')
    parser.add_argument('--source', '-s', required=False,
                        help='Raw data directory')
    parser.add_argument('--dest', '-d', required=False,
                        help='Output directory')
    parser.add_argument('--product-pattern', default=None,
                        help='Regex for product codes (default: %s)'
                             % DEFAULT_PRODUCT_PATTERN)
    parser.add_argument('--machine-pattern', action='append', default=None,
                        help='Regex for machine folder names (repeatable)')
    parser.add_argument('--wafer-pattern', action='append', default=None,
                        help='Regex for wafer folder names (repeatable)')
    parser.add_argument('--interactive', '-i', action='store_true',
                        help='Run in interactive dialog mode')
    parser.add_argument('--dry-run', '-n', action='store_true',
                        help='Preview changes without copying')
    parser.add_argument('--overwrite', action='store_true',
                        help='Overwrite existing targets')
    parser.add_argument('--list-config', action='store_true',
                        help='Show default regex patterns and exit')
    args = parser.parse_args()

    if args.interactive:
        _interactive()
        return

    if args.list_config:
        print("Product: %s" % DEFAULT_PRODUCT_PATTERN)
        print("Machine: %s" % DEFAULT_MACHINE_PATTERN)
        print("Wafer:")
        for p in DEFAULT_WAFER_PATTERNS:
            print("  %s" % p)
        return

    if not args.source or not args.dest:
        parser.print_help()
        return

    source, dest = Path(args.source).resolve(), Path(args.dest).resolve()
    if not source.is_dir():
        print("Error: source not found: %s" % source)
        return

    product_re = re.compile(args.product_pattern or DEFAULT_PRODUCT_PATTERN)
    if args.machine_pattern:
        machine_filters = [re.compile(p) for p in args.machine_pattern]
    else:
        machine_filters = [re.compile(DEFAULT_MACHINE_PATTERN)]
    if args.wafer_pattern:
        wafer_patterns = [re.compile(p, re.IGNORECASE) for p in args.wafer_pattern]
    else:
        wafer_patterns = [re.compile(p, re.IGNORECASE) for p in DEFAULT_WAFER_PATTERNS]

    organize(source, dest, product_re, machine_filters,
             wafer_patterns=wafer_patterns, dry_run=args.dry_run,
             overwrite=args.overwrite)
    print("\nDone." if not args.dry_run else "\nDry-run complete -- no files were copied.")


if __name__ == '__main__':
    main()
