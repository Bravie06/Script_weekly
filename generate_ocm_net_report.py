"""
generate_ocm_net_report.py
Updates OCM Network KPI files (Daily, Weekly, Monthly) from vendor raw files.

Supported vendors/generations:
  Huawei  2G, 3G, 4G
  Nokia   2G, 3G, 4G
  ZTE     2G, 3G, 4G

Sheets updated in each OCM file:
  Avail2G, Avail3G, Avail4G,
  CSSR2G, CSSR3G, CSSR4G,
  DCR2G, DCR3G, DCR4G,
  TRAFFIC2G, TRAFFIC3G, TRAFFIC4G

Note: the former sheets 'DailyCombinedCSTrafic (Kerl)',
'DailyCombinedPSTrafic (TB)', '3G PacketLoss', '4G PacketLoss' and
'Throughput 4G Mbps' no longer exist in the OCM files and are no longer
updated (all related code, including the dedicated packet-loss mode,
has been removed).
"""

from __future__ import annotations

import os
import re
import logging
from datetime import datetime, date
from collections import defaultdict

import openpyxl

logger = logging.getLogger(__name__)

# ─── OCM SHEET NAMES ──────────────────────────────────────────────────────────

SHEETS = [
    'Avail2G',
    'Avail3G',
    'Avail4G',
    'CSSR2G',
    'CSSR3G',
    'CSSR4G',
    'DCR2G',
    'DCR3G',
    'DCR4G',
    'TRAFFIC2G',
    'TRAFFIC3G',
    'TRAFFIC4G',
]

# Sheets where values are SUMMED (not averaged) when aggregating daily
# values to weekly / monthly, and when merging several vendor files:
# traffic volumes add up, unlike rates (availability / CSSR / DCR).
ADDITIVE_SHEETS = {
    'TRAFFIC2G',
    'TRAFFIC3G',
    'TRAFFIC4G',
}

# Sheets where values are rounded to 2 decimal places on output
TRAFFIC_SHEETS = {
    'TRAFFIC2G',
    'TRAFFIC3G',
    'TRAFFIC4G',
}

# ─── KPI MAPPINGS ─────────────────────────────────────────────────────────────
# {sheet_name: {vendor_gen: (column_name_in_source, multiplier)}}
#
# multiplier converts the raw value to the target unit:
#   availability  → % (0-100)
#   CSSR / DCR    → % (0-100)
#   TRAFFIC2G/3G  → Erlang, kept in the raw unit (multiplier 1.0)
#   TRAFFIC4G     → GB, kept in the raw unit (multiplier 1.0)

#
# UNIT NOTES per vendor:
#   Huawei  2G/3G/4G  : all rates already in %; traffic in Erl/MB/kB·s⁻¹
#   Nokia   2G/3G/4G  : rates in % (rows start at index 2, index 1 is KPI codes)
#   ZTE     2G        : availability is ratio (0-1); CSSR/DCR already in %
#   ZTE     3G        : availability, CSSR, DCR all ratios (0-1) → ×100
#   ZTE     4G        : availability and rates already in %

KPI_MAP: dict[str, dict[str, tuple[str, float]]] = {

    'Avail2G': {
        'H_2G': ('RR307:TCH Availability(%)',               1.0),
        'N_2G': ('TCH Availability Normal TRXs',            1.0),    # already %
        'Z_2G': ('ORA_2G_TCH Availability Normal TRXs',    100.0),   # ratio → %
    },

    'Avail3G': {
        'H_3G': ('3G Availability (Group)',                                         1.0),
        'N_3G': ('Cell Availability, excluding blocked by user state (BLU)',        1.0),   # %
        'Z_3G': ('ORA_3G_Cell Availability, excluding BLU (%)',                   100.0),   # ratio → %
    },

    'Avail4G': {
        'H_4G': ('4G Cell Availability (Excluding manual)',         1.0),
        'N_4G': ('Cell Avail',                                      1.0),
        'Z_4G': ('ORA_4G_Cell Availability, excluding BLU_ZTE(%)', 1.0),
    },

    'CSSR2G': {
        'H_2G': ('FT_Call Setup Success Rate-Speech',   1.0),
        'N_2G': ('ORA_2G_CSSR_CS_new',                  1.0),
        'Z_2G': ('ORA_2G_CSSR_CS_New(%)',               1.0),
    },

    'CSSR3G': {
        'H_3G': ('GRP_3G Call Setup Success Rate (CS)(%)',          1.0),
        'N_3G': ('ORA_CSSR_CS_CellPCH_URAPCHnew',                  1.0),
        'Z_3G': ('ORA_3G_CSSR CS with Cell PCH/URA PCH (%)',      100.0),  # ratio → %
    },

    'DCR2G': {
        'H_2G': ('FT_Drop Call Rate-Speech',            1.0),
        'N_2G': ('ORA_2G_Call_Drop_CS_new',             1.0),
        'Z_2G': ('ORA_2G_Call_Drop_CS_New(%)',          1.0),
    },

    'DCR3G': {
        'H_3G': ('Grp_3G Drop Call Rate (CS)%_Update',     1.0),
        'N_3G': ('grp_3G_Drop_Call_CS',                    1.0),
        'Z_3G': ('ORA_3G_Drop Call Rate CS(%)',           100.0),  # ratio → %
    },

    'CSSR4G': {
        'H_4G': ('4G_Grp_4G/LTE CALL SETUP SUCCESS RATE (WITHOUT VOLTE)(%)', 1.0),
        'N_4G': ('LTE_CALL_SETUP_SUCCESS_RATE_NEW_PERC',                     1.0),
        'Z_4G': ('ORA_4G_CALL_SETUP_SUCCESS_RATE_New(%)',                    1.0),
    },

    'DCR4G': {
        'H_4G': ('4G_Grp_4G/LTE DROP CALL RATE (WITHOUT VOLTE)(%)', 1.0),
        # N_4G: Nokia column name not provided yet – to be added when known.
        'Z_4G': ('ORA_4G_LTE_Drop_Call_Rate_WO_VoLTE_New(%)',       1.0),
    },

    'TRAFFIC2G': {
        # CS traffic volume, in Erlang (raw unit kept as-is)
        'H_2G': ('K3014:Traffic Volume on TCH(Erl)',   1.0),
        'N_2G': ('Erlang_Traffic_Carried_2G',          1.0),
        'Z_2G': ('ORA_2G_CS_TRAFFIC',                  1.0),
    },

    'TRAFFIC3G': {
        # CS speech traffic volume, in Erlang (raw unit kept as-is)
        'H_3G': ('Grp_ 3G TRAFFIC – SPEECH(Erl)',       1.0),
        'N_3G': ('grp_traffic_speech_erlangs',         1.0),
        'Z_3G': ('ORA_3G_CS Voice Traffic (Erl)',       1.0),
    },

    'TRAFFIC4G': {
        # Data traffic volume, in GB (raw unit kept as-is)
        # Huawei: DL volume only, per the specified source column
        'H_4G': ('4G_Grp_4G/LTE DL TRAFFIC VOLUME(GB)', 1.0),
        'N_4G': ('Grp_4G_Traffic_New_Gbytes',           1.0),
        'Z_4G': ('ORA_4G_Total TRAFFIC(DL+UL)(GB)',     1.0),
    },
}

# ─── VENDOR FILE READ CONFIGS ─────────────────────────────────────────────────
# date_col       : column header containing the date
# site_col       : column header containing the site/node name
# skip_extra_hdr : Nokia files have a second header row (KPI codes) → skip it
# skip_site_re   : regex to skip aggregate/invalid site name rows

FILE_CONFIGS: dict[str, dict] = {
    'H_2G': {
        'date_col': 'Date',
        'site_col': 'Site Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'H_3G': {
        'date_col': 'Date',
        'site_col': 'NODEBNAME',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'H_4G': {
        'date_col': 'Date',
        'site_col': 'eNodeB Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'N_2G': {
        'date_col': 'Period start time',
        'site_col': 'BCF name',
        'skip_extra_hdr': True,   # row 1 = KPI codes
        'skip_site_re': r'^\d+$', # purely numeric BCF names are non-site rows
    },
    'N_3G': {
        'date_col': 'Period start time',
        'site_col': 'WBTS name',
        'skip_extra_hdr': True,
        'skip_site_re': r'^\d+$', # e.g. '300000000' aggregate rows
    },
    'N_4G': {
        'date_col': 'Period start time',
        'site_col': 'LNBTS name',
        'skip_extra_hdr': True,
        'skip_site_re': None,
    },
    'Z_2G': {
        'date_col': 'Begin Time',
        'site_col': 'SITE Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'Z_3G': {
        'date_col': 'Begin Time',
        'site_col': 'NodeB Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'Z_4G': {
        # ZTE 4G uses a non-breaking space (\xa0) in the column name;
        # headers are normalised before comparison so use a regular space here.
        'date_col': 'Begin Time',
        'site_col': 'Managed Element',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
}


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _normalise(text) -> str:
    """Replace non-breaking spaces and strip; used to normalise column headers."""
    if text is None:
        return ''
    return str(text).replace('\xa0', ' ').strip()


def _to_float(val) -> float | None:
    """Return float or None for missing / Huawei '/0' / 'NIL' values."""
    if val is None:
        return None
    s = str(val).strip()
    if s in ('/0', 'NIL', 'N/A', '-', ''):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _parse_date(val) -> date | None:
    """Convert datetime / date / ISO-string to a date object."""
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val[:10], '%Y-%m-%d').date()
        except ValueError:
            pass
    return None


def _extract_standard_site_code(site_name: str) -> str | None:
    """
    Extract standard XXX_NNN site code for Huawei / Nokia site names.
    Purely numeric strings (aggregate rows) return None.
    """
    if not site_name:
        return None
    parts = _normalise(site_name).split('_')
    if len(parts) >= 2 and re.match(r'^\d+$', parts[1]):
        return f'{parts[0]}_{parts[1]}'
    return None


def _extract_zte_site_code(site_name: str) -> str | None:
    """
    Extract site code for ZTE site names according to Excel formula logic:
    =IF(OR(ISERROR(FIND("_Z_",A3)),ISERROR(FIND("_URZ_",A3))),
        LEFT(A3,FIND("_",A3,FIND("_",A3)+1)-1),
        TRIM(LEFT(A3,MIN(IFERROR(FIND("_Z_",A3),LEN(A3)),IFERROR(FIND("_URZ_",A3),LEN(A3)))-1)))
    """
    if not site_name:
        return None
    s = _normalise(site_name)
    if not s:
        return None

    z_in = '_Z_' in s
    urz_in = '_URZ_' in s

    if not z_in or not urz_in:
        first_u = s.find('_')
        if first_u != -1:
            second_u = s.find('_', first_u + 1)
            if second_u != -1:
                return s[:second_u].strip()
        return s.strip()
    else:
        z_pos = s.find('_Z_')
        urz_pos = s.find('_URZ_')
        min_pos = min(z_pos, urz_pos)
        return s[:min_pos].strip()


def extract_site_code_for_vendor(site_name: str, vendor: str | None) -> str | None:
    """
    Extract site code from site_name based on vendor ('HUAWEI', 'NOKIA', 'ZTE').
    Defaults to standard XXX_NNN if vendor is missing or unrecognized.
    """
    if not site_name:
        return None
    v = (vendor or '').strip().upper()
    if v == 'ZTE':
        return _extract_zte_site_code(site_name)
    else:
        return _extract_standard_site_code(site_name)


def _extract_site_code(site_name: str) -> str | None:
    """Deprecated legacy extraction function; defaults to standard logic."""
    return _extract_standard_site_code(site_name)


def _date_to_week_key(d: date) -> str:
    """ISO week label e.g. '2026W19'."""
    iso = d.isocalendar()
    return f'{iso[0]}W{iso[1]:02d}'


def _date_to_month_key(d: date) -> str:
    """Month label e.g. '2026M05'."""
    return f'{d.year}M{d.month:02d}'


# ─── VENDOR FILE READER ───────────────────────────────────────────────────────

def read_vendor_file(
    filepath: str,
    vendor_gen: str,
) -> dict[str, dict[date, dict[str, float]]]:
    """
    Parse one vendor raw file.
    Returns {site_code: {date: {sheet_name: value}}}.
    """
    cfg = FILE_CONFIGS[vendor_gen]
    logger.debug(f'  Reading {vendor_gen}: {os.path.basename(filepath)}')

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    row_iter = ws.iter_rows(values_only=True)

    # Read header row
    try:
        first_row = next(row_iter)
    except StopIteration:
        wb.close()
        logger.warning(f'  {vendor_gen}: file is empty')
        return {}
    headers = [_normalise(h) for h in first_row]

    # Nokia files have an extra KPI-code row between header and data
    if cfg['skip_extra_hdr']:
        try:
            next(row_iter)  # skip KPI codes row
        except StopIteration:
            pass

    # Locate date and site columns
    date_col = cfg['date_col']
    site_col = cfg['site_col']

    if date_col not in headers:
        logger.warning(f'  {vendor_gen}: date column "{date_col}" not found; '
                       f'available: {headers[:10]}')
        return {}
    if site_col not in headers:
        logger.warning(f'  {vendor_gen}: site column "{site_col}" not found; '
                       f'available: {headers[:10]}')
        return {}

    date_idx = headers.index(date_col)
    site_idx = headers.index(site_col)

    # Build KPI column index → (sheet, multiplier) map  (direct columns)
    kpi_cols: dict[int, tuple[str, float]] = {}
    for sheet_name, vmap in KPI_MAP.items():
        if vendor_gen in vmap:
            col_name, mult = vmap[vendor_gen]
            norm_col = _normalise(col_name)
            if norm_col in headers:
                kpi_cols[headers.index(norm_col)] = (sheet_name, mult)
            else:
                logger.warning(f'  {vendor_gen}: KPI column "{col_name}" '
                               f'not found for sheet "{sheet_name}"')

    skip_re = re.compile(cfg['skip_site_re']) if cfg.get('skip_site_re') else None

    # Use accumulator to handle files with multiple rows per site/day.
    # Structure: {site_code: {date: {sheet_name: [values]}}}
    accum: dict[str, dict[date, dict[str, list[float]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    skipped = 0

    for row in row_iter:
        if not row or all(v is None for v in row):
            continue

        site_raw = row[site_idx] if site_idx < len(row) else None
        if not site_raw:
            continue
        site_str = str(site_raw).strip()
        if skip_re and skip_re.match(site_str):
            skipped += 1
            continue

        vendor_name = 'ZTE' if vendor_gen.startswith('Z_') else ('NOKIA' if vendor_gen.startswith('N_') else 'HUAWEI')
        site_code = extract_site_code_for_vendor(site_str, vendor_name)
        if not site_code:
            skipped += 1
            continue

        d = _parse_date(row[date_idx] if date_idx < len(row) else None)
        if not d:
            continue

        for col_idx, (sheet_name, mult) in kpi_cols.items():
            raw = row[col_idx] if col_idx < len(row) else None
            val = _to_float(raw)
            if val is not None:
                accum[site_code][d][sheet_name].append(round(val * mult, 6))

    wb.close()

    # Collapse accumulator: average multiple values for the same site/date/sheet
    result: dict[str, dict[date, dict[str, float]]] = {}
    for site_code, date_map in accum.items():
        result[site_code] = {}
        for d, sheet_map in date_map.items():
            result[site_code][d] = {
                s: round(sum(vals) / len(vals), 6)
                for s, vals in sheet_map.items()
            }

    logger.info(f'  {vendor_gen}: {len(result)} sites | {skipped} rows skipped')
    return result


# ─── AGGREGATION ──────────────────────────────────────────────────────────────

def _aggregate(
    daily: dict[str, dict[date, dict[str, float]]],
    key_fn,
) -> dict[str, dict[str, dict[str, float]]]:
    """
    Aggregate daily data to weekly or monthly using key_fn(date) → period_key.
    Additive sheets (traffic volumes) are summed; all others are averaged.
    Returns {site_code: {period_key: {sheet: aggregated_value}}}.
    """
    buckets: dict[str, dict[str, dict[str, list[float]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    for site_code, date_map in daily.items():
        for d, sheet_map in date_map.items():
            pk = key_fn(d)
            for sheet, val in sheet_map.items():
                buckets[site_code][pk][sheet].append(val)

    result: dict[str, dict[str, dict[str, float]]] = {}
    for site_code, pk_map in buckets.items():
        result[site_code] = {}
        for pk, sheet_map in pk_map.items():
            result[site_code][pk] = {
                s: round(sum(vals), 4) if s in ADDITIVE_SHEETS
                   else round(sum(vals) / len(vals), 4)
                for s, vals in sheet_map.items()
            }
    return result


def aggregate_weekly(daily):
    return _aggregate(daily, _date_to_week_key)


def aggregate_monthly(daily):
    return _aggregate(daily, _date_to_month_key)


# ─── OCM FILE WRITER ──────────────────────────────────────────────────────────

def build_ocm_reference_table(ws) -> dict[str, dict]:
    """
    Build reference table from OCM worksheet rows (starting at row 3).
    Cols: Col A (index 0) = Site Name, Col B (index 1) = Site Code, Col C (index 2) = Vendor.
    Returns dictionary mapping calculated site_code -> {'row': row_number, 'site_name': name, 'raw_code': raw_code, 'vendor': vendor}.
    """
    ref_table = {}
    for row in ws.iter_rows(min_row=3, max_col=3, values_only=False):
        name_cell = row[0] if len(row) > 0 else None
        code_cell = row[1] if len(row) > 1 else None
        vendor_cell = row[2] if len(row) > 2 else None

        raw_name = str(name_cell.value).strip() if name_cell and name_cell.value is not None else ''
        raw_code = str(code_cell.value).strip() if code_cell and code_cell.value is not None else ''
        vendor = str(vendor_cell.value).strip() if vendor_cell and vendor_cell.value is not None else ''

        if not raw_code and not raw_name:
            continue

        site_code = extract_site_code_for_vendor(raw_code or raw_name, vendor)
        if not site_code and raw_code:
            site_code = raw_code

        if site_code:
            row_num = row[0].row if row else None
            ref_table[site_code] = {
                'row': row_num,
                'site_name': raw_name,
                'raw_code': raw_code,
                'vendor': vendor,
            }
            if raw_code and raw_code not in ref_table:
                ref_table[raw_code] = ref_table[site_code]
    return ref_table


def update_ocm_file(
    filepath: str,
    data: dict,          # {site_code: {period_key: {sheet: value}}}
    period: str,         # 'daily' | 'weekly' | 'monthly'
    log_fn=None,
    sheets_filter: set[str] | None = None,  # if set, only write these sheets
) -> None:
    """
    Write KPI values into the OCM Excel workbook.

    OCM file layout (all three granularities):
      Row 1  (1-based) : empty / title row
      Row 2  (1-based) : header — site metadata in cols A-C, period labels from col D
      Row 3+ (1-based) : one site per row
      Col B             : 'Code du Site'  (used for site lookup)
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)
        else:
            logger.info(msg)

    _log(f'Loading: {os.path.basename(filepath)}')
    wb = openpyxl.load_workbook(filepath)

    sheets_to_write = [s for s in SHEETS
                       if (sheets_filter is None or s in sheets_filter)]

    for sheet_name in sheets_to_write:
        if sheet_name not in wb.sheetnames:
            _log(f'  WARNING: sheet "{sheet_name}" missing – skipped')
            continue

        ws = wb[sheet_name]

        # ── Build period_key → column-number map from header row (row 2) ──────
        period_col: dict = {}
        header_row = ws[2]
        for cell in header_row[3:]:
            hv = cell.value
            if hv is None:
                continue
            if period == 'daily':
                if isinstance(hv, datetime):
                    period_col[hv.date()] = cell.column
                elif isinstance(hv, date):
                    period_col[hv] = cell.column
                elif isinstance(hv, str):
                    d = _parse_date(hv)
                    if d:
                        period_col[d] = cell.column
            elif isinstance(hv, str):
                if period == 'weekly' and re.match(r'^\d{4}W\d{2}$', hv):
                    period_col[hv] = cell.column
                elif period == 'monthly' and re.match(r'^\d{4}M\d{2}$', hv):
                    period_col[hv] = cell.column

        # ── Create columns for periods present in the data but missing ───────
        # in the OCM file (e.g. the raw files cover weeks 2026W27–2026W47
        # while the OCM file only has columns up to 2026W20). Without this,
        # those periods are silently skipped and nothing is written.
        needed_keys = {
            pk
            for period_map in data.values()
            for pk, sheet_map in period_map.items()
            if sheet_name in sheet_map
        }
        missing_keys = sorted(needed_keys - set(period_col))
        if missing_keys:
            last_header_col = max(
                (cell.column for cell in header_row if cell.value is not None),
                default=3,
            )
            next_col = last_header_col + 1
            for pk in missing_keys:
                header_val = (
                    datetime(pk.year, pk.month, pk.day) if period == 'daily' else pk
                )
                ws.cell(row=2, column=next_col, value=header_val)
                period_col[pk] = next_col
                next_col += 1
            _log(f'  {sheet_name}: created {len(missing_keys)} missing '
                 f'period column(s): {missing_keys[0]} … {missing_keys[-1]}')

        # ── Build site_code → row-number reference map using vendor column ─────
        ref_table = build_ocm_reference_table(ws)
        site_row = {code: info['row'] for code, info in ref_table.items()}

        # ── Write values ───────────────────────────────────────────────────────
        written = 0
        matched_sites = 0
        for site_code, period_map in data.items():
            if site_code not in site_row:
                continue
            matched_sites += 1
            rn = site_row[site_code]
            for pk, sheet_map in period_map.items():
                if sheet_name not in sheet_map:
                    continue
                if pk not in period_col:
                    continue
                cn = period_col[pk]
                decimals = 2 if sheet_name in TRAFFIC_SHEETS else 4
                ws.cell(row=rn, column=cn, value=round(sheet_map[sheet_name], decimals))
                written += 1

        _log(f'  {sheet_name}: {written} cells written '
             f'({matched_sites} sites matched)')

    wb.save(filepath)
    _log(f'Saved: {os.path.basename(filepath)}')


# Default OCM filenames (used by CLI and as default browse hints)
OCM_DEFAULT_FILENAMES = {
    'daily':   'OCM Network Key Performances Indicators Sites Per Sites Daily Level.xlsx',
    'weekly':  'OCM Network Key Performances Indicators Sites Per Sites Weekly Level.xlsx',
    'monthly': 'OCM Network Key Performances Indicators Sites Per Sites Monthly Level.xlsx',
}

ALL_GRANULARITIES = {'daily', 'weekly', 'monthly'}


# ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────

def process_vendor_files(
    vendor_files: dict[str, str],
    ocm_files: dict[str, str],
    granularities: set[str] | None = None,
    sheets_filter: set[str] | None = None,
    log_callback=None,
) -> None:
    """
    Main entry point called by the GUI.

    vendor_files  : {vendor_gen: filepath}  e.g. {'H_2G': '...', 'N_3G': '...'}
    ocm_files     : {'daily': path, 'weekly': path, 'monthly': path}
                    Each key is optional – only provide the ones you want to update.
    granularities : subset of {'daily', 'weekly', 'monthly'}.
                    Defaults to all keys present in ocm_files.
    sheets_filter : if set, only these OCM sheets are written and only vendor keys
                    that map to these sheets are read.
    log_callback  : optional callable(str) for UI progress messages.
    """
    if granularities is None:
        granularities = set(ocm_files.keys())
    granularities = granularities & ALL_GRANULARITIES

    # Determine which vendor keys are relevant for sheets_filter
    if sheets_filter is not None:
        relevant_vkeys: set[str] = set()
        for sheet, vmap in KPI_MAP.items():
            if sheet in sheets_filter:
                relevant_vkeys.update(vmap.keys())
        vendor_files = {vg: fp for vg, fp in vendor_files.items()
                        if vg in relevant_vkeys}

    def log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            logger.info(msg)

    if not granularities:
        log('No granularities selected – nothing to do.')
        return

    # ── Validate OCM file paths ───────────────────────────────────────────────
    for gran in granularities:
        fp = ocm_files.get(gran, '')
        if not fp or not os.path.isfile(fp):
            log(f'ERROR: OCM {gran} file not found: {fp}')
            return

    # ── Collect all daily data ────────────────────────────────────────────────
    # Structure: {site_code: {date: {sheet: value}}}
    all_daily: dict[str, dict[date, dict[str, float]]] = defaultdict(
        lambda: defaultdict(dict)
    )

    for vg, fp in vendor_files.items():
        if not fp or not os.path.isfile(fp):
            log(f'  Skipping {vg}: no file provided')
            continue
        log(f'Reading {vg} …')
        try:
            vdata = read_vendor_file(fp, vg)
            for site_code, date_map in vdata.items():
                for d, sheet_map in date_map.items():
                    for sheet, val in sheet_map.items():
                        existing = all_daily[site_code][d].get(sheet)
                        if existing is not None and sheet in ADDITIVE_SHEETS:
                            all_daily[site_code][d][sheet] = round(existing + val, 6)
                        else:
                            all_daily[site_code][d][sheet] = val
        except Exception as exc:
            import traceback
            log(f'  ERROR reading {vg}: {exc}')
            log(traceback.format_exc())

    total_sites = len(all_daily)
    if total_sites == 0:
        log('No data loaded – nothing to write.')
        return
    log(f'Total sites with data: {total_sites}')

    # ── Update selected granularities ─────────────────────────────────────────
    if 'daily' in granularities:
        log('\n--- Updating Daily file ---')
        update_ocm_file(ocm_files['daily'], all_daily, period='daily',
                        log_fn=log_callback, sheets_filter=sheets_filter)

    if 'weekly' in granularities:
        log('\n--- Aggregating to Weekly ---')
        weekly_data = aggregate_weekly(all_daily)
        update_ocm_file(ocm_files['weekly'], weekly_data, period='weekly',
                        log_fn=log_callback, sheets_filter=sheets_filter)

    if 'monthly' in granularities:
        log('\n--- Aggregating to Monthly ---')
        monthly_data = aggregate_monthly(all_daily)
        update_ocm_file(ocm_files['monthly'], monthly_data, period='monthly',
                        log_fn=log_callback, sheets_filter=sheets_filter)

    done = ' + '.join(sorted(granularities))
    log(f'\nDone. Updated: {done}.')


# ─── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import argparse

    logging.basicConfig(level=logging.INFO, format='%(message)s')

    parser = argparse.ArgumentParser(description='Update OCM Network KPI reports from vendor raw files.')
    parser.add_argument('--ocm-dir', default=os.path.join('files', 'ocm net'),
                        help='Folder containing the three OCM Excel files')
    parser.add_argument('--raw-dir', default=os.path.join('files', 'ocm net', 'RAW FILES'),
                        help='Folder containing vendor raw files')

    args = parser.parse_args()

    # Auto-detect files by filename prefix
    prefixes = {
        'H_2G':    ('H 2G', 'H2G'),
        'H_3G':    ('H 3G', 'H3G'),
        'H_4G':    ('H 4G', 'H4G'),
        'N_2G':    ('N 2G', 'N2G'),
        'N_3G':    ('N 3G', 'N3G'),
        'N_4G':    ('N 4G', 'N4G'),
        'Z_2G':    ('Z P',),           # 'Z Performance Management...2G'
        'Z_3G':    ('Z 3G', 'Z3G'),
        'Z_4G':    ('Z-4G', 'Z 4G', 'Z4G'),
    }

    vendor_files: dict[str, str] = {}
    for fname in os.listdir(args.raw_dir):
        if not fname.lower().endswith('.xlsx'):
            continue
        for vg, pfxs in prefixes.items():
            if any(fname.startswith(p) for p in pfxs) and vg not in vendor_files:
                vendor_files[vg] = os.path.join(args.raw_dir, fname)
                break

    logger.info('Detected vendor files:')
    for vg, fp in vendor_files.items():
        logger.info(f'  {vg}: {os.path.basename(fp)}')

    ocm_files = {
        gran: os.path.join(args.ocm_dir, fname)
        for gran, fname in OCM_DEFAULT_FILENAMES.items()
    }
    process_vendor_files(vendor_files, ocm_files)
