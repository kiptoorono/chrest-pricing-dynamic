import pandas as pd
import numpy as np
import yaml
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

CONFIG_FILE = 'flood_pricing_config.yaml'

with open(CONFIG_FILE) as f:
    _cfg = yaml.safe_load(f)

FILE_NAME    = _cfg['file_name']
WINDOW_SIZE  = _cfg['window_size']
TABLE_COLUMNS = _cfg['table_columns']
REGIONS_CONFIG = _cfg['regions']
SHEET_CONFIG = _cfg['sheets']

thin_border   = Side(style='thin')
medium_border = Side(style='medium')

def year_aware_rolling_max(series, year_col, window=14):
    result = []
    vals  = series.values
    years = year_col.values
    for i in range(len(vals)):
        cy = years[i]
        year_start = np.where(years == cy)[0][0]
        start = max(year_start, i - window + 1)
        window_data = vals[start:i + 1]
        if len(window_data):
            mx = np.nanmax(window_data)
            result.append(int(mx) if not np.isnan(mx) else np.nan)
        else:
            result.append(np.nan)
    return result

def identify_flood_events(rolling, threshold):
    events, current, in_event = [], 0, False
    for v in rolling:
        if pd.isna(v):
            events.append(0); continue
        if v > threshold:
            if not in_event:
                current += 1; in_event = True
            events.append(current)
        else:
            in_event = False; events.append(0)
    return events

def mark_event_dates(event_numbers, dates):
    n = len(event_numbers)
    starts, ends, durations = [None]*n, [None]*n, [None]*n
    cur_event = None; start_idx = None; start_date = None
    for i, ev in enumerate(event_numbers):
        if ev > 0:
            if ev != cur_event:
                cur_event = ev; start_idx = i; start_date = dates.iloc[i]
                starts[i] = start_date
        else:
            if cur_event is not None:
                ed = dates.iloc[i - 1]
                ends[i - 1] = ed
                durations[i - 1] = (ed - start_date).days + 1
                cur_event = None
    if cur_event is not None:
        ed = dates.iloc[-1]
        ends[-1] = ed; durations[-1] = (ed - start_date).days + 1
    return starts, ends, durations

def calculate_event_length_tracking(event_numbers, rolling, threshold):
    result = [None] * len(event_numbers)
    for i, ev in enumerate(event_numbers):
        if ev > 0 and not pd.isna(rolling[i]) and rolling[i] > threshold:
            result[i] = rolling[i]
    return result

def calculate_event_block_max(event_numbers, rolling):
    event_max = {}
    for i, ev in enumerate(event_numbers):
        if ev > 0 and not pd.isna(rolling[i]):
            event_max[ev] = max(event_max.get(ev, rolling[i]), rolling[i])
    result = [None] * len(event_numbers)
    for i, ev in enumerate(event_numbers):
        if ev > 0:
            is_end = (i == len(event_numbers) - 1) or (event_numbers[i + 1] != ev)
            if is_end:
                result[i] = event_max.get(ev)
    return result

def calculate_complete_data_series(event_numbers, rolling):
    event_max = {}
    for i, ev in enumerate(event_numbers):
        if ev > 0 and not pd.isna(rolling[i]):
            event_max[ev] = max(event_max.get(ev, rolling[i]), rolling[i])
    result = [None] * len(event_numbers)
    for i, ev in enumerate(event_numbers):
        if ev == 0:
            result[i] = rolling[i]
        else:
            is_end = (i == len(event_numbers) - 1) or (event_numbers[i + 1] != ev)
            if is_end:
                result[i] = event_max.get(ev)
    return result

def calculate_per_event_loss(event_numbers, rolling, attachment, exhaustion):
    event_max = {}
    for i, ev in enumerate(event_numbers):
        if ev > 0 and not pd.isna(rolling[i]):
            event_max[ev] = max(event_max.get(ev, rolling[i]), rolling[i])

    def clamp(v):
        if v is None or pd.isna(v): return None
        return max(0.0, min((attachment - v) / (attachment - exhaustion), 1.0))

    result = [None] * len(event_numbers)
    for i, ev in enumerate(event_numbers):
        if ev == 0:
            result[i] = clamp(rolling[i])
        else:
            is_end = (i == len(event_numbers) - 1) or (event_numbers[i + 1] != ev)
            if is_end:
                result[i] = clamp(event_max.get(ev))
    return result

def calculate_annual_loss(loss_values, years):
    totals = {}
    for i, yr in enumerate(years):
        v = loss_values[i]
        if v is not None and not pd.isna(v):
            totals[yr] = totals.get(yr, 0.0) + v
    return totals


def process_sheet(sheet_name, cfg):
    """
    Returns a dict with all computed series and summary dataframes for
    every region in the sheet.
    """
    df = pd.read_excel(
        FILE_NAME,
        sheet_name=sheet_name,
        header=cfg['header_row'],
        usecols=cfg['data_cols'],
    )
    regions = list(REGIONS_CONFIG.keys())
    col_names = ['date'] + regions
    df.columns = col_names
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    for r in regions:
        df[r] = pd.to_numeric(df[r], errors='coerce')
    df['year'] = df['date'].dt.year

    results = {'df': df, 'regions': regions}

    for r in regions:
        thr = REGIONS_CONFIG[r]['threshold']
        att = REGIONS_CONFIG[r]['attachment']
        exh = REGIONS_CONFIG[r]['exhaustion']

        rolling   = year_aware_rolling_max(df[r], df['year'], WINDOW_SIZE)
        events    = identify_flood_events(pd.Series(rolling), thr)
        starts, ends, durations = mark_event_dates(events, df['date'])

        results[r] = {
            'rolling':          rolling,
            'events':           events,
            'starts':           starts,
            'ends':             ends,
            'durations':        durations,
            'length_tracking':  calculate_event_length_tracking(events, rolling, thr),
            'block_max':        calculate_event_block_max(events, rolling),
            'complete_series':  calculate_complete_data_series(events, rolling),
            'per_event_loss':   calculate_per_event_loss(events, rolling, att, exh),
        }

    # Annual loss summary
    all_years = sorted(df['year'].dropna().unique())
    annual = {'Year': all_years}
    for r in regions:
        ann = calculate_annual_loss(results[r]['per_event_loss'], df['year'].values)
        annual[r.upper()] = [ann.get(y, 0.0) for y in all_years]
    results['annual_loss_df'] = pd.DataFrame(annual)

    return results


def apply_table_outline(ws, col_start, col_end, title_row, last_row):
    cs = column_index_from_string(col_start) - 1
    ce = column_index_from_string(col_end) - 1

    for ci in range(cs, ce + 1):
        cl = get_column_letter(ci + 1)
        c = ws[f'{cl}{title_row}']
        c.border = Border(
            left=medium_border if ci == cs else None,
            right=medium_border if ci == ce else None,
            top=medium_border,
            bottom=None,
        )
        c2 = ws[f'{cl}{last_row}']
        c2.border = Border(
            left=medium_border if ci == cs else None,
            right=medium_border if ci == ce else None,
            top=None,
            bottom=medium_border,
        )

    for row in range(title_row, last_row + 1):
        ws[f'{col_start}{row}'].border = Border(left=medium_border, right=None, top=None, bottom=None)
        ws[f'{col_end}{row}'].border   = Border(left=None, right=medium_border, top=None, bottom=None)


def col_letter(base_letter, offset):
    """Return the column letter that is `offset` positions after base_letter."""
    return get_column_letter(column_index_from_string(base_letter) + offset)


def write_table_header(ws, start_col, title_row, title, region_names, subheader_row=None):
    """Write a section title in title_row and region sub-headers one row below."""
    ws[f'{start_col}{title_row}'] = title
    ws[f'{start_col}{title_row}'].font = Font(bold=True)
    sub_row = subheader_row if subheader_row else title_row + 1
    for i, r in enumerate(region_names):
        ws[f'{col_letter(start_col, i)}{sub_row}'] = r


def write_column_data(ws, start_col, region_idx, data, data_start_row, transform=None):
    """Write a list of values into a single column starting at data_start_row."""
    cl = col_letter(start_col, region_idx)
    for i, v in enumerate(data):
        if v is not None:
            ws[f'{cl}{data_start_row + i}'] = transform(v) if transform else v


def write_sheet_results(ws, results, cfg):
    """Write all computed tables into worksheet ws."""
    regions   = results['regions']
    n         = len(regions)
    df        = results['df']
    nrows     = len(df)
    TITLE_ROW = 1
    SUB_ROW   = 2
    DATA_ROW  = 3
    last_row  = DATA_ROW + nrows - 1

    def sc(key):
        return TABLE_COLUMNS[key]

    # 14-DAY ROLLING WINDOW 
    write_table_header(ws, sc('rolling'), TITLE_ROW, '14-DAY ROLLING WINDOW EVENT LEVEL MAXIMUM', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('rolling'), i, results[r]['rolling'], DATA_ROW)
    apply_table_outline(ws, sc('rolling'), col_letter(sc('rolling'), n - 1), TITLE_ROW, last_row)

    # EVENT IDENTIFICATION 
    write_table_header(ws, sc('events'), TITLE_ROW, 'EVENT IDENTIFICATION', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('events'), i, results[r]['events'], DATA_ROW)
    apply_table_outline(ws, sc('events'), col_letter(sc('events'), n - 1), TITLE_ROW, last_row)

    # EVENT START DATE 
    write_table_header(ws, sc('start_date'), TITLE_ROW, 'EVENT START DATE', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('start_date'), i, results[r]['starts'], DATA_ROW,
                          transform=lambda v: v.date())
    apply_table_outline(ws, sc('start_date'), col_letter(sc('start_date'), n - 1), TITLE_ROW, last_row)

    # EVENT END DATE
    write_table_header(ws, sc('end_date'), TITLE_ROW, 'EVENT END DATE', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('end_date'), i, results[r]['ends'], DATA_ROW,
                          transform=lambda v: v.date())
    apply_table_outline(ws, sc('end_date'), col_letter(sc('end_date'), n - 1), TITLE_ROW, last_row)

    #  DURATION 
    write_table_header(ws, sc('duration'), TITLE_ROW, 'DURATION', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('duration'), i, results[r]['durations'], DATA_ROW)
    apply_table_outline(ws, sc('duration'), col_letter(sc('duration'), n - 1), TITLE_ROW, last_row)

    # EVENT LENGTH TRACKING 
    write_table_header(ws, sc('length_tracking'), TITLE_ROW, 'EVENT LENGTH TRACKING', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('length_tracking'), i, results[r]['length_tracking'], DATA_ROW)
    apply_table_outline(ws, sc('length_tracking'), col_letter(sc('length_tracking'), n - 1), TITLE_ROW, last_row)

    #  EVENT BLOCK MAX 
    write_table_header(ws, sc('block_max'), TITLE_ROW, 'EVENT BLOCK MAX', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('block_max'), i, results[r]['block_max'], DATA_ROW)
    apply_table_outline(ws, sc('block_max'), col_letter(sc('block_max'), n - 1), TITLE_ROW, last_row)

    #  COMPLETE DATA SERIES 
    write_table_header(ws, sc('complete_series'), TITLE_ROW, 'COMPLETE DATA SERIES', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('complete_series'), i, results[r]['complete_series'], DATA_ROW)
    apply_table_outline(ws, sc('complete_series'), col_letter(sc('complete_series'), n - 1), TITLE_ROW, last_row)

    #  YEAR 
    ws[f"{sc('year')}{TITLE_ROW}"] = 'YEAR'
    ws[f"{sc('year')}{TITLE_ROW}"].font = Font(bold=True)
    for i, yr in enumerate(df['year'].values):
        ws[f"{sc('year')}{DATA_ROW + i}"] = yr
    apply_table_outline(ws, sc('year'), sc('year'), TITLE_ROW, last_row)

    # PER EVENT LOSS 
    write_table_header(ws, sc('per_event_loss'), TITLE_ROW, 'PER EVENT LOSS', regions)
    for i, r in enumerate(regions):
        write_column_data(ws, sc('per_event_loss'), i, results[r]['per_event_loss'], DATA_ROW)
    apply_table_outline(ws, sc('per_event_loss'), col_letter(sc('per_event_loss'), n - 1), TITLE_ROW, last_row)

    #  ANNUAL LOSS SUMMARY 
    ann_df = results['annual_loss_df']
    ann_cols = ['YEAR'] + [r.upper() for r in regions]
    write_table_header(ws, sc('annual_summary'), TITLE_ROW, 'ANNUAL LOSS SUMMARY',
                       ann_cols[1:], subheader_row=SUB_ROW)
    ws[f"{sc('annual_summary')}{SUB_ROW}"] = 'YEAR'
    for idx, row in ann_df.iterrows():
        dr = DATA_ROW + idx
        ws[f"{sc('annual_summary')}{dr}"] = int(row['Year'])
        for ci, r in enumerate(regions):
            ws[f"{col_letter(sc('annual_summary'), ci + 1)}{dr}"] = row[r.upper()]
    ann_last = SUB_ROW + len(ann_df)
    apply_table_outline(ws, sc('annual_summary'), col_letter(sc('annual_summary'), n), TITLE_ROW, ann_last)


def write_correlation_sheet(wb, all_results):
    """
    Build the CORRELATION MATRIX sheet using annual loss data from all sheets.
    Calculates correlation from annual loss summary (not per-event loss).
    """
    if 'CORRELATION MATRIX' in wb.sheetnames:
        del wb['CORRELATION MATRIX']
    ws = wb.create_sheet('CORRELATION MATRIX')
    ws.views.sheetView[0].showGridLines = True

    # Use the first sheet's config to get region names (assumed consistent)
    first_key  = list(all_results.keys())[0]
    regions    = all_results[first_key]['regions']
    reg_upper  = [r.upper() for r in regions]

    # Calculate correlation from annual loss data (not per-event loss)
    ann_df = all_results[first_key]['annual_loss_df']
    corr_df = ann_df[reg_upper].corr()
    corr_matrix = corr_df

    TITLE_ROW = 1; SUB_ROW = 2; DATA_ROW = 3
    ann_last = SUB_ROW + len(ann_df)

    # ── Annual loss summary ────────────────────────────────────────────────
    ws[f'A{TITLE_ROW}'] = 'ANNUAL LOSS SUMMARY'
    ws[f'A{TITLE_ROW}'].font = Font(bold=True)
    ws[f'A{SUB_ROW}'] = 'YEAR'
    for ci, r in enumerate(reg_upper):
        ws.cell(row=SUB_ROW, column=ci + 2, value=r)
    for idx, row in ann_df.iterrows():
        dr = DATA_ROW + idx
        ws.cell(row=dr, column=1, value=int(row['Year']))
        for ci, r in enumerate(reg_upper):
            ws.cell(row=dr, column=ci + 2, value=row[r])
    apply_table_outline(ws, 'A', get_column_letter(1 + len(regions)), TITLE_ROW, ann_last)

    # ── Correlation matrix ─────────────────────────────────────────────────
    t1_title = ann_last + 5
    t1_hdr   = t1_title + 1
    t1_start = t1_hdr + 1
    t1_last  = t1_start + len(regions) - 1

    ws.cell(row=t1_title, column=1,
            value='Correlation Matrix on historical payouts for copula application').font = Font(bold=True)

    for ci, r in enumerate(reg_upper):
        ws.cell(row=t1_hdr,          column=ci + 2, value=r).font = Font(bold=True)
        ws.cell(row=t1_start + ci,   column=1,      value=r).font = Font(bold=True)

    for ri, rr in enumerate(reg_upper):
        for ci, cr in enumerate(reg_upper):
            cell = ws.cell(row=t1_start + ri, column=ci + 2,
                           value=float(corr_matrix.loc[rr, cr]))
            cell.number_format = '0.00'

    apply_table_outline(ws, 'A', get_column_letter(1 + len(regions)), t1_title, t1_last)

    for col in range(1, 2 + len(regions)):
        ws.column_dimensions[get_column_letter(col)].width = 12

def main():
    wb = load_workbook(FILE_NAME)

    # Unmerge header rows in all target sheets
    for sheet_name in SHEET_CONFIG:
        sheet_name_str = str(sheet_name)
        if sheet_name_str in wb.sheetnames:
            ws = wb[sheet_name_str]
            to_remove = [m for m in ws.merged_cells.ranges if m.min_row <= 2]
            for m in to_remove:
                ws.unmerge_cells(str(m))

    all_results = {}
    for sheet_name, cfg in SHEET_CONFIG.items():
        sheet_name_str = str(sheet_name)
        if sheet_name_str not in wb.sheetnames:
            print(f"  Sheet '{sheet_name_str}' not found — skipping.")
            continue
        print(f"Processing sheet: {sheet_name_str} ...")
        results = process_sheet(sheet_name_str, cfg)
        all_results[sheet_name_str] = results
        write_sheet_results(wb[sheet_name_str], results, cfg)
        print(f" {sheet_name_str} written")

    write_correlation_sheet(wb, all_results)
    print(" CORRELATION MATRIX written")
    wb.save(FILE_NAME)
    print(f"\n✓ Saved: {FILE_NAME}")


if __name__ == '__main__':
    main()
