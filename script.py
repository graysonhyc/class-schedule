import io
import re
import csv
from typing import Dict, List, Tuple, Set
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment


# ---------- Helpers for File A (timetable_by_day.xlsx) ----------

def _clean_text(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    # treat dashes / empty markers as empty
    if s in {"-", "—", "–", "－"} or len(s.strip(" -—–－")) == 0:
        return ""
    # collapse whitespace
    s = re.sub(r"\s+", " ", s)
    return s

def get_class_columns(ws: Worksheet) -> Dict[str, int]:
    """
    Return a mapping of class_id -> data_col_index.
    Uses merged header ranges on the header row (assumed row 1).
    If header is merged (2+ columns), use the RIGHT-MOST column for data.
    If single cell, use its column.
    Ignores column A (day/period column).
    """
    header_row = 1
    class_cols: Dict[str, int] = {}

    # Track columns already consumed by a merged range so we don't double-count
    consumed_cols: Set[int] = set()

    # 1) scan merged headers on row 1
    for mr in ws.merged_cells.ranges:
        min_row, max_row = mr.min_row, mr.max_row
        min_col, max_col = mr.min_col, mr.max_col
        if min_row == header_row and max_row == header_row:
            header_val = _clean_text(ws.cell(header_row, min_col).value)
            if header_val and min_col != 1:  # skip column A
                class_cols[header_val] = max_col
                for c in range(min_col, max_col + 1):
                    consumed_cols.add(c)

    # 2) handle unmerged, single header cells on row 1
    last_col = ws.max_column
    for col in range(2, last_col + 1):  # start from col 2, skip A
        if col in consumed_cols:
            continue
        header_val = _clean_text(ws.cell(header_row, col).value)
        if header_val:
            class_cols[header_val] = col

    return class_cols

def get_period_blocks(ws: Worksheet) -> List[Tuple[int, int, str]]:
    """
    Find period blocks from column A.
    Primary: detect merged ranges in col A; each block spans 3 rows and has a number (1,2,3,...).
    Fallback: scan column A and infer 3-row groups whenever we see integers, grab the next two rows.
    Returns list of (start_row, end_row, period_str).
    """
    col_a = 1
    blocks: List[Tuple[int, int, str]] = []

    # 1) via merged ranges
    found_any = False
    for mr in ws.merged_cells.ranges:
        if mr.min_col == col_a and mr.max_col == col_a:
            start, end = mr.min_row, mr.max_row
            val = _clean_text(ws.cell(start, col_a).value)
            if val.isdigit() and (end - start + 1) >= 3:
                # Use exactly a 3-row logical block from the start
                blocks.append((start, start + 2, val))
                found_any = True

    if found_any:
        # sort by start row
        blocks.sort(key=lambda t: t[0])
        return blocks

    # 2) fallback: scan col A and group any digit-cell as a 3-row block
    max_row = ws.max_row
    r = 1
    while r <= max_row:
        val = _clean_text(ws.cell(r, col_a).value)
        if val.isdigit() and r + 2 <= max_row:
            blocks.append((r, r + 2, val))
            r += 3
        else:
            r += 1

    return blocks

def extract_mapping(ws: Worksheet,
                    class_cols: Dict[str, int],
                    period_blocks: List[Tuple[int, int, str]]
                   ) -> Dict[str, str]:
    """
    For each class and each 3-row period block:
      - Subject = row1 of block, Teacher = row3 of block in that class's data col
    Build key = f"{class_id}{period}", suffix = f"({subject} {teacher})" (teacher optional).
    """
    mapping: Dict[str, str] = {}

    for class_id, data_col in class_cols.items():
        for start_row, end_row, period in period_blocks:
            # Normalize: ensure 3 rows
            r1, r3 = start_row, start_row + 2
            subject = _clean_text(ws.cell(r1, data_col).value)
            teacher = _clean_text(ws.cell(r3, data_col).value)

            if not subject and not teacher:
                continue

            if subject and teacher:
                suffix = f"({subject} {teacher})"
            elif subject:
                suffix = f"({subject})"
            else:
                # Only teacher (rare) — still include
                suffix = f"({teacher})"

            key = f"{class_id}{period}"
            mapping[key] = suffix
    return mapping

def build_full_mapping(file_a_bytes: bytes) -> Dict[str, str]:
    wb = load_workbook(io.BytesIO(file_a_bytes), data_only=True)
    overall: Dict[str, str] = {}
    for ws in wb.worksheets:
        class_cols = get_class_columns(ws)
        period_blocks = get_period_blocks(ws)
        m = extract_mapping(ws, class_cols, period_blocks)
        # last one wins if duplicates
        overall.update(m)
    return overall


# ---------- Helpers for File B (daily_schedule.xlsx) ----------

TOKEN_RE = re.compile(r"\b(\d+[A-Z])(\d{1,2})\b")

def annotate_schedule(wb: Workbook,
                      mapping: Dict[str, str]
                     ) -> Tuple[int, List[Tuple[str, str, str, str]]]:
    """
    Iterate all sheets and cells; append mapped suffix lines for each matched class+period token.
    Returns (cells_updated_count, unmatched_rows).
    unmatched_rows: list of (sheet, cell_address, missing_key, original_text)
    """
    unmatched: List[Tuple[str, str, str, str]] = []
    cells_updated = 0

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                text = cell.value
                matches = list(TOKEN_RE.finditer(text))
                if not matches:
                    continue

                # Deduplicate while preserving order
                seen_keys: Set[str] = set()
                suffixes: List[str] = []

                for m in matches:
                    class_id = m.group(1)  # e.g., "1A"
                    period = m.group(2)    # e.g., "1"
                    key = f"{class_id}{period}"
                    if key in seen_keys:
                        continue
                    seen_keys.add(key)
                    if key in mapping:
                        suffixes.append(mapping[key])
                    else:
                        unmatched.append((ws.title, cell.coordinate, key, text))

                if suffixes:
                    new_val = text + "\n" + "\n".join(suffixes)
                    # preserve existing alignment props if any
                    alg = cell.alignment or Alignment()
                    cell.value = new_val
                    cell.alignment = Alignment(
                        horizontal=alg.horizontal,
                        vertical=alg.vertical,
                        wrapText=True
                    )
                    cells_updated += 1

    return cells_updated, unmatched


# ---------- Streamlit UI ----------

st.set_page_config(page_title="Schedule Annotator", page_icon="📚", layout="centered")
st.title("📚 Class Schedule Annotator")
st.caption("Upload the weekly timetable (5 sheets) and the daily schedule. "
           "The app will annotate schedule cells like `1A1` with `(科目 老師)` on a new wrapped line.")

st.markdown("**1) Upload File A (weekly timetable with 5 sheets)**")
file_a = st.file_uploader("school_timetable.xlsx", type=["xlsx"], key="file_a")

st.markdown("**2) Upload File B (daily schedule to annotate)**")
file_b = st.file_uploader("september_st_timetable.xlsx", type=["xlsx"], key="file_b")

if st.button("Run Annotation", type="primary", disabled=not (file_a and file_b)):
    try:
        with st.spinner("Parsing File A and building mapping…"):
            mapping = build_full_mapping(file_a.read())

        st.success(f"Built mapping with **{len(mapping)}** class-period entries.")

        # Reload file_b (Streamlit's file-like can be read only once reliably)
        file_b.seek(0)
        wb_b = load_workbook(file_b, data_only=True)

        with st.spinner("Annotating File B…"):
            updated_count, unmatched = annotate_schedule(wb_b, mapping)

        # Prepare outputs
        out_xlsx = io.BytesIO()
        wb_b.save(out_xlsx)
        out_xlsx.seek(0)

        # unmatched CSV
        csv_buf = io.StringIO()
        writer = csv.writer(csv_buf)
        writer.writerow(["sheet", "cell", "missing_key", "cell_text"])
        for row in unmatched:
            writer.writerow(row)
        csv_bytes = csv_buf.getvalue().encode("utf-8")

        st.success(f"Updated **{updated_count}** cells. Unmatched keys: **{len(unmatched)}**")

        st.download_button(
            "⬇️ Download annotated schedule (Excel)",
            data=out_xlsx,
            file_name="daily_schedule_annotated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "⬇️ Download unmatched_keys.csv",
            data=csv_bytes,
            file_name="unmatched_keys.csv",
            mime="text/csv",
        )

        # Optional: small peek at a few mapping samples
        if st.checkbox("Preview a few mapping samples"):
            sample = list(mapping.items())[:25]
            st.write({k: v for k, v in sample})

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)


st.markdown("---")
st.markdown("### Notes")
st.markdown("""
- **File A assumptions**:  
  - Row 1 has class headers (each header may be a merged range of 2 columns).  
  - Column A has the period numbers, merged vertically in **3-row blocks** (1/2/3/4/5…).  
  - Within each 3-row block, **row 1 = Subject**, **row 3 = Teacher** for each class's data column (the right-most col of the merged header).
- **File B**: Any text containing `Year+Class+Period` like `1A1`, `2C19` will get a new wrapped line with the suffix `(科目 老師)`.  
- Cells like `2A/2C` (no period number) are not modified.
""")