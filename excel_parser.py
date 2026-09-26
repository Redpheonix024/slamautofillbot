"""
Excel Parser for Indian Railways SLAM Bookings
"""
import os
import re
import openpyxl
from typing import List, Dict, Any, Optional
from config import SECTION_IDS, DEFAULT_OBS_TYPE, DEFAULT_JOB_KIND

KNOWN_SECTIONS = set(SECTION_IDS.keys())

def extract_loco_from_text(text: str) -> Optional[str]:
    """Extract a 5-digit locomotive number (e.g. 30389, 37931)."""
    if not text:
        return None
    m = re.search(r'\b(3\d{4})\b', text) # Electric locos commonly start with 3xxxx
    if m:
        return m.group(1)
    # Generic 5-digit check
    m2 = re.search(r'\b(\d{5})\b', text)
    if m2:
        return m2.group(1)
    return None

def normalize_sections(sec_str: str) -> List[str]:
    """
    Parse strings like 'E3A & E5A', 'M1 / M2', 'E8 / E4', 'E4/E5B', 'M2 & E3A'
    into validated SLAM section codes.
    """
    if not sec_str:
        return []
    
    # Split by common delimiters: &, /, and, comma, plus
    tokens = re.split(r'[/&+,]|\band\b', str(sec_str), flags=re.IGNORECASE)
    matched = []
    
    for token in tokens:
        cleaned = token.strip().upper()
        # Direct match
        if cleaned in KNOWN_SECTIONS:
            if cleaned not in matched:
                matched.append(cleaned)
        else:
            # Try fuzzy match in known sections
            for ks in sorted(KNOWN_SECTIONS, key=len, reverse=True):
                # E.g. 'E3A' should match before 'E3'
                if re.search(r'\b' + re.escape(ks) + r'\b', cleaned):
                    if ks not in matched:
                        matched.append(ks)
                    break
                    
    # Fallback: if nothing matched but raw text exists
    if not matched:
        raw_upper = sec_str.strip().upper()
        for ks in sorted(KNOWN_SECTIONS, key=len, reverse=True):
            if ks in raw_upper:
                if ks not in matched:
                    matched.append(ks)
                    
    return matched

def parse_excel_bookings(file_path: str) -> Dict[str, Any]:
    """
    Parses a shed bookings Excel file and returns structured details.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    filename = os.path.basename(file_path)
    loco_from_fn = extract_loco_from_text(filename)
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except (PermissionError, IOError, OSError):
        # File might be open with an exclusive lock in Excel; copy to temp file and parse
        import tempfile
        import shutil
        temp_dir = tempfile.gettempdir()
        temp_copy = os.path.join(temp_dir, f"slam_read_{os.getpid()}_{filename}")
        try:
            shutil.copy2(file_path, temp_copy)
            wb = openpyxl.load_workbook(temp_copy, data_only=True)
        finally:
            try:
                if os.path.exists(temp_copy):
                    os.remove(temp_copy)
            except Exception:
                pass
    sheet = wb.active

    loco_number = loco_from_fn
    date_str = ""
    schedule_title = ""
    
    header_row_idx = None
    col_sno = None
    col_desc = None
    col_sec = None

    # Step 1: Scan top 10 rows for Title, Date, and Table Headers
    for r in range(1, min(sheet.max_row + 1, 12)):
        row_vals = [sheet.cell(row=r, column=c).value for c in range(1, min(sheet.max_column + 1, 10))]
        
        # Check for Date and Title in Row 2 or nearby
        for val in row_vals:
            if val is not None:
                val_s = str(val).strip()
                if "date" in val_s.lower():
                    # e.g. "Date \n 31-08-26" or "Date \n15-09-26"
                    dm = re.search(r'(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})', val_s)
                    if dm:
                        date_str = dm.group(1)
                if any(kw in val_s.upper() for kw in ["BOOKING", "CHECKING", "QAI", "LBR", "VOLTAGE"]):
                    schedule_title = val_s
                    if not loco_number:
                        loco_number = extract_loco_from_text(val_s)
                        
        # Check if this row looks like the table header row
        r_sno = None
        r_desc = None
        r_sec = None

        for c, val in enumerate(row_vals, start=1):
            if val is None:
                continue
            v_str = str(val).strip()
            v_clean = re.sub(r'[^A-Z0-9]', '', v_str.upper())

            # Match S.No variations: 'SL. NO', 'SL.NO', 'SL NO', 'S.NO', 'SNO', 'SR. NO', 'SERIAL NO', 'NO'
            if v_clean in {'SNO', 'SLNO', 'SRNO', 'SERIALNO', 'NO'} or v_clean.startswith(('SLNO', 'SRNO', 'SNO')):
                r_sno = c
            # Match Section header
            elif 'SECTION' in v_clean:
                r_sec = c
            # Match Description header (skip titles with loco numbers or long working banners)
            elif any(kw in v_clean for kw in ['BOOKING', 'DESCRIPTION', 'DEFECT', 'OBSERVATION', 'REMARK', 'PARTICULAR']):
                if len(v_str) < 30 and not any(k in v_str.upper() for k in ['WORKING', 'PANTO', 'TFP', 'UNDER', 'LOCO']):
                    r_desc = c

        if (r_sno and r_desc) or (r_sno and r_sec) or (r_desc and r_sec):
            header_row_idx = r
            if r_sno: col_sno = r_sno
            if r_desc: col_desc = r_desc
            if r_sec: col_sec = r_sec
            break

    # Fallback 1: Scan for numeric S.No sequence (1, 2, 3...) if col_sno not identified
    if col_sno is None:
        for c in range(1, min(sheet.max_column + 1, 10)):
            for r in range(1, min(sheet.max_row, 12)):
                val1 = sheet.cell(row=r, column=c).value
                val2 = sheet.cell(row=r+1, column=c).value
                try:
                    if int(float(str(val1).strip())) == 1 and int(float(str(val2).strip())) == 2:
                        col_sno = c
                        if header_row_idx is None:
                            header_row_idx = r - 1
                        break
                except Exception:
                    pass
            if col_sno:
                break

    # Fallback 2: Infer missing col_desc or col_sec from first data row
    if header_row_idx and (col_desc is None or col_sec is None):
        data_row = header_row_idx + 1
        for c in range(1, min(sheet.max_column + 1, 10)):
            if c == col_sno:
                continue
            v = str(sheet.cell(row=data_row, column=c).value or '').strip()
            if not v:
                continue
            if any(sec in v.upper() for sec in KNOWN_SECTIONS) and len(v) < 15 and col_sec is None:
                col_sec = c
            elif len(v) > 8 and col_desc is None:
                col_desc = c

    # Ultimate fallback defaults
    if header_row_idx is None:
        header_row_idx = 3
    if col_sno is None:
        col_sno = 1
    if col_desc is None:
        col_desc = 2
    if col_sec is None:
        col_sec = 3

    # Ensure no column index collisions
    if col_sno == col_desc or col_desc == col_sec or col_sno == col_sec:
        if col_sno == 1:
            col_desc = 2
            col_sec = 3
        else:
            col_sno = 2
            col_desc = 3
            col_sec = 4

    # Extract schedule type if detectable in title or filename
    schedule = ""
    KNOWN_SCHEDULES = {
        "IA", "IB", "IC", "GC", "ET", "TI", "TOH", "IOH", "POH", "AOH",
        "MTR", "QAI", "LBR", "PST", "UNSCH", "M", "M1", "M2", "M12", "M24", "M36"
    }

    # Check filename first, e.g. "30389 - GC - 15-09- 26" -> "GC", "37229 - ET - 20-09-26" -> "ET"
    fn_parts = [p.strip() for p in re.split(r'[-_]', filename) if p.strip()]
    for p in fn_parts:
        up = p.upper()
        if up in KNOWN_SCHEDULES:
            schedule = up
            break
            
    if not schedule and schedule_title:
        title_parts = [p.strip() for p in re.split(r'[-_]', schedule_title) if p.strip()]
        for p in title_parts:
            up = p.upper()
            if up in KNOWN_SCHEDULES:
                schedule = up
                break
        if not schedule and len(title_parts) >= 2:
            schedule = title_parts[-1]

    # Smart fallback: short alphabetic token between hyphens
    if not schedule:
        for p in fn_parts:
            up = p.upper()
            if re.match(r'^[A-Z]{2,4}$', up) and not any(kw in up for kw in ["DATE", "BOOK", "XLSX", "SLIP", "TEST", "LOCO", "SHED"]):
                schedule = up
                break

    bookings = []
    current_category = "General"

    # Step 2: Read booking rows
    for r in range(header_row_idx + 1, sheet.max_row + 1):
        sno_val = sheet.cell(row=r, column=col_sno).value
        desc_val = sheet.cell(row=r, column=col_desc).value
        sec_val = sheet.cell(row=r, column=col_sec).value

        if desc_val is None and sec_val is None and sno_val is None:
            continue

        desc_str = str(desc_val).strip() if desc_val is not None else ""
        sec_str = str(sec_val).strip() if sec_val is not None else ""

        # Category headers (e.g., S.No is empty, but description is like "UG / MISC BOOKING")
        if sno_val is None or str(sno_val).strip() == "":
            if desc_str and not sec_str:
                current_category = desc_str
            continue

        # Check if S.No is numeric
        sno_int = None
        try:
            sno_int = int(float(str(sno_val).strip()))
        except (ValueError, TypeError):
            # Not a data row
            continue

        if not desc_str:
            continue

        # Normalize sections
        matched_secs = normalize_sections(sec_str)

        bookings.append({
            "sno": sno_int,
            "description": desc_str,
            "sections": matched_secs,
            "raw_section": sec_str,
            "category": current_category,
            "obs_type": DEFAULT_OBS_TYPE,
            "kind": DEFAULT_JOB_KIND
        })

    return {
        "file_name": filename,
        "loco_number": loco_number or "Unknown",
        "date": date_str,
        "schedule_title": schedule_title,
        "schedule": schedule,
        "bookings_count": len(bookings),
        "bookings": bookings
    }
