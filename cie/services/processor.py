import zipfile
import os
import re
import io
import pandas as pd

def process_zip(zip_path, threshold_pct=50):
    subject_slow  = {}
    subject_total = {}
    all_slow      = set()
    table_rows    = []
    processed     = 0
    errors        = []

    with zipfile.ZipFile(zip_path, "r") as z:
        for entry in sorted(z.namelist()):
            filename = os.path.basename(entry)
            
            # Skip system files and non-excel files
            if not filename.lower().endswith((".xls", ".xlsx")) or filename.startswith("._") or entry.endswith("/"):
                continue

            try:
                data = z.read(entry)
                if not data:
                    errors.append(f"{filename}: File is empty")
                    continue

                # Use pandas to bypass "seen[2] == 4" corruption errors
                # engine='xlrd' is used for .xls files
                engine = 'xlrd' if filename.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(io.BytesIO(data), engine=engine, header=None)
                
            except Exception as e:
                errors.append(f"{filename}: {str(e)}")
                continue

            # --- Data Extraction Logic ---
            subject = "Unknown Subject"
            try:
                if len(df) > 6:
                    cell_val = str(df.iloc[6, 2])
                    if ":" in cell_val:
                        subject = cell_val.split(":", 1)[1].strip().title()
            except: pass

            cie1_col, cmax = None, 10
            try:
                if len(df) > 10:
                    header_row = df.iloc[10]
                    for j, h in enumerate(header_row):
                        h_str = str(h).strip().upper()
                        if h_str.startswith("CIE1"):
                            cie1_col = j
                            m = re.search(r"\((\d+)\)", h_str)
                            if m: cmax = int(m.group(1))
                            break
            except: pass

            if cie1_col is None:
                errors.append(f"{filename}: CIE1 column not found")
                continue

            cutoff = cmax * (threshold_pct / 100)
            subject_slow.setdefault(subject, set())
            subject_total.setdefault(subject, 0)

            # --- Process Student Rows (Starts at Row 11) ---
            for i in range(11, len(df)):
                try:
                    usn = str(df.iloc[i, 2]).strip()
                    name = str(df.iloc[i, 1]).strip()
                    
                    if not usn or len(usn) < 5 or name.lower() in ("name", "nan", ""):
                        continue

                    raw_val = df.iloc[i, cie1_col]
                    raw_str = str(raw_val).strip().upper()
                    absent = raw_str in ("A", "NE", "NL", "", "NAN")
                    
                    marks = None
                    if not absent:
                        try:
                            marks = float(raw_val)
                        except: continue

                    subject_total[subject] += 1
                    is_slow = absent or (marks is not None and marks < cutoff)

                    if is_slow:
                        subject_slow[subject].add(usn)
                        all_slow.add(usn)

                    table_rows.append({
                        "usn": usn, "name": name, "subject": subject,
                        "marks": marks, "absent": absent, "slow": is_slow, "cmax": cmax,
                    })
                except: continue

            processed += 1

    # Sort results
    table_rows.sort(key=lambda r: (not r["slow"], float(r["marks"]) if r["marks"] is not None else -1))

    return {
        "subject_counts": {s: len(v) for s, v in subject_slow.items()},
        "subject_totals": subject_total,
        "rows": table_rows,
        "total_slow": len(all_slow),
        "processed": processed,
        "errors": errors,
    }