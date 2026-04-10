import zipfile
import os
import re
import pandas as pd
import tempfile
import subprocess


def process_zip(zip_path, threshold_pct=50):
    subject_slow  = {}
    subject_total = {}
    all_slow      = set()
    table_rows    = []
    processed     = 0
    errors        = []

    soffice_path = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

    with zipfile.ZipFile(zip_path, "r") as z:
        for entry in sorted(z.namelist()):
            filename = os.path.basename(entry)

            if not filename.lower().endswith((".xls", ".xlsx")) or filename.startswith("._") or entry.endswith("/"):
                continue

            try:
                data = z.read(entry)
                if not data:
                    errors.append(f"{filename}: File is empty")
                    continue

                # -------- CONVERT USING LIBREOFFICE --------
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        input_path = os.path.join(tmpdir, filename)

                        with open(input_path, "wb") as f:
                            f.write(data)

                        subprocess.run(
                            [
                                soffice_path,
                                "--headless",
                                "--nologo",
                                "--nolockcheck",
                                "--nodefault",
                                "--nofirststartwizard",
                                "--convert-to", "xlsx",
                                input_path,
                                "--outdir", tmpdir
                            ],
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            timeout=10
                        )

                        base = os.path.splitext(filename)[0]
                        converted_path = os.path.join(tmpdir, base + ".xlsx")

                        if not os.path.exists(converted_path):
                            errors.append(f"{filename}: conversion failed")
                            continue

                        df = pd.read_excel(converted_path, engine="openpyxl", header=None)

                except Exception:
                    errors.append(f"{filename}: conversion error")
                    continue

            except Exception as e:
                errors.append(f"{filename}: {str(e)}")
                continue

            # -------- SUBJECT EXTRACTION --------
            subject = "Unknown Subject"
            try:
                for i in range(10):
                    for cell in df.iloc[i]:
                        if isinstance(cell, str) and "course name" in cell.lower():
                            subject = cell.split(":", 1)[1].strip().title()
                            break
            except:
                pass

            # -------- FIND CIE1 COLUMN --------
            cie1_col, cmax = None, 10
            try:
                header_row = df.iloc[10]

                for j, h in enumerate(header_row):
                    h_str = str(h).strip().upper()

                    if h_str.startswith("CIE1"):
                        cie1_col = j

                        m = re.search(r"\((\d+)\)", h_str)
                        if m:
                            cmax = int(m.group(1))
                        break
            except:
                pass

            if cie1_col is None:
                errors.append(f"{filename}: CIE1 column not found")
                continue

            cutoff = cmax * (threshold_pct / 100)

            subject_slow.setdefault(subject, set())
            subject_total.setdefault(subject, 0)

            # -------- PROCESS STUDENTS --------
            for i in range(11, len(df)):
                try:
                    usn  = str(df.iloc[i, 2]).strip().upper()
                    name = str(df.iloc[i, 1]).strip()

                    if not usn or len(usn) < 5 or name.lower() in ("name", "nan", ""):
                        continue

                    raw_val = df.iloc[i, cie1_col]
                    raw_str = str(raw_val).strip().upper()

                    absent = raw_str in ("A", "NE", "NL", "", "NAN")

                    # ❌ Ignore absentees completely
                    if absent:
                        continue

                    try:
                        marks = float(raw_val)
                    except:
                        continue

                    subject_total[subject] += 1

                    is_slow = marks < cutoff

                    if is_slow:
                        subject_slow[subject].add(usn)
                        all_slow.add(usn)

                    table_rows.append({
                        "usn": usn,
                        "name": name,
                        "subject": subject,
                        "marks": marks,
                        "absent": False,
                        "slow": is_slow,
                        "cmax": cmax,
                    })

                except:
                    continue

            processed += 1

    # -------- SORT RESULTS --------
    table_rows.sort(
        key=lambda r: (
            not r["slow"],
            float(r["marks"]) if r["marks"] is not None else -1
        )
    )

    return {
        "subject_counts": {s: len(v) for s, v in subject_slow.items()},
        "subject_totals": subject_total,
        "rows": table_rows,
        "total_slow": len(all_slow),
        "processed": processed,
        "errors": errors,
    }