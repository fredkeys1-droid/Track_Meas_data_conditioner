import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
from pathlib import Path
import io

# ============================================================
# Streamlit UI Setup
# ============================================================
st.set_page_config(page_title="Railway Survey Engine", page_icon="🚆", layout="centered")
st.title("🚆 Railway Survey Consolidation Engine")
st.markdown("""
**Supported Formats:**
* **Sydney Trains:** TCM / FSheets (`.docx` or `.xlsx`)
* **ARTC M-Sheets:** Track Design / Centres / Offset metrics (`.xlsx`)
* **UGL AR Sheets:** 'TCM- Design' Offset and 'Location - KMs' (`.xlsx`)
""")

uploaded_files = st.file_uploader("Upload Survey Files (.docx, .xlsx)", accept_multiple_files=True)

# ============================================================
# Core Processing Logic
# ============================================================
if uploaded_files and st.button("Process Files", type="primary"):
    
    with st.spinner("Analyzing and consolidating track data..."):
        tcm_dfs = {}
        fsheet_dfs = {}
        msheet_flags = {}
        feature_keywords = ["TP", "TRS", "CTRS", "CTP", "Ea", "Bend"]
        
        STANDARD_COLUMNS = [
            "Survey Mark Number", "Point Number", "Mark Type", "Reduced Level",
            "Description", "Metrage", "Offset", "Lift", "Super", "Centres",
            "Track", "Job Reference", "Comments", "Images"
        ]

        def clean_track_name(track):
            track = track.replace("Sydney Platform No5 Road Platform", "Sydney Plat 5 Rd")
            track = track.replace("Sydney Platform No6 Road Platform", "Sydney Plat 6 Rd")
            return track.strip()

        def get_track_name(text):
            if "For " not in text:
                return text
            return text.split("For ")[1].split(" (Basecode")[0]

        def robust_find_column(headers_upper, primary_keywords, secondary_keywords):
            for idx, h in enumerate(headers_upper):
                if any(p in h for p in primary_keywords) and "ERROR" not in h: return idx
            for idx, h in enumerate(headers_upper):
                if any(s in h for s in secondary_keywords) and "ERROR" not in h: return idx
            for idx, h in enumerate(headers_upper):
                if any(p in h for p in primary_keywords): return idx
            return None

        # --- Parse Uploaded Files ---
        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            file_ext = Path(filename).suffix.lower()

            if file_ext == ".xlsx":
                xls = pd.ExcelFile(uploaded_file)
                for sheet in xls.sheet_names:
                    try:
                        raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, dtype=str).fillna("")
                        
                        combined_headers = []
                        if len(raw) > 12:
                            row11 = [str(x).strip() for x in raw.iloc[11]]
                            row12 = [str(x).strip() for x in raw.iloc[12]]
                            current_parent = ""
                            for p_val, c_val in zip(row11, row12):
                                if p_val: current_parent = p_val
                                combined_val = f"{current_parent} - {c_val}" if current_parent and c_val else (p_val or c_val or "")
                                combined_headers.append(combined_val.strip())

                        text_blob = " ".join(raw.astype(str).values.flatten()) + " " + " ".join(combined_headers)
                        text_upper = text_blob.upper()
                        is_ugl_ar = "TCM- DESIGN" in text_upper or "LOCATION - KMS" in text_upper or "TCM - DESIGN" in text_upper
                        is_msheet = "M SHEETS" in text_upper or "TRACK:" in text_upper

                        if is_ugl_ar:
                            track_name = ""
                            for r in range(min(len(raw), 15)):
                                for c in range(len(raw.columns)):
                                    val = str(raw.iat[r, c]).strip().upper()
                                    if "TRACK:" in val or "LINE:" in val:
                                        if c + 1 < len(raw.columns):
                                            track_name = str(raw.iat[r, c+1]).strip()
                                            break
                                    elif len(val) > 0 and (val.startswith("TRW") or "MAIN" in val or "LINE" in val) and c == 0:
                                        track_name = val
                                if track_name: break

                            track = clean_track_name(track_name) if track_name else sheet
                            msheet_flags[track] = True
                            headers_upper = [h.upper() for h in combined_headers]

                            col_mapping = {}
                            for idx, h in enumerate(headers_upper):
                                if "FRAME - POINT" in h: col_mapping["Mark Type"] = idx
                                elif "TCM - MARK" in h or "TCM - OHWS" in h or ("MARK" in h and "TCM" in h):
                                    col_mapping["Survey Mark Number"] = idx
                                    if "Mark Type" not in col_mapping: col_mapping["Mark Type"] = idx
                                elif "LOCATION - KMS" in h or "LOCATION - KM" in h: col_mapping["Metrage"] = idx
                                elif "TCM - DESIGN" in h or "TCM - OFFSET" in h or "TCM - DES" in h: col_mapping["Offset"] = idx
                                elif "SUPERELEVATION - DESIGN" in h or "SUPER - DESIGN" in h or "SUPERELEVATION - DES" in h: col_mapping["Super"] = idx
                                elif "CENTRES - DESIGN" in h or "CENTRES - DES" in h or "TRACK CENTRES" in h: col_mapping["Centres"] = idx

                            for idx, h in enumerate(headers_upper):
                                if "Metrage" not in col_mapping and ("KMS" in h or "KM" in h): col_mapping["Metrage"] = idx
                                if "Offset" not in col_mapping and "OFFSET" in h: col_mapping["Offset"] = idx
                                if "Super" not in col_mapping and "SUPER" in h: col_mapping["Super"] = idx
                                if "Centres" not in col_mapping and "CENTRES" in h: col_mapping["Centres"] = idx
                                if "Survey Mark Number" not in col_mapping and "MARK" in h: col_mapping["Survey Mark Number"] = idx
                                if "Mark Type" not in col_mapping and "POINT" in h: col_mapping["Mark Type"] = idx

                            data_rows = raw.iloc[13:].copy()
                            parsed_df = pd.DataFrame()
                            for std_col in ["Survey Mark Number", "Mark Type", "Metrage", "Super", "Centres", "Offset"]:
                                if std_col in col_mapping: parsed_df[std_col] = data_rows.iloc[:, col_mapping[std_col]]
                                else: parsed_df[std_col] = np.nan

                            parsed_df["Metrage"] = pd.to_numeric(parsed_df["Metrage"], errors="coerce") * 1000.0
                            parsed_df = parsed_df.dropna(subset=["Metrage"])
                            parsed_df["Track"] = track

                            for col in ["Super", "Centres", "Offset"]:
                                if col in parsed_df.columns:
                                    cleaned_series = parsed_df[col].astype(str).str.strip().replace({"?": "0", "": "0", "NaN": "0", "nan": "0"})
                                    num_vals = pd.to_numeric(cleaned_series, errors="coerce").fillna(0)
                                    if col == "Super": parsed_df[col] = num_vals.round()
                                    else: parsed_df[col] = num_vals

                            tcm_dfs[track] = parsed_df
                            continue

                        if is_msheet:
                            track_name = ""
                            for r in range(min(len(raw), 15)):
                                for c in range(len(raw.columns)):
                                    val = str(raw.iat[r, c]).strip().upper()
                                    if "TRACK:" in val:
                                        if c + 1 < len(raw.columns):
                                            track_name = str(raw.iat[r, c+1]).strip()
                                            break
                                if track_name: break

                            track = clean_track_name(track_name) if track_name else sheet
                            msheet_flags[track] = True

                            header_row = None
                            for r in range(len(raw)):
                                row_vals_upper = [str(x).strip().upper() for x in raw.iloc[r]]
                                if any("METRAGE" in val for val in row_vals_upper) and any("MARK" in val or "OHWS" in val for val in row_vals_upper):
                                    header_row = r
                                    break
                            if header_row is None: continue

                            headers_upper = [h.upper() for h in [str(x).strip() for x in raw.iloc[header_row]]]
                            col_mapping = {}
                            for idx, h in enumerate(headers_upper):
                                if "MARK" in h or "OHWS" in h:
                                    col_mapping["Survey Mark Number"] = idx
                                    col_mapping["Mark Type"] = idx
                                elif "METRAGE" in h: col_mapping["Metrage"] = idx

                            super_idx = robust_find_column(headers_upper, ["EA SUPER DESIGN", "SUPER DESIGN", "EA SUPER 1ST", "SUPER 1ST", "SUPER FINAL"], ["SUPER", "EA SUPER"])
                            if super_idx is not None: col_mapping["Super"] = super_idx

                            centres_idx = robust_find_column(headers_upper, ["CTR'S DESIGN", "CENTRES DESIGN", "CTR'S 1ST", "CENTRES 1ST", "CTR'S FINAL"], ["CTR'S", "CENTRES"])
                            if centres_idx is not None: col_mapping["Centres"] = centres_idx

                            offset_idx = robust_find_column(headers_upper, ["TCM TO NEAR RAIL DESIGN", "TCM TO NEAR LOW RAIL", "TAP TO Near Rail Design", "TAP TO NEAR", "TCM TO NEAR RAIL 1ST", "TCM TO NEAR RAIL FINAL"], ["TCM TO NEAR", "TAP TO NEAR", "TAP", "OFFSET"])
                            if offset_idx is not None: col_mapping["Offset"] = offset_idx

                            data_rows = raw.iloc[header_row + 2:].copy()
                            parsed_df = pd.DataFrame()
                            for std_col in ["Survey Mark Number", "Mark Type", "Metrage", "Super", "Centres", "Offset"]:
                                if std_col in col_mapping: parsed_df[std_col] = data_rows.iloc[:, col_mapping[std_col]]
                                else: parsed_df[std_col] = np.nan

                            parsed_df["Metrage"] = pd.to_numeric(parsed_df["Metrage"], errors="coerce")
                            parsed_df = parsed_df.dropna(subset=["Metrage"])
                            parsed_df["Track"] = track

                            for col in ["Super", "Centres", "Offset"]:
                                if col in parsed_df.columns:
                                    cleaned_series = parsed_df[col].astype(str).str.strip().replace({"?": "0", "": "0", "NaN": "0", "nan": "0"})
                                    num_vals = pd.to_numeric(cleaned_series, errors="coerce").fillna(0)
                                    if col == "Super": parsed_df[col] = num_vals.round()
                                    else: parsed_df[col] = num_vals

                            tcm_dfs[track] = parsed_df
                            continue

                        tcm_score = sum(k in text_upper for k in ["SURVEY MARK NUMBER", "POINT NUMBER", "REDUCED LEVEL", "METRAGE"])
                        fs_score = sum(k in text_upper for k in ["SUPER RAMP", "RADIUS", "LOCATION KM", "TP", "TRS"])

                        if tcm_score >= 3: sheet_type = "TCM"
                        elif fs_score >= 3: sheet_type = "FSHEET"
                        else: continue

                        title = None
                        for r in range(min(len(raw), 20)):
                            for c in range(min(len(raw.columns), 10)):
                                val = str(raw.iat[r, c]).strip()
                                if val.startswith("Track Control Marks For") or val.startswith("FSheet For"):
                                    title = val
                                    break
                            if title: break

                        track = clean_track_name(get_track_name(title)) if title else sheet

                        if sheet_type == "TCM":
                            header_row = None
                            for r in range(len(raw)):
                                row_vals = [str(v).strip() for v in raw.iloc[r]]
                                if "Survey Mark Number" in row_vals and "Metrage" in row_vals:
                                    header_row = r
                                    break
                            if header_row is None: continue

                            headers = list(raw.iloc[header_row])
                            unique_headers = []
                            seen = {}
                            for col in headers:
                                col = str(col).strip()
                                if col not in seen:
                                    seen[col] = 1
                                    unique_headers.append(col)
                                else:
                                    seen[col] += 1
                                    unique_headers.append(f"{col}_{seen[col]}")

                            data = raw.iloc[header_row + 1:].copy()
                            data.columns = unique_headers
                            data = data.reset_index(drop=True)

                            if "Metrage" in data.columns: data["Metrage"] = pd.to_numeric(data["Metrage"], errors="coerce")
                            if "Super" in data.columns: data["Super" ] = pd.to_numeric(data["Super"], errors="coerce")
                            tcm_dfs[track] = data

                        elif sheet_type == "FSHEET":
                            header_row = None
                            for r in range(len(raw)):
                                row_vals = [str(v).strip() for v in raw.iloc[r]]
                                if "Super" in row_vals and "KM" in row_vals:
                                    header_row = r
                                    break
                            if header_row is None: continue

                            headers = list(raw.iloc[header_row])
                            unique_headers = []
                            seen = {}
                            for col in headers:
                                if col not in seen:
                                    seen[col] = 1
                                    unique_headers.append(col)
                                else:
                                    seen[col] += 1
                                    unique_headers.append(f"{col}_{seen[col]}")

                            data = raw.iloc[header_row + 2:].copy()
                            data.columns = unique_headers
                            data.columns = ["Feature" if str(col).strip() == "" else col for col in data.columns]
                            data = data.reset_index(drop=True)
                            fsheet_dfs[track] = data

                    except Exception as ex:
                        st.error(f"Error reading sheet {sheet}: {ex}")

            elif file_ext == ".docx":
                doc = Document(uploaded_file)
                tcm_sections = {}
                fsheet_tables = {}

                for i, table in enumerate(doc.tables):
                    try: heading = table.cell(0, 0).text.strip()
                    except: continue

                    if heading.startswith("Track Control Marks For"):
                        track = clean_track_name(get_track_name(heading))
                        if i + 1 < len(doc.tables):
                            parent_table = doc.tables[i + 1]
                            for row in parent_table.rows:
                                for cell in row.cells:
                                    for nested in cell.tables:
                                        rows = []
                                        for r in nested.rows:
                                            rows.append([c.text.replace("\n", " ").strip() for c in r.cells])
                                        tcm_sections.setdefault(track, []).append(rows)

                    elif heading.startswith("FSheet For"):
                        track = clean_track_name(get_track_name(heading))
                        if i + 1 < len(doc.tables):
                            ftable = doc.tables[i + 1]
                            rows = []
                            for r in ftable.rows:
                                rows.append([c.text.replace("\n", " ").strip() for c in r.cells])
                            fsheet_tables[track] = rows

                for track, sections in tcm_sections.items():
                    merged = []
                    first = True
                    for section in sections:
                        if first:
                            merged.extend(section)
                            first = False
                        else: merged.extend(section[4:])
                    if merged:
                        headers = merged[1]
                        data = merged[4:]
                        df = pd.DataFrame(data, columns=headers)
                        if "Metrage" in df.columns: df["Metrage"] = pd.to_numeric(df["Metrage"], errors="coerce")
                        if "Super" in df.columns: df["Super"] = pd.to_numeric(df["Super"], errors="coerce")
                        tcm_dfs[track] = df

                for track, rows in fsheet_tables.items():
                    try:
                        headers = rows[0]
                        unique_headers = []
                        seen = {}
                        for col in headers:
                            if col not in seen:
                                seen[col] = 1
                                unique_headers.append(col)
                            else:
                                seen[col] += 1
                                unique_headers.append(f"{col}_{seen[col]}")
                        data = rows[4:]
                        df = pd.DataFrame(data, columns=unique_headers)
                        fsheet_dfs[track] = df
                    except Exception as ex:
                        st.error(f"Error processing FSheet for '{track}': {ex}")

        # --- Consolidation and Export ---
        all_tracks = sorted(set(tcm_dfs.keys()) | set(fsheet_dfs.keys()))
        
        if not all_tracks:
            st.warning("No track data found in the uploaded files.")
        else:
            st.success("Files processed successfully! Download your generated track files below.")
            
            for idx, track in enumerate(all_tracks):
                is_msheet_source = msheet_flags.get(track, False)

                tcm_df = tcm_dfs[track].copy() if track in tcm_dfs else pd.DataFrame(columns=STANDARD_COLUMNS)
                if "Super" in tcm_df.columns:
                    tcm_df["Super"] = pd.to_numeric(tcm_df["Super"], errors="coerce").fillna(0).round()

                if track in fsheet_dfs and not is_msheet_source:
                    fsheet_df = fsheet_dfs[track].copy()
                    if "Super" in fsheet_df.columns:
                        fsheet_df["Super"] = pd.to_numeric(fsheet_df["Super"], errors="coerce").fillna(0).round()

                    km_columns = [c for c in fsheet_df.columns if str(c).startswith("KM")]
                    if len(km_columns) == 0:
                        combined = tcm_df
                    else:
                        km_col = km_columns[0]
                        fsheet_df[km_col] = pd.to_numeric(fsheet_df[km_col], errors="coerce")
                        fsheet_filtered = fsheet_df[fsheet_df["Super"].notna()].copy()

                        feature_col = "Feature" if "Feature" in fsheet_df.columns else None
                        if feature_col is None:
                            for col in fsheet_df.columns:
                                if fsheet_filtered[col].astype(str).str.strip().isin(feature_keywords).any():
                                    feature_col = col
                                    break

                        base_columns = list(tcm_df.columns) if len(tcm_df.columns) > 0 else STANDARD_COLUMNS
                        mapped_rows = pd.DataFrame(columns=base_columns)
                        mapped_rows["Super"] = fsheet_filtered["Super"]
                        mapped_rows["Metrage"] = fsheet_filtered[km_col] * 1000
                        mapped_rows["Mark Type"] = fsheet_filtered[feature_col].fillna("").astype(str).str.strip() if feature_col else ""

                        transition_rows = []
                        for i in range(len(fsheet_filtered) - 1):
                            try:
                                curr = fsheet_filtered.iloc[i]
                                nxt = fsheet_filtered.iloc[i + 1]
                                start_m = int(float(curr[km_col]) * 1000)
                                end_m = int(float(nxt[km_col]) * 1000)
                                start_s = float(curr["Super"])
                                end_s = float(nxt["Super"])
                                if start_s == end_s or end_m <= start_m: continue

                                next_interval = ((start_m // 5) + 1) * 5
                                while next_interval < end_m:
                                    proportion = (next_interval - start_m) / (end_m - start_m)
                                    new_super = round(start_s + ((end_s - start_s) * proportion))
                                    row = {c: np.nan for c in tcm_df.columns}
                                    row["Metrage"] = next_interval
                                    row["Super"] = new_super
                                    row["Mark Type"] = ""
                                    transition_rows.append(row)
                                    next_interval += 5
                            except: pass

                        interp_df = pd.DataFrame(transition_rows)
                        combined = pd.concat([tcm_df, mapped_rows, interp_df], ignore_index=True).drop_duplicates(subset=["Metrage"], keep="first")
                        combined["Metrage"] = pd.to_numeric(combined["Metrage"], errors="coerce")
                        combined = combined.sort_values(by="Metrage")

                        filler_rows = []
                        rows = combined.to_dict(orient="records")
                        for j in range(len(rows) - 1):
                            current = rows[j]
                            nxt = rows[j + 1]
                            try:
                                start_m = int(float(current["Metrage"]))
                                end_m = int(float(nxt["Metrage"]))
                            except: continue
                            gap = end_m - start_m
                            if gap <= 20: continue

                            next_20 = ((start_m // 20) + 1) * 20
                            while next_20 < end_m:
                                row = {c: np.nan for c in combined.columns}
                                row["Metrage"] = next_20
                                row["Super"] = current.get("Super", np.nan)
                                row["Mark Type"] = ""
                                filler_rows.append(row)
                                next_20 += 20

                        combined = pd.concat([combined, pd.DataFrame(filler_rows)], ignore_index=True)
                else:
                    combined = tcm_df

                if "Metrage" in combined.columns:
                    combined["Metrage"] = pd.to_numeric(combined["Metrage"], errors="coerce")
                    combined = combined.sort_values(by="Metrage", ascending=True)

                for col in STANDARD_COLUMNS:
                    if col not in combined.columns: combined[col] = np.nan

                combined = combined[[c for c in STANDARD_COLUMNS if c in combined.columns] + [c for c in combined.columns if c not in STANDARD_COLUMNS]]

                clean_name = f"Unnamed_Track_{idx + 1}" if not track or str(track).strip() == "" else str(track).replace("/", "_").replace("\\", "_").strip()
                output_filename = f"Track Meas {clean_name} data.xlsx"

                # Write directly to an in-memory buffer instead of saving to disk
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    combined.to_excel(writer, sheet_name="TCM data paste area", index=False, header=True, startrow=2)
                
                # Render a download button for each track
                st.download_button(
                    label=f"💾 Download: {output_filename}",
                    data=output_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
