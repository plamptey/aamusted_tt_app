#!/usr/bin/env python
# coding: utf-8

# In[11]:


# Fol
import pandas as pd
import sys
from IPython.display import display, Image

# ------------------------
# Config / Input paths
# ------------------------
timetable_path = "C:/Users/HP/Documents/AAMUSTED/2025/timetable/input_data/exam_timetable_split_rows.xlsx"
logo_path = "C:/Users/HP/Documents/AAMUSTED/2025/timetable/input_data/AAMUSTED-LOGO.jpg"
developer_info = "👨‍💻 Developed by: Patrick Nii Lante Lamptey | 📞 +233-208 426 593"

# def expand_split_venues(df, venue_col="VENUE"):
#     """
#     Expand rows where the venue cell contains multiple venues separated by comma or slash.
#     Each venue will become its own row, with other data duplicated.
#     """
#     rows = []
#     for _, row in df.iterrows():
#         if pd.notna(row[venue_col]) and ("," in str(row[venue_col]) or "/" in str(row[venue_col])):
#             # split on comma or slash
#             venues = [v.strip() for v in str(row[venue_col]).replace("/", ",").split(",")]
#             for v in venues:
#                 new_row = row.copy()
#                 new_row[venue_col] = v
#                 rows.append(new_row)
#         else:
#             rows.append(row)
#     return pd.DataFrame(rows).reset_index(drop=True)

# # Load timetable
# timetable = pd.read_excel(timetable_path)
timetable = pd.read_excel(timetable_path)
# timetable = expand_split_venues(timetable, venue_col="VENUE")


# A small palette to color groups (will cycle)
GROUP_COLORS = ["#FFF2CC", "#D9EAD3", "#F4CCCC", "#CFE2F3", "#EAD1DC", "#FDEBD0"]

# ------------------------
# Helper: compute group coloring
# ------------------------
def compute_group_row_colors(df, key_cols=None):
    """
    Given a sorted DataFrame df, compute a mapping index -> color
    for groups that represent split-venues (group size > 1).
    key_cols: list of columns to define a group (e.g. ['DAY & DATE','TIME','COURSE CODE'])
    """
    if key_cols is None:
        key_cols = ['DAY & DATE', 'TIME', 'COURSE CODE']

    # Only use keys that exist in df
    key_cols = [c for c in key_cols if c in df.columns]

    # fallback if TIME not present
    if not key_cols:
        return {}

    group_colors = {}
    color_idx = 0

    grouped = df.groupby(key_cols, sort=False)
    for group_key, group in grouped:
        if len(group) > 1:
            color = GROUP_COLORS[color_idx % len(GROUP_COLORS)]
            color_idx += 1
            for idx in group.index:
                group_colors[idx] = color

    return group_colors  # mapping index -> color

# ------------------------
# JUPYTER: display with Styler (merged days + colored groups)
# ------------------------
def display_in_jupyter(df):
    """Display the timetable in Jupyter with merged day effect and colored split-venue groups."""
    try:
        display(Image(filename=logo_path, width=180))
    except Exception:
        pass

    # Sort by day & time (if available)
    sort_cols = [c for c in ["DAY & DATE", "TIME"] if c in df.columns]
    df_sorted = df.sort_values(by=sort_cols).reset_index(drop=True)

    # Compute colors based on original values (before masking)
    row_colors = compute_group_row_colors(df_sorted)

    # Create a styles DataFrame with same index/columns and default empty strings
    styles = pd.DataFrame("", index=df_sorted.index, columns=df_sorted.columns)

    # Fill styles for rows that are in a multi-row group
    for idx, color in row_colors.items():
        styles.loc[idx, :] = f"background-color: {color}"

    # For merged-day visual effect: blank repeated DAY & DATE
    df_display = df_sorted.copy()
    if "DAY & DATE" in df_display.columns:
        df_display["DAY & DATE"] = df_display["DAY & DATE"].mask(df_display["DAY & DATE"].duplicated())

    # Header style
    header_style = [{
        'selector': 'thead th',
        'props': [('background-color', '#4CAF50'),
                  ('color', 'white'),
                  ('font-weight', 'bold'),
                  ('text-align', 'center')]
    }]

    styler = df_display.style.set_table_styles(header_style)
    # apply the precomputed styles
    styler = styler.apply(lambda _: styles, axis=None)

    display(styler)
    print("\n" + developer_info)

# ------------------------
# STREAMLIT: render HTML table with inline row styles so colors persist
# ------------------------
def render_table_html_for_streamlit(df):
    """
    Returns HTML string for an HTML table that:
     - blanks repeated DAY & DATE values (to simulate merged cell)
     - applies row background colors for grouped (split-venue) rows
    """
    # Prefer columns order for display if present
    possible_cols = ["DAY & DATE", "TIME", "CLASS", "COURSE CODE", "COURSE TITLE",
                 "TOTAL STDS", "NO. OF STDS", "VENUE", "INVIG.", "FACULTY", "DEPARTMENT"]
    display_cols = [c for c in possible_cols if c in df.columns]

    # Sort by day & time for consistent grouping
    sort_cols = [c for c in ["DAY & DATE", "TIME"] if c in df.columns]
    df_sorted = df.sort_values(by=sort_cols).reset_index(drop=True)

    # compute row color mapping on the original (unmasked) sorted dataframe
    row_colors = compute_group_row_colors(df_sorted)

    # blank repeated DAY & DATE for display
    df_sorted_display = df_sorted.copy()
    if "DAY & DATE" in df_sorted_display.columns:
        df_sorted_display["DAY & DATE"] = df_sorted_display["DAY & DATE"].mask(df_sorted_display["DAY & DATE"].duplicated())

    # Build HTML
    th_style = "background:#4CAF50;color:white;padding:8px;text-align:center;"
    td_style = "padding:8px;border-bottom:1px solid #ddd;vertical-align:top;"

    html = "<div style='overflow-x:auto;'><table style='border-collapse:collapse;width:100%;'>"
    # header
    html += "<thead><tr>"
    for col in display_cols:
        html += f"<th style='{th_style}'>{col}</th>"
    html += "</tr></thead><tbody>"

    for idx, row in df_sorted_display.iterrows():
        row_color = row_colors.get(idx, "")
        tr_style = f"background:{row_color};" if row_color else ""
        html += f"<tr style='{tr_style}'>"
        for col in display_cols:
            val = row.get(col, "")
            # Show empty string instead of nan
            cell = "" if (pd.isna(val)) else str(val)
            html += f"<td style='{td_style}'>{cell}</td>"
        html += "</tr>"

    html += "</tbody></table></div>"

    return html

# ------------------------
# Combined UI runner (auto-detect environment)
# ------------------------
def run_jupyter_mode():
    print("⚠️ Running in Jupyter preview mode (not Streamlit).")
    display_in_jupyter(timetable)

def run_streamlit_mode():
    import streamlit as st
    from io import BytesIO
    import base64

    st.set_page_config(page_title="University Exam Timetable", layout="wide")

    # Header: fixed header with logo + title
    try:
        st.image(logo_path, width=140)
    except Exception:
        st.warning("Logo not found (check logo_path).")

    st.title("📘 AAMUSTED-M 2nd Semester 2025 Examination Timetable")

    # Sidebar filters
    st.sidebar.header("🔎 Filter Timetable")
    faculties = ["All", "None"] + sorted(timetable["FACULTY"].dropna().unique().tolist()) if "FACULTY" in timetable.columns else ["All", "None"]
    departments = ["All", "None"] + sorted(timetable["DEPARTMENT"].dropna().unique().tolist()) if "DEPARTMENT" in timetable.columns else ["All", "None"]
    levels = ["All", "None"] + sorted(timetable["CLASS"].dropna().unique().tolist()) if "CLASS" in timetable.columns else ["All", "None"]
    days = ["All", "None"] + sorted(timetable["DAY & DATE"].dropna().unique().tolist()) if "DAY & DATE" in timetable.columns else ["All", "None"]
    invigilators = ["All", "None"] + sorted(timetable["INVIG."].dropna().unique().tolist()) if "INVIG." in timetable.columns else ["All", "None"]

    faculty_filter = st.sidebar.selectbox("Select Faculty", faculties)
    dept_filter = st.sidebar.selectbox("Select Department", departments)
    level_filter = st.sidebar.selectbox("Select Level/Class", levels)
    day_filter = st.sidebar.selectbox("Select Day", days)
    inv_filter = st.sidebar.selectbox("Select Invigilator", invigilators)

    # Apply filters
    filtered = timetable.copy()

    if faculty_filter not in ["All", "None"] and "FACULTY" in filtered.columns:
        filtered = filtered[filtered["FACULTY"] == faculty_filter]
    if dept_filter not in ["All", "None"] and "DEPARTMENT" in filtered.columns:
        filtered = filtered[filtered["DEPARTMENT"] == dept_filter]
    if level_filter not in ["All", "None"] and "CLASS" in filtered.columns:
        filtered = filtered[filtered["CLASS"] == level_filter]
    if day_filter not in ["All", "None"] and "DAY & DATE" in filtered.columns:
        filtered = filtered[filtered["DAY & DATE"] == day_filter]
    if inv_filter not in ["All", "None"] and "INVIG." in filtered.columns:
        filtered = filtered[filtered["INVIG."] == inv_filter]

    # Drop columns where "None" was selected
    drop_cols = []
    if faculty_filter == "None": drop_cols.append("FACULTY")
    if dept_filter == "None": drop_cols.append("DEPARTMENT")
    if level_filter == "None": drop_cols.append("CLASS")
    if day_filter == "None": drop_cols.append("DAY & DATE")
    if inv_filter == "None": drop_cols.append("INVIG.")
    filtered = filtered.drop(columns=[c for c in drop_cols if c in filtered.columns], errors="ignore")

    # Render HTML table so colors persist
    html_table = render_table_html_for_streamlit(filtered)

    st.markdown("### 📅 Filtered Timetable", unsafe_allow_html=True)
    st.markdown(html_table, unsafe_allow_html=True)

    # Export to Excel buffer
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Timetable")
    output.seek(0)

    st.download_button(
        label="⬇️ Download Filtered Timetable as Excel",
        data=output,
        file_name="filtered_timetable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Footer
    st.markdown("---")
    st.markdown(f"<div style='text-align:center;color:gray'>{developer_info}</div>", unsafe_allow_html=True)


# Auto-detect environment and run
if "streamlit" in sys.modules:
    run_streamlit_mode()
else:
    run_jupyter_mode()


# In[ ]:




