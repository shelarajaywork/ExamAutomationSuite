"""
seating_plan.py — Seating Plan Allocator
A tool in the Exam Tools suite.

Reads an exam seating-plan workbook in the standard SVKM's CED format
(sheets: "Instructions", "Seating Plan", "ScheduleAssessment", "RoomCodes"),
reproduces the balance / capacity / conflict checks that the original
spreadsheet does with conditional formatting, lets a user edit room
allocations paper-by-paper with live feedback, and exports a freshly
styled workbook (Seating Plan + regenerated ScheduleAssessment) at the end.

Usage (standalone):
    streamlit run seating_plan.py

Usage (as a tool inside app.py, matching the other Exam Tools modules):
    from seating_plan import main
    main()
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

SEATING_SHEET = "Seating Plan"
ROOMS_SHEET = "RoomCodes"
SCHEDULE_SHEET = "ScheduleAssessment"
INSTRUCTIONS_SHEET = "Instructions"

META_COLS = [
    "Exam Date",
    "Exam start time",
    "Exam End Time",
    "Program Code",
    "Programme",
    "Sem",
    "Exam session",
    "Acad. year",
    "Module Code Abbreviation",
    "Module Name",
    "Appraisal Type of E1 Assessment",
    "Location",
    "Exam Plan",
    "Student Count",
    "PWD",
]

# Colors matching the original workbook's conditional formatting
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
GREEN_FONT = Font(color="006100", bold=True)
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006", bold=True)
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C")
AMBER_FONT = Font(color="9C5700", bold=True)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")


def _clean_cols(cols):
    return [c.strip() if isinstance(c, str) else c for c in cols]


# --------------------------------------------------------------------------
# Loading
# --------------------------------------------------------------------------

@st.cache_data(show_spinner="Reading workbook...")
def load_workbook_data(file_bytes: bytes):
    """Parse the uploaded workbook into (rooms_df, plan_df, room_cols)."""
    rooms_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=ROOMS_SHEET, header=0)
    rooms_df.columns = _clean_cols(rooms_df.columns)
    rooms_df["_roomNo"] = rooms_df["_roomNo"].astype(str).str.strip()
    rooms_df["ROOM NO"] = rooms_df["ROOM NO"].astype(str).str.strip()
    if "EXAM CAPACITY" in rooms_df.columns:
        rooms_df["EXAM CAPACITY"] = pd.to_numeric(
            rooms_df["EXAM CAPACITY"], errors="coerce"
        ).fillna(0).astype(int)
    else:
        rooms_df["EXAM CAPACITY"] = 0

    plan_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SEATING_SHEET, header=2)
    plan_df.columns = _clean_cols(plan_df.columns)
    plan_df = plan_df.dropna(axis=1, how="all")
    plan_df = plan_df.dropna(subset=["Exam Date"]).reset_index(drop=True)

    known_ids = set(rooms_df["_roomNo"])
    room_cols = [
        c for c in plan_df.columns
        if c not in META_COLS and c != "Balance" and str(c).strip() in known_ids
    ]

    # Normalize types
    for c in room_cols:
        plan_df[c] = pd.to_numeric(plan_df[c], errors="coerce").fillna(0).astype(int)
    plan_df["Student Count"] = pd.to_numeric(plan_df["Student Count"], errors="coerce").fillna(0).astype(int)
    if "PWD" in plan_df.columns:
        plan_df["PWD"] = pd.to_numeric(plan_df["PWD"], errors="coerce").fillna(0).astype(int)
    plan_df["Exam Date"] = pd.to_datetime(plan_df["Exam Date"])

    return rooms_df, plan_df, room_cols


def capacity_map(rooms_df: pd.DataFrame) -> dict:
    return dict(zip(rooms_df["_roomNo"], rooms_df["EXAM CAPACITY"]))


def room_name_map(rooms_df: pd.DataFrame) -> dict:
    name_col = "ROOM NAME" if "ROOM NAME" in rooms_df.columns else rooms_df.columns[3]
    return dict(zip(rooms_df["_roomNo"], rooms_df[name_col]))


# --------------------------------------------------------------------------
# Core calculations (mirror the spreadsheet's logic)
# --------------------------------------------------------------------------

def with_balance(plan_df: pd.DataFrame, room_cols: list[str]) -> pd.DataFrame:
    df = plan_df.copy()
    df["Allocated"] = df[room_cols].sum(axis=1)
    df["Balance"] = df["Student Count"] - df["Allocated"]
    return df


def row_status(row) -> str:
    """Status for a paper's overall Balance cell."""
    if row["Balance"] == 0 and row["Allocated"] > 0:
        return "green"
    if row["Balance"] < 0:
        return "red"
    return "amber"


def cell_status(allocated: int, capacity: int) -> str:
    """Status for a single paper-in-room cell."""
    if allocated == 0:
        return "none"
    if capacity and allocated > capacity:
        return "red"
    if capacity and allocated == capacity:
        return "green"
    return "amber"


def find_conflicts(df: pd.DataFrame, room_cols: list[str], cap_map: dict) -> pd.DataFrame:
    """
    Flag any room that is booked beyond its exam capacity when summed across
    every paper scheduled on the same date + start time (double-booking
    check, equivalent to the SUMIFS rule in the original sheet).
    """
    rows = []
    for (date, start), grp in df.groupby(["Exam Date", "Exam start time"]):
        for room in room_cols:
            total = grp[room].sum()
            cap = cap_map.get(room, 0)
            if total > 0 and cap and total > cap:
                rows.append({
                    "Exam Date": date,
                    "Start Time": start,
                    "Room": room,
                    "Allocated (all papers)": int(total),
                    "Room Capacity": int(cap),
                    "Over by": int(total - cap),
                    "Papers": ", ".join(
                        grp.loc[grp[room] > 0, "Module Name"].astype(str).tolist()
                    ),
                })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# Styled display helpers (Streamlit)
# --------------------------------------------------------------------------

def style_full_plan(df: pd.DataFrame, room_cols: list[str], cap_map: dict):
    colors = {
        "green": "background-color:#C6EFCE;color:#006100;font-weight:600",
        "red": "background-color:#FFC7CE;color:#9C0006;font-weight:600",
        "amber": "background-color:#FFEB9C;color:#9C5700",
        "none": "",
    }

    def style_cols(frame: pd.DataFrame) -> pd.DataFrame:
        out = pd.DataFrame("", index=frame.index, columns=frame.columns)
        for i, row in frame.iterrows():
            out.loc[i, "Balance"] = colors[row_status(row)]
            for room in room_cols:
                st_ = cell_status(row[room], cap_map.get(room, 0))
                out.loc[i, room] = colors[st_]
        return out

    display_cols = [
        "Exam Date", "Exam start time", "Exam End Time", "Programme", "Sem",
        "Module Name", "Student Count", "PWD", "Allocated", "Balance",
    ] + room_cols
    show = df[display_cols].copy()
    show["Exam Date"] = show["Exam Date"].dt.strftime("%d-%b-%Y")
    return show.style.apply(lambda _: style_cols(df)[display_cols], axis=None)


# --------------------------------------------------------------------------
# Export: rebuild a styled workbook
# --------------------------------------------------------------------------

def build_schedule_assessment(df: pd.DataFrame, room_cols: list[str]) -> pd.DataFrame:
    """Rebuild the long-format ScheduleAssessment sheet from current allocations."""
    records = []
    for _, row in df.iterrows():
        for room in room_cols:
            qty = row[room]
            if qty and qty > 0:
                records.append({
                    "Program Code": row.get("Program Code"),
                    "Academic year": row.get("Acad. year"),
                    "Academic session": row.get("Exam session"),
                    "Module Code Abbreviation": row.get("Module Code Abbreviation"),
                    "Appraisal Type of E1 Assessment": row.get("Exam Plan"),
                    "Location": row.get("Location"),
                    "Room Codes": room,
                    "Exam Date": row["Exam Date"].strftime("%d.%m.%Y"),
                    "Exam start time": row.get("Exam start time"),
                    "Exam End Time": row.get("Exam End Time"),
                    "Capacity": int(qty),
                    "Exam Plan": row.get("Exam Plan"),
                    "Programme": row.get("Programme"),
                    "Sem": row.get("Sem"),
                    "Module Name": row.get("Module Name"),
                })
    return pd.DataFrame(records)


def export_workbook(
    original_bytes: bytes,
    plan_df: pd.DataFrame,
    room_cols: list[str],
    cap_map: dict,
) -> bytes:
    """
    Carry over Instructions + RoomCodes unchanged from the uploaded file,
    rebuild Seating Plan and ScheduleAssessment with current data + styling.
    """
    wb_in = load_workbook(io.BytesIO(original_bytes))
    wb = Workbook()
    wb.remove(wb.active)

    # Copy Instructions + RoomCodes sheets over as-is
    for sheet_name in (INSTRUCTIONS_SHEET, ROOMS_SHEET):
        if sheet_name not in wb_in.sheetnames:
            continue
        src = wb_in[sheet_name]
        dst = wb.create_sheet(sheet_name)
        for row in src.iter_rows():
            for cell in row:
                dst.cell(row=cell.row, column=cell.column, value=cell.value)
        for col, dim in src.column_dimensions.items():
            dst.column_dimensions[col].width = dim.width

    # --- Seating Plan sheet ---
    ws = wb.create_sheet(SEATING_SHEET, 1)
    header = ["Exam Date", "Exam start time", "Exam End Time", "Program Code",
              "Programme", "Sem", "Exam session", "Acad. year",
              "Module Code Abbreviation", "Module Name",
              "Appraisal Type of E1 Assessment", "Location", "Exam Plan",
              "Student Count", "PWD", "Balance"] + room_cols

    ws.cell(row=1, column=10, value=f"Total Student Count: {int(plan_df['Student Count'].sum())}")
    df_bal = with_balance(plan_df, room_cols)
    ws.cell(row=1, column=10).value = f"Total Student Count: {int(plan_df['Student Count'].sum())}"
    ws.cell(row=2, column=10, value=f"Pending Count: {int(df_bal['Balance'].clip(lower=0).sum())}")
    for j, room in enumerate(room_cols, start=17):
        ws.cell(row=2, column=j, value=int(cap_map.get(room, 0)))

    for j, h in enumerate(header, start=1):
        c = ws.cell(row=3, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    status_fill = {"green": GREEN_FILL, "red": RED_FILL, "amber": AMBER_FILL, "none": None}
    status_font = {"green": GREEN_FONT, "red": RED_FONT, "amber": AMBER_FONT, "none": None}

    r = 4
    for _, row in df_bal.iterrows():
        ws.cell(row=r, column=1, value=row["Exam Date"].to_pydatetime()).number_format = "dd-mmm-yyyy"
        ws.cell(row=r, column=2, value=str(row.get("Exam start time", "")))
        ws.cell(row=r, column=3, value=str(row.get("Exam End Time", "")))
        ws.cell(row=r, column=4, value=row.get("Program Code"))
        ws.cell(row=r, column=5, value=row.get("Programme"))
        ws.cell(row=r, column=6, value=row.get("Sem"))
        ws.cell(row=r, column=7, value=row.get("Exam session"))
        ws.cell(row=r, column=8, value=row.get("Acad. year"))
        ws.cell(row=r, column=9, value=row.get("Module Code Abbreviation"))
        ws.cell(row=r, column=10, value=row.get("Module Name"))
        ws.cell(row=r, column=11, value=row.get("Appraisal Type of E1 Assessment"))
        ws.cell(row=r, column=12, value=row.get("Location"))
        ws.cell(row=r, column=13, value=row.get("Exam Plan"))
        ws.cell(row=r, column=14, value=int(row["Student Count"]))
        ws.cell(row=r, column=15, value=int(row.get("PWD", 0)))

        bal_cell = ws.cell(row=r, column=16, value=int(row["Balance"]))
        st_ = row_status(row)
        if status_fill[st_]:
            bal_cell.fill = status_fill[st_]
            bal_cell.font = status_font[st_]

        for j, room in enumerate(room_cols, start=17):
            val = int(row[room])
            cell = ws.cell(row=r, column=j, value=val if val else None)
            cs = cell_status(val, cap_map.get(room, 0))
            if status_fill[cs]:
                cell.fill = status_fill[cs]
                cell.font = status_font[cs]
        r += 1

    ws.freeze_panes = "Q4"
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14

    # --- ScheduleAssessment sheet ---
    sched_df = build_schedule_assessment(df_bal, room_cols)
    ws2 = wb.create_sheet(SCHEDULE_SHEET, 2)
    for j, h in enumerate(sched_df.columns, start=1):
        c = ws2.cell(row=1, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
    for i, row in enumerate(sched_df.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            ws2.cell(row=i, column=j, value=val)
    ws2.freeze_panes = "A2"
    for j in range(1, len(sched_df.columns) + 1):
        ws2.column_dimensions[get_column_letter(j)].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

def main():
    st.subheader("🪑 Seating Plan Allocator")
    st.caption(
        "Upload an exam seating-plan workbook (Instructions / Seating Plan / "
        "ScheduleAssessment / RoomCodes) to review, edit and export room allocations."
    )

    uploaded = st.file_uploader("Upload seating plan workbook (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.info("Waiting for a file — expects the standard SVKM's CED seating-plan template.")
        return

    file_bytes = uploaded.getvalue()
    try:
        rooms_df, plan_df, room_cols = load_workbook_data(file_bytes)
    except Exception as e:
        st.error(f"Couldn't parse this workbook: {e}")
        return

    cap_map = capacity_map(rooms_df)
    name_map = room_name_map(rooms_df)

    # Keep an editable copy in session state, keyed to this upload
    state_key = f"plan_df::{uploaded.name}::{len(file_bytes)}"
    if state_key not in st.session_state:
        st.session_state[state_key] = plan_df.copy()
    df = st.session_state[state_key]

    missing_cap = [r for r in room_cols if cap_map.get(r, 0) == 0 and df[r].sum() > 0]
    if missing_cap:
        st.warning(
            "These rooms are used in the plan but have no Exam Capacity set in "
            f"RoomCodes: {', '.join(missing_cap)}. Capacity checks for them will "
            "be skipped until a capacity is entered."
        )

    df_bal = with_balance(df, room_cols)
    tab_overview, tab_edit, tab_full, tab_conflicts, tab_export = st.tabs(
        ["Overview", "Edit a paper", "Full plan", "Room conflicts", "Export"]
    )

    # ---------------- Overview ----------------
    with tab_overview:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Papers", len(df_bal))
        c2.metric("Total students", int(df_bal["Student Count"].sum()))
        c3.metric("Still pending", int(df_bal["Balance"].clip(lower=0).sum()))
        c4.metric(
            "Fully allocated",
            int((df_bal.apply(row_status, axis=1) == "green").sum()),
        )
        st.dataframe(
            df_bal[df_bal["Balance"] != 0][
                ["Exam Date", "Exam start time", "Programme", "Module Name",
                 "Student Count", "Allocated", "Balance"]
            ].sort_values("Balance", ascending=False),
            use_container_width=True,
            hide_index=True,
        )

    # ---------------- Edit a paper ----------------
    with tab_edit:
        df_bal = df_bal.reset_index(drop=True)
        labels = [
            f"{r['Exam Date'].strftime('%d-%b-%Y')} {r['Exam start time']} — "
            f"{r['Programme']} — {r['Module Name']}"
            for _, r in df_bal.iterrows()
        ]
        idx = st.selectbox("Choose a paper", range(len(labels)), format_func=lambda i: labels[i])
        row = df_bal.loc[idx]

        st.write(
            f"**Student Count:** {int(row['Student Count'])}  |  "
            f"**PWD:** {int(row.get('PWD', 0))}  |  "
            f"**Currently allocated:** {int(row['Allocated'])}  |  "
            f"**Balance:** {int(row['Balance'])}"
        )

        current_rooms = [r for r in room_cols if row[r] > 0]
        chosen_rooms = st.multiselect(
            "Rooms used for this paper", options=room_cols, default=current_rooms,
            format_func=lambda r: f"{r} ({name_map.get(r, '')}, cap {cap_map.get(r, 0)})",
        )

        new_vals = {}
        cols = st.columns(3)
        for i, room in enumerate(chosen_rooms):
            with cols[i % 3]:
                new_vals[room] = st.number_input(
                    f"Room {room} (cap {cap_map.get(room, 0)})",
                    min_value=0, value=int(row[room]), step=1, key=f"room_{idx}_{room}",
                )

        if st.button("Save allocation", type="primary"):
            for room in room_cols:
                st.session_state[state_key].at[idx, room] = new_vals.get(room, 0)
            st.success("Saved. Check the Overview / Full plan tabs for the updated balance.")
            st.rerun()

        live_alloc = sum(new_vals.values())
        live_balance = int(row["Student Count"]) - live_alloc
        if live_balance == 0:
            st.success(f"Balance: 0 — fully allocated.")
        elif live_balance < 0:
            st.error(f"Over-allocated by {-live_balance} students.")
        else:
            st.warning(f"{live_balance} students still need a room.")
        over_cap = [r for r in chosen_rooms if cap_map.get(r, 0) and new_vals[r] > cap_map[r]]
        if over_cap:
            st.error(f"These rooms exceed their capacity: {', '.join(over_cap)}")

    # ---------------- Full plan ----------------
    with tab_full:
        st.caption("Green = room fully used, red = over capacity, amber = under capacity / pending.")
        st.dataframe(
            style_full_plan(with_balance(st.session_state[state_key], room_cols), room_cols, cap_map),
            use_container_width=True,
            hide_index=True,
        )

    # ---------------- Conflicts ----------------
    with tab_conflicts:
        conflicts = find_conflicts(st.session_state[state_key], room_cols, cap_map)
        if conflicts.empty:
            st.success("No room is double-booked beyond capacity in any overlapping time slot.")
        else:
            st.error(f"{len(conflicts)} room/time-slot conflict(s) found.")
            st.dataframe(conflicts, use_container_width=True, hide_index=True)

    # ---------------- Export ----------------
    with tab_export:
        st.write("Export the current allocations as a freshly styled workbook.")
        if st.button("Build export file"):
            out_bytes = export_workbook(file_bytes, st.session_state[state_key], room_cols, cap_map)
            st.download_button(
                "Download Seating Plan.xlsx",
                data=out_bytes,
                file_name=f"Seating_Plan_Updated_{datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()