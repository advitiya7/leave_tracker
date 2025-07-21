import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
import io

# Leave mapping
mapped_code_to_label = {
    "S": "Leave: ILL",
    "S-H": "Leave: ILL Half day",
    "V": "Leave: V01",
    "V-H": "Leave: V01 Half day",
    "C": "Leave: Caregiver Leave",
    "C-H": "Leave: Caregiver Half day",
    "B": "Leave: Bereavement Leave",
    "L-O": "Leave: Others",
    "L": "Leave: LOA",
    "RH": "Holiday: RH",
    "H-WGF": "Holiday: WGF",
    "H-C": "Holiday: Office Closed",
    "T-I": "Others: Induction",
    "T-B": "Others- Training (BCN Academy/ Global)",
    "T": "Others: Training",
    "H": "Work From Home"
}

st.title("Insync Automater")

leave_file = st.file_uploader("Upload Leave Tracker (.xlsx)", type=["xlsx"])

if leave_file:
    try:
        df_raw = pd.read_excel(leave_file, header=None)
        employee_col_index = 1

        # 🔍 Detect month name from top rows of Column B
        col_b_top = df_raw.iloc[:10, employee_col_index].dropna().astype(str).str.strip().tolist()
        month_names = [m for m in calendar.month_name if m]
        month = None

        for val in col_b_top:
            if val.capitalize() in month_names:
                month = datetime.strptime(val, "%B").month
                break

        # Default if not found
        if not month:
            month = datetime.today().month
        year = datetime.today().year

        # Find employee name row
        employee_series = df_raw.iloc[:, employee_col_index].dropna().astype(str).str.strip()
        employee_start_row = employee_series[employee_series.str.lower() == 'employee name'].index[0] + 1
        employee_names = df_raw.iloc[employee_start_row:, employee_col_index].dropna().astype(str).str.strip().tolist()

        # Identify date columns
        header_row = df_raw.iloc[employee_start_row - 1]
        date_cols, date_mapping = [], {}
        for i, val in header_row.items():
            try:
                if pd.notna(val):
                    day = int(float(val))
                    if 1 <= day <= 31:
                        date = datetime(year, month, day)
                        if calendar.weekday(year, month, day) < 5:
                            date_cols.append(i)
                            date_mapping[i] = date
            except:
                continue

        # Map leave entries
        records = []
        for row_idx in range(employee_start_row, len(df_raw)):
            name = str(df_raw.iat[row_idx, employee_col_index]).strip()
            for col in date_cols:
                val = str(df_raw.iat[row_idx, col]).strip().upper()
                if not val or val == 'NAN':
                    continue
                mapped = mapped_code_to_label.get(val)
                if mapped:
                    records.append({'Employee name': name, 'Date': date_mapping[col], 'Mapped Leave Type': mapped})

        if not records:
            st.warning("No mapped leave records found.")
        else:
            working_days = [datetime(year, month, d) for d in range(1, 32)
                            if calendar.monthrange(year, month)[1] >= d and calendar.weekday(year, month, d) < 5]

            num_rows = len(working_days) + 6
            num_cols = 7 + len(employee_names)
            template = pd.DataFrame("", index=range(num_rows), columns=range(num_cols))

            # Headers
            template.iat[3, 2] = "DATE DETAILS"
            template.iloc[4, 2:7] = ["Year", "Month", "Day", "Type", "Date"]
            for i, name in enumerate(employee_names):
                template.iat[4, 7 + i] = name
            for i, date in enumerate(working_days):
                row = 5 + i
                template.iat[row, 2] = date.year
                template.iat[row, 3] = date.strftime("%B")
                template.iat[row, 4] = date.day
                template.iat[row, 5] = date.strftime("%A")
                template.iat[row, 6] = date

            for r in records:
                emp, date, mapped = r["Employee name"], r["Date"], r["Mapped Leave Type"]
                if emp in employee_names and date in working_days:
                    row = 5 + working_days.index(date)
                    col = 7 + employee_names.index(emp)
                    template.iat[row, col] = mapped

            totals_row = 5 + len(working_days)
            template.iat[totals_row, 6] = "Total Entries"
            for i in range(len(employee_names)):
                col_letter = chr(72 + i)
                formula = f'=COUNTA({col_letter}6:{col_letter}{totals_row})'
                template.iat[totals_row, 7 + i] = formula

            # Output to Excel with formatting
            st.success(f"Month Detected: {calendar.month_name[month]} {year} ✅")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter', datetime_format='yyyy-mm-dd') as writer:
                template.to_excel(writer, index=False, header=False)
                wb = writer.book
                ws = writer.sheets['Sheet1']

                border = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                bold = wb.add_format({'bold': True, 'border': 1, 'align': 'center'})
                date_fmt = wb.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center', 'border': 1})

                ws.set_column("C:C", 10, border)
                ws.set_column("D:D", 12, border)
                ws.set_column("E:E", 6, border)
                ws.set_column("F:F", 15, border)
                ws.set_column("G:G", 18, date_fmt)
                ws.set_column("H:ZZ", 22, border)
                for row in range(template.shape[0]):
                    ws.set_row(row, 22, border)
                ws.set_row(4, 26, bold)

                ws.freeze_panes(5, 7)

                for i in range(len(employee_names)):
                    col_num = 7 + i
                    col_letter = chr(65 + col_num) if col_num < 26 else f"{chr(64 + col_num // 26)}{chr(65 + col_num % 26)}"
                    ws.conditional_format(f"{col_letter}6:{col_letter}{totals_row}", {
                        'type': 'text',
                        'criteria': 'containing',
                        'value': 'ILL',
                        'format': wb.add_format({'bg_color': '#FFC7CE', 'border': 1})
                    })
                    ws.conditional_format(f"{col_letter}6:{col_letter}{totals_row}", {
                        'type': 'text',
                        'criteria': 'containing',
                        'value': 'V01',
                        'format': wb.add_format({'bg_color': '#C6EFCE', 'border': 1})
                    })
                    ws.conditional_format(f"{col_letter}6:{col_letter}{totals_row}", {
                        'type': 'text',
                        'criteria': 'containing',
                        'value': 'Work From Home',
                        'format': wb.add_format({'bg_color': '#BDD7EE', 'border': 1})
                    })

            st.download_button(
                label="Download Final Excel File",
                data=buffer.getvalue(),
                file_name=f"insync_output_{calendar.month_name[month].lower()}_{year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Processing failed: {e}")
