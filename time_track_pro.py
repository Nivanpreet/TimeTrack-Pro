# time_track_pro_fixed.py
# Requirements: pandas, openpyxl, reportlab (optional for PDF saving)
# pip install pandas openpyxl reportlab

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from datetime import datetime, timedelta

# ---------- Config & Storage ----------
employee_file = "employee_details.csv"
employee_details = {}  # loaded from file

def load_employee_details():
    if os.path.exists(employee_file):
        try:
            df = pd.read_csv(employee_file)
            if "Employee Name" in df.columns:
                return df.set_index("Employee Name").to_dict("index")
        except Exception:
            return {}
    return {}

def save_employee_details():
    global employee_details
    try:
        df = pd.DataFrame.from_dict(employee_details, orient="index").reset_index().rename(columns={"index":"Employee Name"})
        df.to_csv(employee_file, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save employee details:\n{e}")

employee_details = load_employee_details()

# ---------- Helper functions & core logic ----------

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return
    try:
        # Read the first sheet as a DataFrame
        sheet_data = pd.read_excel(file_path, sheet_name=0, header=None, engine="openpyxl")
        # Use the chunk scheme you used: every employee info is 22 rows + 1 blank row (chunk_size=23)
        chunk_size = 23
        start_row = 2  # data starts from row index 2 (0-based)
        employee_chunks = []
        for i in range(start_row, len(sheet_data), chunk_size):
            end_row = min(i + 22, len(sheet_data))
            chunk = sheet_data.iloc[i:end_row].copy().reset_index(drop=True)
            # employee name typically in first cell of the chunk row 0 col0
            try:
                employee_name = str(chunk.iloc[0, 0]).strip()
            except Exception:
                employee_name = f"Employee_{i}"
            employee_chunks.append({"name": employee_name, "chunk": chunk, "start_row": i, "end_row": end_row})
        if not employee_chunks:
            messagebox.showerror("Error", "No employee chunks found in the selected file. Check file format.")
            return
        display_navigation_window(employee_chunks, file_path, sheet_data)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the file:\n{e}")

# ---------- Navigation window (editing) ----------
def display_navigation_window(employee_chunks, file_path, original_sheet_df):
    navigation_window = tk.Toplevel(root)
    navigation_window.title("Employee Details Navigation")
    navigation_window.geometry("700x500")

    current_index = {"idx": 0}   # use dict to allow nested functions to modify

    # Treeview for dates / in / out
    table_columns = ("Date", "In Time", "Out Time")
    table_frame = tk.Frame(navigation_window)
    table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    table = ttk.Treeview(table_frame, columns=table_columns, show="headings", height=12)
    for col in table_columns:
        table.heading(col, text=col)
    table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=table.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    table.configure(yscrollcommand=scrollbar.set)

    # store edits: dict of emp_index -> {date_str: {"In Time": val, "Out Time": val}}
    edited_data = {}

    def show_employee(idx):
        if idx < 0 or idx >= len(employee_chunks):
            return
        current_index["idx"] = idx
        emp = employee_chunks[idx]
        chunk = emp["chunk"]
        employee_name = emp["name"]
        navigation_window.title(f"Employee: {employee_name} ({idx+1}/{len(employee_chunks)})")
        # find rows labeled 'Date', 'In Time', 'Out Time' in first column of chunk
        table.delete(*table.get_children())
        try:
            date_row = chunk[chunk.iloc[:, 0].astype(str).str.strip().str.lower() == "date"]
            in_row = chunk[chunk.iloc[:, 0].astype(str).str.strip().str.lower() == "in time"]
            out_row = chunk[chunk.iloc[:, 0].astype(str).str.strip().str.lower() == "out time"]
            if date_row.empty or in_row.empty or out_row.empty:
                # fallback: try to read first row as dates and second/third as in/out if labeled differently
                # Here we try to display first 20 columns
                row0 = chunk.iloc[0, :20].astype(str).fillna("")
                row1 = chunk.iloc[1, :20].astype(str).fillna("")
                row2 = chunk.iloc[2, :20].astype(str).fillna("")
                dates = row0[row0 != ""].values
                ins = row1[row1 != ""].values
                outs = row2[row2 != ""].values
            else:
                dates = date_row.iloc[0, 1:].dropna().values
                ins = in_row.iloc[0, 1:].dropna().values
                outs = out_row.iloc[0, 1:].dropna().values
            # insert into table
            for d, it, ot in zip(dates, ins, outs):
                table.insert("", "end", values=(str(d), str(it), str(ot)))
        except Exception:
            messagebox.showwarning("Warning", "Could not parse this employee chunk. It may have an unexpected format.")

    def edit_cell(event):
        selected_item = table.focus()
        if not selected_item:
            return
        vals = table.item(selected_item, "values")
        if not vals:
            return
        date, in_time, out_time = vals
        col = table.identify_column(event.x)
        try:
            col_index = int(col.replace("#", "")) - 1
        except Exception:
            return
        if col_index not in [1, 2]:
            # only allow editing In Time or Out Time
            return
        col_name = "In Time" if col_index == 1 else "Out Time"
        new_val = simpledialog.askstring("Edit Value", f"Enter new {col_name} for {date}:", initialvalue=vals[col_index])
        if new_val is None:
            return
        new_vals = list(vals)
        new_vals[col_index] = new_val
        table.item(selected_item, values=new_vals)
        idx = current_index["idx"]
        if idx not in edited_data:
            edited_data[idx] = {}
        edited_data[idx][date] = {"In Time": new_vals[1], "Out Time": new_vals[2]}

    def apply_edits_to_original():
        # For each edited chunk, write back the changed values into original_sheet_df
        nonlocal original_sheet_df
        for emp_idx, changes in edited_data.items():
            emp = employee_chunks[emp_idx]
            chunk = emp["chunk"]
            start = emp["start_row"]
            # Find column indices that correspond to dates (first row labeled 'Date' expected)
            try:
                date_row_idx = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="date"].index
                in_row_idx = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="in time"].index
                out_row_idx = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="out time"].index
                if not date_row_idx.empty and not in_row_idx.empty and not out_row_idx.empty:
                    date_row_i = date_row_idx[0]
                    in_row_i = in_row_idx[0]
                    out_row_i = out_row_idx[0]
                    # columns in chunk start at 0 -> corresponds to original_sheet_df columns at same offsets
                    for date_val, times in changes.items():
                        # locate which column in chunk has this date string
                        cols = [c for c in chunk.columns if str(chunk.iloc[date_row_i, c]).strip() == str(date_val).strip()]
                        if cols:
                            col = cols[0]
                            orig_col = col  # same column index relative to sheet
                            # write back into original_sheet_df at appropriate row indices
                            orig_row_in = start + in_row_i
                            orig_row_out = start + out_row_i
                            original_sheet_df.iat[orig_row_in, orig_col] = times["In Time"]
                            original_sheet_df.iat[orig_row_out, orig_col] = times["Out Time"]
                else:
                    # fallback: try to match by first three rows as date,in,out pattern
                    for date_val, times in changes.items():
                        # find column by matching date in the first 10 columns
                        for col in chunk.columns[:20]:
                            if str(chunk.iloc[0, col]).strip() == str(date_val).strip():
                                orig_row_in = start + 1
                                orig_row_out = start + 2
                                original_sheet_df.iat[orig_row_in, col] = times["In Time"]
                                original_sheet_df.iat[orig_row_out, col] = times["Out Time"]
            except Exception as e:
                print("Warning while applying edits:", e)

    def save_edits_and_close():
        if edited_data:
            apply_edits_to_original()
            # write back the modified original_sheet_df to the same file
            try:
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    # original_sheet_df may have NaN types; write as headerless to preserve layout
                    original_sheet_df.to_excel(writer, index=False, header=False)
                messagebox.showinfo("Saved", "Edits saved to the source Excel file.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save edits back to file:\n{e}")
        # After saving edits, process the file and display results in main window
        navigation_window.destroy()
        results = process_excel(file_path)
        if results:
            display_results(results)
        else:
            messagebox.showinfo("Info", "Processing finished but no results were generated.")

    def next_employee():
        if current_index["idx"] < len(employee_chunks) - 1:
            show_employee(current_index["idx"] + 1)

    def prev_employee():
        if current_index["idx"] > 0:
            show_employee(current_index["idx"] - 1)

    table.bind("<Double-1>", edit_cell)

    button_frame = tk.Frame(navigation_window)
    button_frame.pack(fill=tk.X, pady=5)
    tk.Button(button_frame, text="Previous", command=prev_employee).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="Next", command=next_employee).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="Save", command=apply_edits_to_original).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="Continue", command=save_edits_and_close).pack(side=tk.RIGHT, padx=10)

    show_employee(0)

# ---------- Processing Excel ----------
def process_excel(file_path):
    """
    Reads the Excel file and computes total hours and salary.
    Returns a list of dicts: [{"Employee Name":..., "Total Monthly Hours":..., "Calculated Salary":...}, ...]
    """
    try:
        calc_method = simpledialog.askstring("Calculation Method", "Choose calculation method:\nType '1' for In/Out Time or '2' for Total Working Hours", parent=root)
        if calc_method not in ("1", "2"):
            messagebox.showerror("Error", "Invalid calculation method. Please type '1' or '2'.")
            return []
        days_in_month = simpledialog.askinteger("Days in month", "Enter total days in the month:", parent=root, minvalue=1, maxvalue=31)
        if days_in_month is None:
            return []
        num_sundays = simpledialog.askinteger("Sundays", "Enter number of Sundays in month:", parent=root, minvalue=0, maxvalue=10)
        if num_sundays is None:
            return []
        num_holidays = simpledialog.askinteger("Company holidays", "Enter number of company holidays:", parent=root, minvalue=0, maxvalue=10)
        if num_holidays is None:
            return []
        hours_per_day = timedelta(hours=9, minutes=30)
        actual_working_days = days_in_month - (num_sundays + num_holidays)

        # read raw sheet
        raw = pd.read_excel(file_path, header=None, engine="openpyxl")
        chunk_size = 23
        start_row = 2
        results = []

        for i in range(start_row, len(raw), chunk_size):
            end_row = min(i + 22, len(raw))
            chunk = raw.iloc[i:end_row].copy().reset_index(drop=True)
            try:
                emp_name = str(chunk.iloc[0, 0]).strip()
            except Exception:
                emp_name = f"Employee_{i}"

            total_working_hours = 0.0

            if calc_method == '1':
                # Parse In Time and Out Time rows
                in_rows = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="in time"]
                out_rows = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="out time"]
                if not in_rows.empty and not out_rows.empty:
                    in_row = in_rows.iloc[0, 1:].dropna()
                    out_row = out_rows.iloc[0, 1:].dropna()
                    total_seconds = 0.0
                    for in_val, out_val in zip(in_row, out_row):
                        try:
                            # accept various formats: H:M, H:M:S, maybe datetime objects
                            if pd.isna(in_val) or pd.isna(out_val) or str(in_val).strip() in ("", "00:00", "0") or str(out_val).strip() in ("", "00:00", "0"):
                                continue
                            in_dt = pd.to_datetime(str(in_val), errors='coerce')
                            out_dt = pd.to_datetime(str(out_val), errors='coerce')
                            if pd.isna(in_dt) or pd.isna(out_dt):
                                continue
                            delta = (out_dt - in_dt).total_seconds()
                            if delta > 0:
                                total_seconds += delta
                        except Exception:
                            continue
                    total_working_hours = total_seconds / 3600.0
                else:
                    # fallback: try "Total Working Hours" row
                    trow = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="total working hours"]
                    if not trow.empty:
                        try:
                            total_data = trow.iloc[0,1:].dropna().astype(str)
                            # try parse strings like '8:30' into hours
                            seconds = 0.0
                            for v in total_data:
                                try:
                                    td = pd.to_timedelta(v)
                                    seconds += td.total_seconds()
                                except Exception:
                                    # try H:M format
                                    try:
                                        parts = str(v).split(":")
                                        h = float(parts[0]); m = float(parts[1]) if len(parts)>1 else 0
                                        seconds += (h*3600 + m*60)
                                    except Exception:
                                        continue
                            total_working_hours = seconds / 3600.0
                        except Exception:
                            total_working_hours = 0.0
            else:
                # calc_method == '2'
                trow = chunk[chunk.iloc[:,0].astype(str).str.strip().str.lower()=="total working hours"]
                if not trow.empty:
                    total_data = trow.iloc[0,1:].dropna().astype(str)
                    seconds = 0.0
                    for v in total_data:
                        try:
                            td = pd.to_timedelta(v)
                            seconds += td.total_seconds()
                        except Exception:
                            try:
                                parts = str(v).split(":"); h = float(parts[0]); m = float(parts[1]) if len(parts)>1 else 0
                                seconds += (h*3600 + m*60)
                            except Exception:
                                continue
                    total_working_hours = seconds / 3600.0
                else:
                    total_working_hours = 0.0

            # add weekend & holiday hours
            total_time_weekoff = (num_sundays * hours_per_day) + (num_holidays * hours_per_day)
            total_working_hours += total_time_weekoff.total_seconds() / 3600.0

            # compute salary if known
            if emp_name in employee_details:
                try:
                    hourly_rate = float(employee_details[emp_name].get("Hourly Salary", 0))
                except Exception:
                    hourly_rate = 0.0
                calc_salary = total_working_hours * hourly_rate
            else:
                hourly_rate = 0.0
                calc_salary = 0.0

            results.append({
                "Employee Name": emp_name,
                "Total Monthly Hours": total_working_hours,
                "Calculated Salary": calc_salary
            })
        return results
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing the file:\n{e}")
        return []

# ---------- Save / Print functions ----------
def save_details(results, option):
    if not results:
        messagebox.showinfo("Info", "No results to save.")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf"), ("Excel Files", "*.xlsx")])
    if not save_path:
        return
    try:
        if save_path.lower().endswith(".pdf"):
            # generate a simple PDF using reportlab if available
            try:
                from reportlab.lib.pagesizes import letter
                from reportlab.pdfgen import canvas
            except Exception:
                messagebox.showerror("Error", "Reportlab not installed. Install with: pip install reportlab")
                return
            c = canvas.Canvas(save_path, pagesize=letter)
            c.setFont("Helvetica", 12)
            w, h = letter
            margin = 50
            y = h - margin
            line_h = 18
            for r in results:
                if y < margin + line_h*4:
                    c.showPage(); c.setFont("Helvetica", 12); y = h - margin
                c.drawString(margin, y, f"Employee: {r['Employee Name']}")
                y -= line_h
                c.drawString(margin+10, y, f"Total Monthly Hours: {r['Total Monthly Hours']:.2f}")
                y -= line_h
                c.drawString(margin+10, y, f"Calculated Salary: {r['Calculated Salary']:.2f}")
                y -= line_h*2
            c.save()
            messagebox.showinfo("Saved", f"Saved PDF to {save_path}")
        else:
            # save to Excel
            df = pd.DataFrame(results)
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Saved", f"Saved Excel to {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save file: {e}")

def print_details(results, option):
    # placeholder — real printing would use OS-specific calls or generate PDF then send to printer
    messagebox.showinfo("Print", "Print functionality sent (placeholder). You can save as PDF then print.")

# ---------- UI: main window and helpers ----------
def display_results(results):
    # clear tree view
    for r in result_tree.get_children():
        result_tree.delete(r)
    for item in results:
        result_tree.insert("", "end", values=(item["Employee Name"], f"{item['Total Monthly Hours']:.2f}", f"{item['Calculated Salary']:.2f}"))
    # enable save/print buttons
    save_btn.config(state=tk.NORMAL)
    print_btn.config(state=tk.NORMAL)
    # store last results globally for saving
    root.last_results = results

def add_or_edit_employee_details(employee_name=None):
    def on_save():
        name = name_var.get().strip()
        hourly = hourly_var.get().strip()
        shift_s = shift_start_var.get().strip()
        shift_e = shift_end_var.get().strip()
        if not (name and hourly and shift_s and shift_e):
            messagebox.showwarning("Warning", "All fields are required.")
            return
        try:
            employee_details[name] = {"Hourly Salary": float(hourly), "Shift Start": shift_s, "Shift End": shift_e}
            save_employee_details()
            messagebox.showinfo("Saved", f"Details for {name} saved.")
            win.destroy()
        except Exception:
            messagebox.showerror("Error", "Hourly salary must be a number.")
    win = tk.Toplevel(root)
    win.title("Add/Edit Employee")
    tk.Label(win, text="Employee Name").pack(pady=4)
    name_var = tk.StringVar(value=employee_name or "")
    name_entry = tk.Entry(win, textvariable=name_var)
    name_entry.pack()
    tk.Label(win, text="Hourly Salary").pack(pady=4)
    hourly_var = tk.StringVar(value=(str(employee_details.get(employee_name, {}).get("Hourly Salary", "")) if employee_name else ""))
    tk.Entry(win, textvariable=hourly_var).pack()
    tk.Label(win, text="Shift Start (HH:MM)").pack(pady=4)
    shift_start_var = tk.StringVar(value=(employee_details.get(employee_name, {}).get("Shift Start", "") if employee_name else ""))
    tk.Entry(win, textvariable=shift_start_var).pack()
    tk.Label(win, text="Shift End (HH:MM)").pack(pady=4)
    shift_end_var = tk.StringVar(value=(employee_details.get(employee_name, {}).get("Shift End", "") if employee_name else ""))
    tk.Entry(win, textvariable=shift_end_var).pack()
    tk.Button(win, text="Save", command=on_save).pack(pady=10)

def view_employee_details():
    win = tk.Toplevel(root)
    win.title("View Employee Details")
    win.geometry("600x400")
    columns = ("Employee Name", "Hourly Salary", "Shift Start", "Shift End")
    tree = ttk.Treeview(win, columns=columns, show="headings")
    for c in columns:
        tree.heading(c, text=c)
    tree.pack(fill=tk.BOTH, expand=True)
    for name, d in employee_details.items():
        tree.insert("", "end", values=(name, d.get("Hourly Salary",""), d.get("Shift Start",""), d.get("Shift End","")))
    def edit_selected():
        sel = tree.selection()
        if not sel:
            return
        name = tree.item(sel[0], "values")[0]
        add_or_edit_employee_details(name)
        win.destroy()
    tk.Button(win, text="Edit Selected", command=edit_selected).pack(pady=6)

# ---------- Main Tkinter setup ----------
root = tk.Tk()
root.title("TimeTrack Pro")
root.geometry("800x700")

tk.Label(root, text="Welcome to TimeTrack Pro", font=("Arial", 20), fg="blue").pack(pady=10)

tk.Button(root, text="Upload Excel File", font=("Arial", 14), command=upload_file).pack(pady=8)
tk.Button(root, text="Enter Employee Details", font=("Arial", 14), command=add_or_edit_employee_details).pack(pady=8)
tk.Button(root, text="View/Edit Employee Details", font=("Arial", 14), command=view_employee_details).pack(pady=8)

columns = ("Employee Name", "Total Monthly Hours", "Calculated Salary")
result_tree = ttk.Treeview(root, columns=columns, show="headings", height=18)
for col in columns:
    result_tree.heading(col, text=col)
result_tree.column("Employee Name", width=300)
result_tree.column("Total Monthly Hours", width=150)
result_tree.column("Calculated Salary", width=150)
result_tree.pack(fill=tk.BOTH, expand=True, pady=10)

bottom_frame = tk.Frame(root)
bottom_frame.pack(fill=tk.X, pady=6)
save_btn = tk.Button(bottom_frame, text="Save Results", font=("Arial", 12), state=tk.DISABLED, command=lambda: save_details(getattr(root,'last_results',[]), "Salary Record"))
save_btn.pack(side=tk.LEFT, padx=10)
print_btn = tk.Button(bottom_frame, text="Print Results", font=("Arial", 12), state=tk.DISABLED, command=lambda: print_details(getattr(root,'last_results',[]), "Salary Record"))
print_btn.pack(side=tk.LEFT, padx=10)

root.mainloop()
