import pandas as pd
from tkinter import Tk, filedialog, messagebox, simpledialog
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import tkinter as tk
from tkinter import ttk

def custom_message(title, message, font_name="Segoe UI", font_size=12):
    win = tk.Toplevel()
    win.title(title)
    win.geometry("500x250")
    win.resizable(False, False)

    # Center window
    win.update_idletasks()
    x = (win.winfo_screenwidth() // 2) - (500 // 2)
    y = (win.winfo_screenheight() // 2) - (250 // 2)
    win.geometry(f"+{x}+{y}")

    lbl = tk.Label(
        win,
        text=message,
        font=(font_name, font_size),
        wraplength=460,
        justify="left",
        padx=20,
        pady=20
    )
    lbl.pack(expand=True, fill="both")

    btn = ttk.Button(win, text="OK", command=win.destroy)
    btn.pack(pady=10)

    win.grab_set()   # Make modal (like messagebox)
    win.wait_window()

# =========================
# 🔹 Excel Formatting (Auto Width + Date Format)
# =========================
def format_excel_columns(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    for column_cells in ws.columns:
        header = str(column_cells[0].value).lower()
        col_letter = column_cells[0].column_letter

        is_date_col = "date" in header
        max_length = 0

        for cell in column_cells[1:]:  # skip header

            # Reset alignment & remove indent
            try:
                cell.alignment = Alignment(horizontal="general", indent=0)
            except:
                pass

            # Strip spaces if text
            if isinstance(cell.value, str):
                cell.value = cell.value.strip()

            # Force Excel date format
            if is_date_col and cell.value:
                try:
                    cell.number_format = "DD-MM-YY"
                except:
                    pass

            # Width calculation
            try:
                if is_date_col:
                    max_length = max(max_length, 8)  # DD-MM-YY
                else:
                    cell_value = str(cell.value)
                    if cell_value:
                        max_length = max(max_length, len(cell_value))
            except:
                pass

        # Include header
        if column_cells[0].value:
            if is_date_col:
                max_length = max(max_length, 8, len(str(column_cells[0].value)))
            else:
                max_length = max(max_length, len(str(column_cells[0].value)))

        ws.column_dimensions[col_letter].width = max_length + 1

    wb.save(file_path)

# =========================
# 🔹 GL Mapping Helpers
# =========================
def select_gl_mapping_file():
    custom_message(
        "👉Select GL Mapping File🙂",
        "👉Please select the GL Mapping Excel file.\n\n"
        "👉It must have exact column names as shown below:\n" 
        "                     GL_Code | Short_Name",
        font_name="Segoe UI",
        font_size=14
    )

    gl_file = filedialog.askopenfilename(
        title="Select GL Mapping Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not gl_file:
        custom_message(
            "Error",
            "😐GL Mapping file is required!",
                font_name="Segoe UI",
                font_size=14
        )
        return None

    return gl_file

def load_gl_mapping_from_file(gl_file_path):
    try:
        df_map = pd.read_excel(gl_file_path, dtype=str)
        return dict(zip(df_map["GL_Code"], df_map["Short_Name"]))
    except Exception as e:
        custom_message(
             "Error",
             "😐Cannot read GL Mapping file.\n\n"
             f"Technical error:\n{e}\n\n"
              "Make sure it has exact column names:\n"
              "GL_Code and Short_Name\n\n"
              "😊 No need to worry! Some GLs may already be stored.\n"
              "   You can continue after fixing the file.\n\n"
              "                👍Any BUGS/issues, please mail to:\n"
              "                          kousik.reddy.r@gmail.com\n"
              "        Credits: R. Kousik Reddy, PA, Nellore Division👍",
             font_name="Segoe UI",
                font_size=14
             )
        return {}

def save_new_gl_to_file(gl_file_path, gl_code, short_name):
    try:
        df_map = pd.read_excel(gl_file_path, dtype=str)

        if gl_code not in df_map["GL_Code"].astype(str).values:
            new_row = pd.DataFrame(
                [[gl_code, short_name]],
                columns=["GL_Code", "Short_Name"]
            )
            df_map = pd.concat([df_map, new_row], ignore_index=True)
            df_map.to_excel(gl_file_path, index=False)
    except Exception as e:
        custom_message(
             "Error",
             f"😐Cannot save new GL to mapping file:\n\n{e}\n\n"
                "Please close the GL mapping file if it is open in Excel.",
                font_name="Segoe UI",
                font_size=14
         )

# =========================
# 🔹 Main Function
# =========================
def merge_and_sort_excels():
    root = Tk()
    root.withdraw()

    # Ask user to select GL mapping file
    gl_mapping_file = select_gl_mapping_file()
    if not gl_mapping_file:
        return

    GL_Head_Name = load_gl_mapping_from_file(gl_mapping_file)

    custom_message(
    "About This Tool",
    "📂 Select the Excel files, you want to merge and sort.\n\n"
    "               👍Any BUGS/issues, please mail to:\n"
    "                         kousik.reddy.r@gmail.com\n"
    "       Credits: R. Kousik Reddy, PA, Nellore Division👍",
    font_name="Segoe UI",
    font_size=14
)

    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files to Merge and Sort",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_paths:
        custom_message(
            "Warning",
            "No files selected!",
            font_name="Segoe UI",
             font_size=14
        )
        return

    merged_df = pd.DataFrame()
    used_category_names = {}

    for file_path in file_paths:
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".xlsx":
                df = pd.read_excel(file_path, engine="openpyxl")
            elif ext == ".xls":
                try:
                    df = pd.read_excel(file_path, engine="xlrd")
                except:
                    df = pd.read_html(file_path)[0]
            else:
                continue
        except Exception as e:
            custom_message(
                "Error",
                  f"😐Cannot read file:\n\n{file_path}\n\nError:\n{e}",
                    font_name="Segoe UI",
                    font_size=14
            )
            continue

        # --- Detect columns ---
        date_col = None
        amount_col = None
        gl_col = None

        for col in df.columns:
            col_lower = str(col).lower()
            if "src_trans_date" in col_lower:
                date_col = col
            if "tot_amount" in col_lower or "amount" in col_lower or "amt" in col_lower:
                amount_col = col
            if gl_col is None and "gl" in col_lower:
                gl_col = col

        if not date_col or not amount_col:
            custom_message(
                 "Warning",
                     f"Skipping file:\n\n{os.path.basename(file_path)}\n\n"
                        "Reason: Missing Date or Amount columns.",
                     font_name="Segoe UI",
                     font_size=14
            )
            continue

        # --- Get GL value & map to category ---
        category_name = "Uncategorized"

        if gl_col:
            try:
                sample_gl = df[gl_col].dropna().iloc[0]
                sample_gl_str = str(int(float(sample_gl))).strip()

                if sample_gl_str in GL_Head_Name:
                    category_name = GL_Head_Name[sample_gl_str]
                else:
                    short_name = simpledialog.askstring(
                        "New GL Detected",
                        f"New GL Code found: {sample_gl_str}\n\n"
                        f"Enter short name (e.g., SB_Dep, RD, SSA):"
                    )

                    if short_name:
                        short_name = short_name.strip()
                        GL_Head_Name[sample_gl_str] = short_name
                        save_new_gl_to_file(gl_mapping_file, sample_gl_str, short_name)
                        category_name = short_name
                    else:
                        category_name = f"GL_{sample_gl_str}"
            except:
                category_name = "Uncategorized"

        # --- Make category name unique if repeated ---
        count = used_category_names.get(category_name, 0) + 1
        used_category_names[category_name] = count
        if count > 1:
            category_name = f"{category_name}_{count}"

        # --- Process data ---
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
        df = df[[date_col, amount_col]].sort_values(by=date_col)

        # --- Rename columns ---
        df.columns = [f"{category_name}_Date", f"{category_name}_Amount"]

        df = df.reset_index(drop=True)
        merged_df = pd.concat([merged_df, df], axis=1)

    if merged_df.empty:
        custom_message(
            "Info",
              "No valid data processed.",
               font_name="Segoe UI",
                 font_size=14
        )
        return

    # --- Ensure all Date columns are datetime ---
    for col in merged_df.columns:
        if "date" in col.lower():
            merged_df[col] = pd.to_datetime(merged_df[col], errors="coerce", dayfirst=True)

    # =========================
    # 🔹 Remove gaps in each Date+Amount pair
    # =========================
    num_cols = merged_df.shape[1]

    for i in range(0, num_cols, 2):
        pair = merged_df.iloc[:, i:i+2]

        # Drop rows where Date is missing
        compact_pair = pair.dropna(subset=[pair.columns[0]]).reset_index(drop=True)

        # Reindex to original length (fills with NaT / NaN correctly)
        new_block = compact_pair.reindex(range(len(merged_df)))

        # Replace original pair safely
        merged_df.iloc[:, i:i+2] = new_block.values

    # --- Save output ---
    save_path = filedialog.asksaveasfilename(
        title="Save Merged and Sorted File As",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not save_path:
        return

    try:
        merged_df.to_excel(save_path, index=False)

        # Auto width + DD-MM-YY date format
        format_excel_columns(save_path)

        custom_message(
    "Success",
    f"✅ File merged and sorted successfully to this path👇:\n\n     👉{save_path}",
    font_name="Segoe UI",
    font_size=14
)
    except Exception as e:
        custom_message(
            "Error",
            f"😐Cannot save file:\n\n{e}",
            font_name="Segoe UI",
            font_size=14
        )

# =========================
# 🔹 Run Program
# =========================
if __name__ == "__main__":
    merge_and_sort_excels()
