import pandas as pd
from tkinter import Tk, filedialog, messagebox
import os

def merge_and_sort_excels():
    root = Tk()
    root.withdraw()

    # 🔔 Custom message on opening
    messagebox.showinfo(
        "About This Tool",
        "This file has been created by R. Kousik Reddy, OA, Nellore Division.\n\n"
        "If you want to merge and sort Excel files that are downloaded from Accounting Details in APT software "
        "as per Date and Amount, this file will be useful.\n\n"
        "📂 Select the Excel files that you want to merge and sort."
    )

    # Step 1: Ask user to select multiple files
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files to Merge and Sort",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_paths:
        messagebox.showwarning("Warning", "No files selected!")
        return

    merged_df = pd.DataFrame()

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
            messagebox.showerror("Error", f"Cannot read {file_path}:\n{e}")
            continue

        # --- Try to find Date and Amount columns automatically ---
        date_col = None
        amount_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if "date" in col_lower:
                date_col = col
            if "amount" in col_lower or "amt" in col_lower:
                amount_col = col

        if not date_col or not amount_col:
            messagebox.showwarning("Warning", f"Skipping {file_path}: missing Date/Amount columns.")
            continue

        # --- Convert date column to datetime for sorting ---
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
        df = df[[date_col, amount_col]].sort_values(by=date_col)

        # --- Rename columns with file name ---
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        df.columns = [f"{file_name}_Date", f"{file_name}_Amount"]

        df = df.reset_index(drop=True)
        merged_df = pd.concat([merged_df, df], axis=1)

    if merged_df.empty:
        messagebox.showinfo("Info", "No valid data processed.")
        return

    # --- Sort each date column individually ---
    for col in merged_df.columns:
        if "date" in col.lower():
            merged_df[col] = pd.to_datetime(merged_df[col], errors="coerce", dayfirst=True)
            merged_df[col] = merged_df[col].dt.strftime("%d-%m-%Y")

    # --- Save the output file ---
    save_path = filedialog.asksaveasfilename(
        title="Save Merged and Sorted File As",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not save_path:
        return

    try:
        merged_df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"✅ File merged and sorted successfully:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot save file:\n{e}")

if __name__ == "__main__":
    merge_and_sort_excels()
