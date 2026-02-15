import pandas as pd
from tkinter import Tk, filedialog, messagebox
import os

def merge_and_sort_excels():
    root = Tk()
    root.withdraw()

    # Step 1: Select multiple Excel files
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files to Merge and Sort",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_paths:
        messagebox.showwarning("Warning", "No files selected!")
        return

    merged_df = pd.DataFrame()

    # Step 2: Merge all files side-by-side
    for file_path in file_paths:
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".xlsx":
                df = pd.read_excel(file_path, engine="openpyxl")
            elif ext == ".xls":
                try:
                    df = pd.read_excel(file_path, engine="xlrd")
                except Exception:
                    df = pd.read_html(file_path)[0]
                    messagebox.showinfo("Info", f"Read {os.path.basename(file_path)} as HTML table.")
            else:
                messagebox.showwarning("Warning", f"Skipping {file_path}: unsupported type.")
                continue
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read {file_path}:\n{e}")
            continue

        # Identify Date and Amount columns
        date_col, amount_col = None, None
        for col in df.columns:
            col_lower = str(col).lower()
            if "date" in col_lower:
                date_col = col
            if "amount" in col_lower or "amt" in col_lower:
                amount_col = col

        if not date_col or not amount_col:
            messagebox.showwarning("Warning", f"Skipping {file_path}: missing Date/Amount.")
            continue

        # Convert date and sort
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
        df = df[[date_col, amount_col]].dropna(subset=[date_col])
        df = df.sort_values(by=date_col, ascending=True)
        df[date_col] = df[date_col].dt.strftime("%d-%m-%Y")

        # Rename columns with file name
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        df.columns = [f"{file_name}_Date", f"{file_name}_Amount"]

        # Reset index and merge
        df = df.reset_index(drop=True)
        merged_df = pd.concat([merged_df, df], axis=1)

    if merged_df.empty:
        messagebox.showinfo("Info", "No valid data processed.")
        return

    # Step 3: Sort each date column individually
    sorted_parts = []
    date_cols = [c for c in merged_df.columns if "date" in c.lower()]

    for date_col in date_cols:
        idx = merged_df.columns.get_loc(date_col)
        amount_col = None
        if idx + 1 < len(merged_df.columns):
            next_col = merged_df.columns[idx + 1]
            if "amount" in next_col.lower():
                amount_col = next_col

        cols_to_sort = [date_col]
        if amount_col:
            cols_to_sort.append(amount_col)

        temp = merged_df[cols_to_sort].copy()
        temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce", dayfirst=True)
        temp = temp.dropna(subset=[date_col]).sort_values(by=date_col, ascending=True)
        temp[date_col] = temp[date_col].dt.strftime("%d-%m-%Y")
        temp = temp.reset_index(drop=True)
        sorted_parts.append(temp)

    final_df = pd.concat(sorted_parts, axis=1)

    # Step 4: Save final file
    save_path = filedialog.asksaveasfilename(
        title="Save Final Merged & Sorted Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not save_path:
        return

    try:
        final_df.to_excel(save_path, index=False, engine="openpyxl")
        messagebox.showinfo("Success", f"✅ Final sorted file saved to:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot save file:\n{e}")

if __name__ == "__main__":
    merge_and_sort_excels()
