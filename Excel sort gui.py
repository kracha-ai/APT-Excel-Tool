import pandas as pd
from tkinter import Tk, filedialog, messagebox

def process_excel():
    # Ask user to choose the source Excel file
    source_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not source_file:
        messagebox.showwarning("Warning", "No file selected!")
        return

    try:
        df = pd.read_excel(source_file)
    except Exception as e:
        messagebox.showerror("Error", f"Cannot read file:\n{e}")
        return

    # ---- change this list to your real column names ----
    required_columns = ["src_trans_date", "tot_amount"]

    # Check if all required columns exist
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        messagebox.showerror("Error", f"Missing columns: {missing}")
        return

    # Filter and sort
    df_filtered = df[required_columns].sort_values(by="src_trans_date")

    # Ask where to save the result
    save_path = filedialog.asksaveasfilename(
        title="Save Output As",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not save_path:
        return

    try:
        df_filtered.to_excel(save_path, index=False)
        messagebox.showinfo("Success", "✅ Data processed and saved!")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot save file:\n{e}")

# --- main window ---
root = Tk()
root.withdraw()   # hides empty Tk window
process_excel()
