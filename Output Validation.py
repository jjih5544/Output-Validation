import smartsheet
import smartsheet.models
import smartsheet.exceptions
import tkinter as tk
from tkinter import messagebox, Toplevel, Label
import threading
import queue
import time

# ========== CONFIGURATION ==========
API_TOKEN = "XXXXX" # Paste API your token
SHEET_ID = 388148  # Keep as integer
CHECK_COLUMN_NAME = "Lot Number" # Paste column name used in the sheet

# ========== GLOBAL STATE ==========
cached_values = set()
column_id = None
submission_queue = queue.Queue()
is_processing = False

# ========== INITIALIZE SMARTSHEET ==========
try:
    client = smartsheet.Smartsheet(API_TOKEN)
except Exception as e:
    print(f"Failed to initialize Smartsheet client: {e}")
    exit(1)

def get_column_id(sheet, col_name):
    for col in sheet.columns:
        if col.title == col_name:
            return col.id
    return None

def preload_values_and_column():
    global cached_values, column_id
    try:
        sheet = client.Sheets.get_sheet(SHEET_ID)
        column_id = get_column_id(sheet, CHECK_COLUMN_NAME)
        if not column_id:
            messagebox.showerror("Error", f"Column '{CHECK_COLUMN_NAME}' not found.")
            return

        for row in sheet.rows:
            for cell in row.cells:
                if cell.column_id == column_id and cell.value is not None:
                    cached_values.add(str(cell.value))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to preload data: {str(e)}")

# ========== SPINNER UI ==========
def show_spinner():
    spinner = Toplevel(root)
    spinner.title("Working...")
    spinner.geometry("200x80")
    spinner.resizable(False, False)
    spinner.attributes("-topmost", True)
    Label(spinner, text="Please wait...", font=("Arial", 12)).pack(pady=20)
    spinner.update()
    return spinner

# ========== BACKGROUND WORKER ==========
def submission_worker():
    global is_processing
    while True:
        value = submission_queue.get()
        is_processing = True

        spinner = None  # Declare in the outer function

        try:
            if value in cached_values:
                root.after(0, lambda: messagebox.showinfo("Info", f"Value '{value}' already exists."))
            else:
                # Create spinner in main thread using an Event to wait until it's ready
                spinner_event = threading.Event()

                def create_spinner():
                    nonlocal spinner  # Declare that we're modifying the outer variable
                    spinner = show_spinner()
                    spinner_event.set()

                root.after(0, create_spinner)
                spinner_event.wait()  # Wait for the spinner to be created

                new_row = smartsheet.models.Row()
                new_row.to_top = True

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = column_id
                new_cell.value = value
                new_row.cells.append(new_cell)

                result = client.Sheets.add_rows(SHEET_ID, [new_row])
                if result.result_code == 0:
                    cached_values.add(value)
                    root.after(0, lambda: messagebox.showinfo("Success", f"Value '{value}' added to the sheet."))
                else:
                    root.after(0, lambda: messagebox.showerror("Error", f"Failed to add row: {result.message}"))
        except smartsheet.exceptions.ApiError as e:
            root.after(0, lambda: messagebox.showerror("API Error", f"Smartsheet API error: {e.error.result.message}"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
        finally:
            if spinner:
                root.after(0, spinner.destroy)
            is_processing = False
            submission_queue.task_done()

# ========== GUI SETUP ==========
root = tk.Tk()
root.title("Output Validation") #Form Title
root.geometry("300x180")

tk.Label(root, text="Scan Lot Number:", font=("Arial", 14)).pack(pady=10, ipady=3)
entry = tk.Entry(root, width=30, font=("Arial", 12), justify="center")
entry.pack(pady=10, ipady=3)

def handle_submit():
    val = entry.get().strip()
    if val:
        submission_queue.put(val)
        entry.delete(0, tk.END)
    else:
        messagebox.showwarning("Warning", "Please enter a value.")

entry.bind("<Return>", lambda event: handle_submit())

submit_btn = tk.Button(root, text="Submit", command=handle_submit,
                       font=("Arial", 12), bg="#4CAF50", fg="white")
submit_btn.pack(pady=10)

# ========== STARTUP ==========
def on_start():
    preload_thread = threading.Thread(target=preload_values_and_column, daemon=True)
    preload_thread.start()

    worker_thread = threading.Thread(target=submission_worker, daemon=True)
    worker_thread.start()

if __name__ == "__main__":
    root.after(100, on_start)
    root.mainloop()