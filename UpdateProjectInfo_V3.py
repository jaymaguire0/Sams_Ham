import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook
import threading
import datetime

# ---------------- Logging ----------------
def log_message(message):
    with open("UpdateProjectInfo_V3.log", "a") as log_file:
        log_file.write(f"{datetime.datetime.now()}: {message}\n")

# ---------------- Excel Update Logic ----------------
def update_excel(file_path, project_name, project_number, issued_for, update_ts, update_es):
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        filename = os.path.basename(file_path)

        if update_ts and filename.startswith("MST-TS"):
            # Technical Submittal file
            ws["B5"] = project_name if project_name else ""
            ws["B6"] = project_number if project_number else ""
            log_message(f"Updated TS file: {filename}")

        if update_es and filename.startswith("ES-"):
            # Equipment Schedule file
            ws["H2"] = project_number if project_number else ""
            ws["H3"] = project_name if project_name else ""
            ws["J4"] = issued_for if issued_for else ""
            log_message(f"Updated ES file: {filename}")

        wb.save(file_path)
    except Exception as e:
        log_message(f"ERROR updating {file_path}: {e}")

# ---------------- Main Processing ----------------
def process_files(project_name, project_number, issued_for, root_folder, update_ts, update_es, progress, status_label):
    if not update_ts and not update_es:
        messagebox.showerror("Error", "Please select at least one file type to update.")
        return

    all_files = []
    for dirpath, _, files in os.walk(root_folder):
        for f in files:
            if update_ts and f.startswith("MST-TS") and f.endswith(".xlsx"):
                all_files.append(os.path.join(dirpath, f))
            if update_es and f.startswith("ES-") and f.endswith(".xlsx"):
                all_files.append(os.path.join(dirpath, f))

    if not all_files:
        messagebox.showinfo("No Files", "No matching Excel files found.")
        return

    progress["maximum"] = len(all_files)

    for i, file in enumerate(all_files, start=1):
        update_excel(file, project_name, project_number, issued_for, update_ts, update_es)
        progress["value"] = i
        status_label.config(text=f"Updated {i}/{len(all_files)} files")
        progress.update_idletasks()

    messagebox.showinfo("Done", "Update completed! Check UpdateProjectInfo_V3.log for details.")
    os.startfile("UpdateProjectInfo_V3.log")  # auto-open log in Notepad

# ---------------- GUI ----------------
def run_gui():
    root = tk.Tk()
    root.title("Update Project Info V3")

    # Checkboxes for file types
    update_ts_var = tk.BooleanVar()
    update_es_var = tk.BooleanVar()

    chk_ts = tk.Checkbutton(root, text="Technical Submittal (TS files)", variable=update_ts_var)
    chk_ts.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)

    chk_es = tk.Checkbutton(root, text="Equipment Schedule (ES files)", variable=update_es_var)
    chk_es.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)

    # Project inputs
    tk.Label(root, text="Project Name:").grid(row=2, column=0, sticky="e", padx=10, pady=5)
    entry_name = tk.Entry(root, width=40)
    entry_name.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(root, text="Project Number:").grid(row=3, column=0, sticky="e", padx=10, pady=5)
    entry_number = tk.Entry(root, width=40)
    entry_number.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(root, text="Issued For:").grid(row=4, column=0, sticky="e", padx=10, pady=5)
    entry_issued_for = tk.Entry(root, width=40, state="disabled")
    entry_issued_for.grid(row=4, column=1, padx=10, pady=5)

    # Enable/disable "Issued For"
    def toggle_issued_for():
        if update_es_var.get():
            entry_issued_for.config(state="normal")
        else:
            entry_issued_for.delete(0, tk.END)
            entry_issued_for.config(state="disabled")

    chk_es.config(command=toggle_issued_for)

    # Folder selection
    tk.Label(root, text="Root Folder:").grid(row=5, column=0, sticky="e", padx=10, pady=5)
    folder_path = tk.StringVar()
    entry_folder = tk.Entry(root, textvariable=folder_path, width=40)
    entry_folder.grid(row=5, column=1, padx=10, pady=5)

    def browse_folder():
        folder_selected = filedialog.askdirectory()
        folder_path.set(folder_selected)

    tk.Button(root, text="Browse", command=browse_folder).grid(row=5, column=2, padx=10, pady=5)

    # Progress bar
    progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress.grid(row=6, column=0, columnspan=3, pady=10)

    status_label = tk.Label(root, text="Waiting...")
    status_label.grid(row=7, column=0, columnspan=3, pady=5)

    # Run button
    def start_update():
        threading.Thread(
            target=process_files,
            args=(
                entry_name.get(),
                entry_number.get(),
                entry_issued_for.get(),
                folder_path.get(),
                update_ts_var.get(),
                update_es_var.get(),
                progress,
                status_label,
            ),
        ).start()

    tk.Button(root, text="Run Update", command=start_update).grid(row=8, column=0, columnspan=3, pady=10)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
