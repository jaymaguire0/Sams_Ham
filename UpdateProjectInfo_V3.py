import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import logging

# Setup logging
logging.basicConfig(
    filename="update_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def update_ts_file(file_path, project_name, project_number):
    """Update Technical Submittal (TS) file."""
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        ws["B5"] = project_name if project_name else ""
        ws["B6"] = project_number if project_number else ""
        wb.save(file_path)
        logging.info(f"Updated TS file: {file_path}")
    except Exception as e:
        logging.error(f"Error updating TS file {file_path}: {e}")

def update_es_file(file_path, project_name, project_number, issued_for):
    """Update Equipment Schedule (ES) file."""
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        ws["H2"] = project_number if project_number else ""
        ws["H3"] = project_name if project_name else ""
        ws["J4"] = issued_for if issued_for else ""
        wb.save(file_path)
        logging.info(f"Updated ES file: {file_path}")
    except Exception as e:
        logging.error(f"Error updating ES file {file_path}: {e}")

def run_update(folder_path, project_name, project_number, issued_for, update_ts, update_es, progress_var, root):
    """Run updates on selected files with progress bar."""
    try:
        files = os.listdir(folder_path)
        tasks = []
        if update_ts:
            tasks.extend([f for f in files if f.startswith("MST-TS") and f.endswith(".xlsx")])
        if update_es:
            tasks.extend([f for f in files if f.startswith("ES-") and f.endswith(".xlsx")])

        total = len(tasks)
        if total == 0:
            messagebox.showwarning("No Files", "No matching files found to update.")
            return

        for i, file in enumerate(tasks, start=1):
            file_path = os.path.join(folder_path, file)
            if file.startswith("MST-TS") and update_ts:
                update_ts_file(file_path, project_name, project_number)
            elif file.startswith("ES-") and update_es:
                update_es_file(file_path, project_name, project_number, issued_for)
            progress_var.set(int((i / total) * 100))
            root.update_idletasks()

        messagebox.showinfo("Success", "Files updated successfully!")
    except Exception as e:
        logging.error(f"Error running update: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

def run_gui():
    """Main GUI."""
    root = tk.Tk()
    root.title("Update Project Info V3")

    # Variables
    project_name_var = tk.StringVar()
    project_number_var = tk.StringVar()
    issued_for_var = tk.StringVar()
    folder_path_var = tk.StringVar()
    update_ts_var = tk.BooleanVar(value=True)
    update_es_var = tk.BooleanVar(value=False)

    # Folder selection
    tk.Label(root, text="Select Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=folder_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: folder_path_var.set(filedialog.askdirectory())).grid(row=0, column=2, padx=5, pady=5)

    # Checkboxes
    tk.Checkbutton(root, text="Technical Submittal", variable=update_ts_var).grid(row=1, column=0, sticky="w", padx=5)
    tk.Checkbutton(root, text="Equipment Schedule", variable=update_es_var).grid(row=1, column=1, sticky="w", padx=5)

    # Notes & Warnings
    notes_text = (
        "Important Notes:\n"
        "- Do NOT run this tool if the Excel files are already open (on your PC or another's).\n"
        "- Ensure the folder contains the correct MST-TS and/or ES files before running.\n"
        "- Empty fields will clear the corresponding Excel cells.\n"
        "- If 'Equipment Schedule' is selected, 'Issued For' must be provided (or left blank intentionally).\n"
        "- Always keep a backup of your files before bulk updates.\n\n"
        "P.S. SAM IS A HAM"
    )

    notes_box = tk.Text(root, height=8, width=70, wrap="word", bg="#f5f5f5")
    notes_box.insert("1.0", notes_text)
    notes_box.tag_add("footer", "end-2l", "end-1c")
    notes_box.tag_config("footer", font=("Arial", 10, "bold"))
    notes_box.config(state="disabled")
    notes_box.grid(row=2, column=0, columnspan=3, padx=10, pady=5)

    # Inputs
    tk.Label(root, text="Project Name:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=project_name_var, width=40).grid(row=3, column=1, padx=5, pady=5)

    tk.Label(root, text="Project Number:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=project_number_var, width=40).grid(row=4, column=1, padx=5, pady=5)

    tk.Label(root, text="Issued For:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
    issued_for_entry = tk.Entry(root, textvariable=issued_for_var, width=40, state="disabled")
    issued_for_entry.grid(row=5, column=1, padx=5, pady=5)

    def toggle_issued_for():
        if update_es_var.get():
            issued_for_entry.config(state="normal")
        else:
            issued_for_entry.config(state="disabled")
            issued_for_var.set("")

    update_es_var.trace_add("write", lambda *args: toggle_issued_for())

    # Progress bar
    progress_var = tk.IntVar()
    progress = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="we")

    # Update button
    tk.Button(
        root,
        text="Update Files",
        command=lambda: run_update(
            folder_path_var.get(),
            project_name_var.get(),
            project_number_var.get(),
            issued_for_var.get(),
            update_ts_var.get(),
            update_es_var.get(),
            progress_var,
            root
        )
    ).grid(row=7, column=1, pady=10)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
