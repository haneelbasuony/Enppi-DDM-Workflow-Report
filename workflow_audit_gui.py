import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread

import compare_workflow_template_vs_ddm as report
import os

# Dark theme colors
BG = "#1e1e1e"
FG = "#ffffff"
ENTRY_BG = "#2d2d2d"
BTN_BG = "#0078d4"
BTN_FG = "#ffffff"
STATUS_COLOR = "#4fc3f7"


def browse_ddm():

    filename = filedialog.askopenfilename(
        title="Select DDM File",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls")
        ]
    )

    if filename:
        ddm_file_var.set(filename)


def browse_output_folder():

    folder = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if folder:
        custom_output_var.set(folder)

def run_report():
    
    if not save_shared_var.get() and not save_custom_var.get():

        messagebox.showerror(
            "Error",
            "Select at least one output location."
        )

        return

    ddm_file = ddm_file_var.get().strip()

    if not ddm_file:
        messagebox.showerror(
            "Error",
            "Please select a DDM file."
        )
        return

    def worker():

        try:

            run_button.config(state="disabled")
            status_var.set("Running report...")

            report.Config.DDM_FILE_PATH = ddm_file

            generated_files = []

            if save_shared_var.get():

                path = report.generate_report(
                    report.Config.OUTPUT_DIR
                )

                generated_files.append(path)

            if save_custom_var.get():

                custom_folder = custom_output_var.get().strip()

                if not custom_folder:

                    raise Exception(
                        "Please select a custom output folder."
                    )

                path = report.generate_report(
                    custom_folder
                )

                generated_files.append(path)

            status_var.set("Completed Successfully")

            if generated_files:
                os.startfile(generated_files[-1])

            messagebox.showinfo(
                "Success",
                "Report generated successfully."
            )

        except Exception as ex:

            status_var.set("Failed")

            messagebox.showerror(
                "Error",
                str(ex)
            )

        finally:

            run_button.config(state="normal")

    Thread(target=worker, daemon=True).start()


# Main Window
root = tk.Tk()
root.title("Workflow Audit Tool")
root.geometry("900x550")
root.resizable(True, True)

root.configure(bg=BG)

ddm_file_var = tk.StringVar()

save_shared_var = tk.BooleanVar(value=True)

save_custom_var = tk.BooleanVar(value=False)

custom_output_var = tk.StringVar()

# Title
title = tk.Label(
    root,
    text="Workflow Audit Tool",
    font=("Segoe UI", 18, "bold"),
    bg=BG,
    fg=FG
)
title.pack(pady=(20, 5))

subtitle = tk.Label(
    root,
    text="Generate Workflow Compliance Reports",
    font=("Segoe UI", 10),
    bg=BG,
    fg="#c0c0c0"
)
subtitle.pack()

# File Selection Frame
frame = tk.Frame(root, bg=BG)
frame.pack(fill="x", padx=20, pady=30)

lbl = tk.Label(
    frame,
    text="DDM Excel File:",
    bg=BG,
    fg=FG,
    font=("Segoe UI", 10, "bold")
)
lbl.pack(anchor="w", pady=(0, 5))

output_frame = tk.LabelFrame(
    root,
    text="Output Options",
    bg=BG,
    fg=FG
)

output_frame.pack(
    fill="x",
    padx=20,
    pady=10
)

tk.Checkbutton(
    output_frame,
    text="Save to Company Shared Folder",
    variable=save_shared_var,
    bg=BG,
    fg=FG,
    selectcolor=ENTRY_BG
).pack(anchor="w")

tk.Checkbutton(
    output_frame,
    text="Save to Custom Folder",
    variable=save_custom_var,
    bg=BG,
    fg=FG,
    selectcolor=ENTRY_BG
).pack(anchor="w")

custom_frame = tk.Frame(
    output_frame,
    bg=BG
)

custom_frame.pack(fill="x")

tk.Entry(
    custom_frame,
    textvariable=custom_output_var,
    width=70
).pack(side="left")

tk.Button(
    custom_frame,
    text="Browse",
    command=browse_output_folder
).pack(side="left", padx=5)


entry = tk.Entry(
    frame,
    textvariable=ddm_file_var,
    width=85,
    bg=ENTRY_BG,
    fg=FG,
    insertbackground=FG,
    relief="flat"
)

entry.pack(side="left", padx=(0, 10), ipady=6)

browse_btn = tk.Button(
    frame,
    text="Browse",
    command=browse_ddm,
    bg=BTN_BG,
    fg=BTN_FG,
    activebackground="#106ebe",
    activeforeground=BTN_FG,
    relief="flat",
    padx=15
)

browse_btn.pack(side="left")

# Run Button
run_button = tk.Button(
    root,
    text="Generate Report",
    command=run_report,
    bg=BTN_BG,
    fg=BTN_FG,
    activebackground="#106ebe",
    activeforeground=BTN_FG,
    relief="flat",
    font=("Segoe UI", 10, "bold"),
    width=25,
    height=2
)

run_button.pack(pady=10)

# Status
status_var = tk.StringVar(value="Ready")

status_label = tk.Label(
    root,
    textvariable=status_var,
    bg=BG,
    fg=STATUS_COLOR,
    font=("Segoe UI", 10)
)

status_label.pack(pady=15)

root.mainloop()