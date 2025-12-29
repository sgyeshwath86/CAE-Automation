#!/usr/bin/env python
# coding: utf-8

# In[2]:


# using a GUI which selects a file, executes the functions and closes once the functions are executed
# import the modules needed
import pandas as pd
import re
import numpy as np
from io import StringIO
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import time  # Only used here to simulate work; remove in real functions

# ===== definition of 2 functions, one to create a dataframe of the participation factors and other to export the modes as an excel =====
# create a function to extract text needed from the dat file, for e.g, Participation Factors
def extract_between(file_path, start_text, end_text, *, include_markers=False):
    with open(file_path, 'r', encoding='utf-8') as f:
        data = f.read()

    start_idx = data.find(start_text)
    if start_idx == -1:
        raise ValueError(f"Start text not found: {start_text!r}")

    # Start slicing after the start marker, unless you want to include it
    content_start = start_idx if include_markers else start_idx + len(start_text)

    end_idx = data.find(end_text, content_start)
    if end_idx == -1:
        raise ValueError(f"End text not found after start: {end_text!r}")

    # Include the end marker if requested
    content_end = end_idx + (len(end_text) if include_markers else 0)

    return data[content_start:content_end]
def create_df_partFact(file_path):
    # set the arguments needed for the function to extract Participation Factors
    start_text = 'P A R T I C I P A T I O N   F A C T O R S'
    end_text = 'E F F E C T I V E   M A S S'
    file = file_path
    # extract and store the participation factors obtained from the function as a string
    partFact_txt = extract_between(file_path,start_text,end_text)
    # use pandas builtin module to read the above extracted text as a dataframe by converting the string into a file using the StringIO module
    partFactFwF = pd.read_fwf(StringIO(partFact_txt))
    # create another dataframe containing only the mode number and translation components
    df_partFact_01 = partFactFwF[['X-COMPONENT','Y-COMPONENT','Z-COMPONENT']]
    # setting the index to start from 1
    df_partFact_01.index = range(1,len(df_partFact_01)+1)
    # creating a new column for mode no
    df_partFact_01['Mode No'] = df_partFact_01.index
    # re-arranging the columns so that mode no column appears first
    df_partFact_01 = df_partFact_01[['Mode No','X-COMPONENT','Y-COMPONENT','Z-COMPONENT']]
    return df_partFact_01
def export_modes(df,component,excel_name):
     # create another dataframe by sorting the full dataframe by the given component
    df_partFact_sorted=df.sort_values(by=component,ascending=False)
    # create a list of top 5 and bottom 5 modes with highest participation factors
    Top5 = df_partFact_sorted['Mode No'][:5].to_list()
    Bot5 = df_partFact_sorted['Mode No'][-5:].to_list()
    # combine the top5 and bottom5 into a single list and sort them
    AllModes = Top5+Bot5
    # create a series from the above list so that we can export it as an excel
    seriesModes=pd.Series(data=AllModes)
    # export the above series to an excel
    seriesModes.to_excel(excel_name+'.xlsx',index=False,header=['Mode'])
    return    
# ===== GUI application =====
class FileProcessorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Text File Processor")
        self.geometry("560x220")
        self.resizable(False, False)

        self.selected_path = None
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        # Instructions
        tk.Label(self, text="1) Choose a .dat file\n2) Click 'Run' to execute").pack(
            anchor="w", padx=12, pady=8
        )

        # Browse button + status label
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", padx=12)

        self.btn_browse = tk.Button(btn_frame, text="Browse...", command=self.choose_text_file)
        self.btn_browse.pack(side="left")

        self.lbl_result = tk.Label(btn_frame, text="No file selected", fg="blue", wraplength=500, justify="left")
        self.lbl_result.pack(side="left", padx=10)

        # Run + Exit buttons
        action_frame = tk.Frame(self)
        action_frame.pack(fill="x", **pad)

        self.btn_run = tk.Button(action_frame, text="Run", command=self.run_pipeline, state="disabled")
        self.btn_run.pack(side="left")

        self.btn_exit = tk.Button(action_frame, text="Exit", command=self.safe_exit, state="disabled")
        self.btn_exit.pack(side="left", padx=10)

        # Progress / status area
        self.status_var = tk.StringVar(value="Waiting for file selection…")
        self.lbl_status = tk.Label(self, textvariable=self.status_var, fg="grey")
        self.lbl_status.pack(anchor="w", padx=12)

    def choose_text_file(self):
        path = filedialog.askopenfilename(
            title="Select the dat file",
            filetypes=[("dat files", "*.dat"), ("All files", "*.*")]
        )
        if path:
            self.selected_path = path
            self.lbl_result.config(text=f"Selected: {path}")
            self.status_var.set("Ready to run.")
            self.btn_run.config(state="normal")   # enable Run after selection
        else:
            self.selected_path = None
            self.lbl_result.config(text="No file selected")
            self.status_var.set("Waiting for file selection…")
            self.btn_run.config(state="disabled")

    def run_pipeline(self):
        if not self.selected_path:
            messagebox.showwarning("No file", "Please choose a file first.")
            return

        # Disable buttons while running
        self.btn_run.config(state="disabled")
        self.btn_browse.config(state="disabled")
        self.status_var.set("Running… Please wait.")

        # Run in a background thread to keep UI responsive
        thread = threading.Thread(target=self._execute_steps, daemon=True)
        thread.start()

    def _execute_steps(self):
        try:
            # Execute your two functions using the selected path
            df_partFactOut = create_df_partFact(self.selected_path)
            components=['X-COMPONENT','Y-COMPONENT','Z-COMPONENT']
            for component in components:
                export_modes(df_partFactOut,component,component[0]+'CompModes')
            # Back to UI thread to notify success
            self.after(0, self._on_success)
        except Exception as e:
            # Back to UI thread to show error
            self.after(0, self._on_failure, e)

    def _on_success(self):
        self.status_var.set("Completed.")
        messagebox.showinfo("Done", "Excel files created successfully.")
        # Enable Exit; keep Browse disabled so user doesn’t re-run accidentally
        self.btn_exit.config(state="normal")
        self.btn_browse.config(state="normal")  # Or keep disabled if you want to lock state

    def _on_failure(self, err: Exception):
        self.status_var.set("Failed. See details.")
        messagebox.showerror("Error", f"An error occurred:\n{err}")
        # Re-enable controls so user can retry or change file
        self.btn_run.config(state="normal" if self.selected_path else "disabled")
        self.btn_browse.config(state="normal")

    def safe_exit(self):
        # Optional: confirm exit only after success
        self.quit()  # or self.destroy()

if __name__ == "__main__":
    app = FileProcessorApp()
    app.mainloop()

    # After GUI closes, you can still access the final chosen path if needed:


# In[ ]:





# In[ ]:





# In[ ]:




