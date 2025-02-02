import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import threading
from typing import List, Dict
import queue
import sys
import os

class ExcelConsolidatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sheet Consolidator")
        self.root.geometry("600x400")
        
        # Create a queue for thread-safe communication
        self.message_queue = queue.Queue()
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.progress_text = tk.StringVar()
        
        self.setup_ui()
        self.check_queue()
    
    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Style configuration
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Helvetica", 16, "bold"))
        
        # Title
        title = ttk.Label(main_frame, text="Excel Sheet Consolidator", style="Title.TLabel")
        title.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file selection
        ttk.Label(main_frame, text="Input File:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.input_file, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_input).grid(row=1, column=2)
        
        # Output file selection
        ttk.Label(main_frame, text="Output File:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(main_frame, textvariable=self.output_file, width=50).grid(row=2, column=1, padx=5, pady=(10, 0))
        ttk.Button(main_frame, text="Browse", command=self.browse_output).grid(row=2, column=2, pady=(10, 0))
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(progress_frame, textvariable=self.status)
        self.status_label.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Progress text
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_text, wraplength=500)
        self.progress_label.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Consolidate button
        self.consolidate_button = ttk.Button(
            main_frame, 
            text="Consolidate Sheets", 
            command=self.start_consolidation,
            style="Accent.TButton"
        )
        self.consolidate_button.grid(row=4, column=0, columnspan=3, pady=(0, 10))
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(0, weight=1)
    
    def browse_input(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            if not self.output_file.get():
                # Automatically set output file path
                input_path = Path(filename)
                output_path = input_path.parent / f"{input_path.stem}_consolidated{input_path.suffix}"
                self.output_file.set(str(output_path))
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
    
    def consolidate_excel_sheets(self, input_file: str, output_file: str):
        try:
            self.message_queue.put(("status", "Reading Excel file..."))
            excel_data = pd.ExcelFile(input_file)
            consolidated_data = []
            
            row_ranges = [
                (56, 177),   # First blog
                (178, 328),  # Second blog
                (329, 479),  # Third blog
                (480, 520)   # Fourth blog
            ]
            
            total_sheets = len([s for s in excel_data.sheet_names if s.startswith('S')])
            processed_sheets = 0
            
            for sheet_name in excel_data.sheet_names:
                if sheet_name.startswith('S'):
                    self.message_queue.put(("progress", f"Processing sheet {sheet_name}..."))
                    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
                    
                    try:
                        if df.iloc[15, 8] == 0:
                            self.message_queue.put(("progress", f"Sheet {sheet_name} has no data"))
                            continue
                        elif df.iloc[15, 8] >= 1:
                            try:
                                province_value = df.iloc[1, 10]  # K2
                                district = df.iloc[1, 16]       # Q2
                                commune = df.iloc[1, 22]        # W2
                                village = df.iloc[1, 29]        # AC2
                                school = df.iloc[1, 39]         # AM2
                                
                            except IndexError:
                                self.message_queue.put(("progress", f"Warning: Could not find some cell values in sheet {sheet_name}"))
                                province_value = district = commune = village = school = None
                            
                            for start_row, end_row in row_ranges:
                                blog_data = df.iloc[start_row:end_row + 1].copy()
                                blog_data = blog_data.dropna(how='all')
                                
                                if not blog_data.empty:
                                    blog_data['SheetName'] = sheet_name
                                    blog_data["province"] = province_value
                                    blog_data["district"] = district
                                    blog_data["commune"] = commune
                                    blog_data["village"] = village
                                    blog_data["school"] = school
                                    consolidated_data.append(blog_data)
                                    
                        processed_sheets += 1
                        self.message_queue.put(("progress", f"Processed {processed_sheets}/{total_sheets} sheets"))
                        
                    except IndexError:
                        self.message_queue.put(("progress", f"Sheet {sheet_name} is not in the expected format"))
                        continue
            
            self.message_queue.put(("status", "Saving consolidated data..."))
            if consolidated_data:
                final_df = pd.concat(consolidated_data, ignore_index=True)
            else:
                final_df = pd.DataFrame()
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='ConsolidatedData')
            
            self.message_queue.put(("complete", "Data has been successfully consolidated!"))
            
        except Exception as e:
            self.message_queue.put(("error", f"Error: {str(e)}"))
    
    def start_consolidation(self):
        if not self.input_file.get() or not self.output_file.get():
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        
        self.consolidate_button.state(['disabled'])
        self.progress_bar.start(10)
        self.status.set("Processing...")
        
        # Start processing in a separate thread
        thread = threading.Thread(
            target=self.consolidate_excel_sheets,
            args=(self.input_file.get(), self.output_file.get())
        )
        thread.daemon = True
        thread.start()
    
    def check_queue(self):
        try:
            while True:
                msg_type, message = self.message_queue.get_nowait()
                
                if msg_type == "status":
                    self.status.set(message)
                elif msg_type == "progress":
                    self.progress_text.set(message)
                elif msg_type == "complete":
                    self.progress_bar.stop()
                    self.status.set("Complete")
                    self.consolidate_button.state(['!disabled'])
                    messagebox.showinfo("Success", message)
                elif msg_type == "error":
                    self.progress_bar.stop()
                    self.status.set("Error")
                    self.consolidate_button.state(['!disabled'])
                    messagebox.showerror("Error", message)
        
        except queue.Empty:
            pass
        
        self.root.after(100, self.check_queue)

def main():
    root = tk.Tk()
    app = ExcelConsolidatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()