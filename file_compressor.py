import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import zipfile
import winsound
import threading
import openpyxl
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P

class FileCompressor:
    def __init__(self, root):
        self.root = root
        self.root.title("File Compressor")
        self.root.geometry("600x400")
        
        # Create main frame
        self.main_frame = tk.Frame(self.root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create listbox to display selected files
        self.files_listbox = tk.Listbox(self.main_frame, height=10, width=50)
        self.files_listbox.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create buttons frame
        self.buttons_frame = tk.Frame(self.main_frame)
        self.buttons_frame.pack(fill=tk.X, pady=10)
        
        # Add file button
        self.add_button = tk.Button(self.buttons_frame, text="Add Excel/OpenOffice File", command=self.add_file)
        self.add_button.pack(side=tk.LEFT, padx=5)
        
        # Compress button
        self.compress_button = tk.Button(self.buttons_frame, text="Compress Files", command=self.start_compression)
        self.compress_button.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress_frame = tk.Frame(self.main_frame)
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_label = tk.Label(self.progress_frame, text="Progress:")
        self.progress_label.pack(side=tk.LEFT)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Status label
        self.status_label = tk.Label(self.main_frame, text="", wraplength=550)
        self.status_label.pack(fill=tk.X, pady=10)
        
        # Store selected files
        self.selected_files = []
        self.is_compressing = False

    def add_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel or OpenOffice File",
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"),
                ("OpenOffice files", "*.ods"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.endswith(('.xlsx', '.xls')):
                self.read_excel_file(file_path)
            elif file_path.endswith('.ods'):
                self.read_ods_file(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")

    def read_excel_file(self, file_path):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        # Clear existing files
        self.selected_files.clear()
        self.files_listbox.delete(0, tk.END)
        
        # Read paths from second row onwards
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row and row[0]:  # Check if first column has a value
                path = str(row[0]).strip()
                if os.path.exists(path):
                    self.selected_files.append(path)
                    self.files_listbox.insert(tk.END, path)
                else:
                    self.update_status(f"Warning: Path not found - {path}")
        
        if not self.selected_files:
            messagebox.showwarning("Warning", "No valid file paths found in the Excel file")
        else:
            self.update_status(f"Loaded {len(self.selected_files)} files from Excel")

    def read_ods_file(self, file_path):
        doc = load(file_path)
        table = doc.getElementsByType(Table)[0]
        
        # Clear existing files
        self.selected_files.clear()
        self.files_listbox.delete(0, tk.END)
        
        # Read paths from second row onwards
        for row in table.getElementsByType(TableRow)[1:]:  # Skip header row
            cell = row.getElementsByType(TableCell)[0]
            path = ""
            for p in cell.getElementsByType(P):
                path += p.firstChild.data if p.firstChild else ""
            
            path = path.strip()
            if path and os.path.exists(path):
                self.selected_files.append(path)
                self.files_listbox.insert(tk.END, path)
            elif path:
                self.update_status(f"Warning: Path not found - {path}")
        
        if not self.selected_files:
            messagebox.showwarning("Warning", "No valid file paths found in the OpenOffice file")
        else:
            self.update_status(f"Loaded {len(self.selected_files)} files from OpenOffice")

    def start_compression(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please add a file with paths first")
            return
            
        if self.is_compressing:
            return
            
        self.is_compressing = True
        self.compress_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        # Start compression in a separate thread
        threading.Thread(target=self.compress_files, daemon=True).start()

    def compress_files(self):
        total_files = 0
        processed_files = 0
        
        # First, count total files across all paths
        for path in self.selected_files:
            if os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    total_files += len(files)
            elif os.path.isfile(path):
                total_files += 1
        
        for path in self.selected_files:
            try:
                if os.path.isdir(path):
                    # Process all files in the directory
                    for root, dirs, files in os.walk(path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            # Skip if it's already a zip file
                            if file_path.endswith('.zip'):
                                continue
                                
                            # Get filename without extension
                            file_name = os.path.splitext(file)[0]
                            # Create zip file in the same directory as the source file
                            zip_path = os.path.join(root, f"{file_name}.zip")
                            
                            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                                zipf.write(file_path, file)
                            
                            processed_files += 1
                            # Update progress
                            progress = (processed_files / total_files) * 100
                            self.progress_var.set(progress)
                            self.root.update_idletasks()
                            
                            self.update_status(f"Compressed: {file}")
                elif os.path.isfile(path):
                    # Process single file
                    file_name = os.path.splitext(os.path.basename(path))[0]
                    zip_path = os.path.join(os.path.dirname(path), f"{file_name}.zip")
                    
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        zipf.write(path, os.path.basename(path))
                    
                    processed_files += 1
                    # Update progress
                    progress = (processed_files / total_files) * 100
                    self.progress_var.set(progress)
                    self.root.update_idletasks()
                    
                    self.update_status(f"Compressed: {os.path.basename(path)}")
            except Exception as e:
                self.update_status(f"Error compressing {path}: {str(e)}")
                messagebox.showerror("Error", f"Failed to compress {path}: {str(e)}")
        
        # Play completion sound
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        
        # Reset UI state
        self.is_compressing = False
        self.compress_button.config(state=tk.NORMAL)
        self.progress_var.set(100)
        self.update_status("All files have been compressed successfully!")

    def update_status(self, message):
        self.status_label.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileCompressor(root)
    root.mainloop() 