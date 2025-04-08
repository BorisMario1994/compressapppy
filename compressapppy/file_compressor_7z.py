import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import winsound
import threading
import openpyxl
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P

class FileCompressor:
    def __init__(self, root):
        self.root = root
        self.root.title("File Compressor (7-Zip)")
        self.root.geometry("600x400")
        
        # List of compressed file extensions to skip
        self.compressed_extensions = {
            '.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz', '.iso',
            '.cab', '.arj', '.lzh', '.lha', '.ace', '.tar.gz', '.tar.bz2',
            '.tar.xz', '.tgz', '.tbz2', '.txz', '.z', '.zipx', '.war', '.jar',
            '.ear', '.sar', '.apk', '.ipa', '.msi', '.msp', '.msm', '.mst'
        }
        
        # Create main frame
        self.main_frame = tk.Frame(self.root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create file input frame
        self.file_frame = tk.Frame(self.main_frame)
        self.file_frame.pack(fill=tk.X, pady=10)
        
        self.file_label = tk.Label(self.file_frame, text="Select Excel/OpenOffice File:")
        self.file_label.pack(side=tk.LEFT)
        
        self.file_var = tk.StringVar()
        self.file_entry = tk.Entry(self.file_frame, textvariable=self.file_var, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_button = tk.Button(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT)
        
        # Create listbox to display items to compress
        self.items_listbox = tk.Listbox(self.main_frame, height=10, width=50)
        self.items_listbox.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create buttons frame
        self.buttons_frame = tk.Frame(self.main_frame)
        self.buttons_frame.pack(fill=tk.X, pady=10)
        
        # Compress button
        self.compress_button = tk.Button(self.buttons_frame, text="Compress Files/Folders", command=self.start_compression)
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
        
        self.is_compressing = False
        self.paths_to_compress = []
        
        # Check if 7-Zip is installed
        self.check_7zip()

    def check_7zip(self):
        try:
            # Try to find 7-Zip in common installation locations
            possible_paths = [
                r"C:\Program Files\7-Zip\7z.exe",
                r"C:\Program Files (x86)\7-Zip\7z.exe"
            ]
            
            self.seven_zip_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    self.seven_zip_path = path
                    break
            
            if not self.seven_zip_path:
                messagebox.showerror("Error", "7-Zip not found. Please install 7-Zip from https://www.7-zip.org/")
                self.compress_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to check 7-Zip installation: {str(e)}")
            self.compress_button.config(state=tk.DISABLED)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel or OpenOffice File",
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"),
                ("OpenOffice files", "*.ods"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_var.set(file_path)
            self.read_file(file_path)

    def read_file(self, file_path):
        try:
            self.paths_to_compress.clear()
            self.items_listbox.delete(0, tk.END)
            
            if file_path.endswith(('.xlsx', '.xls')):
                self.read_excel_file(file_path)
            elif file_path.endswith('.ods'):
                self.read_ods_file(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format")
                return
                
            self.update_items_list()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")

    def read_excel_file(self, file_path):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        # Read paths from second row onwards
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row and row[0]:  # Check if first column has a value
                path = str(row[0]).strip()
                if os.path.exists(path):
                    if os.path.isdir(path):
                        # Add all items in the directory to paths_to_compress
                        for item in os.listdir(path):
                            item_path = os.path.join(path, item)
                            self.paths_to_compress.append(item_path)
                            self.update_status(f"Found item in directory: {item_path}")
                    else:
                        self.paths_to_compress.append(path)
                        self.update_status(f"Found valid path: {path}")
                else:
                    self.update_status(f"Warning: Path not found - {path}")
        
        if not self.paths_to_compress:
            messagebox.showwarning("Warning", "No valid file paths found in the Excel file")
        else:
            self.update_status(f"Loaded {len(self.paths_to_compress)} items to process")

    def read_ods_file(self, file_path):
        doc = load(file_path)
        table = doc.getElementsByType(Table)[0]
        
        # Read paths from second row onwards
        for row in table.getElementsByType(TableRow)[1:]:  # Skip header row
            cell = row.getElementsByType(TableCell)[0]
            path = ""
            for p in cell.getElementsByType(P):
                path += p.firstChild.data if p.firstChild else ""
            
            path = path.strip()
            if path and os.path.exists(path):
                if os.path.isdir(path):
                    # Add all items in the directory to paths_to_compress
                    for item in os.listdir(path):
                        item_path = os.path.join(path, item)
                        self.paths_to_compress.append(item_path)
                        self.update_status(f"Found item in directory: {item_path}")
                else:
                    self.paths_to_compress.append(path)
                    self.update_status(f"Found valid path: {path}")
            elif path:
                self.update_status(f"Warning: Path not found - {path}")
        
        if not self.paths_to_compress:
            messagebox.showwarning("Warning", "No valid file paths found in the OpenOffice file")
        else:
            self.update_status(f"Loaded {len(self.paths_to_compress)} items to process")

    def update_items_list(self):
        self.items_listbox.delete(0, tk.END)
        for path in self.paths_to_compress:
            if os.path.isdir(path):
                self.items_listbox.insert(tk.END, f"[Folder] {os.path.basename(path)}")
            elif os.path.isfile(path):
                self.items_listbox.insert(tk.END, f"[File] {os.path.basename(path)}")

    def start_compression(self):
        if not self.paths_to_compress:
            messagebox.showwarning("Warning", "Please select a file with paths first")
            return
            
        if self.is_compressing:
            return
            
        if not self.seven_zip_path:
            messagebox.showerror("Error", "7-Zip not found. Please install 7-Zip from https://www.7-zip.org/")
            return
            
        self.is_compressing = True
        self.compress_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        # Start compression in a separate thread
        threading.Thread(target=self.compress_items, daemon=True).start()

    def compress_items(self):
        total_items = len(self.paths_to_compress)
        processed_items = 0
        
        self.update_status(f"Starting compression of {total_items} items...")
        
        for path in self.paths_to_compress:
            try:
                self.update_status(f"Processing: {path}")
                
                # Skip if it's a compressed file
                if os.path.isfile(path) and any(path.lower().endswith(ext) for ext in self.compressed_extensions):
                    self.update_status(f"Skipping already compressed item: {os.path.basename(path)}")
                    processed_items += 1
                    continue
                
                if os.path.isdir(path):
                    self.update_status(f"Compressing folder: {path}")
                    # Check if folder is already compressed
                    if os.path.exists(f"{path}.7z"):
                        self.update_status(f"Skipping already compressed folder: {os.path.basename(path)}")
                        processed_items += 1
                        continue
                        
                    # Create 7z file for the folder
                    seven_zip_path = f"{path}.7z"
                    self.update_status(f"Creating 7z file: {seven_zip_path}")
                    
                    # Use 7-Zip to compress the folder
                    cmd = [
                        self.seven_zip_path,
                        'a',  # Add files to archive
                        '-t7z',  # Use 7z format
                        '-mx=9',  # Maximum compression
                        '-m0=lzma2',  # Use LZMA2 compression method
                        '-mfb=64',  # Set number of fast bytes
                        '-md=32m',  # Set dictionary size
                        '-ms=on',  # Enable solid mode
                        seven_zip_path,
                        path
                    ]
                    
                    process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    stdout, stderr = process.communicate()
                    
                    if process.returncode != 0:
                        raise Exception(f"7-Zip error: {stderr.decode()}")
                    
                    self.update_status(f"Successfully compressed folder: {os.path.basename(path)}")
                    
                elif os.path.isfile(path):
                    self.update_status(f"Compressing file: {path}")
                    # Check if 7z file already exists
                    file_name = os.path.splitext(os.path.basename(path))[0]
                    seven_zip_path = os.path.join(os.path.dirname(path), f"{file_name}.7z")
                    if os.path.exists(seven_zip_path):
                        self.update_status(f"Skipping already compressed file: {os.path.basename(path)}")
                        processed_items += 1
                        continue
                    
                    self.update_status(f"Creating 7z file: {seven_zip_path}")
                    
                    # Use 7-Zip to compress the file
                    cmd = [
                        self.seven_zip_path,
                        'a',  # Add files to archive
                        '-t7z',  # Use 7z format
                        '-mx=9',  # Maximum compression
                        '-m0=lzma2',  # Use LZMA2 compression method
                        '-mfb=64',  # Set number of fast bytes
                        '-md=32m',  # Set dictionary size
                        '-ms=on',  # Enable solid mode
                        seven_zip_path,
                        path
                    ]
                    
                    process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    stdout, stderr = process.communicate()
                    
                    if process.returncode != 0:
                        raise Exception(f"7-Zip error: {stderr.decode()}")
                    
                    self.update_status(f"Successfully compressed file: {os.path.basename(path)}")
                
                processed_items += 1
                progress = (processed_items / total_items) * 100
                self.progress_var.set(progress)
                self.root.update_idletasks()
                
            except Exception as e:
                self.update_status(f"Error compressing {os.path.basename(path)}: {str(e)}")
                messagebox.showerror("Error", f"Failed to compress {os.path.basename(path)}: {str(e)}")
        
        # Play completion sound
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        
        self.is_compressing = False
        self.compress_button.config(state=tk.NORMAL)
        self.update_status("Compression completed!")

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = FileCompressor(root)
    root.mainloop() 