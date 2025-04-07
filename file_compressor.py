import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import zipfile
from pathlib import Path
import winsound
import threading

class FileCompressor:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Compressor")
        self.root.geometry("600x400")
        
        # Create main frame
        self.main_frame = tk.Frame(self.root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create listbox to display selected folders
        self.folders_listbox = tk.Listbox(self.main_frame, height=10, width=50)
        self.folders_listbox.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create buttons frame
        self.buttons_frame = tk.Frame(self.main_frame)
        self.buttons_frame.pack(fill=tk.X, pady=10)
        
        # Add folder button
        self.add_button = tk.Button(self.buttons_frame, text="Add Folder", command=self.add_folder)
        self.add_button.pack(side=tk.LEFT, padx=5)
        
        # Remove folder button
        self.remove_button = tk.Button(self.buttons_frame, text="Remove Selected", command=self.remove_folder)
        self.remove_button.pack(side=tk.LEFT, padx=5)
        
        # Compress button
        self.compress_button = tk.Button(self.buttons_frame, text="Compress Folders", command=self.start_compression)
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
        
        # Store selected folders
        self.selected_folders = []
        self.is_compressing = False

    def start_compression(self):
        if not self.selected_folders:
            messagebox.showwarning("Warning", "Please select at least one folder to compress")
            return
            
        if self.is_compressing:
            return
            
        self.is_compressing = True
        self.compress_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        # Start compression in a separate thread
        threading.Thread(target=self.compress_folders, daemon=True).start()

    def compress_folders(self):
        total_folders = len(self.selected_folders)
        total_files = 0
        processed_files = 0
        
        # First, count total files across all folders
        for folder_path in self.selected_folders:
            for root, dirs, files in os.walk(folder_path):
                total_files += len(files)
        
        for folder_path in self.selected_folders:
            try:
                for root, dirs, files in os.walk(folder_path):
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
                            zipf.write(file_path, file)  # Keep original filename in the zip
                        
                        processed_files += 1
                        # Update progress
                        progress = (processed_files / total_files) * 100
                        self.progress_var.set(progress)
                        self.root.update_idletasks()
                        
                        self.update_status(f"Compressed: {file}")
            except Exception as e:
                self.update_status(f"Error compressing files in {folder_path}: {str(e)}")
                messagebox.showerror("Error", f"Failed to compress files in {folder_path}: {str(e)}")
        
        # Play completion sound
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        
        # Reset UI state
        self.is_compressing = False
        self.compress_button.config(state=tk.NORMAL)
        self.progress_var.set(100)
        self.update_status("All files have been compressed successfully!")

    def add_folder(self):
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path and folder_path not in self.selected_folders:
            self.selected_folders.append(folder_path)
            self.folders_listbox.insert(tk.END, folder_path)
            self.update_status(f"Added folder: {os.path.basename(folder_path)}")

    def remove_folder(self):
        selected_index = self.folders_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            removed_folder = self.selected_folders.pop(index)
            self.folders_listbox.delete(index)
            self.update_status(f"Removed folder: {os.path.basename(removed_folder)}")

    def update_status(self, message):
        self.status_label.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileCompressor(root)
    root.mainloop() 