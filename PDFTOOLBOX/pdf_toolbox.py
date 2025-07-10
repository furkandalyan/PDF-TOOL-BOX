import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import ttk, filedialog, messagebox, simpledialog
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import os
from typing import List, Tuple
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pytesseract
import threading

class PDFToolbox:
    def __init__(self):
        self.window = tb.Window(themename="darkly")
        self.window.title("PDF Toolbox - Advanced PDF Operations")
        self.window.geometry("700x600")
        self.window.resizable(True, True)
        self.history = []
        self.history_pointer = -1
        self.create_undo_redo_buttons()
        self.selected_files = []
        self.setup_ui()
        self.create_progress_bar()
        
    def setup_ui(self):
        # Create notebook for tabs
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.merge_tab = ttk.Frame(notebook)
        self.delete_tab = ttk.Frame(notebook)
        self.split_tab = ttk.Frame(notebook)
        self.encrypt_tab = ttk.Frame(notebook)
        self.rotate_tab = ttk.Frame(notebook)
        self.metadata_tab = ttk.Frame(notebook)
        
        notebook.add(self.merge_tab, text="Merge PDFs")
        notebook.add(self.delete_tab, text="Delete Pages")
        notebook.add(self.split_tab, text="Split PDF")
        notebook.add(self.encrypt_tab, text="Encrypt/Decrypt")
        notebook.add(self.rotate_tab, text="Rotate Pages")
        notebook.add(self.metadata_tab, text="Metadata Editor")
        
        self.setup_merge_tab()
        self.setup_delete_tab()
        self.setup_split_tab()
        self.setup_encrypt_tab()
        self.setup_rotate_tab()
        self.setup_metadata_tab()
    
    def setup_merge_tab(self):
        # Title
        title_label = tb.Label(self.merge_tab, text="Merge PDFs with Reordering", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)
        
        # File selection frame
        file_frame = tb.Labelframe(self.merge_tab, text="Select PDF Files", padding=10)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        select_btn = tb.Button(file_frame, text="Select PDF Files", command=self.select_files_for_merge)
        select_btn.pack(pady=5)
        
        # Files listbox with scrollbar
        list_frame = tb.Frame(file_frame)
        list_frame.pack(fill='both', expand=True, pady=5)
        
        self.files_listbox = tk.Listbox(list_frame, height=8, bg="#222", fg="#eee", selectbackground="#444", selectforeground="#fff", highlightbackground="#333", relief="flat")
        scrollbar = tb.Scrollbar(list_frame, orient="vertical", command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.files_listbox.pack(side="left", fill="both", expand=True, padx=(0, 2))
        scrollbar.pack(side="right", fill="y")
        
        # Reorder buttons
        reorder_frame = tb.Frame(file_frame)
        reorder_frame.pack(pady=5)
        
        tb.Button(reorder_frame, text="Move Up", command=self.move_file_up).pack(side="left", padx=2)
        tb.Button(reorder_frame, text="Move Down", command=self.move_file_down).pack(side="left", padx=2)
        tb.Button(reorder_frame, text="Remove", command=self.remove_file).pack(side="left", padx=2)
        tb.Button(reorder_frame, text="Clear All", command=self.clear_files).pack(side="left", padx=2)
        
        # Merge button
        merge_btn = tb.Button(self.merge_tab, text="Merge PDFs", command=self.merge_pdfs_with_reordering)
        merge_btn.pack(pady=10)
    
    def setup_delete_tab(self):
        # Title
        title_label = tb.Label(self.delete_tab, text="Delete Pages from PDF", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)
        
        # File selection
        file_frame = tb.Labelframe(self.delete_tab, text="Select PDF File", padding=10)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        self.delete_file_path = tb.StringVar()
        tb.Button(file_frame, text="Select PDF", command=self.select_file_for_delete).pack(pady=5)
        tb.Label(file_frame, textvariable=self.delete_file_path, wraplength=400).pack(pady=5)
        
        # PDF Preview (thumbnail)
        self.delete_preview_label = tb.Label(self.delete_tab)
        self.delete_preview_label.pack(pady=5)
        # OCR button
        self.ocr_btn = tb.Button(self.delete_tab, text="🧠 OCR This Page", command=self.ocr_current_page)  # type: ignore
        self.ocr_btn.pack(pady=2)
        
        # Multi-page preview navigation
        nav_frame = tb.Frame(self.delete_tab)
        nav_frame.pack(pady=2)
        self.delete_page_num = tb.IntVar(value=1)
        self.delete_total_pages = 1
        self.delete_pdf_doc = None
        self.delete_pdf_path = None
        self.delete_prev_btn = tb.Button(nav_frame, text="Previous", command=self.delete_prev_page, state="disabled")
        self.delete_prev_btn.pack(side="left", padx=2)
        self.delete_page_label = tb.Label(nav_frame, text="Page 1/1")
        self.delete_page_label.pack(side="left", padx=4)
        self.delete_next_btn = tb.Button(nav_frame, text="Next", command=self.delete_next_page, state="disabled")
        self.delete_next_btn.pack(side="left", padx=2)
        self.delete_slider = ttk.Scale(nav_frame, from_=1, to=1, orient="horizontal", command=self.delete_slider_move, state="disabled", length=200)
        self.delete_slider.pack(side="left", padx=8)
        
        # Page selection
        page_frame = tb.Labelframe(self.delete_tab, text="Select Pages to Delete", padding=10)
        page_frame.pack(fill='x', padx=10, pady=5)
        
        tb.Label(page_frame, text="Enter page numbers to delete (e.g., 1,3,5 or 2-4):").pack(pady=5)
        self.pages_to_delete = tb.StringVar()
        tb.Entry(page_frame, textvariable=self.pages_to_delete, width=40).pack(pady=5)
        
        # Delete button
        delete_btn = tb.Button(self.delete_tab, text="Delete Pages", command=self.delete_pages)
        delete_btn.pack(pady=10)
    
    def setup_split_tab(self):
        # Title
        title_label = tb.Label(self.split_tab, text="Split PDF by Page Range or Smart Suggestions", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)
        
        # File selection
        file_frame = tb.Labelframe(self.split_tab, text="Select PDF File", padding=10)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        self.split_file_path = tb.StringVar()
        tb.Button(file_frame, text="Select PDF", command=self.select_file_for_split).pack(pady=5)
        tb.Label(file_frame, textvariable=self.split_file_path, wraplength=400).pack(pady=5)
        
        # Split mode selection
        mode_frame = tb.Labelframe(self.split_tab, text="Split Mode", padding=10)
        mode_frame.pack(fill='x', padx=10, pady=5)
        self.split_mode = tb.StringVar(value="range")
        tb.Radiobutton(mode_frame, text="By Page Range", variable=self.split_mode, value="range", command=self.update_split_mode).pack(anchor="w")
        tb.Radiobutton(mode_frame, text="Every N Pages", variable=self.split_mode, value="every_n", command=self.update_split_mode).pack(anchor="w")
        tb.Radiobutton(mode_frame, text="Into N Equal Parts", variable=self.split_mode, value="equal_n", command=self.update_split_mode).pack(anchor="w")
        tb.Radiobutton(mode_frame, text="By Bookmarks (if present)", variable=self.split_mode, value="bookmarks", command=self.update_split_mode).pack(anchor="w")
        
        # Page range
        self.range_frame = tb.Labelframe(self.split_tab, text="Page Range", padding=10)
        self.range_frame.pack(fill='x', padx=10, pady=5)
        range_input_frame = tb.Frame(self.range_frame)
        range_input_frame.pack(pady=5)
        tb.Label(range_input_frame, text="From page:").pack(side="left", padx=5)
        self.start_page = tb.StringVar()
        tb.Entry(range_input_frame, textvariable=self.start_page, width=10).pack(side="left", padx=5)
        tb.Label(range_input_frame, text="To page:").pack(side="left", padx=5)
        self.end_page = tb.StringVar()
        tb.Entry(range_input_frame, textvariable=self.end_page, width=10).pack(side="left", padx=5)
        
        # Every N pages
        self.every_n_frame = tb.Labelframe(self.split_tab, text="Split Every N Pages", padding=10)
        self.every_n_frame.pack(fill='x', padx=10, pady=5)
        self.every_n_label = tb.Label(self.every_n_frame, text="N:")
        self.every_n_label.pack(side="left", padx=5)
        self.every_n_var = tb.StringVar()
        self.every_n_entry = tb.Entry(self.every_n_frame, textvariable=self.every_n_var, width=10)
        self.every_n_entry.pack(side="left", padx=5)
        
        # Equal N parts
        self.equal_n_frame = tb.Labelframe(self.split_tab, text="Split Into N Equal Parts", padding=10)
        self.equal_n_frame.pack(fill='x', padx=10, pady=5)
        self.equal_n_label = tb.Label(self.equal_n_frame, text="N:")
        self.equal_n_label.pack(side="left", padx=5)
        self.equal_n_var = tb.StringVar()
        self.equal_n_entry = tb.Entry(self.equal_n_frame, textvariable=self.equal_n_var, width=10)
        self.equal_n_entry.pack(side="left", padx=5)
        
        # Bookmarks info
        self.bookmarks_frame = tb.Labelframe(self.split_tab, text="Split by Bookmarks", padding=10)
        self.bookmarks_frame.pack(fill='x', padx=10, pady=5)
        self.bookmarks_label = tb.Label(self.bookmarks_frame, text="Splits at top-level bookmarks if present.")
        self.bookmarks_label.pack(padx=5, pady=5)
        
        # Split button
        split_btn = tb.Button(self.split_tab, text="Split PDF", command=self.smart_split_pdf)
        split_btn.pack(pady=10)
        self.update_split_mode()
    
    def setup_encrypt_tab(self):
        # Title
        title_label = tb.Label(self.encrypt_tab, text="Encrypt/Decrypt PDF", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)
        
        # File selection
        file_frame = tb.Labelframe(self.encrypt_tab, text="Select PDF File(s)", padding=10)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        self.encrypt_file_paths = []
        tb.Button(file_frame, text="Select PDF(s)", command=self.select_files_for_encrypt).pack(pady=5)
        self.encrypt_files_label = tb.Label(file_frame, text="No file selected", wraplength=400)
        self.encrypt_files_label.pack(pady=5)
        
        # Password
        password_frame = tb.Labelframe(self.encrypt_tab, text="Password", padding=10)
        password_frame.pack(fill='x', padx=10, pady=5)
        self.password_var = tb.StringVar()
        self.owner_password_var = tb.StringVar()
        tb.Label(password_frame, text="User Password:").pack(pady=2)
        tb.Entry(password_frame, textvariable=self.password_var, show="*", width=30).pack(pady=2)
        tb.Label(password_frame, text="Owner Password (optional):").pack(pady=2)
        tb.Entry(password_frame, textvariable=self.owner_password_var, show="*", width=30).pack(pady=2)
        # Permissions
        perm_frame = tb.Labelframe(self.encrypt_tab, text="Restrict Permissions", padding=10)
        perm_frame.pack(fill='x', padx=10, pady=5)
        self.perm_print = tb.BooleanVar(value=False)
        self.perm_copy = tb.BooleanVar(value=False)
        self.perm_edit = tb.BooleanVar(value=False)
        tb.Checkbutton(perm_frame, text="Disallow Printing", variable=self.perm_print).pack(anchor='w')
        tb.Checkbutton(perm_frame, text="Disallow Copying", variable=self.perm_copy).pack(anchor='w')
        tb.Checkbutton(perm_frame, text="Disallow Editing", variable=self.perm_edit).pack(anchor='w')
        
        # Buttons
        button_frame = tb.Frame(self.encrypt_tab)
        button_frame.pack(pady=10)
        
        tb.Button(button_frame, text="Encrypt PDF(s)", command=self.encrypt_pdf_batch).pack(side="left", padx=5)
        tb.Button(button_frame, text="Decrypt PDF", command=self.decrypt_pdf).pack(side="left", padx=5)
    
    def setup_rotate_tab(self):
        # Title
        title_label = tb.Label(self.rotate_tab, text="Rotate PDF Pages (Batch)", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)
        
        # File selection
        file_frame = tb.Labelframe(self.rotate_tab, text="Select PDF File(s)", padding=10)
        file_frame.pack(fill='x', padx=10, pady=5)

        self.rotate_file_paths = []
        tb.Button(file_frame, text="Select PDF(s)", command=self.select_files_for_rotate if hasattr(self, 'select_files_for_rotate') else lambda: None).pack(pady=5)
        self.rotate_files_label = tb.Label(file_frame, text="No file selected", wraplength=400)
        self.rotate_files_label.pack(pady=5)

        # Page selection
        page_frame = tb.Labelframe(self.rotate_tab, text="Page Selection", padding=10)
        page_frame.pack(fill='x', padx=10, pady=5)
        
        tb.Label(page_frame, text="Enter page numbers to rotate (e.g., 1,3,5 or 2-4, leave empty for all):").pack(pady=5)
        self.pages_to_rotate = tb.StringVar()
        tb.Entry(page_frame, textvariable=self.pages_to_rotate, width=40).pack(pady=5)
        
        # Rotation angle
        angle_frame = tb.Labelframe(self.rotate_tab, text="Rotation Angle", padding=10)
        angle_frame.pack(fill='x', padx=10, pady=5)
        
        self.rotation_angle = tb.StringVar(value="90")
        tb.Radiobutton(angle_frame, text="90° Clockwise", variable=self.rotation_angle, value="90").pack(anchor="w")
        tb.Radiobutton(angle_frame, text="180°", variable=self.rotation_angle, value="180").pack(anchor="w")
        tb.Radiobutton(angle_frame, text="270° Clockwise (90° Counter-clockwise)", variable=self.rotation_angle, value="270").pack(anchor="w")
        
        # Rotate button
        rotate_btn = tb.Button(self.rotate_tab, text="Rotate PDF(s)", command=self.rotate_pages_batch)
        rotate_btn.pack(pady=10)
    
    def setup_metadata_tab(self):
        # Title
        title_label = tb.Label(self.metadata_tab, text="PDF Metadata Editor", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)
        # File selection
        file_frame = tb.Labelframe(self.metadata_tab, text="Select PDF File", padding=10)
        file_frame.pack(fill='x', padx=10, pady=5)
        self.meta_file_path = tb.StringVar()
        tb.Button(file_frame, text="Select PDF", command=self.select_file_for_metadata).pack(pady=5)
        tb.Label(file_frame, textvariable=self.meta_file_path, wraplength=400).pack(pady=5)
        # Metadata fields
        meta_frame = tb.Labelframe(self.metadata_tab, text="Edit Metadata", padding=10)
        meta_frame.pack(fill='x', padx=10, pady=5)
        self.meta_title = tb.StringVar()
        self.meta_author = tb.StringVar()
        self.meta_subject = tb.StringVar()
        self.meta_keywords = tb.StringVar()
        tb.Label(meta_frame, text="Title:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        tb.Entry(meta_frame, textvariable=self.meta_title, width=40).grid(row=0, column=1, padx=5, pady=2)
        tb.Label(meta_frame, text="Author:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        tb.Entry(meta_frame, textvariable=self.meta_author, width=40).grid(row=1, column=1, padx=5, pady=2)
        tb.Label(meta_frame, text="Subject:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        tb.Entry(meta_frame, textvariable=self.meta_subject, width=40).grid(row=2, column=1, padx=5, pady=2)
        tb.Label(meta_frame, text="Keywords:").grid(row=3, column=0, sticky='e', padx=5, pady=2)
        tb.Entry(meta_frame, textvariable=self.meta_keywords, width=40).grid(row=3, column=1, padx=5, pady=2)
        # Buttons
        btn_frame = tb.Frame(self.metadata_tab)
        btn_frame.pack(pady=10)
        tb.Button(btn_frame, text="Save Metadata", command=self.save_metadata).pack(side='left', padx=5)
        tb.Button(btn_frame, text="Clear Metadata", command=self.clear_metadata).pack(side='left', padx=5)
    
    def create_undo_redo_buttons(self):
        # Place Undo/Redo buttons at the top left
        top_frame = ttk.Frame(self.window)
        top_frame.pack(fill='x', side='top', anchor='nw')
        self.undo_btn = tb.Button(top_frame, text="Undo", command=self.undo_action)
        self.undo_btn.pack(side='left', padx=5, pady=5)
        self.redo_btn = tb.Button(top_frame, text="Redo", command=self.redo_action)
        self.redo_btn.pack(side='left', padx=5, pady=5)
        self.update_undo_redo_buttons()
    def update_undo_redo_buttons(self):
        self.undo_btn.config(state="normal" if self.history_pointer >= 0 else "disabled")
        self.redo_btn.config(state="normal" if self.history_pointer < len(self.history) - 1 else "disabled")
    def add_history(self, action, file_path):
        # Remove any redo history
        self.history = self.history[:self.history_pointer+1]
        self.history.append((action, file_path))
        self.history_pointer += 1
        self.update_undo_redo_buttons()
    def undo_action(self):
        if self.history_pointer >= 0:
            action, file_path = self.history[self.history_pointer]
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    messagebox.showinfo("Undo", f"Undid {action}. Deleted {os.path.basename(file_path)}.")
                except Exception as e:
                    messagebox.showerror("Undo Error", f"Failed to delete {file_path}: {e}")
            self.history_pointer -= 1
            self.update_undo_redo_buttons()
    def redo_action(self):
        if self.history_pointer < len(self.history) - 1:
            self.history_pointer += 1
            action, file_path = self.history[self.history_pointer]
            messagebox.showinfo("Redo", f"Redo: {action} (file {os.path.basename(file_path)})")
            self.update_undo_redo_buttons()
    
    def create_progress_bar(self):
        bottom_frame = ttk.Frame(self.window)
        bottom_frame.pack(fill='x', side='bottom', anchor='s')
        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate')
        self.progress.pack(fill='x', padx=10, pady=2)
        self.status_label = ttk.Label(bottom_frame, text="Ready")
        self.status_label.pack(side='left', padx=10)
    def start_progress(self, status="Processing..."):
        self.progress.start(10)
        self.status_label.config(text=status)
        self.window.update_idletasks()
    def stop_progress(self, status="Done"):
        self.progress.stop()
        self.status_label.config(text=status)
        self.window.update_idletasks()
    
    # File selection methods
    def select_files_for_merge(self):
        files = filedialog.askopenfilenames(
            title="Select PDF files to merge",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            self.selected_files.extend(files)
            self.update_files_listbox()
    
    def select_file_for_delete(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.delete_file_path.set(file_path)
            self.delete_pdf_path = file_path
            try:
                self.delete_pdf_doc = fitz.open(file_path)
                self.delete_total_pages = self.delete_pdf_doc.page_count
            except Exception:
                self.delete_pdf_doc = None
                self.delete_total_pages = 1
            self.delete_page_num.set(1)
            self.update_delete_preview()
    
    def select_file_for_split(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.split_file_path.set(file_path)
    
    def select_files_for_encrypt(self):
        files = filedialog.askopenfilenames(
            title="Select PDF file(s)",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            self.encrypt_file_paths = list(files)
            self.encrypt_files_label.config(text="\n".join([os.path.basename(f) for f in self.encrypt_file_paths]))
        else:
            self.encrypt_file_paths = []
            self.encrypt_files_label.config(text="No file selected")
    
    def select_file_for_rotate(self):  # type: ignore
        file_path = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.rotate_file_path = tb.StringVar(value=file_path)  # type: ignore
    
    def select_file_for_metadata(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.meta_file_path.set(file_path)
            try:
                reader = PdfReader(file_path)
                info = reader.metadata
                self.meta_title.set(info.title if info.title else "")  # type: ignore
                self.meta_author.set(info.author if info.author else "")  # type: ignore
                self.meta_subject.set(info.subject if info.subject else "")  # type: ignore
                self.meta_keywords.set(info.keywords if info.keywords else "")  # type: ignore
            except Exception:
                self.meta_title.set("")
                self.meta_author.set("")
                self.meta_subject.set("")
                self.meta_keywords.set("")
    
    # Merge with reordering methods
    def update_files_listbox(self):
        self.files_listbox.delete(0, tk.END)
        for file_path in self.selected_files:
            self.files_listbox.insert(tk.END, os.path.basename(file_path))
    
    def move_file_up(self):
        selection = self.files_listbox.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            self.selected_files[index], self.selected_files[index-1] = self.selected_files[index-1], self.selected_files[index]
            self.update_files_listbox()
            self.files_listbox.selection_set(index-1)
    
    def move_file_down(self):
        selection = self.files_listbox.curselection()
        if selection and selection[0] < len(self.selected_files) - 1:
            index = selection[0]
            self.selected_files[index], self.selected_files[index+1] = self.selected_files[index+1], self.selected_files[index]
            self.update_files_listbox()
            self.files_listbox.selection_set(index+1)
    
    def remove_file(self):
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            del self.selected_files[index]
            self.update_files_listbox()
    
    def clear_files(self):
        self.selected_files.clear()
        self.update_files_listbox()
    
    def merge_pdfs_with_reordering(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please select at least one PDF file.")
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            title="Save Merged PDF As"
        )
        
        if save_path:
            def do_merge():
                self.start_progress("Merging PDFs...")
                try:
                    merger = PdfMerger()
                    for pdf in self.selected_files:
                        merger.append(pdf)
                    merger.write(save_path)
                    merger.close()
                    messagebox.showinfo("Success", "PDFs merged successfully!")
                    self.add_history('Merge PDFs', save_path)
                except Exception as e:
                    messagebox.showerror("Error", f"Error merging PDFs: {str(e)}")
                self.stop_progress("Ready")
            threading.Thread(target=do_merge).start()
    
    # Delete pages method
    def delete_pages(self):
        file_path = self.delete_file_path.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        
        pages_input = self.pages_to_delete.get().strip()
        if not pages_input:
            messagebox.showwarning("Warning", "Please enter page numbers to delete.")
            return
        
        try:
            # Parse page numbers
            pages_to_delete = self.parse_page_numbers(pages_input)
            
            # Read PDF
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            
            # Validate page numbers
            for page_num in pages_to_delete:
                if page_num < 1 or page_num > total_pages:
                    messagebox.showerror("Error", f"Page {page_num} is out of range (1-{total_pages})")
                    return
            
            # Create new PDF without deleted pages
            writer = PdfWriter()
            for i in range(total_pages):
                if i + 1 not in pages_to_delete:
                    writer.add_page(reader.pages[i])
            
            # Save
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Save PDF As"
            )
            
            if save_path:
                with open(save_path, 'wb') as output_file:
                    writer.write(output_file)
                messagebox.showinfo("Success", f"Pages {pages_to_delete} deleted successfully!")
                self.add_history('Delete Pages', save_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error deleting pages: {str(e)}")
    
    # Split PDF method
    def update_split_mode(self):
        mode = self.split_mode.get()
        self.range_frame.pack_forget()
        self.every_n_frame.pack_forget()
        self.equal_n_frame.pack_forget()
        self.bookmarks_frame.pack_forget()
        if mode == "range":
            self.range_frame.pack(fill='x', padx=10, pady=5)
        elif mode == "every_n":
            self.every_n_frame.pack(fill='x', padx=10, pady=5)
        elif mode == "equal_n":
            self.equal_n_frame.pack(fill='x', padx=10, pady=5)
        elif mode == "bookmarks":
            self.bookmarks_frame.pack(fill='x', padx=10, pady=5)
    
    def smart_split_pdf(self):
        file_path = self.split_file_path.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        mode = self.split_mode.get()
        try:
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            results = []
            if mode == "range":
                start = self.start_page.get().strip()
                end = self.end_page.get().strip()
                if not start or not end:
                    messagebox.showwarning("Warning", "Please enter start and end page numbers.")
                    return
                start_page = int(start)
                end_page = int(end)
                if start_page < 1 or end_page > total_pages or start_page > end_page:
                    messagebox.showerror("Error", f"Invalid page range. PDF has {total_pages} pages.")
                    return
                writer = PdfWriter()
                for i in range(start_page - 1, end_page):
                    writer.add_page(reader.pages[i])
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Files", "*.pdf")],
                    title="Save Split PDF As"
                )
                if save_path:
                    with open(save_path, 'wb') as output_file:
                        writer.write(output_file)
                    results.append(f"Pages {start_page}-{end_page}: Success")
                    self.add_history('Split PDF (Range)', save_path)
            elif mode == "every_n":
                n = int(self.every_n_var.get())
                part = 1
                for i in range(0, total_pages, n):
                    writer = PdfWriter()
                    for j in range(i, min(i + n, total_pages)):
                        writer.add_page(reader.pages[j])
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        filetypes=[("PDF Files", "*.pdf")],
                        title=f"Save Part {part} (pages {i+1}-{min(i+n, total_pages)})"
                    )
                    if save_path:
                        with open(save_path, 'wb') as output_file:
                            writer.write(output_file)
                        results.append(f"Part {part} (pages {i+1}-{min(i+n, total_pages)}): Success")
                        self.add_history('Split PDF (Every N)', save_path)
                    part += 1
            elif mode == "equal_n":
                n = int(self.equal_n_var.get())
                pages_per_part = total_pages // n
                extra = total_pages % n
                start = 0
                for part in range(1, n+1):
                    end = start + pages_per_part + (1 if part <= extra else 0)
                    writer = PdfWriter()
                    for j in range(start, end):
                        writer.add_page(reader.pages[j])
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        filetypes=[("PDF Files", "*.pdf")],
                        title=f"Save Part {part} (pages {start+1}-{end})"
                    )
                    if save_path:
                        with open(save_path, 'wb') as output_file:
                            writer.write(output_file)
                        results.append(f"Part {part} (pages {start+1}-{end}): Success")
                        self.add_history('Split PDF (Equal N)', save_path)
                    start = end
            elif mode == "bookmarks":
                # Find top-level bookmarks (outlines)
                outlines = []
                try:
                    outlines = reader.outline
                except Exception:
                    pass
                if not outlines:
                    messagebox.showinfo("No Bookmarks", "No bookmarks found in this PDF.")
                    return
                # Only handle top-level bookmarks
                page_indices = []
                for item in outlines:
                    if hasattr(item, 'page') and item.page is not None:  # type: ignore
                        page_indices.append(item.page)  # type: ignore
                page_indices = sorted(set(page_indices))
                page_indices.append(total_pages)  # Add end
                for idx in range(len(page_indices)-1):
                    start = page_indices[idx]
                    end = page_indices[idx+1]
                    writer = PdfWriter()
                    for j in range(start, end):
                        writer.add_page(reader.pages[j])
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        filetypes=[("PDF Files", "*.pdf")],
                        title=f"Save Section {idx+1} (pages {start+1}-{end})"
                    )
                    if save_path:
                        with open(save_path, 'wb') as output_file:
                            writer.write(output_file)
                        results.append(f"Section {idx+1} (pages {start+1}-{end}): Success")
                        self.add_history('Split PDF (Bookmarks)', save_path)
            if results:
                messagebox.showinfo("Split Results", "\n".join(results))
        except Exception as e:
            messagebox.showerror("Error", f"Error splitting PDF: {str(e)}")
    
    # Encrypt/Decrypt methods
    def encrypt_pdf(self):
        file_path = self.encrypt_file_path.get()  # type: ignore
        if not file_path:
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        
        password = self.password_var.get()
        if not password:
            messagebox.showwarning("Warning", "Please enter a password.")
            return
        
        try:
            # Read PDF
            reader = PdfReader(file_path)
            writer = PdfWriter()
            
            # Add all pages
            for page in reader.pages:
                writer.add_page(page)
            
            # Encrypt
            writer.encrypt(password)
            
            # Save
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Save Encrypted PDF As"
            )
            
            if save_path:
                with open(save_path, 'wb') as output_file:
                    writer.write(output_file)
                messagebox.showinfo("Success", "PDF encrypted successfully!")
                self.add_history('Encrypt PDF', save_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error encrypting PDF: {str(e)}")
    
    def decrypt_pdf(self):
        file_path = self.encrypt_file_path.get()  # type: ignore
        if not file_path:
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        
        password = self.password_var.get()
        if not password:
            messagebox.showwarning("Warning", "Please enter the password.")
            return
        
        try:
            # Read PDF
            reader = PdfReader(file_path)
            
            # Try to decrypt
            if reader.is_encrypted:
                reader.decrypt(password)
            else:
                messagebox.showinfo("Info", "This PDF is not encrypted.")
                return
            
            # Create new PDF
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            
            # Save
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Save Decrypted PDF As"
            )
            
            if save_path:
                with open(save_path, 'wb') as output_file:
                    writer.write(output_file)
                messagebox.showinfo("Success", "PDF decrypted successfully!")
                self.add_history('Decrypt PDF', save_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error decrypting PDF: {str(e)}")
    
    def encrypt_pdf_batch(self):
        file_paths = self.encrypt_file_paths
        if not file_paths:
            messagebox.showwarning("Warning", "Please select at least one PDF file.")
            return
        user_password = self.password_var.get()
        owner_password = self.owner_password_var.get() or None
        if not user_password:
            messagebox.showwarning("Warning", "Please enter a user password.")
            return
        # Permissions
        from PyPDF2.constants import Permissions  # type: ignore
        permissions = set()  # type: ignore
        if not self.perm_print.get():
            permissions.add(Permissions.PRINTING)  # type: ignore
        if not self.perm_copy.get():
            permissions.add(Permissions.COPYING)  # type: ignore
        if not self.perm_edit.get():
            permissions.add(Permissions.MODIFYING)  # type: ignore
        results = []
        for file_path in file_paths:
            try:
                reader = PdfReader(file_path)
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                # Advanced encryption
                writer.encrypt(
                    user_password,
                    owner_password,
                    permissions=permissions if permissions else None  # type: ignore
                )  # type: ignore
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Files", "*.pdf")],
                    title=f"Save Encrypted PDF As (for {os.path.basename(file_path)})"
                )
                if save_path:
                    with open(save_path, 'wb') as output_file:
                        writer.write(output_file)
                    results.append(f"{os.path.basename(file_path)}: Success")
                    self.add_history('Batch Encrypt PDF', save_path)
                else:
                    results.append(f"{os.path.basename(file_path)}: Skipped (no save path)")
            except Exception as e:
                results.append(f"{os.path.basename(file_path)}: Error - {str(e)}")
        messagebox.showinfo("Batch Encryption Results", "\n".join(results))
    
    # Rotate pages method
    def rotate_pages(self):
        file_path = self.rotate_file_path.get()  # type: ignore
        if not file_path:
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        
        pages_input = self.pages_to_rotate.get().strip()
        angle = int(self.rotation_angle.get())
        
        try:
            # Read PDF
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            
            # Determine which pages to rotate
            if pages_input:
                pages_to_rotate = self.parse_page_numbers(pages_input)
                # Validate page numbers
                for page_num in pages_to_rotate:
                    if page_num < 1 or page_num > total_pages:
                        messagebox.showerror("Error", f"Page {page_num} is out of range (1-{total_pages})")
                        return
            else:
                # Rotate all pages
                pages_to_rotate = list(range(1, total_pages + 1))
            
            # Create new PDF with rotated pages
            writer = PdfWriter()
            for i in range(total_pages):
                page = reader.pages[i]
                if i + 1 in pages_to_rotate:
                    page.rotate(angle)
                writer.add_page(page)
            
            # Save
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Save Rotated PDF As"
            )
            
            if save_path:
                with open(save_path, 'wb') as output_file:
                    writer.write(output_file)
                messagebox.showinfo("Success", f"Pages rotated successfully!")
                self.add_history('Rotate PDF', save_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error rotating pages: {str(e)}")
    
    def rotate_pages_batch(self):
        file_paths = self.rotate_file_paths
        if not file_paths:
            messagebox.showwarning("Warning", "Please select at least one PDF file.")
            return
        pages_input = self.pages_to_rotate.get().strip()
        angle = int(self.rotation_angle.get())
        results = []
        for file_path in file_paths:
            try:
                reader = PdfReader(file_path)
                total_pages = len(reader.pages)
                if pages_input:
                    pages_to_rotate = self.parse_page_numbers(pages_input)
                    for page_num in pages_to_rotate:
                        if page_num < 1 or page_num > total_pages:
                            results.append(f"{os.path.basename(file_path)}: Page {page_num} out of range")
                            continue
                else:
                    pages_to_rotate = list(range(1, total_pages + 1))
                writer = PdfWriter()
                for i in range(total_pages):
                    page = reader.pages[i]
                    if i + 1 in pages_to_rotate:
                        page.rotate(angle)
                    writer.add_page(page)
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Files", "*.pdf")],
                    title=f"Save Rotated PDF As (for {os.path.basename(file_path)})"
                )
                if save_path:
                    with open(save_path, 'wb') as output_file:
                        writer.write(output_file)
                    results.append(f"{os.path.basename(file_path)}: Success")
                    self.add_history('Batch Rotate PDF', save_path)
                else:
                    results.append(f"{os.path.basename(file_path)}: Skipped (no save path)")
            except Exception as e:
                results.append(f"{os.path.basename(file_path)}: Error - {str(e)}")
        messagebox.showinfo("Batch Rotation Results", "\n".join(results))
    
    # Utility method to parse page numbers
    def parse_page_numbers(self, input_str: str) -> List[int]:
        """Parse page numbers from string like '1,3,5' or '2-4' or '1,3-5,7'"""
        pages = set()
        parts = input_str.split(',')
        
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                start = int(start.strip())
                end = int(end.strip())
                pages.update(range(start, end + 1))
            else:
                pages.add(int(part))
        
        return sorted(list(pages))
    
    def update_delete_preview(self):
        page_num = self.delete_page_num.get()
        total = self.delete_total_pages
        self.delete_page_label.config(text=f"Page {page_num}/{total}")
        self.delete_slider.config(from_=1, to=total, state="normal" if total > 1 else "disabled")
        self.delete_slider.set(page_num)
        self.delete_prev_btn.config(state="normal" if page_num > 1 else "disabled")
        self.delete_next_btn.config(state="normal" if page_num < total else "disabled")
        if self.delete_pdf_doc:
            self.show_pdf_preview(self.delete_pdf_path, self.delete_preview_label, page_num-1)
        else:
            self.delete_preview_label.config(text="Preview unavailable")
    def delete_prev_page(self):
        if self.delete_page_num.get() > 1:
            self.delete_page_num.set(self.delete_page_num.get() - 1)
            self.update_delete_preview()
    def delete_next_page(self):
        if self.delete_page_num.get() < self.delete_total_pages:
            self.delete_page_num.set(self.delete_page_num.get() + 1)
            self.update_delete_preview()
    def delete_slider_move(self, val):
        val = int(float(val))
        if val != self.delete_page_num.get():
            self.delete_page_num.set(val)
            self.update_delete_preview()
    
    def show_pdf_preview(self, file_path, label_widget, page_number=0):
        try:
            doc = fitz.open(file_path)
            page = doc.load_page(page_number)
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(0.3, 0.3))  # type: ignore
            except AttributeError:
                pix = page.getPixmap(matrix=fitz.Matrix(0.3, 0.3))  # type: ignore
            mode = "RGB" if pix.n < 4 else "RGBA"
            img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
            img_tk = ImageTk.PhotoImage(img)
            label_widget.configure(image=img_tk, text="")  # type: ignore
            label_widget.image = img_tk  # type: ignore
        except Exception as e:
            label_widget.configure(text=f"Preview unavailable: {e}", image=None)  # type: ignore
            label_widget.image = None  # type: ignore
    
    def ocr_current_page(self):
        file_path = self.delete_pdf_path
        page_num = self.delete_page_num.get() - 1
        try:
            doc = fitz.open(file_path)
            page = doc.load_page(page_num)
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            except AttributeError:
                pix = page.getPixmap(matrix=fitz.Matrix(2, 2))
            mode = "RGB" if pix.n < 4 else "RGBA"
            img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
            text = pytesseract.image_to_string(img)
            self.show_ocr_result(text)
        except Exception as e:
            messagebox.showerror("OCR Error", f"Failed to extract text: {e}")

    def show_ocr_result(self, text):
        ocr_win = tb.Toplevel(self.window)
        ocr_win.title("OCR Result")
        ocr_win.geometry("600x400")
        text_widget = tb.Text(ocr_win, wrap="word")
        text_widget.insert("1.0", text)
        text_widget.pack(fill="both", expand=True)
        def save_txt():
            save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")], title="Save OCR Text As")
            if save_path:
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(text_widget.get("1.0", "end-1c"))
        save_btn = tb.Button(ocr_win, text="Save as .txt", command=save_txt)
        save_btn.pack(pady=5)
    
    def save_metadata(self):
        file_path = self.meta_file_path.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            metadata = {
                '/Title': self.meta_title.get(),
                '/Author': self.meta_author.get(),
                '/Subject': self.meta_subject.get(),
                '/Keywords': self.meta_keywords.get(),
            }
            writer.add_metadata(metadata)
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Save PDF with New Metadata"
            )
            if save_path:
                with open(save_path, 'wb') as output_file:
                    writer.write(output_file)
                messagebox.showinfo("Success", "Metadata saved successfully!")
                self.add_history('Save Metadata', save_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error saving metadata: {e}")
    def clear_metadata(self):
        self.meta_title.set("")
        self.meta_author.set("")
        self.meta_subject.set("")
        self.meta_keywords.set("")
    
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = PDFToolbox()
    app.run()
