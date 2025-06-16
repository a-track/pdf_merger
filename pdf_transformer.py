import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import io
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
import threading

class DragDropListbox(tk.Listbox):
    """A listbox that supports drag and drop reordering"""
    
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        
        self.bind('<Button-1>', self.on_click)
        self.bind('<B1-Motion>', self.on_drag)
        self.bind('<ButtonRelease-1>', self.on_drop)
        
        self.drag_start_index = None
    
    def on_click(self, event):
        self.drag_start_index = self.nearest(event.y)
    
    def on_drag(self, event):
        pass  # Visual feedback could be added here
    
    def on_drop(self, event):
        if self.drag_start_index is not None:
            drop_index = self.nearest(event.y)
            if drop_index != self.drag_start_index:
                # Get the item being moved
                item = self.get(self.drag_start_index)
                # Delete from old position
                self.delete(self.drag_start_index)
                # Insert at new position
                self.insert(drop_index, item)
                # Select the moved item
                self.selection_clear(0, tk.END)
                self.selection_set(drop_index)
        
        self.drag_start_index = None

class PDFPageSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("🐹 Andrin's PDF Page Selector & Merger")
        self.root.geometry("900x600")
        
        # Try to set a hamster icon for the title bar
        try:
            # If you have a hamster.ico file, uncomment the next line:
            # self.root.iconbitmap('hamster.ico')
            pass
        except:
            pass  # If icon setting fails, just continue without it
        
        # Variables
        self.input_path = ""
        self.output_path = ""
        self.all_pages = []  # List of (pdf_file, page_number, page_object)
        self.selected_pages = []  # List of selected pages for merging
        
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title with hamster emoji
        title_label = ttk.Label(main_frame, text="🐹 Andrin's PDF Page Selector & Merger", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Path selection frame
        path_frame = ttk.LabelFrame(main_frame, text="Paths", padding="10")
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Input path
        input_frame = ttk.Frame(path_frame)
        input_frame.pack(fill=tk.X, pady=5)
        ttk.Label(input_frame, text="Input Folder:").pack(side=tk.LEFT)
        self.input_label = ttk.Label(input_frame, text="No folder selected", 
                                    foreground='gray', relief='sunken', padding="5")
        self.input_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 5))
        ttk.Button(input_frame, text="Browse", command=self.browse_input).pack(side=tk.RIGHT)
        
        # Output path
        output_frame = ttk.Frame(path_frame)
        output_frame.pack(fill=tk.X, pady=5)
        ttk.Label(output_frame, text="Output File:").pack(side=tk.LEFT)
        self.output_label = ttk.Label(output_frame, text="No file selected", 
                                     foreground='gray', relief='sunken', padding="5")
        self.output_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 5))
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.RIGHT)
        
        # Main content frame
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Left side - All pages
        left_frame = ttk.LabelFrame(content_frame, text="All Available Pages", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # All pages listbox
        self.all_pages_listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED)
        all_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.all_pages_listbox.yview)
        self.all_pages_listbox.configure(yscrollcommand=all_scrollbar.set)
        
        self.all_pages_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        all_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Middle - Control buttons
        middle_frame = ttk.Frame(content_frame)
        middle_frame.pack(side=tk.LEFT, padx=10, pady=50)
        
        ttk.Button(middle_frame, text="Add →", command=self.add_pages).pack(pady=5)
        ttk.Button(middle_frame, text="← Remove", command=self.remove_pages).pack(pady=5)
        ttk.Button(middle_frame, text="↑ Up", command=self.move_up).pack(pady=5)
        ttk.Button(middle_frame, text="↓ Down", command=self.move_down).pack(pady=5)
        
        # Right side - Selected pages (with drag & drop)
        right_frame = ttk.LabelFrame(content_frame, text="Selected Pages (Drag to Reorder)", padding="10")
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Selected pages listbox with drag & drop
        self.selected_pages_listbox = DragDropListbox(right_frame, selectmode=tk.EXTENDED)
        selected_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.selected_pages_listbox.yview)
        self.selected_pages_listbox.configure(yscrollcommand=selected_scrollbar.set)
        
        self.selected_pages_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        selected_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bottom frame - Action button
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        self.create_btn = ttk.Button(bottom_frame, text="🐹 Create Merged PDF", 
                                    command=self.create_merged_pdf, style='Action.TButton')
        self.create_btn.pack(pady=10)
        
        # Status
        self.status_var = tk.StringVar(value="Select input folder to start")
        status_label = ttk.Label(bottom_frame, textvariable=self.status_var, foreground='blue')
        status_label.pack()
        
        # Configure button style
        style = ttk.Style()
        style.configure('Action.TButton', font=('Arial', 12, 'bold'))
        
        # Initially disable create button
        self.create_btn.configure(state='disabled')
    
    def browse_input(self):
        """Browse for input folder containing PDF files"""
        folder = filedialog.askdirectory(title="Select folder containing PDF files")
        if folder:
            self.input_path = folder
            self.input_label.config(text=folder, foreground='black')
            self.load_pdf_pages()
    
    def browse_output(self):
        """Browse for output PDF file location"""
        file = filedialog.asksaveasfilename(
            title="Save merged PDF as...",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.output_path = file
            self.output_label.config(text=file, foreground='black')
            self.update_create_button_state()
    
    def load_pdf_pages(self):
        """Load all pages from all PDF files in the input folder"""
        try:
            self.all_pages = []
            self.all_pages_listbox.delete(0, tk.END)
            
            if not self.input_path:
                return
            
            input_folder = Path(self.input_path)
            pdf_files = sorted([f for f in input_folder.glob("*.pdf") if f.is_file()])
            
            if not pdf_files:
                self.status_var.set("No PDF files found in selected folder")
                return
            
            total_pages = 0
            for pdf_file in pdf_files:
                try:
                    # Read entire file content into memory
                    with open(pdf_file, 'rb') as file:
                        pdf_content = file.read()
                    
                    # Create reader from memory content
                    reader = PdfReader(io.BytesIO(pdf_content))
                    
                    for page_num, page in enumerate(reader.pages, 1):
                        # Store the page data with file content
                        page_info = {
                            'pdf_file': pdf_file,
                            'pdf_content': pdf_content,  # Keep content in memory
                            'page_number': page_num,
                            'page_index': page_num - 1,  # 0-based index
                            'display_name': f"{pdf_file.stem} - Page {page_num}"
                        }
                        self.all_pages.append(page_info)
                        
                        # Add to listbox
                        self.all_pages_listbox.insert(tk.END, page_info['display_name'])
                        total_pages += 1
                        
                except Exception as e:
                    print(f"Error reading {pdf_file.name}: {e}")
            
            self.status_var.set(f"Loaded {total_pages} pages from {len(pdf_files)} PDF files")
            self.update_create_button_state()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading PDF files: {str(e)}")
    
    def add_pages(self):
        """Add selected pages to the merge list"""
        selection = self.all_pages_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select pages to add")
            return
        
        for index in selection:
            page_info = self.all_pages[index]
            # Check if already selected
            if not any(p['pdf_file'] == page_info['pdf_file'] and 
                      p['page_number'] == page_info['page_number'] 
                      for p in self.selected_pages):
                self.selected_pages.append(page_info)
                self.selected_pages_listbox.insert(tk.END, page_info['display_name'])
        
        self.update_create_button_state()
    
    def remove_pages(self):
        """Remove selected pages from the merge list"""
        selection = self.selected_pages_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select pages to remove")
            return
        
        # Remove in reverse order to maintain indices
        for index in reversed(selection):
            del self.selected_pages[index]
            self.selected_pages_listbox.delete(index)
        
        self.update_create_button_state()
    
    def move_up(self):
        """Move selected page up in the merge list"""
        selection = self.selected_pages_listbox.curselection()
        if not selection or selection[0] == 0:
            return
        
        index = selection[0]
        # Swap in data list
        self.selected_pages[index], self.selected_pages[index-1] = \
            self.selected_pages[index-1], self.selected_pages[index]
        
        # Update listbox
        item = self.selected_pages_listbox.get(index)
        self.selected_pages_listbox.delete(index)
        self.selected_pages_listbox.insert(index-1, item)
        self.selected_pages_listbox.selection_set(index-1)
    
    def move_down(self):
        """Move selected page down in the merge list"""
        selection = self.selected_pages_listbox.curselection()
        if not selection or selection[0] == self.selected_pages_listbox.size()-1:
            return
        
        index = selection[0]
        # Swap in data list
        self.selected_pages[index], self.selected_pages[index+1] = \
            self.selected_pages[index+1], self.selected_pages[index]
        
        # Update listbox
        item = self.selected_pages_listbox.get(index)
        self.selected_pages_listbox.delete(index)
        self.selected_pages_listbox.insert(index+1, item)
        self.selected_pages_listbox.selection_set(index+1)
    
    def update_create_button_state(self):
        """Enable/disable create button based on selections"""
        if self.output_path and self.selected_pages:
            self.create_btn.configure(state='normal')
        else:
            self.create_btn.configure(state='disabled')
    
    def create_merged_pdf(self):
        """Create the merged PDF with selected pages in order"""
        if not self.selected_pages or not self.output_path:
            messagebox.showwarning("Warning", "Please select pages and output location")
            return
        
        def create_worker():
            try:
                self.create_btn.configure(state='disabled')
                self.status_var.set("Creating merged PDF...")
                
                writer = PdfWriter()
                
                # Add pages in the order they appear in the selected list
                for page_info in self.selected_pages:
                    # Create a new reader from the stored content for each page
                    reader = PdfReader(io.BytesIO(page_info['pdf_content']))
                    page = reader.pages[page_info['page_index']]
                    writer.add_page(page)
                
                # Write to output file
                with open(self.output_path, 'wb') as output_file:
                    writer.write(output_file)
                
                self.status_var.set(f"Success! Created PDF with {len(self.selected_pages)} pages")
                messagebox.showinfo("Success", 
                    f"🐹 Merged PDF created successfully!\n\n"
                    f"File: {self.output_path}\n"
                    f"Pages: {len(self.selected_pages)}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error creating merged PDF: {str(e)}")
                self.status_var.set("Error occurred during PDF creation")
            finally:
                self.create_btn.configure(state='normal')
        
        threading.Thread(target=create_worker, daemon=True).start()

def main():
    # Check if PyPDF2 is installed
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Missing Dependency", 
            "PyPDF2 is required but not installed.\n\n"
            "Please install it using:\npip install PyPDF2")
        return
    
    root = tk.Tk()
    app = PDFPageSelector(root)
    root.mainloop()

if __name__ == "__main__":
    main()