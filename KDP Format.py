import os
import configparser
from tkinter import filedialog, StringVar, DoubleVar, messagebox, Listbox
import customtkinter as ctk
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from PIL import Image
from reportlab.pdfgen import canvas
from threading import Thread

# Initialize config file
config_file = 'settings.ini'
config = configparser.ConfigParser()

# Load configuration if it exists
if os.path.exists(config_file):
    config.read(config_file)
else:
    config['Settings'] = {
        'input_folder': '',
        'output_filename': 'Output.docx',
        'file_type': 'DOCX',
        'page_width': '8.27',
        'page_height': '11.69',
        'top_margin': '0',
        'bottom_margin': '0',
        'left_margin': '0',
        'right_margin': '0',
        'gutter': '0',
        'image_sequence': '',
        'deleted_items': ''
    }
    with open(config_file, 'w') as configfile:
        config.write(configfile)

# Initialize CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class DraggableListbox(Listbox):
    def __init__(self, master, *args, **kwargs):
        kwargs['selectmode'] = "extended"
        super().__init__(*args, **kwargs)
        self.master = master
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.start_index = None
        self.selected_before_drag = []

    def on_click(self, event):
        clicked_index = self.nearest(event.y)
        self.start_index = clicked_index
        
        # Store the current selection before any changes
        self.selected_before_drag = list(self.curselection())
        
        # If the clicked item is not in the current selection and no modifiers are pressed,
        # clear selection and select only the clicked item
        if clicked_index not in self.selected_before_drag and not (event.state & 0x0004 or event.state & 0x0001):
            self.selection_clear(0, 'end')
            self.selection_set(clicked_index)
            self.selected_before_drag = [clicked_index]

    def on_drag(self, event):
        if self.start_index is None:
            return
            
        current_index = self.nearest(event.y)
        if current_index == self.start_index:
            return
            
        # Use the selection that was active when we started dragging
        selected_indices = self.selected_before_drag
        if not selected_indices:
            return
            
        # Don't allow drag to positions within the selection
        if min(selected_indices) <= current_index <= max(selected_indices):
            return

        # Get all items to move
        items_to_move = [self.get(i) for i in selected_indices]
        
        # Remove items from their current positions
        for index in sorted(selected_indices, reverse=True):
            self.delete(index)
        
        # Calculate new insertion position
        num_items = len(items_to_move)
        
        if current_index > max(selected_indices):
            # Moving down - adjust for items we removed
            insert_index = current_index - num_items + 1
        else:
            # Moving up
            insert_index = current_index
        
        # Insert items at new position
        new_indices = []
        for i, item in enumerate(items_to_move):
            pos = insert_index + i
            self.insert(pos, item)
            new_indices.append(pos)
        
        # Update selection to new positions
        self.selection_clear(0, 'end')
        for idx in new_indices:
            self.selection_set(idx)
        
        # Update the image files in the main app
        self.master.update_image_files()
        
        # Update start index for continuous dragging
        self.start_index = current_index
        # Update the stored selection for the next drag operation
        self.selected_before_drag = new_indices



class ImageDocxApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.title("KDP Auto Formatting Tool - For Sangi")
        self.iconbitmap("icon.ico")
        self.geometry("1000x700")
        self.resizable(False, False)  # Make window non-resizable

        # Grid layout settings
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure((0, 1), weight=1)

        # Load configuration values
        self.input_folder = StringVar(value=config['Settings']['input_folder'])
        self.output_filename = StringVar(
            value=config['Settings']['output_filename'])
        self.file_type = StringVar(value=config['Settings']['file_type'])
        self.page_width = DoubleVar(
            value=float(config['Settings']['page_width']))
        self.page_height = DoubleVar(
            value=float(config['Settings']['page_height']))
        self.top_margin = DoubleVar(
            value=float(config['Settings']['top_margin']))
        self.bottom_margin = DoubleVar(
            value=float(config['Settings']['bottom_margin']))
        self.left_margin = DoubleVar(
            value=float(config['Settings']['left_margin']))
        self.right_margin = DoubleVar(
            value=float(config['Settings']['right_margin']))
        self.gutter = DoubleVar(value=float(
            config['Settings']['gutter']))  # Default gutter value

        # Folder selection
        ctk.CTkLabel(self, text="Select Image Folder:").grid(
            row=0, column=0, padx=10, pady=(20, 5), sticky="w")
        self.folder_entry = ctk.CTkEntry(
            self, textvariable=self.input_folder, width=300)
        self.folder_entry.grid(row=0, column=1, padx=10, pady=(20, 5))
        self.select_folder_btn = ctk.CTkButton(
            self, text="Browse", command=self.select_folder, fg_color="#71198c")
        self.select_folder_btn.grid(row=0, column=2, padx=10, pady=(20, 5))

        self.save_in_same_folder = ctk.BooleanVar(value=True)
        self.save_in_same_folder_checkbox = ctk.CTkCheckBox(
            self, text="Save In Same Folder", variable=self.save_in_same_folder)
        self.save_in_same_folder_checkbox.grid(
            row=0, column=4, padx=10, pady=(20, 5), sticky="w")

        # Output filename
        ctk.CTkLabel(self, text="Output File Name:").grid(
            row=2, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(
            self, textvariable=self.output_filename, width=300)
        self.output_entry.grid(row=2, column=1, padx=10, pady=5)

        # File type selection
        self.file_type_option = ctk.CTkOptionMenu(
            self, values=["DOCX", "PDF"], variable=self.file_type)
        self.file_type_option.grid(row=3, column=1, padx=10, pady=5)

        # Keep DOCX option
        self.keep_docx = ctk.BooleanVar(value=True)
        self.keep_docx_checkbox = ctk.CTkCheckBox(
            self, text="Keep DOCX file", variable=self.keep_docx)
        self.keep_docx_checkbox.grid(
            row=3, column=0, padx=10, pady=5, sticky="w")

        # Bind the method to the file type variable
        self.file_type.trace_add("write", self.update_keep_docx_visibility)

        # Page size label
        ctk.CTkLabel(self, text="Page Size (inches):").grid(
            row=4, column=0, columnspan=2, padx=10, pady=(20, 5))

        book_size_label = ctk.CTkLabel(
            self, text="KDP Size Template:")
        book_size_label.grid(row=4, column=2, padx=10, pady=2)

        # Common KDP sizes
        self.book_size = StringVar(value='8.5 x 11 in')
        self.common_sizes = [
            "5 x 8 in",
            "5.25 x 8 in",
            "5.5 x 8.5 in",
            "6 x 9 in",
            "5.06 x 7.81 in",
            "6.14 x 9.21 in",
            "6.69 x 9.61 in",
            "7 x 10 in",
            "7.44 x 9.69 in",
            "7.5 x 9.25 in",
            "8 x 10 in",
            "8.5 x 11 in",
            "8.27 x 11.69 in",
            "8.25 x 6 in",
            "8.25 x 8.25 in",
            "8.5 x 8.5 in"
        ]
        self.book_size_option = ctk.CTkOptionMenu(
            self, values=self.common_sizes, variable=self.book_size, command=self.update_size_on_change)
        self.book_size_option.grid(row=5, column=2, padx=10, pady=2)

        # Width and height with labels
        ctk.CTkLabel(self, text="Page Width:").grid(
            row=5, column=0, padx=10, pady=2, sticky="w")
        self.page_width_entry = ctk.CTkEntry(
            self, textvariable=self.page_width)
        self.page_width_entry.grid(row=5, column=1, padx=10, pady=2)

        ctk.CTkLabel(self, text="Page Height:").grid(
            row=6, column=0, padx=10, pady=2, sticky="w")
        self.page_height_entry = ctk.CTkEntry(
            self, textvariable=self.page_height)
        self.page_height_entry.grid(row=6, column=1, padx=10, pady=2)

        # Margins label
        ctk.CTkLabel(self, text="Margins (inches):").grid(
            row=7, column=0, columnspan=2, padx=10, pady=(20, 5))

        # Individual margin fields with labels
        ctk.CTkLabel(self, text="Top Margin:").grid(
            row=8, column=0, padx=10, pady=2, sticky="w")
        self.top_margin_entry = ctk.CTkEntry(
            self, textvariable=self.top_margin)
        self.top_margin_entry.grid(row=8, column=1, padx=10, pady=2)

        ctk.CTkLabel(self, text="Bottom Margin:").grid(
            row=9, column=0, padx=10, pady=2, sticky="w")
        self.bottom_margin_entry = ctk.CTkEntry(
            self, textvariable=self.bottom_margin)
        self.bottom_margin_entry.grid(row=9, column=1, padx=10, pady=2)

        ctk.CTkLabel(self, text="Left(inside) Margin:").grid(
            row=10, column=0, padx=10, pady=2, sticky="w")
        self.left_margin_entry = ctk.CTkEntry(
            self, textvariable=self.left_margin)
        self.left_margin_entry.grid(row=10, column=1, padx=10, pady=2)

        ctk.CTkLabel(self, text="Right(outside) Margin:").grid(
            row=11, column=0, padx=10, pady=2, sticky="w")
        self.right_margin_entry = ctk.CTkEntry(
            self, textvariable=self.right_margin)
        self.right_margin_entry.grid(row=11, column=1, padx=10, pady=2)

        ctk.CTkLabel(self, text="Gutter:").grid(
            row=12, column=0, padx=10, pady=2, sticky="w")
        self.gutter_entry = ctk.CTkEntry(
            self, textvariable=self.gutter)
        self.gutter_entry.grid(row=12, column=1, padx=10, pady=2)

        self.bleed_mode = StringVar(value='Bleed')

        self.bleed_option = ctk.CTkOptionMenu(
            self, values=["Bleed", "No Bleed"], variable=self.bleed_mode)
        self.bleed_option.grid(row=13, column=0, padx=10, pady=5)

        # Create document button
        self.create_doc_btn = ctk.CTkButton(
            self, text="Create Document", command=self.document_creator_thread, width=200, fg_color="#4a7a25")
        self.create_doc_btn.grid(row=13, column=1, pady=20)

        self.image_files = []

        self.image_listbox_label = ctk.CTkLabel(
            self, text="Page Serial(Click to select):")
        self.image_listbox_label.grid(row=6, column=2, pady=10)

        self.add_custom_image_btn = ctk.CTkButton(
            self, text="Add Custom Image", command=self.add_custom_image, fg_color="#1f6aa5")
        self.add_custom_image_btn.grid(row=7, column=2, padx=5, pady=5)

        # Image List Column (New)
        #  show vertical scrollbar
        self.image_listbox = DraggableListbox(
            self, selectmode="extended", height=15, exportselection=False, 
            font=("Arial", 12), background="#343638", foreground="#ffffff", 
            activestyle="dotbox", selectbackground="#1f6aa5", selectforeground="#ffffff")
        self.scrollbar = ctk.CTkScrollbar(
            self, orientation="vertical", command=self.image_listbox.yview)
        self.image_listbox.config(yscrollcommand=self.scrollbar.set)
        self.image_listbox.grid(row=8, column=2, rowspan=4, pady=10)
        self.scrollbar.grid(row=8, column=3, rowspan=4,
                            pady=10, sticky="ns")
        self.image_listbox.bind("<<ListboxSelect>>", self.update_preview)

        self.deleted_items = []  # To store deleted items

        self.delete_button = ctk.CTkButton(
            self, text="Delete Selected Page", command=self.delete_item, fg_color="#c73126")
        self.delete_button.grid(row=12, column=2, padx=5, pady=5)

        self.undo_button = ctk.CTkButton(
            self, text="Undo Delete", command=self.undo_delete, fg_color="#96156d")
        self.undo_button.grid(row=13, column=2, padx=1, pady=1)

        self.up_button = ctk.CTkButton(
            self, text="Move Page Up", command=self.move_up)
        self.up_button.grid(row=12, column=4, padx=5, pady=5)

        self.down_button = ctk.CTkButton(
            self, text="Move Page Down", command=self.move_down)
        self.down_button.grid(row=13, column=4, padx=5, pady=5)

        self.duplicate_button = ctk.CTkButton(
            self, text="Duplicate", command=self.duplicate_item, fg_color="#c76000")
        self.duplicate_button.grid(row=11, column=4, padx=5, pady=5)

        # Preview window (New)
        self.preview_label = ctk.CTkLabel(self, text="Image Preview:")
        self.preview_label.grid(row=5, column=4, padx=10, pady=10)

        self.preview_canvas = ctk.CTkLabel(
            self, text="", width=200)
        self.preview_canvas.grid(row=6, column=4, rowspan=5, padx=10, pady=10)
        folder_path = self.input_folder.get()
        if os.path.exists(folder_path):
            self.image_files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith(
                ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))])
            try:
                self.image_files = sorted(self.image_files, key=lambda x: float(
                    x.split(' ')[-1].split('.')[0].replace('-', '.')))
            except:
                pass

        # Load image sequence from config if it exists
        self.image_files = []
        if 'image_sequence' in config['Settings'] and config['Settings']['image_sequence']:
            saved_sequence = config['Settings']['image_sequence'].split('|')
            folder_path = self.input_folder.get()
            if os.path.exists(folder_path):
                # Filter only files that exist in the folder
                existing_files = [f for f in os.listdir(folder_path) if f.lower().endswith(
                    ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))]
                
                # Load deleted items FIRST
                self.deleted_items = []  # To store deleted items
                if 'deleted_items' in config['Settings'] and config['Settings']['deleted_items']:
                    deleted_data = config['Settings']['deleted_items'].split('|')
                    for item_data in deleted_data:
                        if item_data:
                            try:
                                filename, index_str = item_data.split(':', 1)
                                index = int(index_str)
                                self.deleted_items.append((filename, index))
                            except:
                                pass
                
                # Get list of deleted filenames for filtering
                deleted_filenames = [filename for filename, index in self.deleted_items]
                
                # Filter saved sequence to exclude deleted items
                self.image_files = [f for f in saved_sequence if f in existing_files and f not in deleted_filenames]
                
                # Add any new files that weren't in the saved sequence and aren't deleted
                new_files = [f for f in existing_files if f not in self.image_files and f not in deleted_filenames]
                self.image_files.extend(new_files)
        else:
            # Fallback to original loading method
            folder_path = self.input_folder.get()
            if os.path.exists(folder_path):
                self.image_files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith(
                    ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))])
                try:
                    self.image_files = sorted(self.image_files, key=lambda x: float(
                        x.split(' ')[-1].split('.')[0].replace('-', '.')))
                except:
                    pass

        # Load images from folder
        self.update_image_list()

    def add_custom_image(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Custom Images",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff")]
        )
        
        if file_paths:
            selection = self.image_listbox.curselection()
            insert_index = selection[0] if selection else 0
            
            for file_path in file_paths:
                filename = os.path.basename(file_path)
                # Copy file to input folder
                dest_path = os.path.join(self.input_folder.get(), filename)
                
                # Handle duplicate filenames
                counter = 1
                name, ext = os.path.splitext(filename)
                while os.path.exists(dest_path):
                    filename = f"{name}_{counter}{ext}"
                    dest_path = os.path.join(self.input_folder.get(), filename)
                    counter += 1
                
                # Copy the file
                import shutil
                shutil.copy2(file_path, dest_path)
                
                # Insert into listbox and image_files
                self.image_listbox.insert(insert_index, filename)
                self.image_files.insert(insert_index, filename)
                insert_index += 1
            
            self.update_image_list()
            self.save_config()  # Save immediately after adding images
            
            # Select the first added image
            if selection:
                self.image_listbox.select_clear(0, 'end')
                self.image_listbox.select_set(selection[0])
                self.update_preview()

    def duplicate_item(self):
        selected_indices = self.image_listbox.curselection()
        if selected_indices:
            # Sort in normal order to duplicate in sequence
            for selected_index in sorted(selected_indices):
                selected_item = self.image_listbox.get(selected_index)
                # Insert duplicate right after the original
                self.image_listbox.insert(selected_index + 1, selected_item)
            self.update_image_files()  # This now calls save_config automatically

    def update_image_files(self):
        self.image_files = list(self.image_listbox.get(0, self.image_listbox.size()))
        self.save_config()  # Save immediately after sequence change

    def delete_item(self):
        selected_indices = self.image_listbox.curselection()
        if selected_indices:
            # Sort in reverse order to delete from bottom to top
            for selected_index in sorted(selected_indices, reverse=True):
                selected_item = self.image_listbox.get(selected_index)
                self.deleted_items.append((selected_item, selected_index))
                self.image_listbox.delete(selected_index)
            self.update_image_files()  # This now calls save_config automatically
            self.save_deleted_items()  # Save deleted items immediately

    def undo_delete(self):
        if self.deleted_items:
            last_deleted_item, index = self.deleted_items.pop()
            self.image_listbox.insert(index, last_deleted_item)
            self.update_image_files()  # This now calls save_config automatically
            self.save_deleted_items()  # Save deleted items immediately

    def save_deleted_items(self):
        # Convert deleted_items to string format: "filename:index|filename:index"
        deleted_str = '|'.join([f"{filename}:{index}" for filename, index in self.deleted_items])
        config['Settings']['deleted_items'] = deleted_str
        self.save_config()

    def update_selection_view(self):
        # Get the index of the selected item
        selected_index = self.image_listbox.curselection()
        if selected_index:
            # Scroll to the selected item
            self.image_listbox.yview_moveto(
                selected_index[0] / float(self.image_listbox.size()))

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder.set(folder_selected)
            folder_path = self.input_folder.get()
            
            # Clear deleted items when changing folders
            self.deleted_items = []
            self.save_deleted_items()
            
            if os.path.exists(folder_path):
                # Try to load saved sequence first
                if 'image_sequence' in config['Settings'] and config['Settings']['image_sequence']:
                    saved_sequence = config['Settings']['image_sequence'].split('|')
                    existing_files = [f for f in os.listdir(folder_path) if f.lower().endswith(
                        ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))]
                    self.image_files = [f for f in saved_sequence if f in existing_files]
                    
                    # Add any new files that weren't in the saved sequence
                    new_files = [f for f in existing_files if f not in self.image_files]
                    self.image_files.extend(new_files)
                else:
                    # Fallback to original loading
                    self.image_files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith(
                        ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))])
                    try:
                        self.image_files = sorted(self.image_files, key=lambda x: float(
                            x.split(' ')[-1].split('.')[0].replace('-', '.')))
                    except:
                        pass
                self.save_config()  # Save the updated sequence
            self.update_image_list()

    def update_image_list(self):
        folder_path = self.input_folder.get()
        if not os.path.exists(folder_path):
            return

        self.image_listbox.delete(0, "end")
        for img in self.image_files:
            self.image_listbox.insert("end", img)

    def update_preview(self, event=None):
        selection = self.image_listbox.curselection()
        if selection:
            img_path = os.path.join(
                self.input_folder.get(), self.image_files[selection[0]])
            image = Image.open(img_path)
            image.thumbnail((200, 200))
            self.preview_img = ctk.CTkImage(
                light_image=image, dark_image=image, size=(200, 200))
            self.preview_canvas.configure(image=self.preview_img)

    def move_up(self):
        selected_indices = self.image_listbox.curselection()
        if selected_indices and min(selected_indices) > 0:
            # Move all selected items up
            for index in sorted(selected_indices):
                if index > 0:
                    self.image_files[index], self.image_files[index-1] = self.image_files[index-1], self.image_files[index]
            
            self.update_image_list()
            # Reselect the moved items (they all moved up by 1)
            self.image_listbox.selection_clear(0, 'end')
            for new_index in [i-1 for i in selected_indices]:
                self.image_listbox.selection_set(new_index)
            self.update_selection_view()
            self.save_config()  # Save immediately after move

    def move_down(self):
        selected_indices = self.image_listbox.curselection()
        if selected_indices and max(selected_indices) < len(self.image_files) - 1:
            # Move all selected items down (from bottom to top)
            for index in sorted(selected_indices, reverse=True):
                if index < len(self.image_files) - 1:
                    self.image_files[index], self.image_files[index+1] = self.image_files[index+1], self.image_files[index]
            
            self.update_image_list()
            # Reselect the moved items (they all moved down by 1)
            self.image_listbox.selection_clear(0, 'end')
            for new_index in [i+1 for i in selected_indices]:
                self.image_listbox.selection_set(new_index)
            self.update_selection_view()
            self.save_config()  # Save immediately after move

    def update_listbox_selection(self, new_index):
        self.update_image_list()
        self.image_listbox.select_set(new_index)
        self.image_listbox.activate(new_index)
        self.update_preview()

    def update_size_on_change(self, item):
        if self.bleed_mode.get() == "Bleed":
            width, height = item.split(" x ")
            height = height.replace(" in", "")
            self.page_width.set(float(width)+0.125)
            self.page_height.set(float(height)+0.25)
        else:
            width, height = item.split(" x ")
            height = height.replace(" in", "")
            self.page_width.set(float(width))
            self.page_height.set(float(height))

    def update_keep_docx_visibility(self, *args):
        if self.file_type.get() == "PDF":
            self.keep_docx_checkbox.grid(
                row=3, column=0, padx=10, pady=5, sticky="w")
        else:
            self.keep_docx_checkbox.grid_forget()  # Hides the checkbox

    def save_config(self):
        config['Settings']['input_folder'] = self.input_folder.get()
        config['Settings']['output_filename'] = self.output_filename.get()
        config['Settings']['file_type'] = self.file_type.get()
        config['Settings']['page_width'] = str(self.page_width.get())
        config['Settings']['page_height'] = str(self.page_height.get())
        config['Settings']['top_margin'] = str(self.top_margin.get())
        config['Settings']['bottom_margin'] = str(self.bottom_margin.get())
        config['Settings']['left_margin'] = str(self.left_margin.get())
        config['Settings']['right_margin'] = str(self.right_margin.get())
        config['Settings']['gutter'] = str(self.gutter.get())
        
        # Save the current image sequence
        config['Settings']['image_sequence'] = '|'.join(self.image_files)
        
        # Note: deleted_items is saved separately via save_deleted_items()

        with open(config_file, 'w') as configfile:
            config.write(configfile)

    def document_creator_thread(self):
        self.create_doc_btn.configure(state="disabled", text="Processing...")
        thread = Thread(target=self.create_document)
        thread.daemon = True
        thread.start()

    def create_document(self):
        self.save_config()  # Save config when creating the document
        folder_path = self.input_folder.get()
        if not os.path.exists(folder_path):
            messagebox.showerror(
                "Error", "Selected folder does not exist, please select input folder again.")
            self.create_doc_btn.configure(
                state="normal", text="Create Document")
            return
        output_name = self.output_filename.get()
        file_type = self.file_type.get()

        doc = Document()
        section = doc.sections[0]
        sectPr = section._sectPr

        # Create the mirrorMargins element
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.append(mirror_margins)
        section.page_width = Inches(self.page_width.get())
        section.page_height = Inches(self.page_height.get())
        section.top_margin = Inches(self.top_margin.get())
        section.bottom_margin = Inches(self.bottom_margin.get())
        section.left_margin = Inches(
            self.left_margin.get())  # left is inside margin
        section.right_margin = Inches(
            self.right_margin.get())  # right is outside margin
        section.gutter = Inches(self.gutter.get())

        page_width, page_height = section.page_width, section.page_height

        if self.bleed_mode.get() == "Bleed":
            print("Bleed mode")
            available_width = page_width
            available_height = page_height
            section.left_margin = Inches(0)
            section.right_margin = Inches(0)
            section.top_margin = Inches(0)
            section.bottom_margin = Inches(0)
            section.gutter = Inches(0)
        else:
            available_width = page_width - section.left_margin - \
                section.right_margin - section.gutter
            available_height = page_height - section.top_margin - \
                section.bottom_margin - section.gutter

        # Check if there are no images to process
        if not self.image_files:
            messagebox.showwarning(
                "No Images Found", "No image files found in the selected folder.")
            self.create_doc_btn.configure(
                state="normal", text="Create Document")
            return

        # Flag to track if any valid images were added
        any_images_added = False

        if self.save_in_same_folder.get():
            target_folder = folder_path
        else:
            target_folder = "OUTPUT"
        # Create PDF using reportlab
        pdf_file_path = os.path.join(target_folder, f"{output_name}.pdf")
        if os.path.exists(pdf_file_path):
            os.remove(pdf_file_path)
        pdf_canvas = canvas.Canvas(
            pdf_file_path, pagesize=(page_width.pt, page_height.pt))

        target_ppi = 330

        for idx, filename in enumerate(self.image_files):
            for _try in range(3):
                try:
                    file_path = os.path.join(folder_path, filename)
                    image = Image.open(file_path)

                    # Ensure image has content before proceeding
                    if image.size[0] == 0 or image.size[1] == 0:
                        continue  # Skip blank images

                    resized_image = image.resize((int(page_width.pt * target_ppi / 72), int(
                        page_height.pt * target_ppi / 72)), Image.LANCZOS)  # Resize for 300 PPI

                    temp_image_path = os.path.join(
                        folder_path, f'temp_{filename}')
                    resized_image.save(temp_image_path, format='PNG')

                    # Add image to the document
                    doc.add_picture(temp_image_path, width=available_width,
                                    height=available_height)
                    break
                except:
                    pass

            pdf_canvas.drawImage(temp_image_path, 0, 0, width=int(
                page_width.pt), height=int(page_height.pt))
            pdf_canvas.showPage()
            any_images_added = True  # Mark that an image has been added
            os.remove(temp_image_path)

        # Save the document only if there are images added
        if any_images_added:
            if os.path.exists(os.path.join(target_folder, output_name + '.docx')):
                os.remove(os.path.join(target_folder, output_name + '.docx'))
            doc.save(os.path.join(target_folder, output_name + '.docx'))
            

            if file_type == "PDF":
                pdf_canvas.save()
                if not self.keep_docx.get():
                    os.remove(os.path.join(target_folder, output_name + '.docx'))

        messagebox.showinfo("Document Created Successfully!", f"Document saved as {
                            output_name}.{file_type.lower()}")
        self.create_doc_btn.configure(
            state="normal", text="Create Document")


# Run the application
if __name__ == "__main__":
    if not os.path.exists("OUTPUT"):
        os.makedirs("OUTPUT")
    app = ImageDocxApp()
    app.mainloop()
