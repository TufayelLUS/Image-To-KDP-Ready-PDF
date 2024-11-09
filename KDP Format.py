import os
import configparser
from tkinter import filedialog, StringVar, DoubleVar, messagebox
import customtkinter as ctk
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from PIL import Image
from docx2pdf import convert
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
        'gutter': '0'
    }
    with open(config_file, 'w') as configfile:
        config.write(configfile)

# Initialize CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class ImageDocxApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.title("KDP Auto Formatting Tool - For Sangi")
        self.iconbitmap("icon.ico")
        self.geometry("800x600")
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
            self, text="Browse", command=self.select_folder)
        self.select_folder_btn.grid(row=0, column=2, padx=10, pady=(20, 5))

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
            self, text="Create Document", command=self.document_creator_thread, width=200)
        self.create_doc_btn.grid(row=13, column=1, pady=20)

    def update_keep_docx_visibility(self, *args):
        if self.file_type.get() == "PDF":
            self.keep_docx_checkbox.grid(
                row=3, column=0, padx=10, pady=5, sticky="w")
        else:
            self.keep_docx_checkbox.grid_forget()  # Hides the checkbox

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder.set(folder_selected)

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
        section.left_margin = Inches(self.left_margin.get()) # left is inside margin
        section.right_margin = Inches(self.right_margin.get()) # right is outside margin
        section.gutter = Inches(self.gutter.get())

        page_width, page_height = section.page_width, section.page_height
        image_files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith(
            ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))])
        # sort images based on numbering
        image_files = sorted(image_files, key=lambda x: int(x.split(' ')[-1].split('.')[0]))
        # print(image_files)
        

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
        if not image_files:
            messagebox.showwarning(
                "No Images Found", "No image files found in the selected folder.")
            self.create_doc_btn.configure(
                state="normal", text="Create Document")
            return

        # Flag to track if any valid images were added
        any_images_added = False

        # Create PDF using reportlab
        pdf_file_path = os.path.join("OUTPUT", f"{output_name}.pdf")
        if os.path.exists(pdf_file_path):
            os.remove(pdf_file_path)
        pdf_canvas = canvas.Canvas(
            pdf_file_path, pagesize=(page_width.pt, page_height.pt))

        target_ppi = 330

        for idx, filename in enumerate(image_files):
            file_path = os.path.join(folder_path, filename)
            image = Image.open(file_path)

            # Ensure image has content before proceeding
            if image.size[0] == 0 or image.size[1] == 0:
                continue  # Skip blank images

            resized_image = image.resize((int(page_width.pt * target_ppi / 72), int(
                page_height.pt * target_ppi / 72)), Image.LANCZOS)  # Resize for 300 PPI

            temp_image_path = os.path.join(folder_path, f'temp_{filename}')
            resized_image.save(temp_image_path, format='PNG')

            # Add image to the document
            doc.add_picture(temp_image_path, width=available_width,
                            height=available_height)
            pdf_canvas.drawImage(temp_image_path, 0, 0, width=int(
                page_width.pt), height=int(page_height.pt))
            any_images_added = True  # Mark that an image has been added
            os.remove(temp_image_path)

        # Save the document only if there are images added
        if any_images_added:
            if os.path.exists(os.path.join("OUTPUT", output_name + '.docx')):
                os.remove(os.path.join("OUTPUT", output_name + '.docx'))
            doc.save(os.path.join("OUTPUT", output_name + '.docx'))
            pdf_canvas.save()

        if file_type == "PDF":
            convert(os.path.join("OUTPUT", output_name + '.docx')
                    ), os.path.join("OUTPUT", output_name + '.pdf')
            if not self.keep_docx.get():
                os.remove(os.path.join("OUTPUT", output_name + '.docx'))

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
