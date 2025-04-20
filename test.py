import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import pytesseract
from PIL import ImageDraw

# Set Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Rene\Desktop\BUCKAP\projects ojt\New folder\darang tess\tesseract.exe'

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wamy's RLM Scanner  ദ്ദി(˵ •̀ ᴗ - ˵ ) ✧")
        self.root.geometry("1000x600")

        self.top_frame = tk.Frame(root, padx=10, pady=10)
        self.top_frame.pack(fill=tk.X)

        self.left_frame = tk.Frame(root, padx=10, pady=10)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.right_frame = tk.Frame(root, padx=10, pady=10)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.folder_path = tk.StringVar()
        self.remarks_var = tk.StringVar(value="NEW")

        self.create_widgets()
        self.pdf_files = []
        self.current_file_index = 0
        self.images = []  # List to store all loaded images

    def create_widgets(self):
        tk.Label(self.top_frame, text="Select Folder:", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        tk.Entry(self.top_frame, textvariable=self.folder_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(self.top_frame, text="Browse", command=self.select_folder).pack(side=tk.LEFT, padx=5)

        self.canvas = tk.Canvas(self.right_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        nav_frame = tk.Frame(self.right_frame)
        nav_frame.pack(pady=10)

        self.prev_btn = tk.Button(nav_frame, text="Previous", command=self.show_previous_pdf)
        self.prev_btn.pack(side=tk.LEFT, padx=10)

        self.next_btn = tk.Button(nav_frame, text="Next", command=self.show_next_pdf)
        self.next_btn.pack(side=tk.LEFT, padx=10)


        self.create_input_fields()
        self.create_remarks_section()

        tk.Button(self.left_frame, text="Process PDFs", command=self.process_pdfs, font=("Arial", 10, "bold")).pack(fill=tk.X, pady=10)

    def create_input_fields(self):
        tk.Label(self.left_frame, text="Name:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        self.name_entry = tk.Entry(self.left_frame, width=30)
        self.name_entry.pack(fill=tk.X)

        tk.Label(self.left_frame, text="O.R. Number:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        self.or_entry = tk.Entry(self.left_frame, width=30)
        self.or_entry.pack(fill=tk.X)

        tk.Label(self.left_frame, text="Valid Until:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        self.valid_until_entry = tk.Entry(self.left_frame, width=30)
        self.valid_until_entry.pack(fill=tk.X)

        tk.Label(self.left_frame, text="Control Code:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        self.control_entry = tk.Entry(self.left_frame, width=30)
        self.control_entry.pack(fill=tk.X)

    def create_remarks_section(self):
        tk.Label(self.left_frame, text="Remarks:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        remarks_options = ["NEW", "REN", "RENDUP", "RENMOD", "MOD"]
        for remark in remarks_options:
            tk.Radiobutton(self.left_frame, text=remark, variable=self.remarks_var, value=remark).pack(anchor="w")

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
            self.current_file_index = 0
            self.images = []  # Reset images list
            self.preload_pdfs(folder)  # Preload all PDFs and their images

    def preload_pdfs(self, folder):
        for pdf_file in self.pdf_files:
            pdf_path = os.path.join(folder, pdf_file)
            images = convert_from_path(pdf_path, dpi=500)  # ✅ Optimized DPI (reduced to 500)
            self.images.append(images[0])  # Store only the first page of each PDF for now

        # Start processing the first PDF after preload
        self.process_next_pdf()

    def find_valid_until_by_location(self, image):
        gray = image.convert("L")

        crop_area = (5300, 3100, 6700, 3600)

        scaling_factor = 200 / 500

        crop_area_scaled = tuple(int(coordinate * scaling_factor) for coordinate in crop_area)

        cropped_image = gray.crop(crop_area_scaled)

        valid_until_text = pytesseract.image_to_string(cropped_image, config="--psm 6").strip()
        valid_until_text = re.sub(r"Valid\s*Until[:\-]?\s*", "", valid_until_text).strip()

        match = re.match(r"\d{2}-[A-Za-z]{3}-(\d{2}|\d{4})", valid_until_text)
        if match:
            valid_until_year = match.group(1)
            if len(valid_until_year) == 4:
                valid_until_year = valid_until_year[-2:]
            return valid_until_year

        return "Not Found"

    def find_control_code_by_location(self, text):
        control_match = re.search(r"\b(RLMP|NTCNCR)(?:-[A-Za-z0-9]+)+\b", text)
        return control_match.group(0) if control_match else ""

    def process_next_pdf(self, index=None):
        if index is not None:
            self.current_file_index = index

        if 0 <= self.current_file_index < len(self.pdf_files):
            image = self.images[self.current_file_index]

            pdf_path = os.path.join(self.folder_path.get(), self.pdf_files[self.current_file_index])
            extracted_text = self.extract_text_from_pdf(pdf_path)
            name, or_number, _, control_code = self.extract_info(extracted_text)
            valid_until = self.find_valid_until_by_location(image)
            control_code = self.find_control_code_by_location(extracted_text)

            self.name_entry.delete(0, tk.END)
            self.name_entry.insert(0, name)

            self.or_entry.delete(0, tk.END)
            self.or_entry.insert(0, or_number)

            self.valid_until_entry.delete(0, tk.END)
            self.valid_until_entry.insert(0, valid_until)

            self.control_entry.delete(0, tk.END)
            self.control_entry.insert(0, control_code)

            self.display_image(image)
        else:
            messagebox.showinfo("End", "No more PDFs in this direction.")
            
    def show_next_pdf(self):
        if self.current_file_index < len(self.pdf_files) - 1:
            if self.confirm_navigation("next"):
                self.current_file_index += 1
                self.process_next_pdf()

    def show_previous_pdf(self):
        if self.current_file_index > 0:
            if self.confirm_navigation("previous"):
                self.current_file_index -= 1
                self.process_next_pdf()

    def extract_text_from_pdf(self, pdf_path):
        images = convert_from_path(pdf_path, dpi=200)
        text = ""
        for img in images:
            text += pytesseract.image_to_string(img, config="--psm 6") + "\n"
        return text.strip()

    def extract_info(self, text):
        cleaned_text = text.replace("___", "").replace("_", " ").strip()
        name_match = re.search(r"Name[:\-]?\s*(.*)", cleaned_text, re.IGNORECASE)
        or_match = re.search(r"O\.R\. No\.:?\s*(\d+)", cleaned_text, re.IGNORECASE)
        valid_match = re.search(r"(?:Valid\s*Until[:\-]?\s*)?(\d{2}-[A-Za-z]{3}-\d{2})", cleaned_text, re.IGNORECASE)
        valid_year = valid_match.group(1).split('-')[-1] if valid_match else "00"
        control_code = self.find_control_code_by_location(cleaned_text)

        name = name_match.group(1) if name_match else "Unknown"
        or_number = or_match.group(1) if or_match else "000000"

        name_parts = name.split()
        if len(name_parts) >= 2:
            last_name = name_parts[-1]
            first_name = name_parts[0]
            middle_initial = name_parts[1][0] + '.' if len(name_parts) > 2 else ''
            name = f"{last_name}, {first_name} {middle_initial}".strip()

        return name, or_number, valid_year, control_code

    def display_image(self, img):
        img.thumbnail((800, 800))
        img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(self.canvas.winfo_width() // 2, self.canvas.winfo_height() // 2, image=img)
        self.canvas.image = img

    def process_pdfs(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Error", "Please select a folder first.")
            return

        pdf_path = os.path.join(folder, self.pdf_files[self.current_file_index])
        name = self.name_entry.get()
        or_number = self.or_entry.get()
        valid_until = self.valid_until_entry.get()
        control_code = self.control_entry.get()
        remarks = self.remarks_var.get()

        valid_year = valid_until.split('/')[-1] if '/' in valid_until else valid_until
        new_filename = f"{or_number}_RML_{remarks}_{name}_20{valid_year}_{control_code}.pdf"
        new_filepath = os.path.join(folder, new_filename)

        confirm = messagebox.askyesno("Confirm Rename", f"Are you sure you want to rename:\n\n{self.pdf_files[self.current_file_index]}\n\nto\n\n{new_filename}?")
        if not confirm:
            return

        try:
            if not os.path.exists(new_filepath):
                os.rename(pdf_path, new_filepath)
                messagebox.showinfo("Success", f"Renamed: {new_filename}")
            else:
                messagebox.showerror("Error", f"File with name {new_filename} already exists.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {str(e)}")

        self.current_file_index += 1
        self.process_next_pdf()


if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
