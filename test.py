import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import pytesseract
from pytesseract import Output
import re
from PIL import Image

# Set Tesseract OCR path (Modify if needed)
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Rene\Desktop\BUCKAP\projects ojt\New folder\darang tess\tesseract.exe'

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wamy's RLM Scanner  ദ്ദി(˵ •̀ ᴗ - ˵ ) ✧")
        self.root.geometry("1000x600")  # Set window size

        # Main Frames
        self.top_frame = tk.Frame(root, padx=10, pady=10)
        self.top_frame.pack(fill=tk.X)

        self.left_frame = tk.Frame(root, padx=10, pady=10)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.right_frame = tk.Frame(root, padx=10, pady=10)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Variables
        self.folder_path = tk.StringVar()
        self.remarks_var = tk.StringVar(value="NEW")

        # UI components
        self.create_widgets()

        # PDF files to process
        self.pdf_files = []
        self.current_file_index = 0

    def create_widgets(self):
        """Create and arrange all widgets on the GUI"""
        # Folder Selection Section
        tk.Label(self.top_frame, text="Select Folder:", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        tk.Entry(self.top_frame, textvariable=self.folder_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(self.top_frame, text="Browse", command=self.select_folder).pack(side=tk.LEFT, padx=5)

        # Canvas for image display
        self.canvas = tk.Canvas(self.right_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Input Fields and Remarks Section
        self.create_input_fields()
        self.create_remarks_section()

        # Buttons
        tk.Button(self.left_frame, text="Process PDFs", command=self.process_pdfs, font=("Arial", 10, "bold")).pack(fill=tk.X, pady=10)

    def create_input_fields(self):
        """Create input fields for name, O.R. number, valid until, and control code"""
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
        """Create radio buttons for remarks options"""
        tk.Label(self.left_frame, text="Remarks:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        remarks_options = ["NEW", "REN", "RENDUP", "RENMOD", "MOD"]
        for remark in remarks_options:
            tk.Radiobutton(self.left_frame, text=remark, variable=self.remarks_var, value=remark).pack(anchor="w")

    def select_folder(self):
        """Open a dialog to select a folder"""
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
            self.current_file_index = 0
            self.process_next_pdf()

    def find_valid_until_by_location(self, image):
        """Extract the 'Valid Until' date by cropping the relevant area before OCR processing."""
        # Convert image to grayscale for better OCR accuracy
        gray = image.convert("L")

        # Define the starting point of "Valid" at (x=2175, y=1257)
        crop_area = (2175, 1257, 2600, 1500)  # (x1, y1, x2, y2) Adjust these values for your image

        # Crop the image to the area below the word "Valid"
        cropped_image = gray.crop(crop_area)

        # Optional: Preprocess the cropped image to improve OCR accuracy
        # Apply thresholding to make the text stand out more
        threshold_image = cropped_image.point(lambda p: p > 200 and 255)  # Simple threshold

        # Display the thresholded image for verification (optional)
        """threshold_image.show()"""

        # Run OCR on the cropped image (now thresholded for better contrast)
        valid_until_text = pytesseract.image_to_string(threshold_image, config="--psm 6").strip()

        print(f"Extracted OCR Text: {valid_until_text}")
        
        # Remove any unwanted text (like "Valid Until") that might appear in the OCR output
        valid_until_text = re.sub(r"Valid\s*Until[:\-]?\s*", "", valid_until_text).strip()

        # Validate and clean extracted text to match date formats (e.g., 30-Jul-20 or 30-Jul-2020)
        match = re.match(r"\d{2}-[A-Za-z]{3}-(\d{2}|\d{4})", valid_until_text)
        
        if match:
            # Extract year part
            valid_until_year = match.group(1)
            
            # If the year is in 4-digit format (e.g., 2020), extract only the last two digits
            if len(valid_until_year) == 4:
                valid_until_year = valid_until_year[-2:]  # Keep only last two digits (e.g., 2020 -> 20)

            print(f"Extracted Valid Until Year: {valid_until_year}")
            return valid_until_year
        else:
            print("OCR failed to detect the correct date format. Trying alternative processing.")
            return "Not Found"
        
    def find_control_code_by_location(self, text):
        """Extract the Control Code starting with 'RLMP' or 'NTCNCR'."""
        # Updated regex for Control Code: Match 'RLMP' or 'NTCNCR' followed by optional hyphens and alphanumeric characters
        control_match = re.search(r"\b(RLMP|NTCNCR)(?:-[A-Za-z0-9]+)+\b", text)

        if control_match:
            return control_match.group(0)  # Return the full match (e.g., 'RLMP-CC-02523-17')
        else:
            return ""  # Return 'Unknown' if no control code is found


    def process_next_pdf(self):
        """Process the next PDF file in the folder"""
        if self.current_file_index < len(self.pdf_files):
            pdf_path = os.path.join(self.folder_path.get(), self.pdf_files[self.current_file_index])
            images = convert_from_path(pdf_path, dpi=500)

            # Get the first image from the PDF (assuming it is the correct page)
            image = images[0]

            # Extract text from the PDF and other necessary fields (Name, O.R. Number, Control Code)
            extracted_text = self.extract_text_from_pdf(pdf_path)
            name, or_number, _, control_code = self.extract_info(extracted_text)

            # Extract the Valid Until date based on location
            valid_until = self.find_valid_until_by_location(image)

            # Extract Control Code by location
            control_code = self.find_control_code_by_location(extracted_text)


            # Populate the input fields with extracted values
            self.name_entry.delete(0, tk.END)
            self.name_entry.insert(0, name)

            self.or_entry.delete(0, tk.END)
            self.or_entry.insert(0, or_number)

            self.valid_until_entry.delete(0, tk.END)
            self.valid_until_entry.insert(0, valid_until)

            self.control_entry.delete(0, tk.END)
            self.control_entry.insert(0, control_code)

            # Display the image for preview
            self.display_image(image)

        else:
            messagebox.showinfo("Done", "All PDFs have been processed.")

    def extract_text_from_pdf(self, pdf_path):
        """Convert PDF to images and extract text using OCR"""
        images = convert_from_path(pdf_path, dpi=300)
        text = ""
        for img in images:
            text += pytesseract.image_to_string(img, config="--psm 6") + "\n"
        return text.strip()

    def extract_info(self, text):
        """Extract the necessary information from OCR text"""
        # Clean up the text globally (removing '___' and replacing underscores with spaces)
        cleaned_text = text.replace("___", "").replace("_", " ").strip()

        # Perform regex searches on the cleaned text
        name_match = re.search(r"Name[:\-]?\s*(.*)", cleaned_text, re.IGNORECASE)
        
        # Updated regex for O.R. Number to match 'Paid under O.R. No.: 123456'
        or_match = re.search(r"O\.R\. No\.:?\s*(\d+)", cleaned_text, re.IGNORECASE)

        # Regex to capture Valid Until in format 'dd-MMM-yy' (e.g., 07-Sep-23)
        valid_match = re.search(r"Valid\s*Until[:\-]?\s*(\d{2}-[A-Za-z]{3}-\d{2})", cleaned_text, re.IGNORECASE)
        
        # If the date is not found immediately, look for a date after the Valid Until field
        if not valid_match:
            valid_match = re.search(r"(?:Valid\s*Until[:\-]?\s*)?(\d{2}-[A-Za-z]{3}-\d{2})", cleaned_text, re.IGNORECASE)

        print(f"Valid Until Match: {valid_match}")

        # Extract just the year from the Valid Until field (e.g., "23" from "07-Sep-23")
        valid_year = valid_match.group(1).split('-')[-1] if valid_match else "00"

        # Find Control Code using the new function
        control_code = self.find_control_code_by_location(cleaned_text)

        # Extracted information with fallback values
        name = name_match.group(1) if name_match else "Unknown"
        or_number = or_match.group(1) if or_match else "000000"
        valid_until = valid_year  # Only use the year

        # Format name to Lastname, Firstname M.I. (if applicable)
        name_parts = name.split()
        if len(name_parts) >= 2:
            last_name = name_parts[-1]
            first_name = name_parts[0]
            middle_initial = name_parts[1][0] + '.' if len(name_parts) > 2 else ''
            name = f"{last_name}, {first_name} {middle_initial}".strip()

        return name, or_number, valid_until,control_code

    

    def display_image(self, img):
        """Display the image on the canvas with a larger size"""
        # Resize the image for a larger preview
        img.thumbnail((800, 800))  # Increase thumbnail size for better preview
        img = ImageTk.PhotoImage(img)
        
        # Clear any previous image on the canvas
        self.canvas.delete("all")
        
        # Display the image at the center of the canvas
        self.canvas.create_image(self.canvas.winfo_width() // 2, self.canvas.winfo_height() // 2, image=img)
        self.canvas.image = img

    def process_pdfs(self):
        """Process the selected PDF files and rename them based on the extracted information"""
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

        # Extract the year for naming
        valid_year = valid_until.split('/')[-1] if '/' in valid_until else valid_until
        new_filename = f"{or_number}_RML_{remarks}_{name}_20{valid_year}_{control_code}.pdf"
        new_filepath = os.path.join(folder, new_filename)

        # Rename the PDF file if it doesn't already exist
        if not os.path.exists(new_filepath):
            os.rename(pdf_path, new_filepath)
            messagebox.showinfo("Success", f"Renamed: {new_filename}")
        else:
            messagebox.showerror("Error", f"File with name {new_filename} already exists.")

        # Move to the next PDF
        self.current_file_index += 1
        self.process_next_pdf()


if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
