# OCR-to-Excel Tool

This project automates the extraction of handwritten and printed text from scanned delivery sheets and saves the structured data into an Excel file.

## 📌 What it does

- Processes scanned sheets in JPG/PNG format
- Extracts key data from the **first (title) page**:
  - Document number
  - Transport info
  - Date (if machine-printed)
- Extracts tabular data from the **following pages**:
  - Roll number
  - Format (mm)
  - Weight (kg)
  - Grammage (g/m²)
  - Comments
- Saves everything to a structured `.xlsx` file

## 🛠 Technologies Used

- Python 3.11
- EasyOCR
- OpenCV
- Pillow
- Pandas
- OpenPyXL
- Tkinter (planned for GUI)

## 📁 Project Structure

ocr_to_excel/
├── input/ # Folder for input images (JPG/PNG)
├── output/ # Output Excel files
├── ocr_to_excel/
│ ├── ocr_engine.py # OCR logic
│ ├── excel_writer.py # Excel file creator
│ └── init.py
├── main.py # Entry point
├── requirements.txt
├── README.md
└── .gitignore


## 🚀 How to Run

1. Convert scanned PDF to JPG or PNG files
2. Place image files into the `input/` folder
3. Make sure virtual environment is active
4. Run:

```bash
python main.py

## The output Excel will be saved in the output/ folder

🧾 Notes

Currently works best with clearly written numbers and printed tables

Designed for internal use during my internship

GUI version (Tkinter) planned for ease of use