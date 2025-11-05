OCR-to-Excel Tool
This project automates the extraction of printed and handwritten text from scanned paper sheets and saves the recognized data into a structured Excel file.

It was developed and tested during an internship as a prototype OCR pipeline.

📌 What It Does
Processes scanned warehouse forms in .jpg / .png format.

Extracts tabular data from the roll list pages, including:

Roll number (e.g., B12947)

Format (mm)

Weight (kg)

Grammage (g/m²)

Comment (if handwritten)

Saves the structured data into output/results.xlsx.
Produces debug images and logs in /output/.

⚠️ Title pages (the first page of each delivery sheet) are intentionally not processed — the program only reads the tables with roll data.

🛠 Technologies Used
Python 3.11
EasyOCR — multilingual text detection
OpenCV — image preprocessing and table segmentation
Pandas + OpenPyXL — Excel export
Pillow — image handling
Tesseract OCR — for mixed Cyrillic/Latin text

⚙️ Installation

Install Tesseract OCR
Download: Tesseract at UB Mannheim
Install to the default path: C:\Program Files\Tesseract-OCR
During installation, include English and Russian language packs.
Add to PATH: C:\Program Files\Tesseract-OCR
Verify installation:
tesseract --version

Clone Repository & Create Virtual Environment
git clone <your-repo-url>
cd ocr_to_excel
python -m venv venv
venv\Scripts\activate # on Windows

Install Dependencies
pip install -r requirements.txt

Run the Tool
python test_table.py

The processed Excel file will appear in the output/ folder as results.xlsx.

📁 Project Structure
ocr_to_excel/
├── input/ # Folder for input images (JPG/PNG)
├── output/ # Debug images + Excel results
├── venv/ # Virtual environment
├── ocr_engine.py # Core OCR logic
├── test_table.py # Main script for testing OCR
├── requirements.txt # Dependencies
├── README.md # Documentation
└── .gitignore

🧪 Experimental Results (Handwritten + Printed Tables)
Field | Accuracy | Notes
Roll Number | ~30% | Some numbers correctly detected (e.g., B12952), others missed or misread
Format (mm) | ~45% | Detects printed numbers, but columns sometimes shift
Weight (kg) | ~25% | Often confused with grammage
Grammage (g/m²) | ~20% | Rarely recognized correctly
Comment (handwritten) | <15% | OCR fails on cursive handwriting
Overall Accuracy: ≈ 27%

📉 Limitations & Observations
Handwritten Cyrillic text is rarely recognized — both EasyOCR and Tesseract fail on cursive styles.
Table segmentation with OpenCV works on clean scans but struggles when grid lines are faint or broken.
Mixed Cyrillic and Latin text (e.g., B vs В) often leads to character confusion.
Windows.Media.Ocr (via winsdk) gave higher accuracy but is unreliable across systems and versions.

💡 Future Improvements
Use a hybrid approach:

Detect table layout via machine learning (e.g., Detectron or YOLO layout models)

Combine OCR from Google Vision or ABBYY Cloud for higher accuracy

Build a simple Tkinter GUI for file selection and batch processing

Add a post-processing correction module using regex validation and fuzzy matching for roll numbers

📋 Status
🚧 Prototype Stage (Internship Project)
Recognizes around 25–30% of text fields correctly on real scanned forms.
Suitable for further research and integration testing — not for production use yet.