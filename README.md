# 📄 OCR-to-Excel Tool

This project automates the extraction of handwritten and printed text from scanned delivery sheets and saves the structured data into an Excel file.

---

## 📌 What It Does

- Processes scanned sheets in **JPG/PNG** format  

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

- Saves everything to a structured **.xlsx file**

---

## 🛠 Technologies Used

- **Python 3.11**  

- **EasyOCR**  

- **OpenCV**  

- **Pillow**  

- **Pandas**  

- **OpenPyXL**  

- **Tkinter** *(planned for GUI)*

---

## 📁 Project Structure

ocr_to_excel/

├── input/ # Folder for input images (JPG/PNG)

├── output/ # Output Excel files

├── ocr_engine.py # OCR logic

├── excel_writer.py # Excel file creator

├── main.py # Entry point (optional older version)

├── test_table.py # Active script for processing

├── requirements.txt # Dependencies

├── README.md

└── .gitignore

yaml

Copy code

---

## 🚀 How to Run

1\. Convert scanned PDF to **JPG or PNG** files  

2\. Place image files into the **input/** folder  

3\. Make sure your virtual environment is active  

4\. Run:

```bash

python test_table.py

✅ The processed Excel file will be saved in the output/ folder.

🧾 Notes

Works best with clearly written numbers and printed tables

Designed for internal use during my internship

GUI version (Tkinter) planned for future use

📊 Recognition Quality (Current Results)

Based on testing with real delivery sheets:

Metric  Approximate Accuracy

Roll Number recognition  ~40--50%

Numeric fields (format/weight/grammage)  ~50--60%

Comment field  ~30%

Overall structured accuracy  ~45%

Summary:

The prototype successfully segments tables and extracts partial data,

but text accuracy remains limited. Recommended next step: test Google Cloud Vision or ABBYY OCR SDK.

🔖 Status

🟢 Stable prototype --- basic OCR-to-Excel pipeline works

🟡 Accuracy requires further improvement

🔵 Next version planned with cloud OCR integration