import os
import pandas as pd
from ocr_engine import extract_table_rows

INPUT_DIR = "input"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def main():
    files = [f for f in os.listdir(INPUT_DIR)
             if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    if not files:
        print("⚠️ There are no images in the input/ folder.")
        return

    all_rows = []
    for fname in sorted(files):
        image_path = os.path.join(INPUT_DIR, fname)
        print(f"\n📄 Processing: {image_path}")
        rows = extract_table_rows(image_path, debug=True, verbose=True)
        for r in rows:
            r["Source File"] = fname
            all_rows.append(r)

    if all_rows:
        df = pd.DataFrame(all_rows)
        out_xlsx = os.path.join(OUTPUT_DIR, "results.xlsx")
        df.to_excel(out_xlsx, index=False)
        print(f"\n✅ Results saved to {out_xlsx}")
        print("🧩 Debug files: output/debug_log.csv, output/debug_table1.jpg")

if __name__ == "__main__":
    main()
