import os
import re
import csv
import cv2
import easyocr
import pytesseract

# ---------- Tesseract path (Windows) ----------
# Edit the path if Tesseract is installed elsewhere on your PC.
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# ---------- Readers ----------
# Roll Number / text: ru+en helps with Cyrillic/Latin confusion
easy_reader_text = easyocr.Reader(['ru', 'en'], gpu=False)
# Digits reader for numeric fields
easy_reader_digits = easyocr.Reader(['en'], gpu=False)

# ---------- Columns ----------
COLUMNS = ["Roll Number", "Format (mm)", "Weight (kg)", "Grammage (g/m2)", "Comment"]

# ---------- OCR helpers ----------
def easyocr_text(img_bgr, digits=False):
    reader = easy_reader_digits if digits else easy_reader_text
    res = reader.readtext(img_bgr)
    return " ".join([t[1] for t in res]).strip()

def tesseract_text(img_bgr):
    config = "--oem 3 --psm 6"
    return pytesseract.image_to_string(img_bgr, lang="rus+eng", config=config).strip()

# ---------- Cleaning helpers ----------
def clean_roll_number(raw: str) -> str:
    """
    Normalize and validate a roll number like B12345 (one letter + 5 digits).
    Fix frequent OCR substitutions (e.g., Cyrillic 'В'->'B', '8'->'B' before 5 digits).
    """
    if not raw:
        return ""
    s = raw.strip()
    s = (s.replace("В", "B").replace("в", "B")
           .replace("Ь", "B").replace("ь", "B"))
    # Replace leading '8' with 'B' if followed by 5 digits
    s = re.sub(r"8(?=\s*\d{5})", "B", s)
    # Remove separators/spacers
    s = re.sub(r"[\s|/\\_,.:;~-]+", "", s)
    m = re.search(r"[A-Za-zА-Яа-я]\d{5}", s)
    if not m:
        return ""
    token = m.group(0).upper()
    # Force Latin B if the first letter is non-Latin
    if token[0] not in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        token = "B" + token[1:]
    return token

def digits_in_range(raw: str, lo: int, hi: int) -> str:
    """Extract digits from raw and keep only if lo <= value <= hi."""
    if not raw:
        return ""
    d = "".join(ch for ch in raw if ch.isdigit())
    if not d:
        return ""
    try:
        v = int(d)
    except ValueError:
        return ""
    return str(v) if lo <= v <= hi else ""

# ---------- Cell preprocessing ----------
def preprocess_cell(cell_bgr):
    """Light binarization helps both EasyOCR and Tesseract."""
    gray = cv2.cvtColor(cell_bgr, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return cv2.cvtColor(th, cv2.COLOR_GRAY2BGR)

# ---------- Table segmentation ----------
def _find_boxes(image_bgr):
    """Detect grid lines and return bounding boxes for candidate cells."""
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    _, bin_inv = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Kernels for horizontal and vertical lines
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))

    # Correct order of arguments
    h_lines = cv2.morphologyEx(bin_inv, cv2.MORPH_OPEN, h_kernel, iterations=2)
    v_lines = cv2.morphologyEx(bin_inv, cv2.MORPH_OPEN, v_kernel, iterations=2)

    grid = cv2.add(h_lines, v_lines)
    contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(c) for c in contours if cv2.contourArea(c) > 1000]
    boxes.sort(key=lambda b: (b[1], b[0]))
    return boxes

def _group_rows(boxes, y_tol=30):
    """Group boxes into rows using a vertical tolerance."""
    rows, cur, cur_y = [], [], None
    for (x, y, w, h) in boxes:
        if cur_y is None:
            cur, cur_y = [(x, y, w, h)], y
            continue
        if abs(y - cur_y) > y_tol:
            cur.sort(key=lambda b: b[0])
            rows.append(cur)
            cur, cur_y = [(x, y, w, h)], y
        else:
            cur.append((x, y, w, h))
    if cur:
        cur.sort(key=lambda b: b[0])
        rows.append(cur)
    return rows

# ---------- Smart assignment per row ----------
def _assign_row_fields(img, row_boxes, verbose=False):
    """Assign OCR'd cells to logical fields based on content."""
    cells = []
    for (x, y, w, h) in row_boxes:
        crop = img[y:y+h, x:x+w]
        proc = preprocess_cell(crop)
        raw_text = easyocr_text(proc, digits=False)
        raw_digits = easyocr_text(proc, digits=True)
        cells.append({
            "box": (x, y, w, h),
            "raw_text": raw_text,
            "raw_digits": raw_digits
        })

    # 1) Roll Number
    roll_idx, roll_val = -1, ""
    for i, c in enumerate(cells):
        rn = clean_roll_number(c["raw_text"])
        if rn:
            roll_idx, roll_val = i, rn
            break

    # 2) Other numeric fields
    remaining = [i for i in range(len(cells)) if i != roll_idx]

    def pick_best(choices, lo, hi):
        for i in choices:
            v = digits_in_range(cells[i]["raw_digits"], lo, hi)
            if v:
                return i, v
        return -1, ""

    fmt_i, fmt_v = pick_best(remaining, 300, 1800)
    rem2 = [i for i in remaining if i != fmt_i]

    w_i, w_v = pick_best(rem2, 20, 3000)
    rem3 = [i for i in rem2 if i != w_i]

    g_i, g_v = pick_best(rem3, 20, 400)
    rem4 = [i for i in rem3 if i != g_i]

    # 3) Comment
    comment_v = ""
    if rem4:
        widest = max(rem4, key=lambda i: cells[i]["box"][2])
        x, y, w, h = cells[widest]["box"]
        crop = img[y:y+h, x:x+w]
        comment_v = tesseract_text(preprocess_cell(crop))
    if not comment_v and rem4:
        comment_v = " ".join([cells[i]["raw_text"] for i in rem4]).strip()

    row = {c: "" for c in COLUMNS}
    row["Roll Number"] = roll_val
    row["Format (mm)"] = fmt_v
    row["Weight (kg)"] = w_v
    row["Grammage (g/m2)"] = g_v
    row["Comment"] = comment_v

    if verbose:
        print(f"    -> Assigned: RN={roll_val} | F={fmt_v} | W={w_v} | G={g_v} | C='{comment_v[:40]}'")
    return row

# ---------- Public API ----------
def extract_table_rows(image_path, debug=False, verbose=False):
    """Extract table rows using OCR."""
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"Cannot load image: {image_path}")

    if verbose:
        print(f"📄 Processing: {image_path}")

    os.makedirs("output", exist_ok=True)
    boxes = _find_boxes(img)
    rows_boxes = _group_rows(boxes, y_tol=30)

    base = os.path.splitext(os.path.basename(image_path))[0]
    debug_img = img.copy()
    debug_log = os.path.join("output", "debug_log.csv")
    debug_image = os.path.join("output", f"debug_{base}.jpg")
    if base.lower().startswith("table"):
        debug_image = os.path.join("output", "debug_table1.jpg")

    results = []
    with open(debug_log, "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["Row", "Column", "Raw Text (note)", "Value"])

        for r_idx, row_boxes in enumerate(rows_boxes, start=1):
            if debug:
                for (x, y, w, h) in row_boxes:
                    cv2.rectangle(debug_img, (x, y), (x+w, y+h), (0, 255, 0), 2)

            row = _assign_row_fields(img, row_boxes, verbose=verbose)
            results.append(row)

            wr.writerow([r_idx, "Roll Number", "", row["Roll Number"]])
            wr.writerow([r_idx, "Format (mm)", "", row["Format (mm)"]])
            wr.writerow([r_idx, "Weight (kg)", "", row["Weight (kg)"]])
            wr.writerow([r_idx, "Grammage (g/m2)", "", row["Grammage (g/m2)"]])
            wr.writerow([r_idx, "Comment", "", (row["Comment"][:100] if row["Comment"] else "")])

    if debug:
        cv2.imwrite(debug_image, debug_img)
        if verbose:
            print(f"🔎 Debug image saved: {debug_image}")
            print(f"📑 Debug log saved: {debug_log}")

    return results

if __name__ == "__main__":
    rows = extract_table_rows("input/table1.jpg", debug=True, verbose=True)
    for r in rows:
        print(r)
