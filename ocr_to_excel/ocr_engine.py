import os
import easyocr
from PIL import Image

reader = easyocr.Reader(['ru', 'en'], gpu=False)

def extract_text_from_image(image_path):
    results = reader.readtext(image_path, detail=0, paragraph=False)
    return results

def extract_all_images(input_folder='input'):
    all_results = {}
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            path = os.path.join(input_folder, filename)
            text = extract_text_from_image(path)
            all_results[filename] = text
    return all_results
