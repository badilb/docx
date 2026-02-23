import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

QR_CODE_PATH = os.path.join(UPLOAD_DIR, "2025-03-13 14.34.05.jpg")
LOGO_PATH = os.path.join(UPLOAD_DIR, "dcx logo png.png")