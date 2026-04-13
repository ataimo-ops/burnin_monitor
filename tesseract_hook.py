import os
import sys

# Set Tesseract path for pytesseract
tesseract_exe = os.path.join(sys._MEIPASS, 'tesseract.exe')
if os.path.exists(tesseract_exe):
    # Set the tesseract command path
    try:
        import pytesseract
        pytesseract.pytesseract.tesseract_cmd = tesseract_exe
    except ImportError:
        pass

# Set tessdata path
tessdata_dir = os.path.join(sys._MEIPASS, 'tessdata')
if os.path.exists(tessdata_dir):
    os.environ['TESSDATA_PREFIX'] = tessdata_dir