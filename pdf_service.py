import os
from fpdf import FPDF
from datetime import datetime

PDF_PATH = 'pdf'

if not os.path.exists(PDF_PATH):
    os.makedirs(PDF_PATH)
    
