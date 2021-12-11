import os
import sys
import glob
import pdfkit
import PyPDF2
import base64
from pdfkit import source
from pygments import highlight
from pygments.lexers.python import Python3Lexer
from pygments.formatters.html import HtmlFormatter

# Function to merge all pdf files in a target directory.
def merge_pdf_in_dir(dir_path, dst_path):
    l = glob.glob(os.path.join(dir_path, '*.pdf'))
    l.sort()

    merger = PyPDF2.PdfFileMerger()
    for p in l:
        if not PyPDF2.PdfFileReader(p).isEncrypted:
            merger.append(p)

    merger.write(dst_path)
    merger.close()
    return len(l)

# Set target directory
target_dir = os.path.abspath("./sample_report")
output_dir = os.path.abspath("./")

# Initialize html source
output_html = ""
processed_file_num = 0

# Select .c source and text files
target_files = glob.glob(target_dir+"/*.c") + glob.glob(target_dir+"/*.txt")
processed_file_num += len(target_files)
for source_fn in target_files:
    f = open(source_fn, 'r', encoding="UTF-8")
    code = f.read()
    f.close()

    output_html += "<!DOCTYPE html><head><style type='text/css'>" \
                + HtmlFormatter(linenos=True, cssclass="source").get_style_defs('.highlight')\
                + "</style></head></body>" \
                + "<h2>"+ source_fn +"</h2>" \
                + highlight(code, Python3Lexer(), HtmlFormatter())\
                + "</body></html>"

# Select image files
target_files = glob.glob(target_dir+"/*.jpg") + glob.glob(target_dir+"/*.jpeg") \
             + glob.glob(target_dir+"/*.gif") + glob.glob(target_dir+"/*.png") + glob.glob(target_dir+"/*.bmp")
processed_file_num += len(target_files)

for image_fn in target_files:
    encoded_str = base64.b64encode(open(image_fn,'rb').read())
    data_uri = 'data:image/jpg;base64,'+encoded_str.decode()
    output_html += "<!DOCTYPE html></body>" \
                + "<h2>"+ image_fn +"</h2>" \
                + "<img src='"+ data_uri +"' width='100%'>"\
                + "</body></html>"

# Output a pdf file.
options = {
    'page-size': 'A4',
    'orientation': 'Portrait',
    'margin-top': '0.4in',
    'margin-right': '0.4in',
    'margin-bottom': '0.4in',
    'margin-left': '0.4in',
    'encoding': "UTF-8",
    'no-outline': None
}
print("Making pdf...")
pdfkit.from_string(output_html, os.path.join(target_dir,"auto_gen_output.pdf"), options=options)

# Merge all pdf files in the target dir
processed_file_num += merge_pdf_in_dir(target_dir, os.path.join(output_dir, "auto_united_report.pdf"))

# Check if other type of files are not included.
total_file_num = len(glob.glob(target_dir+"/*"))
if total_file_num != processed_file_num : 
    print(f"total_file_num={total_file_num}, but processed_file_num={processed_file_num}")
    print(f"##### Error! Unsupported file format has found! Check inside {target_dir} #####")
    sys.exit()
print(f"All files in {target_dir} processed successfully.")