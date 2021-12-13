import os
import sys
import glob
import pdfkit
import PyPDF2
import mammoth
import base64
from PIL import Image
from pdfkit import source
import subprocess
from subprocess import PIPE
from pygments import highlight
from pygments.lexers.python import Python3Lexer
from pygments.formatters.html import HtmlFormatter

# Function to merge all pdf files in a target directory.
def merge_pdf_in_dir(dir_path, dst_path):

    # list up all pdf files
    l = glob.glob(os.path.join(dir_path, '*.pdf'))
    l.sort()

    merger = PyPDF2.PdfFileMerger()
    for p in l:
        if not PyPDF2.PdfFileReader(p).isEncrypted:
            merger.append(p)

    merger.write(dst_path)
    merger.close()
    return len(l)

# remove file if 
def removefiles(abs_paths):
    for abs_path in abs_paths:
        if os.path.isfile(abs_path):
            os.remove(abs_path)

# merge pdf
def mergepdf(target_dir):
    # Set target directory
    print(f"Start processing {target_dir}")

    # remove trashes
    removefiles(glob.glob(os.path.join(target_dir, '*.exe')))
    removefiles(glob.glob(os.path.join(target_dir, '*.out')))
    removefiles(glob.glob(os.path.join(target_dir, 'auto_united_report.pdf')))

    # Initialize html source
    output_html = ""
    processed_file_num = 0

    # Select source and text files
    target_files = glob.glob(target_dir+"/*.c") + glob.glob(target_dir+"/*.txt") + glob.glob(target_dir+"/*.md")

    # Add non-extention file
    target_files += [nonext for nonext in glob.glob(target_dir+"/*") if os.path.splitext(nonext)[-1] == '' ]

    processed_file_num += len(target_files)
    for source_fn in target_files:
        try:
            f = open(source_fn, 'r', encoding="UTF-8")
            code = f.read()
            f.close()
        except:
            f = open(source_fn, 'r', encoding="shift_jis")
            code = f.read()
            f.close()

        output_html += "<!DOCTYPE html><head><style type='text/css'>" \
                    + HtmlFormatter(linenos=True, cssclass="source").get_style_defs('.highlight')\
                    + "</style></head></body>" \
                    + "<h2>"+ source_fn +"</h2>" \
                    + highlight(code, Python3Lexer(), HtmlFormatter())\
                    + "</body></html>"

    # Select source and text files
    target_files = glob.glob(target_dir+"/*.f*") + glob.glob(target_dir+"/*.cc") + glob.glob(target_dir+"/*.cpp")
    processed_file_num += len(target_files)
    for source_fn in target_files:
        try:
            f = open(source_fn, 'r', encoding="UTF-8")
            code = f.read()
            f.close()
        except:
            f = open(source_fn, 'r', encoding="shift_jis")
            code = f.read()
            f.close()

        output_html += "<!DOCTYPE html><head><style type='text/css'>" \
                    + HtmlFormatter(linenos=True, cssclass="source").get_style_defs('.highlight')\
                    + "</style></head></body>" \
                    + "<h1 style='color:red;'> ### EXTENTION WARNING ### </h1>" \
                    + "<h2>" + source_fn + "</h2>"\
                    + highlight(code, Python3Lexer(), HtmlFormatter())\
                    + "</body></html>"


    # TODO : deal with vector graphics
    # Select EPS or SVG files
    # vector_images = glob.glob(target_dir+"/*.eps") + glob.glob(target_dir+"/*.svg")

    # for vecimg in vector_images : 
    #     print(vecimg)
    #     print(f"magick convert {os.path.basename(vecimg)} {os.path.splitext(os.path.basename(vecimg))[0]}.png")
    #     popen = subprocess.Popen(f"magick convert {os.path.basename(vecimg)} {os.path.splitext(os.path.basename(vecimg))[0]}.png", shell=True, stdout=PIPE, stderr=PIPE, text=True, cwd=target_dir)
    #     popen.wait() # wait during conversion
        
    # print(target_dir + "/" + os.path.splitext(os.path.basename(vecimg))[0]+".png")
    # im = Image.open(vecimg)
    # im = im.thumbnail((2000,2000), Image.ANTIALIAS)
    # fig = im.convert("RGBA")
    # fig.save(target_dir + "/" + os.path.splitext(os.path.basename(vecimg))[0]+".png", lossless=True)


    # Select image files
    target_files = glob.glob(target_dir+"/*.jpg") + glob.glob(target_dir+"/*.jpeg") \
                + glob.glob(target_dir+"/*.gif") + glob.glob(target_dir+"/*.png") + glob.glob(target_dir+"/*.bmp") + glob.glob(target_dir+"/*.eps")
    processed_file_num += len(target_files)

    for image_fn in target_files:
        encoded_str = base64.b64encode(open(image_fn,'rb').read())
        data_uri = 'data:image/jpg;base64,'+encoded_str.decode()
        output_html += "<!DOCTYPE html></body>" \
                    + "<h2>"+ image_fn +"</h2>" \
                    + "<img src='"+ data_uri +"' width='100%'>"\
                    + "</body></html>"

    # Select word files
    target_files = glob.glob(target_dir+"/*.docx") + glob.glob(target_dir+"/*.doc")
    processed_file_num += len(target_files)
    for word_fn in target_files:
        with open(word_fn, 'rb') as document:
            word_doc = mammoth.convert_to_html(document)
            output_html += "<h2>"+ word_fn +"</h2>" + word_doc.value
            messages = word_doc.messages
            if messages : print(f"Word to html converter : {messages}")

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
    output_path = os.path.join(target_dir,"auto_gen_output.pdf")
    print(f"Making pdf at {output_path}")
    pdfkit.from_string(output_html, output_path, options=options)

    # Merge all pdf files in the target dir
    processed_file_num += merge_pdf_in_dir(target_dir, os.path.join(target_dir, "auto_united_report.pdf")) + 1

    # Check if other type of files are not included.
    total_file_num = len(glob.glob(target_dir+"/*"))
    if total_file_num != processed_file_num : 
        print(f"total_file_num={total_file_num}, but processed_file_num={processed_file_num}")
        print(f"##### Error! Unsupported file format has found! Check inside {target_dir} #####")
        sys.exit()
    print(f"All files in {target_dir} processed successfully.")

def mergepdf_each_dir(target_folders, subdir):

    # Run merging process for each directories.
    for dirpath in target_folders:
        filesndirs = os.listdir(dirpath)
        dirs = [f for f in filesndirs if os.path.isdir(os.path.join(dirpath,f))]
        for dir in dirs :
            mergepdf(os.path.join(os.path.join(dirpath,dir), subdir))

if __name__ == '__main__':
    root_dir = "C:/Users/mizuho/Desktop/成績2021"
    target_folders = [
                    os.path.join(root_dir,"課題１＆２"), 
                    os.path.join(root_dir,"課題３"), 
                    os.path.join(root_dir,"中間レポート")
                   ]
    print(target_folders)
    mergepdf_each_dir(target_folders, subdir="提出物の添付ファイル")