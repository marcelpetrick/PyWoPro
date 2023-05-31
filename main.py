import argparse
import re
from docx import Document
import subprocess
from PyPDF2 import PdfReader, PdfWriter
import os

def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", help="Template DOCX file", required=True)
    parser.add_argument("--articleNumber", help="Article number to be replaced in the DOCX file", required=True)
    parser.add_argument("--logo", help="Logo PNG file to be replaced in the DOCX file", required=True)
    parser.add_argument("--manufacturerSpec", help="Manufacturer spec PDF file", required=True)

    args = parser.parse_args()

    return args.template, args.articleNumber, args.logo, args.manufacturerSpec

def edit_docx(template, article_number, logo):
    doc = Document(template)

    for paragraph in doc.paragraphs:
        if "Article number" in paragraph.text:
            paragraph.text = re.sub('Article number', article_number, paragraph.text)

    for rel in doc.part.rels.values():
        if "Logo" in rel.reltype:
            rel.reltype = logo

    doc.save('modified.docx')

def convert_to_pdf(docx_file):
    command = f'libreoffice --headless --convert-to pdf:writer_pdf_Export --outdir . {docx_file}'
    subprocess.run(command, shell=True, stdout=subprocess.DEVNULL)


def split_pdf(manufacturer_spec, article_number):
    inputpdf = PdfReader("modified.pdf")

    with open("front.pdf", "wb") as outputStream:
        output = PdfWriter()
        output.add_page(inputpdf.pages[0])
        output.write(outputStream)

    with open("back.pdf", "wb") as outputStream:
        output = PdfWriter()
        output.add_page(inputpdf.pages[1])
        output.write(outputStream)

    output = PdfWriter()
    output.add_page(PdfReader("front.pdf").pages[0])
    for page in PdfReader(manufacturer_spec).pages:
        output.add_page(page)
    output.add_page(PdfReader("back.pdf").pages[0])

    with open(f"Spec_{article_number}.pdf", "wb") as outputStream:
        output.write(outputStream)

def cleanup():
    files_to_delete = ['modified.pdf', 'modified.docx', 'front.pdf', 'back.pdf']
    for filename in files_to_delete:
        try:
            os.remove(filename)
        except FileNotFoundError:
            print(f"File {filename} not found.")


def main():
    template, article_number, logo, manufacturer_spec = get_args()
    edit_docx(template, article_number, logo)
    convert_to_pdf('modified.docx')
    split_pdf(manufacturer_spec, article_number)
    cleanup()

if __name__ == "__main__":
    main()

# python3 main.py --template=template_front_and_back.docx --articleNumber=3.1415 --logo=fancy_logo.png --manufacturerSpec=manufacturer_spec.pdf


# javaldx: Could not find a Java Runtime Environment!
# Warning: failed to read path from javaldx
# Solved by:
# sudo pacman -S jre-openjdk-headless jre-openjdk jdk-openjdk openjdk-doc openjdk-src
