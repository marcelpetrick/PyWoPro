import sys
import argparse
import re
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfFileMerger, PdfFileReader
from PIL import Image

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

def convert_to_pdf():
    convert('modified.docx', 'modified.pdf')

def split_pdf(manufacturer_spec, article_number):
    inputpdf = PdfFileReader(open("modified.pdf", "rb"))

    with open("front.pdf", "wb") as outputStream:
        output = PdfFileMerger()
        output.addPage(inputpdf.getPage(0))
        output.write(outputStream)

    with open("back.pdf", "wb") as outputStream:
        output = PdfFileMerger()
        output.addPage(inputpdf.getPage(1))
        output.write(outputStream)

    output = PdfFileMerger()
    output.append("front.pdf")
    output.append(manufacturer_spec)
    output.append("back.pdf")
    output.write(f"Spec_{article_number}.pdf")

def main():
    template, article_number, logo, manufacturer_spec = get_args()
    edit_docx(template, article_number, logo)
    convert_to_pdf()
    split_pdf(manufacturer_spec, article_number)

if __name__ == "__main__":
    main()

# python3 main.py --template=template_front_and_back.docx --articleNumber=3.1415 --logo=fancy_logo.png --manufacturerSpec=manufacturer_spec.pdf