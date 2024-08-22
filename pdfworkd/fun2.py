import os
from pdf2docx import Converter
from docx import Document
from docx2pdf import convert

def replace_text(words_to_replace, replace_word, foldername):
    docxname = convert_to_docx('22Q71A4422.pdf')
    print(docxname)
    for word in words_to_replace:
        doc = Document(docxname)
        for paragraph in doc.paragraphs:
            if replace_word in paragraph.text:
                print('Replacing text in docx')
                paragraph.text = paragraph.text.replace(replace_word, word)
        output_file_name = os.path.join(foldername, f"{word}.docx")
        doc.save(output_file_name)
        convert_to_pdf(output_file_name, word)

def convert_to_pdf(filename, name):
    print('Converting to docx')
    pdf_folder = 'finalpdfs'
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)
    pdf_name = os.path.join(pdf_folder, f"{name}.pdf")
    convert(filename, pdf_name)

def remove_pdf_extension(filename):
    if filename.endswith('.pdf'):
        print(filename[:-4])
        return filename[:-4]  
    else:
        return filename

def convert_to_docx(filename):
    print(f"The filename in the docx conversion function is: {filename}")
    converter = Converter(filename)

    docx_folder = 'docx_folder'
    if not os.path.exists(docx_folder):
        os.makedirs(docx_folder)

    filename2 = remove_pdf_extension(filename)

    docx_name = os.path.join(docx_folder, f"{filename2}.docx")

    converter.convert(docx_name, start=0, end=None)
    converter.close()

    return docx_name

search_word = '22Q71A4454'
list_of_roll = ['22Q71A4434', '22Q71A4422', '22Q71A4494']
foldername = 'output_folder'

if not os.path.exists(foldername):
    os.makedirs(foldername)

replace_text(list_of_roll, search_word, foldername)
