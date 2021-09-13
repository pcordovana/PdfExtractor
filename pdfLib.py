import sys
import os
import easygui
import getpass
from easygui import *

from PyPDF2 import PdfFileReader, PdfFileWriter

def extract_information(pdf_path):
    try:
        with open(pdf_path, 'rb') as f:
            pdf = PdfFileReader(f)
            information = pdf.getDocumentInfo()
            number_of_pages = pdf.getNumPages()

        txt = f"""
        Information about {pdf_path}: 

        Author: {information.author}
        Creator: {information.creator}
        Producer: {information.producer}
        Subject: {information.subject}
        Title: {information.title}
        Number of pages: {number_of_pages}
        """

        print(txt)
        return information

    except FileNotFoundError:
        print(f"File {pdf_path} not found.  Aborting")
        sys.exit(1)
    except OSError:
        print(f"OS error occurred trying to open {pdf_path}")
        sys.exit(1)
    except Exception as err:
        print(f"Unexpected error opening {pdf_path} is",repr(err))
        sys.exit(1)  # or replace this with "raise" ?

def merge_pdfs(paths, output):
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            # Add each page to the writer object
            pdf_writer.addPage(pdf_reader.getPage(page))

    # Write out the merged PDF
    with open(output, 'wb') as out:
        pdf_writer.write(out)


def rotate_pages(pdf_path):
    pdf_writer = PdfFileWriter()
    pdf_reader = PdfFileReader(pdf_path)
    # Rotate page 90 degrees to the right
    page_1 = pdf_reader.getPage(0).rotateClockwise(90)
    pdf_writer.addPage(page_1)
    # Rotate page 90 degrees to the left
    page_2 = pdf_reader.getPage(1).rotateCounterClockwise(90)
    pdf_writer.addPage(page_2)
    # Add a page in normal orientation
    pdf_writer.addPage(pdf_reader.getPage(2))

    with open('rotate_pages.pdf', 'wb') as fh:
        pdf_writer.write(fh)

def split(path, name_of_split):
    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        output = f'{name_of_split}{page}.pdf'
        with open(output, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

def add_encryption(input_pdf, output_pdf, password):
    pdf_writer = PdfFileWriter()
    pdf_reader = PdfFileReader(input_pdf)
    for page in range(pdf_reader.getNumPages()):
        pdf_writer.addPage(pdf_reader.getPage(page))
    pdf_writer.encrypt(user_pwd=password, owner_pwd=None, use_128bit=True)
    with open(output_pdf, 'wb') as fh:
        pdf_writer.write(fh)


def multi_choice():
    msg = "Inserisci le scelte seguenti:"
    title = "ESTRAZIONE PAGINE DA DOCUMENTO PDF"
    fieldNames = ["Prima pagina", "Ultima pagina", "Nome del nuovo documento"]
    fieldValues = multenterbox(msg, title, fieldNames)
    if fieldValues is None:
        sys.exit(0)
    # make sure that none of the fields were left blank
    while 1:
        errmsg = ""
        for i, name in enumerate(fieldNames):
            if fieldValues[i].strip() == "":
                errmsg += "{} is a required field.\n\n".format(name)
        if errmsg == "":
            break # no problems found
        fieldValues = multenterbox(errmsg, title, fieldNames, fieldValues)
        if fieldValues is None:
            break
    print("Reply was:{}".format(fieldValues))
    return fieldValues

if __name__ == '__main__':
    msgCreate= ""
    fileChoosed = easygui.fileopenbox(msg="Scegli il documento pdf da analizzare", title="PDF EXTRACT", default='*.pdf', filetypes="*.pdf", multiple=False)
    if fileChoosed != None:
        print("file scelto = ", fileChoosed)

        drive, pathAndFile = os.path.splitdrive(fileChoosed)
        nameFile = os.path.basename(pathAndFile).strip()
        pathInput = os.path.dirname(pathAndFile)

        fieldValues = multi_choice()
        firstPage = int (fieldValues[0])
        if firstPage and firstPage != 1:
            firstPage -= 1
        lastPage = int (fieldValues[1])
        if lastPage and firstPage != 1:
            lastPage -= 1

        if firstPage and lastPage:
            # Create a PdfFileWriter object
            out = PdfFileWriter()
      
            # Open PDF file with the PdfFileReader
            file = PdfFileReader(fileChoosed)
            information = file.getDocumentInfo()

            for idx in range(file.numPages):
                
                # Get the page at index idx
                #scarta la prima pagina vuota e prende idx+1
                if idx >= firstPage and idx <=lastPage:
                    page = file.getPage(idx)
                    # Add it to the output file
                    out.addPage(page)

                if idx >= lastPage:
                    break

            out.addMetadata({
                    '/Author': getpass.getuser(),
                })
            nameFile = fieldValues[2]+".pdf"
            pathAndNameFile = os.path.join(pathInput, nameFile)
            
            with open(pathAndNameFile, "wb") as f:
                
                # Write our decrypted PDF to this file
                out.write(f)
                msgCreate= "Creato documento {} - Pagine totali {} ".format(pathAndNameFile, lastPage-firstPage+1)
                msgCreate += "\nAutore: {} ".format(getpass.getuser())

    msgbox(msg=msgCreate+'\n\n\n                              PROGRAMMA TERMINATO', title=' PDF EXTRACT ', ok_button='OK', image=None, root=None)
