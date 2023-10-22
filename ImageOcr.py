import glob
import cv2
import pytesseract
from odf.opendocument import OpenDocumentText
from odf.style import Style, ParagraphProperties
from odf.text import H, P, Span
import tkinter as tk
from tkinter import *

window=tk.Tk()
window.title('OCR Img to OpenDoc')
window.geometry('600x400')
infoLabel = tk.Label(text = 'Places your files in .jpg format inside Images folder then press the button to start creating a open document file')
infoLabel.grid(row = 0, column = 0, sticky = W, pady = 2)

def ocrImgToOpenDoc():
    print('start')
    logLabel = tk.Label(text = 'Please wait until the file `myOcrDoc.odt` is created...')
    logLabel.grid(row = 1, column = 0, sticky = W, pady = 2)
    # Open Office GitHub documentation = https://github.com/eea/odfpy/wiki/Introduction
    # Tesseract GitHub documentation = https://github.com/tesseract-ocr/tesseract

    # Create a open document file
    textdoc = OpenDocumentText()

    # Create a style for the paragraph with page-break
    withbreak = Style(name="WithBreak", parentstylename="Standard", family="paragraph")
    withbreak.addElement(ParagraphProperties(breakbefore="page"))
    textdoc.automaticstyles.addElement(withbreak)

    # Open the image files
    imgs = [cv2.imread(file) for file in glob.glob("Images/*.jpg")]

    # Perform OCR using PyTesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\zanni\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

    pageNumber = 0
    for img in imgs:
        pageNumber = pageNumber + 1
        ocrText = pytesseract.image_to_string(img)
        print('----------')
        print('Page Number being crawled is : ' + str(pageNumber))
        p = P(text=ocrText, stylename=withbreak)
        textdoc.text.addElement(p)
        # Alternative text saving, every page is saved to a different .txt file
        # with open(str(x)+".txt", mode = 'w') as f:
        # f.write(text)

    textdoc.save("myOcrDoc.odt")
    print('end')

B = Button(window, text ="Create OpenDoc File", command = ocrImgToOpenDoc)
B.place(x=50,y=50)

window.mainloop()

