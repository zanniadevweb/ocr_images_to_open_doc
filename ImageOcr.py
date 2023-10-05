import glob
import cv2
import pytesseract
from odf.opendocument import OpenDocumentText
from odf.style import Style, ParagraphProperties
from odf.text import H, P, Span
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
    print('Page Number being crawled is :' + str(pageNumber))
    p = P(text=ocrText, stylename=withbreak)
    textdoc.text.addElement(p)
    # Alternative text saving, every page is saved to a different .txt file
    # with open(str(x)+".txt", mode = 'w') as f:
    # f.write(text)

textdoc.save("myOcrDoc.odt")

