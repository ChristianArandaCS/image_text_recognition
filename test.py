# import libraries
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import xlsxwriter
import os


# set pytesseract location via variable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# create output excel file
workbook = xlsxwriter.Workbook('tester.xlsx')
worksheet = workbook.add_worksheet()

header1 = 'File name'
header2 = 'Output file'

worksheet.write(0,0,header1)
worksheet.write(0,1,header2)

row = 1
col = 0
#for x in img_name:
#    worksheet.write(row, col, img_name[x])
#    col = col +1

for filename in os.listdir(r'C:\Users\chris\Desktop\scipts\python\test\barcodes'):

    # get img name
    worksheet.write(row, col, filename)

    col = col +1

    # extracts text from img
    im = Image.open(filename) # the second one 
    im = im.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(im)
    im = enhancer.enhance(2)
    im = im.convert('1')
    #im.save(filename)
    text = pytesseract.image_to_string(Image.open(filename))

    worksheet.write(row, col, text)

    row = row +1
    col = col - 1

    
workbook.close()

quit()





for filename in os.listdir(r'C:\Users\chris\Desktop\scipts\python\test\barcodes'):

    # get all
    im = Image.open(filename) # the second one 
    im = im.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(im)
    im = enhancer.enhance(2)
    im = im.convert('1')
    im.save(img)
    text = pytesseract.image_to_string(Image.open(img))

    row = 2
    col = 0

    worksheet.write(row, col, text)

    col = col + 1

workbook.close()

