import PyPDF2
import openpyxl
import os

#list to hold all the pdf files in directory
filenames = []
#Path for where pdf files are located on my local machine
directory = r"C:\Users\llb20\PDFMultiplePages"
#loops through all the files that end with .pdf in the listed directory above
for filename in os.listdir(directory):
    if filename.lower().endswith(".pdf"):
        filenames.append(os.path.join(directory, filename))
#Reads all the pdf files in filesnames list
for filename in filenames:
    x = PyPDF2.PdfReader(open(filename, "rb"))
#loop to scan all the pages in a pdf file and extract text
    num_pages = len(x.pages)
    for page_num in range(num_pages):
        page = x.pages[page_num]
        text = page.extract_text()
        print(f"Scanning file: {filename} page: {page_num + 1}")
#extract lines not all the text
        line = text.split(' ')
#Pdfs are all structured differently. This was the data I needed from my resume to mimic what we were doing for our printing company.
#It will not work for all pdfs we will have to look at the structure and change it accordingly.
        data = [
            [line[0]+ ' ' + line[1],line[2]+ ' ' + line[3]+ ' ' + line[4], line[5].removesuffix(','),line[6],line[7],line[8]]
            ]
#Locate excel file that data should be loaded to and opens it
        wb = openpyxl.load_workbook('C:/Users/llb20/Excel/DataExtracted2.xlsx')
        sheet = wb.active
        sheet.title = "MYPDF"
#Adding data to rows in the excel file
        for row in data:
            sheet.append(row)
#save data to
# excel workbook
        wb.save("C:/Users/llb20/Excel/DataExtracted2.xlsx")