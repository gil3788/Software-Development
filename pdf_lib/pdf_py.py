import PyPDF2
import re

def pdfRepetitions(pattern):
   r = re.compile(r"(.+?)\1+")
   for match in r.finditer(pattern):
    yield (match.group(1), len(match.group(0))/len(match.group(1)))


def pdfPageFunction (pdfFile, pdfPageNum):
    counter = 0
    page = {}
    page["text"] = ""
    page["numPages"] = 0
    pdfFileObj = open(pdfFile,"rb")
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    #print "Number of Pages in PDF File: ", pdfReader.numPages
    while counter < pdfReader.numPages:
        pageObj = pdfReader.getPage(counter)
        page["text"] = page["text"] + pageObj.extractText()
        counter = counter + 1

    page["numPages"] = pdfReader.numPages


    return page

def pdfValueFinder (pdfPage, findVal):
    valuePosition = pdfPage.find(findVal, 10)
    return valuePosition
