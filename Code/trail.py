# -*- coding: utf-8 -*-
import PyPDF2
import textract
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

# This function will extract and return the pdf file text content.
def extractPdfText(filePath=''):

    # Open the pdf file in read binary mode.
    fileObject = open(filePath, 'rb')

    # Create a pdf reader .
    pdfFileReader = PyPDF2.PdfFileReader(fileObject)

    # Get total pdf page number.
    totalPageNumber = pdfFileReader.numPages

    # Print pdf total page number.
    print('This pdf file contains totally ' + str(totalPageNumber) + ' pages.')

    currentPageNumber = 0
    text = ''

    # Loop in all the pdf pages.
    while(currentPageNumber < totalPageNumber ):

        # Get the specified pdf page object.
        pdfPage = pdfFileReader.getPage(currentPageNumber)

        # Get pdf page text.
        text = text + pdfPage.extractText()

        # Process next page.
        currentPageNumber += 1

    if(text == ''):
        # If can not extract text then use ocr lib to extract the scanned pdf file.
        text = textract.process(filePath, method='tesseract', encoding='utf-8')
       
    return text

# This function will remove all stop words and punctuations in the text and return a list of keywords.
def extractKeywords(text):
    # Split the text words into tokens
    wordTokens = word_tokenize(text)

    # Remove blow punctuation in the list.
    punctuations = ['(',')',';',':','[',']',',']

    # Get all stop words in english.
    stopWords = stopwords.words('english')

    # Below list comprehension will return only keywords tha are not in stop words and  punctuations
    keywords = [word for word in wordTokens if not word in stopWords and not word in punctuations]
   
    return keywords

if __name__ == '__main__': 

    pdfFilePath = '/Users/zhaosong/Documents/WorkSpace/e-book/Mastering-Node.js.pdf'
   
    pdfText = extractPdfText(pdfFilePath)
    print('There are ' + str(pdfText.__len__()) + ' word in the pdf file.')
    #print(pdfText)

    keywords = extractKeywords(pdfText)
    print('There are ' + str(keywords.__len__()) + ' keyword in the pdf file.')
    #print(keywords)    