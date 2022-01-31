import os
import win32com.client
import glob
import tqdm
import time
import re

NOWPATH = os.getcwd() + "/"
DOCFILES = [doc for doc in glob.glob(
    NOWPATH+'/*') if re.search('/*\.(doc|docx)$', str(doc))]


def main():

    print("Converting WORD to PDF....\n")
    
    for doc in tqdm.tqdm(DOCFILES):
        wdFormatPDF = 17
        #outputFile = doc.rstrip(".doc")+".pdf"
        outputFile = doc+".pdf"
        file = open(outputFile, "w")
        file.close()
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(doc)
        doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        
    print("\nCompleted!\n")
    time.sleep(2)

if __name__ == "__main__":
    main()




