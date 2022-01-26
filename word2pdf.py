import os
import win32com.client
import glob
import tqdm
import time

NOWPATH = os.getcwd() + "/"
docs = glob.glob(NOWPATH+"*.doc")


def convert():
    for doc in tqdm.tqdm(docs):
        wdFormatPDF = 17
        outputFile = doc.rstrip(".doc")+".pdf"
        file = open(outputFile, "w")
        file.close()
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(doc)
        doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


if __name__ == "__main__":
    
    print("Converting WORD to PDF....\n")
    convert()
    print("\nCompleted!\n")
    time.sleep(2)

