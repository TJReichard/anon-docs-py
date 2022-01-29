from pydoc import Doc
from tkinter import N
from docx import Document
import csv

document = Document('testdoc.docx')


result = [p.text for p in document.paragraphs]

with open('nameslist.csv', newline='', encoding='utf-8') as csvfile:
    namesList = csv.reader(csvfile, delimiter=',', quotechar='|')
    
    cleanNamesList = []
    for row in namesList:
        cleanNamesList.append(row[3].strip('""'))
        cleanNamesList.append(row[4].strip('""'))

def clean(text, cleanNamesList, run):
    cleaned = Document()
    if run <2:
        for par in text:
            if any(name in par for name in cleanNamesList):
                for name in cleanNamesList:
                    if par.find(name) != -1:
                        newSen = par.replace(name, name[0]+'.')
                        cleaned.add_paragraph(newSen)
            else:
                cleaned.add_paragraph(par)
        # print(run)
        cleanedResult = [p.text for p in cleaned.paragraphs]
        # print(f"""run: {run}
        # cleanedResult: {cleanedResult}""")
        run += 1
        clean(cleanedResult, cleanNamesList, run)
    else:
        finalResult = text
        end(finalResult)

def end(final):
    finishedDocument = Document()
    finalText = ""
    for idx, f in enumerate(final):
        try:
            if final[idx] == final[idx+1]:
                pass
            else:
                finishedDocument.add_paragraph(f)
                # finalText += f + '\n'
        except:
            pass

    finishedDocument.save('newandanonDoc.docx')   
    # print(finalText)

clean(result,cleanNamesList, 0)