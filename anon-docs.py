from tkinter import Label, Scale, StringVar, filedialog
import tkinter as tk
from docx import Document
import csv


def startup():
    document = Document('testdoc.docx')


    result = [p.text for p in document.paragraphs]

    with open('nameslist.csv', newline='', encoding='utf-8') as csvfile:
        namesList = csv.reader(csvfile, delimiter=',', quotechar='|')
        
        cleanNamesList = []
        for row in namesList:
            cleanNamesList.append(row[3].strip('""'))
            cleanNamesList.append(row[4].strip('""'))
    
    clean(result,cleanNamesList, 0)


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
                #swap to create Doc instead of just string
                finishedDocument.add_paragraph(f)
                # finalText += f + '\n'
        except:
            pass
    # swap change and print to save as document
    # print(finalText)
    try:
        filename = tk.filedialog.asksaveasfilename(defaultextension=".docx")
        finishedDocument.save(filename)
        print(f"succesfully saved {filename}")
    except: 
        print("you broke something")   


# startup()

#start tkinter 
root = tk.Tk()
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
canvas = tk.Canvas(root)
canvas.grid(columnspan=6, rowspan=7)

#outer layers
top_layer= tk.Label(root)
top_layer.grid(columnspan=6, column=0, row=0)
btm_layer = tk.Label(root)
btm_layer.grid(columnspan=6, column=0, row=7)
left_layer = tk.Label(root)
left_layer.grid(rowspan=7, column=0, row=0)
right_layer = tk.Label(root)
right_layer.grid(rowspan=7, column=5, row=0)

#url entry box
text_line = tk.Entry(root, width=45)
text_line.grid(columnspan=3, column=2, row=2, sticky="w", padx=20, pady=20)
#url entry box instruction
label_text= tk.StringVar()
label_text.set("Url:")
label_info=Label(root, textvariable=label_text)
label_info.grid(column=1, row=2)

#clean doc button
create_btn = tk.Button(root, command=lambda:startup(), text="Clean and Save", height=2, width=15, pady=5)
create_btn.grid(columnspan=2, column=2, row=3)


root.mainloop()