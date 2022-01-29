from email import message
from tkinter import Label, messagebox, Scale, StringVar, filedialog
import tkinter as tk
from docx import Document
import csv


def choose_file():
    doc_name = tk.filedialog.askopenfilename()
    text_line_var.set(doc_name) 

def startup():
    try:
        doc_name = text_line.get()
        document = Document(doc_name)
    except: 
        tk.messagebox.showerror(title="Wrong Format", message="Please choose .docx")
    
    prepData(document)


def choose_csv_list():
    csv_name = tk.filedialog.askopenfilename()
    csv_line_var.set(csv_name)


# prep CSV Data and call clean, receives document from startup()
def prepData(document):
   
    result = [p.text for p in document.paragraphs]

    try:
        csv_name = csv_line_var.get()
        with open(csv_name, newline='', encoding='utf-8') as csvfile:
            names_list = csv.reader(csvfile, delimiter=',', quotechar='|')
            clean_names_list = []
            for row in names_list:
                clean_names_list.append(row[3].strip('""'))
                clean_names_list.append(row[4].strip('""'))
        
        clean(result,clean_names_list, 0)
    except:
        tk.messagebox.showerror(title="Wrong Format", message="Please choose .csv")

#clean document function, compare nameslist to text, do it twice to catch everything
def clean(text, clean_names_list, run):
    cleaned = Document()
    if run <2:
        for par in text:
            if any(name in par for name in clean_names_list):
                for name in clean_names_list:
                    if par.find(name) != -1:
                        new_sentence = par.replace(name, name[0]+'.')
                        cleaned.add_paragraph(new_sentence)
            else:
                cleaned.add_paragraph(par)
        # print(run)
        cleaned_result = [p.text for p in cleaned.paragraphs]
        # print(f"""run: {run}
        # cleanedResult: {cleanedResult}""")
        run += 1
        clean(cleaned_result, clean_names_list, run)
    else:
        final_result = text
        end(final_result)


#clean duplicate entries from list due to first- lastname idx issue, create doc and save doc
def end(final):
    finished_document = Document()
    # final_text = ""
    for idx, f in enumerate(final):
        try:
            if final[idx] == final[idx+1]:
                pass
            else:
                #swap to create Doc instead of just string
                finished_document.add_paragraph(f)
                # finalText += f + '\n'
        except:
            pass
    # swap change and print to save as document
    # print(finalText)
    try:
        filename = tk.filedialog.asksaveasfilename(defaultextension=".docx")
        finished_document.save(filename)
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


#CSV path box
csv_line_var = tk.StringVar()
csv_line = tk.Entry(root, width=45, textvariable=csv_line_var)
csv_line.grid(columnspan=2, column=2, row=2, sticky="w", padx=20, pady=20)
#CSV path box instruction
csv_text= tk.StringVar()
csv_text.set("Datensatz:")
csv_info=Label(root, textvariable=csv_text)
csv_info.grid(column=1, row=2)

#choose csv button
csv_btn = tk.Button(root, command=lambda:choose_csv_list(), text="Choose CSV", height=2, width=15, pady=5)
csv_btn.grid(columnspan=2, column=5, row=2)

#document path box
text_line_var = tk.StringVar()
text_line = tk.Entry(root, width=45, textvariable=text_line_var)
text_line.grid(columnspan=2, column=2, row=4, sticky="w", padx=20, pady=20)
#document path box instruction
label_text= tk.StringVar()
label_text.set("Dokument:")
label_info=Label(root, textvariable=label_text)
label_info.grid(column=1, row=4)

#choose docx button
docx_btn = tk.Button(root, command=lambda:choose_file(), text="Choose Docx", height=2, width=15, pady=5)
docx_btn.grid(columnspan=2, column=5, row=4)

#clean doc button
clean_btn = tk.Button(root, command=lambda:startup(), text="Clean and Save", height=2, width=15, pady=5)
clean_btn.grid(columnspan=2, column=2, row=6)


root.mainloop()