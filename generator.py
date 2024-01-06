# ---imports---
import docx
from docx2pdf import convert
import os
import pandas as pd
# ---imports---

# ---variables---
p_numbers = 3
placeholders = ["測試名", "測試編號", "測試試場"]
p_title = ["姓名", "編號", "試場"]
p_paragraph = []
save_with = "姓名"
# ---variables---

# ---generating functions---
def export(pdf):
    doc = docx.Document("test.docx")
    data = pd.read_excel("test.xlsx")
    for i in range(len(data[p_title[0]])):
        for j in range(p_numbers):
            inline = doc.paragraphs[p_paragraph[j]].runs
            for k in range(len(inline)):
                if placeholders[j] in inline[k].text:
                    text = inline[k].text.replace(placeholders[j], data[p_title[j]][i])
                    inline[k].text = text
            print(doc.paragraphs[p_paragraph[j]].text)
        docx_name = "test_" + data[save_with][i].replace(" ", "_") + ".docx"
        doc.save(docx_name)
        if pdf:
            pdf_name = "test_" + data[save_with][i].replace(" ", "_") + ".pdf"
            convert(docx_name, pdf_name)
            os.remove(docx_name)
    return None
# ---generating functions---

# ---placeholder function---
def find_words():
    doc = docx.Document("test.docx")
    p_location = [0] * p_numbers
    for i in range(len(doc.paragraphs)):
        for j in range(p_numbers):
            if doc.paragraphs[i].text.find(placeholders[j]) != -1:
                p_location[j] = i
    return p_location
# ---placeholder function---

# ---main function---
p_paragraph = find_words()
print(p_paragraph)
export(pdf = True)
print("done")
# ---main function---