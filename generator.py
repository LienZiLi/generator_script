# ---imports---
import docx
from docx.oxml.ns import qn
from docx2pdf import convert
import os
import pandas as pd
# ---imports---

# ---variables---
p_numbers = 3
placeholders = ["測試名", "測試編號", "測試試場"]
p_title = ["姓名", "編號", "試場"]
p_paragraph = []
# ---variables---

# ---generating functions---
def export_docx():
    return None

def export_pdf():
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
print("done")
# ---main function---