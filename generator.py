# ---imports---
import docx
from docx2pdf import convert
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd
# ---imports---

# ---variables---
p_numbers = 3
placeholders = ["測試名", "測試編號", "測試試場"]
p_title = ["姓名", "編號", "試場"]
p_paragraph = []
save_with = "姓名"
save_name = "test"
save_folder = "./"
doc_path = ""
data_path = ""
# ---variables---

# ---generating functions---
def export(pdf):
    global p_numbers, p_paragraph, placeholders, p_title, save_with, save_name, save_folder, doc_path, data_path, output
    doc = docx.Document(doc_path)
    data = pd.read_excel(data_path)
    for i in range(len(data[p_title[0]])):
        for j in range(p_numbers):
            inline = doc.paragraphs[p_paragraph[j]].runs
            for k in range(len(inline)):
                if placeholders[j] in inline[k].text:
                    text = inline[k].text.replace(placeholders[j], data[p_title[j]][i])
                    inline[k].text = text
            print(doc.paragraphs[p_paragraph[j]].text)
        docx_name = save_folder + save_name + "_" + data[save_with][i].replace(" ", "_") + ".docx"
        doc.save(docx_name)
        if pdf:
            pdf_name = save_folder + save_name + "_" + data[save_with][i].replace(" ", "_") + ".pdf"
            convert(docx_name, pdf_name)
            os.remove(docx_name)
    return None
# ---generating functions---

# ---placeholder function---
def find_words():
    doc = docx.Document(doc_path)
    p_location = [0] * p_numbers
    for i in range(len(doc.paragraphs)):
        for j in range(p_numbers):
            if doc.paragraphs[i].text.find(placeholders[j]) != -1:
                p_location[j] = i
    return p_location
# ---placeholder function---

# ---choose doc function---
def select_doc():
    global doc_path, show_doc
    filetypes = (('document files', '*.docx'),)
    doc_path = fd.askopenfilename(title='Choose a file', initialdir='./', filetypes=filetypes)
    show_doc.set(doc_path.split("/")[-1])
    return None
# ---choose doc function---
    
# ---choose data function---
def select_data():
    global data_path, show_data
    filetypes = (('Excel files', '*.xlsx'),)
    data_path = fd.askopenfilename(title='Choose a file', initialdir='./', filetypes=filetypes)
    show_data.set(data_path.split("/")[-1])
    return None
# ---choose data function---
    
# ---set word function---
def set_words():
    global p_numbers, placeholders, p_title, output
    p_numbers = int(i3_1.get())
    placeholders = i3_2.get().split(",")
    p_title = i3_3.get().split(",")
    if (p_numbers == len(placeholders) and p_numbers == p_title):
        return None
    else:
        err = "error: 3. 替代數量:" + p_numbers + "，替代文字數量:" + len(placeholders) + "，標題名稱數量:" + len(p_title) + "。三者不相符。"
        output.set(err)
        return None
# ---set word function---
    
# ---set save function---
def set_save():
    global save_with, save_name, output
    save_name = i4_1.get()
    save_with = i4_2.get()
    if save_with in p_title:
        return None
    else:
        err = "error: 4. 命名標題不存在於標題名稱中。"
        output.set(err)
        return None
# ---set save function---

# ---set save folder function---
def select_folder():
    global save_folder, show_folder
    save_folder = fd.askdirectory(title="select a directory", initialdir='./')
    show_folder.set(save_folder)
# ---set save folder function---
    
# ---execute function---
def execute():
    global p_numbers, p_paragraph, placeholders, p_title, save_with, save_name, save_folder, doc_path, data_path, typeVar, output
    #validity check
    if os.path.isfile(doc_path) == False:
        err = "error: 1. 範本檔案(.docx)不存在，請重新選擇。"
        output.set(err)
        return None
    if os.path.isfile(data_path) == False:
        err = "error: 2. 資料檔案(.xlsx)不存在，請重新選擇。"
        output.set(err)
        return None
    if (p_numbers == len(placeholders) and p_numbers == p_title) == False:
        err = "error: 3. 替代數量:" + p_numbers + "，替代文字數量:" + len(placeholders) + "，標題名稱數量:" + len(p_title) + "。三者不相符。"
        output.set(err)
        return None
    if save_with in p_title == False:
        err = "error: 4. 命名標題不存在於標題名稱中。"
        output.set(err)
        return None
    #end of validity check

    p_paragraph = find_words()
    for i in range(len(p_paragraph)):
        if p_paragraph[i] == 0:
            err = "範例檔案中該關鍵字不存在:" + placeholders[i]
            output.set(err)
            return None
    export(typeVar)
    output.set("產生完成。")
    return None
# ---execute function---

# ---main function---
window = tk.Tk()
window.title("通知／證明產生器")
window.minsize(width=300, height=600)
window.resizable(False, False)

p1 = tk.Label(text="1. 選取範本檔案(.docx):", font=("Microsoft JhengHei UI", 12))
p1.pack(anchor="w")
b1 = tk.Button(text="選取檔案", command=select_doc)
b1.pack(anchor="w")
show_doc = tk.StringVar()
p1_1 = tk.Label(textvariable=show_doc, font=("Microsoft JhengHei UI", 10))
p1_1.place(x=60, y=28)
pm_1 = tk.Label(text="", font=(10))
pm_1.pack()

p2 = tk.Label(text="2. 選取資料檔案(.xlsx):", font=("Microsoft JhengHei UI", 12))
p2.pack(anchor="w")
b2 = tk.Button(text="選取檔案", command=select_data)
b2.pack(anchor="w")
show_data = tk.StringVar()
p1_1 = tk.Label(textvariable=show_data, font=("Microsoft JhengHei UI", 10))
p1_1.place(x=60, y=106)
pm_2 = tk.Label(text="", font=(10))
pm_2.pack()

p3 = tk.Label(text="3. 設定關鍵字:", font=("Microsoft JhengHei UI", 12))
p3.pack(anchor="w")
p3_1 = tk.Label(text="替代數量:", font=("Microsoft JhengHei UI", 10))
p3_1.pack(anchor="w")
i3_1 = tk.Entry()
i3_1.place(x=60, y=185)
p3_2 = tk.Label(text="替代文字:", font=("Microsoft JhengHei UI", 10))
p3_2.pack(anchor="w")
i3_2 = tk.Entry()
i3_2.place(x=60, y=208)
p3_3 = tk.Label(text="標題名稱:", font=("Microsoft JhengHei UI", 10))
p3_3.pack(anchor="w")
i3_3 = tk.Entry()
i3_3.place(x=60, y=231)
b3 = tk.Button(text="儲存", command=set_words)
b3.pack(anchor="w")
pm_3 = tk.Label(text="", font=(10))
pm_3.pack()

p4 = tk.Label(text="4. 設定輸出檔案命名:", font=("Microsoft JhengHei UI", 12))
p4.pack(anchor="w")
p4_1 = tk.Label(text="檔案名稱:", font=("Microsoft JhengHei UI", 10))
p4_1.pack(anchor="w")
i4_1 = tk.Entry()
i4_1.place(x=60, y=332)
p4_2 = tk.Label(text="命名標題:", font=("Microsoft JhengHei UI", 10))
p4_2.pack(anchor="w")
i4_2 = tk.Entry()
i4_2.place(x=60, y=355)
b4_1 = tk.Button(text="選取儲存資料夾", command=select_folder)
b4_1.pack(anchor="w")
show_folder = tk.StringVar()
p4_3 = tk.Label(textvariable=show_folder, font=("Microsoft JhengHei UI", 10))
p4_3.place(x=100, y=378)
b4_2 = tk.Button(text="儲存", command=set_save)
b4_2.pack(anchor="w")
pm_4 = tk.Label(text="", font=(10))
pm_4.pack()

p5 = tk.Label(text="5. 選擇輸出檔案類型:", font=("Microsoft JhengHei UI", 12))
p5.pack(anchor="w")
typeVar = tk.BooleanVar()
radio1 = tk.Radiobutton(text=".docx檔", variable=typeVar, value=False)
radio2 = tk.Radiobutton(text=".pdf檔", variable=typeVar, value=True)
radio1.place(x=0, y=480)
radio2.place(x=80, y=480)
pm_5 = tk.Label(text="\n", font=(10))
pm_5.pack()

b6 = tk.Button(text="開始執行", command=execute)
b6.pack()
output = tk.StringVar()
p6 = tk.Label(textvariable=output, font=("Microsoft JhengHei UI", 10))
p6.pack(anchor="w")

window.mainloop()
# ---main function---