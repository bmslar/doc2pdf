import os
import win32com.client
word = win32com.client.Dispatch('Word.Application')

folder = "fwdstatement"
files = os.listdir(folder)
word_files = [f for f in files if f.endswith((".doc", ".docx"))]
for word_file in word_files:
    new_name = word_file.replace(".doc", ".pdf")
    in_file =(os.getcwd() + '//fwdstatement//'+ word_file)
    new_file =(os.getcwd() + '//pdf//' + new_name)
    doc = word.Documents.Open(in_file)
    doc.SaveAs(new_file, FileFormat = 17)
    doc.Close()
word.Quit()
