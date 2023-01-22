import re
import docx
import sys
import os
import win32com.client
import PyPDF2

wdFormatPDF = 17

filename = 'Resolution1229516004_1108.doc'
filenamePDF = filename.split('.')[0]
path = os.getcwd()
in_file = f"{path}\{filename}"
out_file = f"{path}\{filenamePDF}"

word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()




reader = PyPDF2.PdfReader(f"{out_file}.pdf")

texte =reader.pages[0].extract_text()

num_resol = re.search(r"CA[\d]{2}\s[\d]{2}\s[\d]{2,4}", texte)
date = re.search(r"(?:lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\s[\d]{1,2}\s(?:janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s[\d]{4}", texte)

print('La résolution n°', num_resol.group())

print('en date du', date.group())