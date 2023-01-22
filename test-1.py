import re
import docx

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

text = getText('Resolution1229516001_1106.docx')
pattern = '^CA\d{2} \d{2} \d{2,3}'
result = re.match(pattern, text)
print(getText('Resolution1229516001_1106.doc'))