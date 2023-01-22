import PyPDF2
import re

# open the pdf file
reader = PyPDF2.PdfReader("Resolution1229516001_1106.pdf")

# get number of pages
num_pages = len(reader.pages)

# define key terms
pattern = '^CA\d{2} \d{2} \d{2,3}'


# extract text and do the search
for page in reader.pages:
    text = page.extract_text() 
    # print(text)
    res_search = re.search(pattern, text)
    print(res_search)