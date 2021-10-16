from docx import Document

print(dir())

# добавить другие доки, найти общее между заголовками
doc = Document('1.docx')
# nlp = Russian()
text = ""

"""
for par in doc.paragraphs:
    if "адрес" in par.text.lower():
        print(par.text)
"""