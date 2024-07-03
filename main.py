from docx.api import Document

# Open the .docx file
doc = Document('test.docx')

# Create an empty list to store all paragraphs
paraText = ""
tableText = ""
# Iterate through each paragraph and append its text to the list
for p in doc.paragraphs:
    paraText = paraText + p.text

for table in doc.tables:
    for row in table.rows:
        tableText = tableText + " ".join([cell.text for cell in row.cells])

fileName = (tableText + paraText).replace("\n", " ")
# remove double space
fileName = ' '.join(fileName.split())
mid = int(len(fileName)/2)
q1 = int(len(fileName)/4)
q3 = int(len(fileName)*3/4)
if len(fileName) > 190:
    fileName = fileName[0:60] + fileName[q1:q1+30] + fileName[mid:mid+35] + fileName[q3:q3+30] + fileName[-35:]

print(fileName)
