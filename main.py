from docx.api import Document
import os

source_folder = r'C:\Users\matra\PycharmProjects\AutoFileName'

for item in os.listdir(source_folder):
    # check if an item is file or not
    if os.path.isfile(os.path.join(source_folder, item)):
        if item.endswith('.docx'):
            try:
                # Open the .docx file
                doc = Document(item)

                # Create an empty list to store all paragraphs
                paraText = ""
                tableText = ""

                # Iterate through each paragraph and append its text to the list
                for p in doc.paragraphs:
                    paraText = paraText + p.text

                for table in doc.tables:
                    for row in table.rows:
                        tableText = tableText + " ".join([cell.text for cell in row.cells])

                # remove forbidden ASCII characters for file name
                fileName = ((tableText + paraText).replace("\n", " ")
                            .replace("/", "").replace("\\", "")
                            .replace("<", "").replace(">", "")
                            .replace(":", "").replace(";", "")
                            .replace("|", "").replace("?", "")
                            .replace("&", "").replace("!", "")
                            .replace("@", "").replace("#", "")
                            .replace("$", "").replace("^", "")
                            .replace("+", "").replace("-", "")
                            .replace("~", "").replace("`", "")
                            .replace(".", "").replace("*", ""))

                # remove double space
                fileName = ' '.join(fileName.split())

                mid = int(len(fileName) / 2)
                q1 = int(len(fileName) / 4)
                q3 = int(len(fileName) * 3 / 4)

                if len(fileName) > 200:
                    fileName = fileName[:200]
                print(fileName)
                os.rename(
                    os.path.join(source_folder, item),
                    os.path.join(source_folder, fileName + '.docx')
                )
            except PermissionError:
                continue
            except Exception as e:
                raise Exception(e)
