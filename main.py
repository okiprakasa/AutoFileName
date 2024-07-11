import docx2txt
from tkinter import *
from tkinter import filedialog
from spire.doc import *
from spire.doc.common import *


def filename_cleaner(target):
    # remove forbidden ASCII characters and unused words for file name
    target = (target.replace("\n", " ")
              .replace("/", "").replace("\\", "")
              .replace("<", "").replace(">", "")
              .replace(":", "").replace(";", "")
              .replace("|", "").replace("?", "")
              .replace("&", "").replace("!", "")
              .replace("@", "").replace("#", "")
              .replace("$", "").replace("^", "")
              .replace("+", "").replace("-", "")
              .replace("~", "").replace("`", "")
              .replace(".", "").replace("*", "")
              .replace("direktorat", "").replace("jenderal", "")
              .replace("evaluation warning the document was created with spiredoc for python ", "")
              .replace("balai pengujian dan identifikasi barang", "bpib")
              .replace("balai laboratorium bea dan cukai", "blbc")
              .replace("kementerian keuangan", "")
              .replace("bea dan cukai", "bc")
              .replace("ii", "").replace("blbc", ""))

    # remove double space
    target = ' '.join(target.split())

    return target[:225]


def save_as_docx(target):
    # Create an object of the Document class
    document = Document()

    # Load a Word DOC file
    try:
        document.LoadFromFile(target)
    except:
        document.LoadFromFile(target, FileFormat.Docx, "iso17025")
        # document.RemoveEncryption()

    # Save the DOC file to DOCX format
    document.SaveToFile(target+"x", FileFormat.Docx)

    # Close the Document object
    document.Close()


root = Tk()
root.withdraw()
source_folder = filedialog.askdirectory()

fileName = ""
for item in os.listdir(source_folder):
    # check if an item is file or not
    item_path = os.path.join(source_folder, item)
    new_path = ''
    if os.path.isfile(item_path):
        if item.lower().endswith('.docx'):
            try:
                # Read the .docx file
                text = docx2txt.process(item_path).lower()
                fileName = filename_cleaner(text)
                new_path = os.path.join(source_folder, fileName + '.docx')
                os.rename(item_path, new_path)
            except PermissionError:
                continue
            except FileExistsError:
                os.remove(new_path)
                os.rename(item_path, new_path)
            except Exception as e:
                print(0)
                print(e)
                continue
        elif item.lower().endswith('.doc'):
            try:
                # Read the .docx file
                save_as_docx(item_path)
                text = docx2txt.process(item_path + 'x').lower()
                fileName = filename_cleaner(text)
                new_path = os.path.join(source_folder, fileName + '.docx')
                os.rename(item_path, new_path)
            except PermissionError:
                continue
            except FileExistsError:
                os.remove(new_path)
                os.rename(item_path, new_path)
            except Exception as e:
                print(1)
                print(e)
                continue
