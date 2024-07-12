from tkinter import *
from tkinter import filedialog

from spire.doc import *
from spire.doc.common import *

root = Tk()
root.withdraw()
source_folder = filedialog.askdirectory()
sorted_folder = r"C:\Users\matra\Documents\sort"


def filename_cleaner(target):
    # remove forbidden ASCII characters and unused words for file name
    target = (target.replace("\n", " ")
              .replace("/", "").replace("\\", "")
              .replace("\"", "").replace("\'", "")
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
              .replace("evaluation warning the document was created with spiredoc for python", "")
              .replace("balai pengujian dan identifikasi barang", "bpib")
              .replace("balai laboratorium", "")
              .replace("kementerian keuangan", "")
              .replace("bea dan cukai", "")
              .replace("bpib tipe b", "")
              .replace("republik indonesia", "")
              .replace("republik indonesia", "")
              .replace("(031) 3286492", "")
              .replace("(031) 3284154", "")
              .replace("kantor wilayah", "")
              .replace("jawa timur i", "")
              .replace("surabaya", "").replace("60165", "")
              .replace("kelas", "").replace("djbc", "")
              .replace("jalan perak timur no 498", "")
              .replace("telepon", "").replace("faksimile", "")
              .replace("surat elektronik", "").replace("laman", "")
              .replace("bpibsurabayayahoocom", "").replace("bpibyahoocom", "")
              .replace("wwwbeacukaigoid", "")
              .replace("pusat kontak layanan", "")
              .replace("1500225", "").replace("surel", "")
              .replace("bcbpibcustomsgoid", "").replace("nota dinas", "nd")
              .replace("ii", "").replace("yth", ""))

    # remove double space
    target = ' '.join(target.split())
    return target[:225]


def read_ms_word(target):
    # Create an object of the Document class
    document = Document()
    x = ""
    try:
        if target.lower().endswith(".docx"):
            document.LoadFromFile(target, FileFormat.Docx)
            x = "x"
        else:
            document.LoadFromFile(target, FileFormat.Doc)
    except:
        if target.lower().endswith(".docx"):
            document.LoadFromFile(target, FileFormat.Docx, "iso17025")
            x = "x"
        else:
            document.LoadFromFile(target, FileFormat.Doc, "iso17025")
        document.RemoveEncryption()
    document.Close()
    return os.path.join(sorted_folder, filename_cleaner(document.GetText().lower()) + '.doc' + x)


for item in os.listdir(source_folder):
    # check if an item is file or not
    item_path = os.path.join(source_folder, item)
    if os.path.isfile(item_path):
        try:
            updated_path = read_ms_word(item_path)
            try:
                os.rename(item_path, updated_path)
            except FileExistsError:
                os.remove(updated_path)
                os.rename(item_path, updated_path)
            os.remove(item_path)
        except Exception as e:
            print(item_path, e)
            continue
