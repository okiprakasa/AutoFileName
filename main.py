from tkinter import *
from tkinter import filedialog
from spire.doc import *
from spire.doc.common import *
import time
import pandas as pd

sorted_folder = r"C:\1"
sorted_ksbuki = r"C:\KSBUKI"
sorted_kabalai = r"C:\Kabalai"
sorted_kspe = r"C:\KSPE"
sorted_kstl = r"C:\KSTL"
sorted_satpel = r"C:\Satpel"
root = Tk()
root.withdraw()
source_folder = filedialog.askdirectory()
start = time.perf_counter()
counter = 0


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
              .replace("balai laboratorium bea dan cukai", "blbc")
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
              .replace("seksi program dan evaluasi", "spe")
              .replace("seksi teknis laboratorium", "stl")
              .replace("subbagian umum dan kepatuhan internal", "sbuki")
              .replace("sub bagian umum dan kepatuhan internal", "sbuki")
              .replace("pusat kontak layanan", "")
              .replace("1500225", "").replace("surel", "")
              .replace("bcbpibcustomsgoid", "").replace("nota dinas", "nd").replace(" NaN ", "")
              .replace("tanggal", "tgl").replace("dengan", "dgn").replace("  ", " ")
              .replace("ii", "").replace("yth", "").replace("nomor", "no"))

    # remove double space
    target = ' '.join(target.split())
    return target[:215]


def read_ms_word(target):
    # Create an object of the Document class
    document = Document()
    x = ""
    # noinspection PyBroadException
    try:
        if target.lower().endswith(".docx"):
            document.LoadFromFile(target, FileFormat.Docx)
            x = "x"
        else:
            document.LoadFromFile(target, FileFormat.Doc)
    except Exception:
        if target.lower().endswith(".docx"):
            document.LoadFromFile(target, FileFormat.Docx, "iso17025")
            x = "x"
        else:
            document.LoadFromFile(target, FileFormat.Doc, "iso17025")
        document.RemoveEncryption()
    words = document.GetText().lower()
    document.Close()
    return filename_cleaner(words) + '.doc' + x


def read_ms_excel(target):
    x = ""
    # noinspection PyBroadException
    if target.lower().endswith(".xlsx"):
        x = "x"
    cells = pd.read_excel(target)
    cells.to_string('my_file.txt')
    with open('my_file.txt') as f:
        names = f.read()

    # Print the names
    print(type(names))
    return filename_cleaner(names) + '.xls' + x


for item in os.listdir(source_folder):
    # check if an item is file or not
    item_path = os.path.join(source_folder, item)
    if os.path.isfile(item_path):
        counter += 1
        try:
            # updated_path = os.path.join(source_folder, item.replace("jl tanjung perak timur no 498  ", "").replace(
            # ", ", "") .replace("blbc , ", ""))
            # filename = read_ms_word(item_path)
            filename = read_ms_excel(item_path)
            updated_path = os.path.join(sorted_folder, filename)
            # item_path_new = item_path.replace("blbc , ", "")
            # if "dari kepala spe sifat" in item_path:
            #     updated_path = os.path.join(sorted_kspe, item)
            #     os.rename(item_path, updated_path)
            #     print(item_path)
            # elif "dari kepala sbuki sifat" in item_path:
            #     updated_path = os.path.join(sorted_ksbuki, item)
            #     os.rename(item_path, updated_path)
            # elif "dari kepala stl sifat" in item_path:
            #     updated_path = os.path.join(sorted_kstl, item)
            #     os.rename(item_path, updated_path)
            # elif "dari kepala blbc sifat" in item_path:
            #     updated_path = os.path.join(sorted_kabalai, item)
            #     os.rename(item_path, updated_path)
            # elif "dari penyelia satuan pelayanan" in item_path:
            #     updated_path = os.path.join(sorted_satpel, item)
            #     os.rename(item_path, updated_path)
            try:
                os.rename(item_path, updated_path)
            except FileExistsError:
                os.remove(updated_path)
                os.rename(item_path, updated_path)
        except Exception as e:
            print(item_path, e)
            continue

end = time.perf_counter()
print(f"Processed files: {counter} files")
print(f"Processing time: {end - start:0.4f} seconds")
