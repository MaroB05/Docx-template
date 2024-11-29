from spire.doc import *
from spire.doc.common import *


def loadFile():
    while True:
        try:
            fileName = input("الموضوع:").strip()
            doc = Document()
            doc.LoadFromFile(f"data/{fileName}.docx")
            return doc, fileName
        except:
            print("Error! File not found!")



if __name__ == "__main__":
    doc, fileName = loadFile()
    
    pattern = Regex("\*.*?\*")
    fields = doc.FindAllPattern(pattern)
    fields = [ field.SelectedText.strip("*") for field in fields ]
    
    name = ""

    for field in fields:
        info = input(f"{field}:")
        doc.Replace(f"*{field}*", info, False, False)
        if field == "اسم صاحب/ة الطلب":
            name = info.split()[0] + " "
    
    doc.SaveToFile(f"output/{name + fileName}.docx", FileFormat.Docx2016)
    doc.Close()