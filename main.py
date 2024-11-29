# import os 
# import sys
# from PyQt6.QtWidgets import QApplication, QLabel, QWidget, QComboBox

# subjects = [ sub.strip(".docx") for sub in os.listdir("data/") ]

# app = QApplication([])

# window = QWidget()

# lst = QComboBox(parent=window)
# lst.addItems(subjects)

# window.show()
# sys.exit(app.exec())


import os
import sys
from spire.doc import *
from functools import partial
from spire.doc.common import *
from PyQt6.QtWidgets import QApplication, QLabel, QWidget, QComboBox, QPushButton, QLineEdit

class Graphics:
    

    def __init__(self, window):
        self.X = 10
        self.Y = 40
        self.doc = Document()
        self.elements = []
        self.window = window
        
    def loadFile(self, fileName):
        try:
            doc = Document()
            doc.LoadFromFile(f"data/{fileName}.docx")
            self.doc = doc
            return fileName
        except:
            print(fileName)
            print("Error! File not found!")

    def addButton(self, text, func, x, y):
        choose = QPushButton(parent=self.window, text=text)
        choose.clicked.connect(func)
        choose.move(x, y)
        return choose

    def addList(self, elmnts, x, y):
        combo = QComboBox(parent=self.window)
        combo.addItems(elmnts)
        combo.move(x, y)
        return combo

    # doc, fileName = loadFile()
    def updateInfo(self):
        for label , box in self.elements[:-2]:
            self.doc.Replace(f"*{label.text()}*", box.text(), False, False)
        self.doc.SaveToFile(f"output/{self.elements[-1][0].text()}.docx", FileFormat.Docx2016)


    def updateGUI(self):
        pattern = Regex("\*.*?\*")
        fields = self.doc.FindAllPattern(pattern)
        fields = [ field.SelectedText.strip("*") for field in fields ]
        elements = []

        for i in range(len(fields)):
            box = QLineEdit(parent=self.window)
            box.move(self.X, self.Y + 30 * i)
            box.show()
            label = QLabel(parent=self.window, text=fields[i])
            label.move(self.X + box.width() + 10, self.Y + 30 * i)
            label.show()
            elements.append( (label, box) )


        outputName = QLineEdit(parent=self.window)

        button = self.addButton("Submit", self.updateInfo, self.X, self.Y + 30 * (i + 1) )
        button.show()

        outputName.move(self.X + button.width() + 30, self.Y + 30 * (i + 1) )
        outputName.resize(200, outputName.height() - 5)
        outputName.show()

        elements.append( (button, button) )
        elements.append( (outputName, outputName) )

        self.elements = elements

    def loadAndUpdate(self, comboBox):
        self.loadFile(comboBox.currentText())
        self.updateGUI()
        

name = ""

subjects = [""] + [ sub.strip(".docx") for sub in os.listdir("data/") ]

app = QApplication([])

window = QWidget()
window.setWindowTitle("PyQt App")
window.setGeometry(100, 100, 400, 400)

widget = Graphics(window)

subjects = widget.addList(subjects, 150, 10)
subjects.resize(200, subjects.height())
button = widget.addButton("choose", partial( widget.loadAndUpdate, subjects), 50, 10)

window.show()

sys.exit(app.exec())