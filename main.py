from PyQt5 import QtGui
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMainWindow,QApplication,QFileDialog,QMessageBox
from PyQt5.uic import loadUi
from os import path
from docx import Document
import pytesseract as pt
from sys import argv
pt.pytesseract.tesseract_cmd=r"Tesseract-OCR/tesseract"

# the Thread class for gettig the images 
class Thread(QThread):
    # the signal in this case is list of all images paths
    signal= pyqtSignal(str)
    def __init__(self,paths,names):
        super(Thread,self).__init__()
        self.__paths =paths
        self.__imgnames = names
        self.text_content = ""
    def run(self):
        try:
            for i,value in enumerate(self.__paths):
                content= pt.image_to_string(value)
                self.text_content+=f" Image-[{i+1}] "+ self.__imgnames[i]+"\n"
                self.text_content+="-"*30+"\n"
                self.text_content+=content
                self.text_content+="-"*60+'\n'
            self.signal.emit(self.text_content)
        except:
            QMessageBox.warning(self,"Error","There is some error happend, Try agin.")
    
class MainApp(QMainWindow):
    def __init__(self, parent=None):
        super(MainApp,self).__init__(parent)
        QMainWindow.__init__(self)
        # initiate the gui 
        loadUi("Ui Files/style.ui",self)
        # intiate the ui setting
        self.ui_setting()
        self.buttons()
        self.setup_icons()
        self.images=None
        self.text_content =""
    def ui_setting(self):
        self.setWindowTitle("Text Extractor")
        self.setWindowIcon(QtGui.QIcon("icons/icon.png"))
    def setup_icons(self):
         self.actionSave_as_2.setIcon(QtGui.QIcon("icons/save.png"))
         self.actionSelect_Image.setIcon(QtGui.QIcon("icons/select.png"))
         self.pushButton.setIcon(QtGui.QIcon("icons/get.png"))
    def buttons(self):
        self.actionSelect_Image.triggered.connect(self.select_images)
        self.actionSave_as_2.triggered.connect(self.savefile)
        self.pushButton.clicked.connect(self.get_text)
    def select_images(self):
        try:
            directory=""
            fileter_mast= "PNG/JPG/JIF files (*.png *.jpg *.gif)"
            self.images= QFileDialog.getOpenFileNames(None,"Select images",directory=directory,filter=fileter_mast)
            self.images_paths= self.images[0]
            self.images_names= [x.split("/")[-1] for x in self.images_paths]
            self.lcdNumber.display(len(self.images_paths))
        except:
            QMessageBox.warning(self,"Error","There is some error happend, try select again")
    def get_text(self):
        try:
            if self.images:
                self.th = Thread(paths=self.images_paths,names=self.images_names) 
                self.th.start()
                # set the connection 
                self.th.signal.connect(self.textEdit.setText)
            else:
               QMessageBox.warning(self,"Warning"," Select the images first.") 
        except:
            QMessageBox.warning(self,"Error","There is some error happend, Try agin.")
        
    def savefile(self):
        try:
            if self.textEdit.toPlainText() !="":
                options = QFileDialog.Options()
                # options |= QFileDialog.DontUseNativeDialog
                file_name=QFileDialog.getSaveFileName(self, 'Save File',"./","Text Files(*.txt);; Word files (*.docx)",options=options)
                extension = file_name[0].split("/")[-1].split(".")[-1]
                # get valid charcters
                def valid_xml_char_ordinal(c):
                    codepoint = ord(c)
                    # conditions ordered by presumed frequency
                    return (
                        0x20 <= codepoint <= 0xD7FF or
                        codepoint in (0x9, 0xA, 0xD) or
                        0xE000 <= codepoint <= 0xFFFD or
                        0x10000 <= codepoint <= 0x10FFFF
                        )
                text=" ".join(c for c in self.textEdit.toPlainText() if valid_xml_char_ordinal(c))
                if extension == "docx":
                    doc= Document()
                    doc.add_paragraph(text)
                    doc.save(file_name[0]) 
                    QMessageBox.information(self,"Success"," The file made.")
                elif extension=="txt":
                    with open(file_name[0],"w") as file:
                        file.write(text)
                    QMessageBox.information(self,"Success"," The file made.")
            else:
                QMessageBox.warning(self,"Warning"," Extract the text first.") 
        except:
            QMessageBox.warning(self,"Error"," There was an error saving the file.")
def main():
    app =QApplication(argv)
    window = MainApp()
    window.show()
    app.exec_()
if __name__=="__main__":
    main()