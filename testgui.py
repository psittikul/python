import sys
import tkinter
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtGui import QIcon
from flask import Flask, render_template, request
from werkzeug import secure_filename
from tkinter import Tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
UPLOAD_FOLDER = 'C:/Users/Bridget Velez/'
ALLOWED_EXTENSIONS = 'xls'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


class Example(QWidget):

    def getFile(self):
        print ("getFile called!")
        Tk().withdraw()
        filename = askopenfilename()
        print(filename)
        data_wb = openpyxl.load_workbook(filename)
    
    def __init__(self):
        super().__init__()
        
        self.initUI()        
        
    def initUI(self):

        btn = QPushButton('Upload file to prep for mail merge', self)
        btn.resize(btn.sizeHint())
        btn.move(50,50)
        btn.clicked.connect(self.getFile)
        updateBtn = QPushButton('Upload file to update contact information', self)
        updateBtn.resize(updateBtn.sizeHint())
        updateBtn.move(350, 50)
        self.setGeometry(300, 300, 800, 500)
        self.setWindowTitle('Orio Email Automation Program')
        self.setWindowIcon(QIcon('logo.png'))        
    
        self.show()


            

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
