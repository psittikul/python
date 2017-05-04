import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtGui import QIcon

class Example(QWidget):
    
    def __init__(self):
        super().__init__()
        
        self.initUI()
        
        
    def initUI(self):

        btn = QPushButton('Upload file to prep for mail merge', self)
        btn.resize(btn.sizeHint())
        btn.move(50,50)
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
