import sys
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

def window():
   app = QApplication(sys.argv)
   w = QWidget()
   b = QPushButton(w)
   b.setText("Show message!")

   b.move(50,50)
   b.clicked.connect(showdialog)
   w.setWindowTitle("PyQt Dialog demo")
   w.show()
   sys.exit(app.exec_())
	
def showdialog():
   msg = QMessageBox()
   msg.setIcon(QMessageBox.Question)

   msg.setText("Are the fields matched to the correct columns?")
   msg.setInformativeText("Click 'Show Details' to check the detected merge fields. If the pairings below are incorrect, you can manually enter the correct column for each merge field.")
   msg.setWindowTitle("Confirm Fields to Merge")
   msg.setDetailedText("A: OSC Code | B: DBA")
   msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
   msg.buttonClicked.connect(msgbtn)
	
   retval = msg.exec_()
   print ("value of pressed message box button:", retval)
	
def msgbtn(i):
   print ("Button pressed is:",i.text())
	
if __name__ == '__main__': 
   window()
