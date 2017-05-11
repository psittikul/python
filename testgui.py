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
from collections import defaultdict
UPLOAD_FOLDER = 'C:/Users/Bridget Velez/'
ALLOWED_EXTENSIONS = set({'xls', 'xlsx'})
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


class Example(QWidget):
    
    #Function to prompt user to upload the file with the updated objectives, and then collect and store the relevant data in an array, which will later be prepped to mail merge.
    def getFile(self):
        print("Getting file")
        Tk().withdraw()
        filename = askopenfilename()
        data_wb = openpyxl.load_workbook(filename)

        #Go to the correct worksheet
        sheet = data_wb.active
        objectives_update = {}
        code_list = []

        #Cycle through each row. If the row is an OSC, gather this data to later be stored in the array "objectives_update"
        print ("Entering for loop to gather data")
        for row in range(25, 422):
            print(sheet['A' + str(row)].value)
            if str(sheet['B' + str(row)].value) == "OSC":
                print("This is an OSC")
                code = str(sheet['A' + str(row)].value)
                #Somehow this variable will have to accommodate the changing columns with Aero status changes between quarters
                status = str(sheet['G' + str(row)].value)
                if status != "#N/A" and not "-" in status and not "-" in code:
                    reward = int(sheet['H' + str(row)].value)
                    DBA = str(sheet['J' + str(row)].value)
                    objective = int(sheet['L' + str(row)].value)
                    MTD = int(sheet['M' + str(row)].value)
                    percent = int(sheet['N' + str(row)].value)
                    purchasesToGo = int(sheet['Q' + str(row)].value)
                    
                    #Store our values in our dictionary now
                    code_list.append(code)
                    data = [code, status, reward, DBA, objective, MTD, percent, purchasesToGo]
                    objectives_update[code] = data
                    
                else:
                    continue

            else:
                continue
        print("Ended for loop")
        print("Writing results")
        print(len(objectives_update))
        merge_wb = openpyxl.Workbook()
        merge_sheet = merge_wb.active
        #This would create a new MailMerge.xlsx file everytime...is there a way we can just update it?
        merge_sheet.title = "Mail Merge"
        merge_wb.save('MailMerge.xlsx')
        print("Adding column headers")
        merge_sheet['A1'].value = "OSC Code"
        merge_sheet['B1'].value = "OSC Name"
        merge_sheet['C1'].value = "Contact First Name"
        merge_sheet['D1'].value = "Contact Last Name"
        merge_sheet['E1'].value = "Contact Email"
        merge_sheet['F1'].value = "Aero Status"
        merge_sheet['G1'].value = "Objective"
        merge_sheet['H1'].value = "MTD"
        merge_sheet['I1'].value = "% of Goal"
        merge_sheet['J1'].value = "Purchases to Go"
        merge_sheet['K1'].value = "Reward"
        merge_wb.save('MailMerge.xlsx')
        pop_row = 2
        print(len(code_list))
        for i in range(0, len(code_list)):
            current_code = str(code_list.pop(i))
            print(current_code)
            print(str(objectives_update[current_code][0]))
            merge_sheet['A' + str(pop_row)].value = str(objectives_update[current_code][0])
            merge_wb.save('MailMerge.xlsx')
            merge_sheet['B' + str(i)].value = DBA
            #merge_sheet['F' + str(i)].value = status
            #merge_sheet['G' + str(i)].value = objective
            #merge_sheet['H' + str(i)].value = MTD
            #merge_sheet['I' + str(i)].value = percent
            #merge_sheet['J' + str(i)].value = purchasesToGo
            #merge_sheet['K' + str(i)].value = reward
            pop_row += 1
        
        print("Finished copying over values")    
    def __init__(self):
        super().__init__()
        
        self.initUI()        
        
    def initUI(self):

        btn = QPushButton('Upload sales objective file to prep for mail merge', self)
        btn.resize(btn.sizeHint())
        btn.move(50,50)
        #When this button is clicked, function getFile will be called
        btn.clicked.connect(self.getFile)
        rewardsBtn = QPushButton('Upload end of month file to calculate rewards', self)
        rewardsBtn.resize(rewardsBtn.sizeHint())
        rewardsBtn.move(350, 50)
        updateBtn = QPushButton('Upload file to update contact information', self)
        updateBtn.resize(updateBtn.sizeHint())
        updateBtn.move(350, 10)
        self.setGeometry(300, 300, 800, 500)
        self.setWindowTitle('Orio Email Automation Program')
        self.setWindowIcon(QIcon('logo.png'))        
    
        self.show()


            

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
