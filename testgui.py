import sys
import os
import errno
import tkinter
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from flask import Flask, render_template, request
from werkzeug import secure_filename
from tkinter import Tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from collections import defaultdict
UPLOAD_FOLDER = os.getcwd()
ALLOWED_EXTENSIONS = set({'xls', 'xlsx'})
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


class Example(QWidget):
    
    # --- Function to prompt user to upload the file with the updated objectives, and then collect and store the relevant data in an array, which will later be prepped to mail merge. ---
    # !! Add code in to make sure it's an accepted file type later
##    def allowed_file(filename):
##        return '.' in filename and \
##               filename.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS
        
    
    def getFile(self):
        def formatCurrency(amt):
            if len(amt) > 3:
                hundreds = amt[len(amt)-3:len(amt)+1]
                thousands = amt[:len(amt)-3]
                return "$"+thousands+","+hundreds
            else:
                return "$"+amt
    
        print("Getting file")
        Tk().withdraw()
        filename = askopenfilename()
        print(filename)
        # Only continue if an "ONA - Sales by Channel Assignment..." file is being uploaded. If not, prompt again.
        print("Uploading to "+ str(UPLOAD_FOLDER)) 
        data_wb = openpyxl.load_workbook(filename)
        print("Opening this uploaded file")
        # Go to the correct worksheet
        sheet = data_wb.active
        print(sheet)
        # Before we go through the objectives file, create a Mail Merge workbook if it doesn't already exist, and just open it to update if it does
        print("Checking if Mail Merge workbook exists or not")
        mergeFilePath = str(os.getcwd() + "\MailMerge.xlsx")
        print(mergeFilePath)
        if os.path.exists(mergeFilePath):
            print("Opening existing Mail Merge workbook")
            merge_wb = openpyxl.load_workbook('MailMerge.xlsx')
            merge_sheet = merge_wb.active
        else:
            # if the file doesn't exist yet, create a new mail merge file
            merge_wb = openpyxl.Workbook()
            merge_sheet = merge_wb.active
            print("Adding column headers")
            merge_sheet['A1'] = "OSC Code"
            merge_sheet['B1'] = "DBA"
            merge_sheet['C1'] = "Contact First Name"
            merge_sheet['D1'] = "Contact Last Name"
            merge_sheet['E1'] = "Contact Email"
            merge_sheet['F1'] = "Aero Status"
            merge_sheet['G1'] = "Objective"
            merge_sheet['H1'] = "MTD"
            merge_sheet['I1'] = "% of Goal"
            merge_sheet['J1'] = "Purchases to Go"
            merge_sheet['K1'] = "Reward"
            print("All headers added")
            merge_wb.save('MailMerge.xlsx')
            
##        # Check that the columns are lined up correctly
##        confirm = QMessageBox()
##        confirm.setIcon(QMessageBox.Question)
##        confirm.SetText("Are the columns matched up correctly?")
##        confirm.setWindowTitle("Confirm Merge Fields")
##        confirm.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
##        QMessageBox.Yes.buttonClicked 
        
        # Reactivate the objectives update workbook
        print("Reactivating objectives workbook now")
        sheet = data_wb.active
        # Cycle through each row. If the row is an OSC, gather this data to later be stored in the array "objectives_update"
        print ("Entering for loop to gather data")
        # We'll start populating the mail merge worksheet at row 2
        mergeRow = 2
        # Search for the cell headers and take note of the column index
        for column in range(1, sheet.max_column + 1):
            if sheet.cell(row = 24, column = column).value == "Rwd.":
                rewardColumn = column
                break
            else:
                continue
        aeroColumn = rewardColumn - 1
        DBAColumn = rewardColumn + 2
        objectiveColumn = DBAColumn + 2
        PTDColumn = objectiveColumn + 1
        percentColumn = PTDColumn + 1
        PTGColumn = percentColumn + 3
            
            # !!!!! At this point, prompt user to confirm that the fields are matched up (create little table in the message box maybe?)
            # --------------------------------------------------------------------------------------------------------------------------
            # Only pay attention to rows labeled "OSC" that have a valid reward amountf
        for row in range(23, sheet.max_row + 1):
            if str(sheet.cell(row = row, column = 2).value) == "OSC" and type(sheet.cell(row = row, column = rewardColumn).value) == int:
                code = str(sheet.cell(row = row, column = 1).value)
                if "-" in code:
                    code = code[1:len(code)+1]
                else:
                    code = code
                # Reward will be in the column labeled "Rwd."
                rewardCell = sheet.cell(row = row, column = rewardColumn)
                reward = "$" + str(rewardCell.value)
                # Aero status will be one cell to the left of reward
                aeroCell = sheet.cell(row = row, column = aeroColumn)
                aeroStatus = str(aeroCell.value)
                # DBA will be 2 cells to the right of reward
                DBACell = sheet.cell(row = row, column = DBAColumn)
                DBA = str(DBACell.value)
                # Objective will be 2 cells to the right of DBA
                print("Getting objective value")
                objectiveCell = sheet.cell(row = row, column = objectiveColumn)
                print(objectiveCell.value)
                objective = str(objectiveCell.value)
                objective = formatCurrency(objective)
                # Purchases to date will be 1 cell to the right of objective
                PTDCell = sheet.cell(row = row, column = PTDColumn)
                PTD = str(PTDCell.value)
                PTD = formatCurrency(PTD)
                # % of objective will be 1 cell to the right of PTD
                percentCell = sheet.cell(row = row, column = percentColumn)
                print(percentCell.value)
                percent = int(percentCell.value*100)
                print(percent)
                percent = str(percent)+"%"
                # Purchases to go will be 3 cells to the right of % of objective
                PTGCell = sheet.cell(row = row, column = PTGColumn)
                PTG = str(PTGCell.value)
                PTG = formatCurrency(PTG)
                print(code + " " + DBA + " " + aeroStatus + " " + str(objective) + " " + reward)
                # Now that you've gotten the information, copy it over to Mail Merge
                merge_sheet = merge_wb.active
                # Go through all the info columns in Mail Merge and copy over the values
                merge_sheet.cell(row = mergeRow, column = 1).value = code
                merge_sheet.cell(row = mergeRow, column = 2).value = DBA
                merge_sheet.cell(row = mergeRow, column = 6).value = aeroStatus
                merge_sheet.cell(row = mergeRow, column = 7).value = objective
                merge_sheet.cell(row = mergeRow, column = 8).value = PTD
                merge_sheet.cell(row = mergeRow, column = 9).value = percent
                merge_sheet.cell(row = mergeRow, column = 10).value = PTG
                merge_sheet.cell(row = mergeRow, column = 11).value = reward
                # Now that we've copied the information to this row, get the next merge row ready and reactivate the objectives workbook and worksheet
                merge_wb.save("MailMerge.xlsx")
                mergeRow = mergeRow + 1
                sheet = data_wb.active
            else:
                continue
            merge_sheet = merge_wb.active
            merge_wb.save("MailMerge.xlsx")
            


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
