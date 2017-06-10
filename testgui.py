import sys
import os
import errno
import tkinter
import openpyxl
from openpyxl.utils import get_column_letter
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from flask import Flask, render_template, request
from werkzeug import secure_filename
import tkinter as tk
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
        
    def getEmails(self):
        print("Getting OSC Master file")
        Tk().withdraw()
        filename = askopenfilename()
        if not filename:
            self.e.config(state=NORMAL)
            self.e.delete(0,END)
            self.e.insert(0,"...")
            self.e.config(state="readonly")
        # Only continue if an "ONA - Sales by Channel Assignment..." file is being uploaded. If not, prompt again.
        print("Uploading to "+ str(UPLOAD_FOLDER)) 
        osc_wb = openpyxl.load_workbook(filename)
        print("Opening this uploaded file")
        # Go to the correct worksheet
        osc_sheet = osc_wb.active
        print(osc_sheet)

        
    def getFile(self):

        # Function to format the currency values with $ and comma separation
        def formatCurrency(amt):
            if len(amt) > 3:
                hundreds = amt[len(amt)-3:len(amt)+1]
                thousands = amt[:len(amt)-3]
                return "$"+thousands+","+hundreds
            else:
                return "$"+amt

            

        # Function to confirm the fields are matched properly to merge    
##        def confirmFields(label_list):
##            print("Confirm fields!")
##            confirm = QMessageBox()
##            confirm.setIcon(QMessageBox.Question)
##            confirm.setText("Are the fields matched to the correct columns?")
##            confirm.setInformativeText("Click 'Show Details' to check the detected merge fields. If the pairings are not correct, you can manually enter the correct columns for each field.")
##            confirm.setDetailedText(label_list[0] + ": OSC Number \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[1] + ": Aero Status \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[2] + ": DBA \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[3] + ": Objective \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[4] + ": MTD \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[5] + ": % of Goal \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[6] + ": PurchasesToGo \n"\
##                                    "-------------------------------- \n"\
##                                    "" + label_list[7] + ": Reward")
##            confirm.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
##            confirm.exec_()
##            if QMessageBox.No:
##                print("Correct fields")
            
        print("Getting file")
        Tk().withdraw()
        self.filename = askopenfilename()
        if not self.filename:
            print("Cancel pressed")
            self.close()
            
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
            merge_sheet['A1'] = "OSC Number"
            merge_sheet['B1'] = "DBA"
            merge_sheet['C1'] = "Contact First Name"
            merge_sheet['D1'] = "Contact Last Name"
            merge_sheet['E1'] = "Contact Email"
            merge_sheet['F1'] = "Aero Status"
            merge_sheet['G1'] = "Objective"
            merge_sheet['H1'] = "MTD"
            merge_sheet['I1'] = "% of Goal"
            merge_sheet['J1'] = "PurchasesToGo"
            merge_sheet['K1'] = "Reward"
            print("All headers added")
            merge_wb.save('MailMerge.xlsx')

        
        # Reactivate the objectives update workbook
        print("Reactivating objectives workbook now")
        sheet = data_wb.active
        # Cycle through each row. If the row is an OSC, gather this data to later be stored in the array "objectives_update"
        print ("Entering for loop to gather data")
        # We'll start populating the mail merge worksheet at row 2
        mergeRow = 2
        # Search for the cell headers and take note of the column index
        for row in range(23,30):                
            for column in range(1, sheet.max_column + 1):
                if sheet.cell(row = row, column = column).value == "Rwd.":
                    rewardColumn = column
                    rewardLabel = "Rwd."
                    labelRow = row
                    break
                else:
                    continue
        codeLabel = str(sheet.cell(row = labelRow, column = 1).value)
        aeroColumn = rewardColumn - 1
        aeroLabel = str(sheet.cell(row = labelRow, column = aeroColumn).value)
        DBAColumn = rewardColumn + 2
        DBALabel = str(sheet.cell(row = labelRow, column = DBAColumn).value)
        objectiveColumn = DBAColumn + 2
        objectiveLabel = str(sheet.cell(row = labelRow, column = objectiveColumn).value)
        PTDColumn = objectiveColumn + 1
        PTDLabel = str(sheet.cell(row = labelRow, column = PTDColumn).value)
        percentColumn = PTDColumn + 1
        percentLabel = str(sheet.cell(row = labelRow, column = percentColumn).value)
        PTGColumn = percentColumn + 3
        PTGLabel = str(sheet.cell(row = labelRow, column = PTGColumn).value)
        labels = (codeLabel, aeroLabel, DBALabel, objectiveLabel, PTDLabel, percentLabel, PTGLabel, rewardLabel)

##        confirmFields(labels)
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
            if merge_sheet.cells(row = 2, column = 5)== None:
                self.getEmails
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
##        rewardsBtn = QPushButton('Upload end of month file to calculate rewards', self)
##        rewardsBtn.resize(rewardsBtn.sizeHint())
##        rewardsBtn.move(350, 50)
        updateBtn = QPushButton('Upload file to update contact information', self)
        updateBtn.resize(updateBtn.sizeHint())
        updateBtn.move(350, 50)
        updateBtn.clicked.connect(self.getEmails)
        self.setGeometry(100, 100, 800, 500)
        self.setWindowTitle('Orio Email Automation Program')
        self.setWindowIcon(QIcon('logo.png'))        
    
        self.show()


            

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
