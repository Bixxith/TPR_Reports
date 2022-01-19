import pandas as pd
from datetime import date
from tkinter import Tk, Frame, Button, Label, messagebox
import os
import pendulum
from openpyxl import load_workbook, styles


class TPR_Reporter:
    
    def __init__(self):
        self.nextSaturday = self.getNextSaturday(pendulum.now())
        self.nextThreeSaturdays = self.getNextThreeSaturdays()
        self.nextSaturdayDateFormatted = self.formatDates(self.nextSaturday)
        self.brdFile = os.path.join(os.path.expanduser("~/Desktop"),
                                    "BRdata_Prices.xlsx")

    def formatDates(self, date):
        dateFormat = '%#m/%#d/%Y'
        return date.strftime(dateFormat)
    
    def getNextSaturday(self, date):
        return date.next(pendulum.SATURDAY)
    
    def checkUpdated(self):
        fileModifiedDate = date.fromtimestamp(os.path.getmtime(self.brdFile))
        return bool(fileModifiedDate == date.today())
    
    def getData(self):
        tprColumn = 'TPR\nPrior'
        self.dateColumnName = 'TPR To'
        tprPriority = 99
        columnsToUse = "B,C,D,F,G,H,I,J,K,T"
        columnRenaming = {'Description': 'Item Description',
                               'Reg\nPM': ' ', 
                               'Reg\nPrice': 'Regular Price', 
                               'TPR\nPM': '  ',
                               'TPR\nPrice': 'TPR Price'}
        rawFile = pd.read_excel(self.brdFile, usecols=columnsToUse)
        filterOld = rawFile[rawFile[self.dateColumnName] >= self.nextSaturdayDateFormatted]
        filterFirstSat = filterOld[filterOld[self.dateColumnName] != self.nextThreeSaturdays[0]]
        filterSecondSat = filterFirstSat[filterFirstSat[self.dateColumnName] != self.nextThreeSaturdays[1]]
        filterThirdSat = filterSecondSat[filterSecondSat[self.dateColumnName] != self.nextThreeSaturdays[2]]
        tprFilter = filterThirdSat[filterThirdSat[tprColumn] == tprPriority]
        finalDataFile = tprFilter
        finalDataFile.rename(columns=columnRenaming,
                                        inplace=True)
        self.dataFile = finalDataFile

    def getNextThreeSaturdays(self):
        saturdays = []
        nextSat = self.nextSaturday
        for i in range(1,4):
            nextSat = self.getNextSaturday(nextSat)
            saturdays.append(self.formatDates(nextSat))
        return(saturdays)
                
    def setupFiles(self):
        fileDateFormat ='%#m-%#d-%Y'
        nxtSatDate = self.nextSaturday.strftime(fileDateFormat)
        self.reportFolder = os.path.expanduser("~\Desktop\TPR Report")
        if not os.path.exists(self.reportFolder):
            os.mkdir(self.reportFolder)
        self.reportFile = os.path.join(self.reportFolder,
                                       f'TPRreport{nxtSatDate}.xlsx')
        self.reportWriter = pd.ExcelWriter(self.reportFile,
                                           engine='xlsxwriter') 
              
    def createReport(self):
        notFoundError = """
            BRdata_Prices.xlsx does not exist.
            Please run "Parse BRdata Prices" and try again.
                        """
        notUpdatedError =   """
            BRdata_Prices.xlsx exists but has not been updated.
            Please run "Parse BRdata Prices" and try again.  
                            """
        try:
            self.brdUpdated = self.checkUpdated()
        except FileNotFoundError:
            messagebox.showerror(title="File Not Found",
                                   message=notFoundError)
            os._exit(0)
        if self.brdUpdated:
            self.getData()
            self.setupFiles()
            self.createSheets()
            self.reportWriter.close()
            self.postProcessing()
            
            self.completedReports()  
            quit()
        else:
            messagebox.showerror(title="File Not Up to Date",
                                message=notUpdatedError)
            return
    
    def completedReports(self):
        infoText = """
            Report created and stored in the
            TPR Report folder located on your desktop.
                    """
        messagebox.showinfo(title="Report Compiled",
                            message=infoText)

    def createSheets(self):
        upcString = '[<=99999]#;[<=9999999999]#####-#####;###-#####-#####'
        workbookFormats = {'upc':{'num_format': upcString},
                           'num':{'num_format': '$#.00'}}
        departments = {'Produce':[20,21,22,23,24,25,26,27,28,29,100],
                            'Meat':[30,31,32,33,34,35,36,37,38,39],
                            'Frozen':[40,41,42,43,44,45,46,47,48,49],
                            'Dairy':[50,51,52,53,54,55,56,57,58,59],
                            'Deli & Bakery':[60,61,62,63,64,65,66,67,68,69,
                                             80,81,82,83,84,85,86,87,88,89],
                            'GM & HBC':[70,71,72,73,74,75,76,77,78,79,
                                        90,91,92,93,94,95,96,97,98,99],
                            'Grocery':[200,201,202,203,204,205,
                                       206,207,208,209,210],
                            'Stray TPRs':range(1,211)}
        columnsList = ["UPC",
                       "Item Description",
                       " ",
                       "Regular Price",
                       "  ",
                       "TPR Price"]
        for dept in departments.keys():
            if dept != 'Stray TPRs':
                self.processDepts(dept, departments,columnsList, workbookFormats)
            else:
                self.processDepts(dept, departments,columnsList, workbookFormats, strays=True)
                
    def processDepts(self, dept, departments,columnsList, workbookFormats, strays=False):
        regularTPRs = self.dataFile[self.dataFile[self.dateColumnName] == self.nextSaturdayDateFormatted]
        print(regularTPRs)
        strayTPRS = self.dataFile[self.dataFile[self.dateColumnName] != self.nextSaturdayDateFormatted]
        print(strayTPRS)
        numList = departments[dept]
        if not strays:
            departmentTPRs = regularTPRs[regularTPRs['Dept'].isin(numList)]
        else:
            departmentTPRs = strayTPRS[strayTPRS['Dept'].isin(numList)]
        if not departmentTPRs.empty:
            sortedTPRs = departmentTPRs.sort_values(by=['UPC'])
            rows = sortedTPRs.index
            sortedTPRs.to_excel(self.reportWriter, 
                        sheet_name=dept,
                        index=False,
                        columns=columnsList)
            reportWorkbook = self.reportWriter.book
            moneyFormat = reportWorkbook.add_format(workbookFormats['num'])
            upcFormat = reportWorkbook.add_format(workbookFormats['upc'])
            reportWorksheet = self.reportWriter.sheets[dept]
            if dept != 'Stray TPRs':
                headerFormat = (f'&C&20TPR Report  |'
                                f'|  {dept} Department  |'
                                f'|  {self.nextSaturdayDateFormatted}')
            else:
                headerFormat = (f'&C&20TPR Report  |'
                                f'|  {dept} All Departments  |'
                                f'|  {self.nextSaturdayDateFormatted}')
            reportWorksheet.set_header(headerFormat)
            moneyFormat.set_align('center')
            reportWorksheet.set_column('A:A', 14.86, upcFormat)
            reportWorksheet.set_column('B:B', 38)
            reportWorksheet.set_column('C:C', 2.29)
            reportWorksheet.set_column('D:D', 11.86, moneyFormat)
            reportWorksheet.set_column('E:E', 2.29)
            reportWorksheet.set_column('F:F', 8.43, moneyFormat)
            
    def addBorders(self):
        self.processSheets()
           
    def processSheets(self):
        workbook = load_workbook(self.reportFile)
        for sheets in workbook.sheetnames:
            sheet = workbook[sheets]
            self.setBorder(sheet)
        workbook.save(self.reportFile )
        workbook.close()
      
    def setBorder(self, worksheet):
        workSheetRange = ['A', 'F']
        thinFormat = styles.Side(border_style="dotted", color="000000")
        addBorder = styles.Border(bottom=thinFormat)
        currentSheet = worksheet[workSheetRange[0]:workSheetRange[1]]
        for row in currentSheet:
            for cell in row: 
                if cell.row % 2 != 0 and cell.row != 1:
                    cell.border = addBorder

    def postProcessing(self):
        self.addBorders()
     
        
class TPR_Reporter_GUI:
    
    def __init__(self):
        self.window = Tk()
        self.mainFrame = Frame(self.window)
        self.windowSettings()
        self.widgets()
        self.packWidgets()
        self.mainFrame.pack()
        self.window.mainloop()
    
    def windowSettings(self):
        self.window.title("TPR Report")
        self.window.geometry("200x100")
        
    def widgets(self):
        nxtSat = TPR_Reporter().nextSaturdayDateFormatted
        self.lblTopTxt = Label(self.mainFrame, 
                           text="Press Compile Report to get started")
        self.lblBtmTxt = Label(self.mainFrame,
                               text=f"Next Saturday is: {nxtSat}")
        self.btnCompile = Button(self.mainFrame, text="Process Report")
        self.btnCompile.bind("<Button-1>", self.startProgram)
    
    def startProgram(self, event):
        TPR_Reporter().createReport()
        
    def packWidgets(self):
        self.lblTopTxt.pack()
        self.lblBtmTxt.pack()
        self.btnCompile.pack()


if __name__ == "__main__":
    TPR_Reporter_GUI()
    