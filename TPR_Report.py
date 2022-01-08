import pandas as pd
from datetime import date
from tkinter import *
import os
import pendulum

deptCount = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
deptCounter = 0


# finds the current day of the week and checks to see how many days until the closest saturday so that way it can get the exact date of the upcoming saturday
def getSaturday():
    dayGet = date.today().weekday()
    daysInc = 0
    nextSelector = date.today()
    while dayGet != 5:
        if dayGet == 6:
            dayGet = 5
            daysInc = 6
        else:
            dayGet += 1
            daysInc += 1
    todayPlusIncDay = date.today().day + daysInc

    nextSaturday = nextSelector.replace(day=todayPlusIncDay)

    formatSaturday = nextSaturday

    return formatSaturday

# sets up the file path and location. dictates the path to the folder and file
username = os.getlogin()
path = f'C:\\Users\\{username}\\Desktop\\TPR Report'
filename = f'{path}\\TPRreport{getSaturday()}''.xlsx'


# gets the current day of the week and checks to see how many days until saturday in order to find the date of the
# nearest saturday


class reportGUI:
    def __init__(self):
        window = Tk()
        window.title("TPR Report Compiler")
        window.geometry("200x200")
        frame1 = Frame(window)
        frame1.grid(row=3, column=3, padx=5, pady=5)
        btnCompile = Button(frame1, text="Process Report", width=15, height=5)
        lblInstructions1 = Label(frame1, text="Press Compile Report to get started")
        lblInstructions3 = Label(frame1, text=f" ")
        lblInstructions2 = Label(frame1, text=f"Next Saturday is: {getSaturday()}")

        def handle_click(event):
            import os
            seperateDepartments()
            os._exit(0)

        btnCompile.bind("<Button-1>", handle_click)
        lblInstructions3.grid(row=1, column=0)
        lblInstructions2.grid(row=2, column=0)
        lblInstructions1.grid(row=0, column=0)
        btnCompile.grid(row=3, column=0)
        frame1.pack()
        window.mainloop()


# checks to make sure the file was created today so that way we get the most accurate report
def load_if_modified_today(xlsx):
    print("Modified!")
    modx = os.path.getmtime(xlsx)
    xmod = date.fromtimestamp(modx)
    rawSaturday = getSaturday()
    saturday = rawSaturday.strftime("%#m/%#d/%Y")
    # need to add logic for handling instances where the BRData Prices has not been updated.
    if date.today() == xmod:
        x = 99
        df = pd.read_excel(xlsx, usecols="B,C,D,F,G,H,I,J,K,O,T")
        df = df[df['TPR\nPrior'].isin([x])]
        #df = df[df['TPR To'].isin([saturday])]
        print(df)
        return df


# static file name that never changes
desktop = os.path.expanduser("~/Desktop")
filePath = os.path.join(desktop, "BRdata_Prices.xlsx")



# calls the check to make sure file was updated today
df = load_if_modified_today(filePath)


# creates the file incase it doesn't already exist and assigns writer to reference it
def createTPRReport():
    while True:
        if os.path.exists(path):
            break
        else:
            os.mkdir(path) 
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    return writer


writer = createTPRReport()


# assigns departments to worksheets and formats them accordingly
def createSheets(dept, deptName):
    if len(dept) > 0:
        deptWriter = dept.to_excel(writer, sheet_name=deptName, index=False,
                                   columns=["UPC", "Item Description", " ", "Regular Price", "  ", "TPR Price"])
        workbook = writer.book
        worksheet = writer.sheets[deptName]
        print(len(dept))
        worksheet.set_header(
            f'&C&20TPR Report  ||  {deptName} Department  ||  {date.strftime(getSaturday(), "%m %d %Y")}')
        new_format = workbook.add_format({'num_format': '$#.##'})
        new_format.set_align('center')
        worksheet.set_column('A:A', 12, )
        worksheet.set_column('B:B', 38)
        worksheet.set_column('C:C', 2.29)
        worksheet.set_column('D:D', 11.86, new_format)
        worksheet.set_column('E:E', 2.29)
        worksheet.set_column('F:F', 8.43, new_format)


# seperates all of the departments and assigns them to worksheets
def seperateDepartments():
    print('seperateDepart')
    df2 = df.rename(
        columns={'Description': 'Item Description', 'Reg\nPM': ' ', 'Reg\nPrice': 'Regular Price', 'TPR\nPM': '  ',
                 'TPR\nPrice': 'TPR Price'}, inplace=True)

    result = df.sort_values('UPC')

    produce = df[df.Dept.between(20, 29)].sort_values('UPC')
    deptProduce = 'Produce'
    createSheets(produce, deptProduce)

    meat = df[df.Dept.between(30, 39)].sort_values('UPC')
    deptMeat = 'Meat'
    createSheets(meat, deptMeat)

    frozen = df[df.Dept.between(40, 49)].sort_values('UPC')
    deptFrozen = 'Frozen'
    createSheets(frozen, deptFrozen)

    dairy = df[df.Dept.between(50, 59)].sort_values('UPC')
    deptDairy = 'Dairy'
    createSheets(dairy, deptDairy)

    bakery = df[df.Dept.between(60, 69)].sort_values('UPC')
    deptBakery = 'Bakery'
    createSheets(bakery, deptBakery)

    gm1 = df[df.Dept.between(70, 79)].sort_values('UPC')
    deptGM1 = 'GMHBC1'
    createSheets(gm1, deptGM1)

    deli = df[df.Dept.between(80, 89)].sort_values('UPC')
    deptDeli = 'Deli'
    createSheets(deli, deptDeli)

    gm2 = df[df.Dept.between(90, 99)].sort_values('UPC')
    deptGM2 = 'GMHBC2'
    createSheets(gm2, deptGM2)

    grocery = df[df.Dept.between(200, 210)].sort_values('UPC')
    deptGrocery = 'Grocery'
    createSheets(grocery, deptGrocery)

    writer.save()

    # # opens the sheets and then tries to print them
    # import win32com.client

    # o = win32com.client.Dispatch('Excel.Application')
    # o.visible = True
    # wb = o.Workbooks.Open(filename)
    # ws = wb.Worksheets
    # print(ws)

    # # returns a TypeError stating something about a bool.  It still prints and the solution is attained.  this
    # # bypasses the error.

    # try:
    #     ws.printout()
    #     print('yes')
    # except:
    #     return {
    #         closeProgram()
    #     }
    # closeProgram()


# kills the excel program
def closeProgram():
    os.system("taskkill /pid " + str('EXCEL.EXE'))


reportGUI()
