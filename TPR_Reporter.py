import pandas as pd
from datetime import date
from tkinter import *
import os
import pendulum


class TPR_Reporter:
    
    def __init__(self):
        self.nxtSat = pendulum.now().next(pendulum.SATURDAY).strftime('%#m/%#d/%Y')
        self.brdFile = os.path.join(os.path.expanduser("~/Desktop"),
                                    "BRdata_Prices.xlsx")
        self.brdUpdated = self.checkUpdated()
    
    def checkUpdated(self):
        fileModifiedDate = date.fromtimestamp(os.path.getmtime(self.brdFile))
        return bool(fileModifiedDate == date.today())
            

class TPR_Reporter_GUI:
    
    def __init__(self):
        self.window = Tk()
        self.mainFrame = Frame(self.window)
        self.widgets()
        self.packWidgets()
        self.mainFrame.pack()
        self.window.mainloop()
    
    def windowSettings(self):
        self.window.title("TPR Report")
        self.window.geometry("200x200")
        
    def widgets(self):
        self.lblTopText = Label(self.mainFrame, 
                           text="Press Compile Report to get started")
        self.lblBottomText = Label(self.mainFrame, 
                                   text=f"Next Saturday is: {TPR_Reporter().nxtSat}")
        self.btnCompile = Button(self.mainFrame, text="Process Report")
    
    def packWidgets(self):
        self.lblTopText.pack()
        self.lblBottomText.pack()
        self.btnCompile.pack()


if __name__ == "__main__":
    TPR_Reporter_GUI()
    