
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from  datetime import datetime, timedelta
import time
import openpyxl as excel
import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog as fd
from tkinter import messagebox

class App:
    def __init__(self, root):
        self.filename =""
        #setting title
        root.title("WhatsApp Web Automator")
        #setting window size
        width=600
        height=500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_762=tk.Label(root)
        GLabel_762["activebackground"] = "#90f090"
        GLabel_762["activeforeground"] = "#90f090"
        GLabel_762["bg"] = "#36dc36"
        GLabel_762["cursor"] = "trek"
        ft = tkFont.Font(family='Times',size=28)
        GLabel_762["font"] = ft
        GLabel_762["fg"] = "#ffffff"
        GLabel_762["justify"] = "center"
        GLabel_762["text"] = "WhatsApp Bot"
        #GLabel_762["anchor"]="nw"
        GLabel_762["relief"] = "flat"
        GLabel_762.place(x=0,y=0,width=598,height=30)

        # GLabel_795=tk.Label(root)
        # GLabel_795["bg"] = "#90ee90"
        # ft = tkFont.Font(family='Times',size=12)
        # GLabel_795["font"] = ft
        # GLabel_795["fg"] = "#333333"
        # GLabel_795["justify"] = "center"
        # GLabel_795["text"] = "Instructions:"
        # GLabel_795.place(x=280,y=40,width=319,height=30)

        GMessage_993=tk.Message(root)
        GMessage_993["bg"] = "#90f090"
        ft = tkFont.Font(family='Times',size=12)
        GMessage_993["font"] = ft
        GMessage_993["fg"] = "#333333"
        GMessage_993["justify"] = "left"
        GMessage_993["text"] = "Click on Select File button and select Excel template file:"
        GMessage_993.place(x=0,y=40,width=278,height=118)

        GButton_383=tk.Button(root)
        GButton_383["bg"] = "#efefef"
        ft = tkFont.Font(family='Times',size=12)
        GButton_383["font"] = ft
        GButton_383["fg"] = "#000000"
        GButton_383["justify"] = "center"
        GButton_383["text"] = "Select File"
        GButton_383.place(x=100,y=130,width=70,height=25)
        GButton_383["command"] = self.GButton_383_command

        GMessage_860=tk.Message(root)
        GMessage_860["bg"] = "#90f090"
        ft = tkFont.Font(family='Times',size=12)
        GMessage_860["font"] = ft
        GMessage_860["fg"] = "#333333"
        GMessage_860["justify"] = "left"
        GMessage_860["text"] = "Click on Run button to execute the bot"
        GMessage_860.place(x=0,y=160,width=278,height=119)

        GButton_826=tk.Button(root)
        GButton_826["bg"] = "#efefef"
        ft = tkFont.Font(family='Times',size=12)
        GButton_826["font"] = ft
        GButton_826["fg"] = "#000000"
        GButton_826["justify"] = "center"
        GButton_826["text"] = "Run"
        GButton_826.place(x=100,y=250,width=70,height=25)
        GButton_826["command"] = self.run

        self.GMessage_886=tk.Message(root)
        self.GMessage_886["bg"] = "#90f090"
        ft = tkFont.Font(family='Times',size=10)
        self.GMessage_886["font"] = ft
        self.GMessage_886["fg"] = "#333333"
        self.GMessage_886["justify"] = "left"
        self.GMessage_886["text"] = ""
        self.GMessage_886["anchor"]="nw"
        self.GMessage_886.place(x=280,y=40,width=318,height=456)

        GLabel_683=tk.Label(root)
        GLabel_683["bg"] = "#90f090"
        ft = tkFont.Font(family='Times',size=10)
        GLabel_683["font"] = ft
        GLabel_683["fg"] = "#333333"
        GLabel_683["justify"] = "center"

        GLabel_683["text"] = ""
        GLabel_683.place(x=0,y=280,width=278,height=236)
    def findLinkinthesheet(self,file,sheetname,day):
        sheet = file[sheetname]
        ColWText= sheet['C']
        ColYLink= sheet['D']
        icounter =0
        for cell in range(len(ColWText)):
            if icounter>0:

                dayText = str(ColWText[cell].value).split("|")[0]
                print("1"+ dayText)
                print("2 Day "+ str(int(day)))
                if  str(dayText).strip() == "Day "+ str(int(day)) :
                    #print(str(ColWText[cell].value),str(ColYLink[cell].value))
                    return str(ColWText[cell].value),str(ColYLink[cell].value)
                    break
            icounter +=1
        return "NA","NA"
    def readContacts(self,fileName):
        #lstGroupName = []
        #lstSheetName=[]
        #lstLastDay =[]
        dictInformation={}
        file = excel.load_workbook(fileName,  data_only=True)
        sheet = file['Index']
        ColGroupName = sheet['B']
        ColSheetName = sheet['E']
        ColDay= sheet['F']
        ColStatus= sheet['C']
        icounter =0
        for cell in range(len(ColGroupName)):
            if icounter > 0:
                if str(ColGroupName[cell].value)=="None":
                    break
                if str(ColGroupName[cell].value)!="None" and str(ColStatus[cell].value)=="Open":
                    contact = str(ColGroupName[cell].value)
                    SheetName = str(ColSheetName[cell].value)
                    LastDay = str(ColDay[cell].value)
                    print("0" + str(LastDay))
                    WText,YLink = self.findLinkinthesheet(file,SheetName,LastDay)
                    #contact = "\"" + contact + "\""
                    #lstGroupName.append(contact)
                    #lstSheetName.append(SheetName)
                    #lstLastDay.append(LastDay)
                    if WText !="NA":
                        dictInformation[contact+ "||" + SheetName]= WText
            icounter +=1
        return dictInformation
    def run(self):
        targets = self.readContacts(self.filename)
        print(targets)
        driver = webdriver.Chrome('chromedriver.exe')
        
        driver.get("https://web.whatsapp.com/")
        print("Page Opened")
        end_time = datetime.now()+timedelta(minutes=5)
        while end_time>= datetime.now():
            wait = WebDriverWait(driver, 10)
            try:
                searBoxPath = '//*[@title="Search input textbox"]'
                inputSearchBox = driver.find_element(by=By.XPATH,value=searBoxPath)
                if inputSearchBox:
                    break
            except:
                pass
            
        success = 0
        sNo = 1
        failList = []
        count =0
        wait = WebDriverWait(driver, 10)
        wait5 = WebDriverWait(driver, 5)
        print("ok")
        for target in targets.keys():
            msg = targets[target]

            target = '"'+ str(target).split("||")[0] +'"'
            print(sNo, ". Target is: " + target)
            sNo+=1
            try:
                # Select the target
                x_arg = '//*[contains(@title,' + target + ')]'
                print(x_arg)
                try:
                    wait5.until(EC.presence_of_element_located((
                        By.XPATH, x_arg
                    )))
                except:
                    # If contact not found, then search for it
                    searBoxPath = '//*[@title="Search input textbox"]'
                    wait5.until(EC.presence_of_element_located((
                        By.XPATH, searBoxPath
                    )))
                    inputSearchBox = driver.find_element(by=By.XPATH,value=searBoxPath)
                    print("Search")
                    time.sleep(0.5)
                    # click the search button
                    #driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/div[2]/div/button').click()
                    inputSearchBox.click()
                    print("clicked")
                    time.sleep(1)
                    inputSearchBox.clear()
                    inputSearchBox.send_keys(target[1:len(target) - 1])
                    print('Target Searched')
                    # Increase the time if searching a contact is taking a long time
                    time.sleep(4)

                # Select the target
                #driver.find_element_by_xpath(x_arg).click()
                driver.find_element(by=By.XPATH, value=x_arg).click()
                print("Target Successfully Selected")
                time.sleep(2)

                # Select the Input Box
                inp_xpath = '//*[@title="Type a message"]'
                #inp_xpath = "//div[@contenteditable='true']"
                input_box = wait.until(EC.presence_of_element_located((
                    By.XPATH, inp_xpath)))
                time.sleep(1)
                print(input_box)
                # Send message
                # target is your target Name and msgToSend is you message
                input_box.send_keys( msg ) # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
                # Link Preview Time, Reduce this time, if internet connection is Good
                time.sleep(2)
                input_box.send_keys(Keys.ENTER)
                print("Successfully Send Message to : "+ target + '\n')
                success+=1
                time.sleep(0.5)
                self.GMessage_886["text"]= self.GMessage_886["text"]+"\nTo: " + target + "\nStatus: Sent\n--------------------------------------------------" 
            except:
                # If target Not found Add it to the failed List
                print("Cannot find Target: " + target)
                failList.append(target)
                self.GMessage_886["text"]= self.GMessage_886["text"]+"\nTo: " + target + "\nStatus: Fail\n--------------------------------------------------" 
                pass

        print("\nSuccessfully Sent to: ", success)
        print("Failed to Sent to: ", len(failList))
        print(failList)
        print('\n\n')
        count+=1
        driver.quit()
        messagebox.showinfo("showinfo", "Messages Sent")
        
    def GButton_383_command(self):
        self.filename =  fd.askopenfilename()
        if self.filename:
            messagebox.showinfo("showinfo", "File Selected")


if __name__ == "__main__":
    root = tk.Tk()
    print("1")
    app = App(root)
    print("2")
    root.mainloop()
    print("3")
