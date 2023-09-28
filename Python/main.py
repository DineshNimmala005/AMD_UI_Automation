import tkinter
from tkinter import ttk
from tkinter import messagebox
from tkinter import *
import openpyxl
import os

def JSONConversion():
     print("Converting excel into JSON")
     
# Performing logical executional code funtionality to save data 
# when we click on 'Save the data' button this logic will execute
# This logic will save the 
def SaveData():

    # This step creates the work book in respective path
    # If file already exist this step is going to excemption

    filePath="C:\\Users\\CQE\\Desktop\\Python\\Game.xlsx"

    # When we click on save button first it will generates Workbook.
    # If file already exist then, It will leave the method

    if not os.path.exists(filePath):
                    Workbook=openpyxl.Workbook()
                    sheet=Workbook.active
                    Heading=["Game name","Game application","Steam support","EPIC support","EA Sports support"]
                    sheet.append(Heading)
                    Workbook.save(filePath)
                    print(" Just created the excel workbook ")
    
    # Executional steps for UI starts from here.

    Userpermision=Save_Info_var.get()
    if Userpermision == "Accepted":
        GameName=Game_Name_entry.get()
        GameEXE=Game_Exe_entry.get()
        SteamValue=Steam_Dropdown.get()
        EpicValue=EPIC_Dropdown.get()
        EAValue=EA_Dropdown.get()
        if GameName and GameEXE:
            if SteamValue and EpicValue and EAValue:
                
                # Open the excel sheet and added the values in excel
                Workbook=openpyxl.load_workbook(filePath)
                sheet=Workbook.active
                sheet.append([GameName,GameEXE,SteamValue,EpicValue,EAValue])
                Workbook.save(filePath)

                # Sample should be print in Console
                print("Game : ",GameName)
                print("Application name : ",GameEXE)
                print("User permissions : ",Userpermision)
                print("Steam support : ",SteamValue)
                print("EPIC support : ",EpicValue)
                print("EA sports support : ",EAValue)
                print("------------------------------------------")
                # Providing conformation popup after clicking on save button
                tkinter.messagebox.showwarning(title="Successfully",message=" Data saved Successfully ")
            else:
                tkinter.messagebox.showwarning(title=" Warning ! ",message=" Game launchers data required ")
        else:
            tkinter.messagebox.showwarning(title=" Warning ! ",message=" Game name and application path required ")
# Printing the out put logic.
    else:

        tkinter.messagebox.showwarning(title=" Warning ! ",
                                       message="Please accept the check box to save the data ")
        print("ERROR : Data not saved")

# Window is called overall window which contains all elements inside it
Window=tkinter.Tk()
# Window Name like Overal UI name in header
Window.title(" AMD Game Data UI")
Window.geometry("1200x700")
Window.iconbitmap('C:/Users/CQE/Desktop/Python/AMD.ico')
#icon=PhotoImage("C:\\Users\\CQE\\Desktop\\Python\\AMD.jpg")
##Window.iconphoto(False,icon)
#icon=Window.iconbitmap('AMD.ico')
# Window.iconphoto(True,icon)

# Frame is a sub window under the window which keeps all the elements 
# in an order
Frame=tkinter.Frame(Window)
# Inside a frame there are many sections like Div in HTML under HTML body
Frame.pack()
#
# Saving Game information like game name and FGame exe name
Game_info_frame=tkinter.LabelFrame(Frame,text="Game Information",font=('Calibri',16))
Game_info_frame.grid(row=0,column=0,padx=20,pady=20)

Game_Name_label=tkinter.Label(Game_info_frame,text="Game name")
Game_Name_label.grid(row=0,column=0)

Game_Exe_label=tkinter.Label(Game_info_frame,text="Game application path")
Game_Exe_label.grid(row=0,column=1)


Game_Name_entry=tkinter.Entry(Game_info_frame)
Game_Exe_entry=tkinter.Entry(Game_info_frame)
Game_Name_entry.grid(row=1,column=1)
Game_Exe_entry.grid(row=1,column=0)

# Padding and spacing for all Child elements of Game_info_frame Frames
for widget in Game_info_frame.winfo_children():
    widget.grid_configure(padx=5,pady=10)

# Saving Game Launchers information like Epic launcher 
# Stream launcher and EA launcher etc,
Game_Launchers_frame=tkinter.LabelFrame(Frame,text="Game launchers",font=('Calibri',16))
Game_Launchers_frame.grid(row=1,column=0,padx=20,pady=20)

# Steam launcher values 
Steam_Launcher=tkinter.Label(Game_Launchers_frame,text="Steam")
Steam_Dropdown=ttk.Combobox(Game_Launchers_frame,values=["True","False"])
Steam_Launcher.grid(row=0,column=0)
Steam_Dropdown.grid(row=1,column=0)

# Epic launcher values 
EPIC_Launcher=tkinter.Label(Game_Launchers_frame,text="Epic Games")
EPIC_Dropdown=ttk.Combobox(Game_Launchers_frame,values=["True","False"])
EPIC_Launcher.grid(row=0,column=1)
EPIC_Dropdown.grid(row=1,column=1)

# EA launcher values 
EA_Launcher=tkinter.Label(Game_Launchers_frame,text="EA Sports")
EA_Dropdown=ttk.Combobox(Game_Launchers_frame,values=["True","False"])
EA_Launcher.grid(row=0,column=2)
EA_Dropdown.grid(row=1,column=2)


# Padding and spacing for all Child elements of Game_Launchers_frame Frames
for widget in Game_Launchers_frame.winfo_children():
    widget.grid_configure(padx=5,pady=10)

#
# Save conformation frame
Conformation_frame=tkinter.LabelFrame(Frame,text="Terms and conditions")
Conformation_frame.grid(row=2,column=0,padx=20,pady=10)

# Like terms and condition checkbox
Save_Info_var=tkinter.StringVar(value="Denined")
Save_Check=tkinter.Checkbutton(Conformation_frame,text="Save the game data.",
                               variable=Save_Info_var,onvalue="Accepted",offvalue="Denined",padx=50)

Save_Check.grid(row=0,column=0)

# Adding save and conformation button
# command descibes that it need to perform action after clicking button
# All logic is writen in save data method

Buttons_frame=tkinter.LabelFrame(Frame)
Buttons_frame.grid(row=3,column=0,padx=20,pady=20)


Button=tkinter.Button(Buttons_frame,text="Save The Data",command=SaveData,fg="White",highlightthickness=5,font=('Times New Roman',8),bg="Red",height=1,width=20)
Button.grid(row=0,column=3,sticky="news",padx=45,pady=20)

Button=tkinter.Button(Buttons_frame,text="Create JSON",command=JSONConversion,fg="White",highlightthickness=5,font=('Times New Roman',8),bg="Gray",height=1,width=20)
Button.grid(row=0,column=0,sticky="news",padx=45,pady=20)



#which creates a close button for UI like window close option
Window.mainloop()