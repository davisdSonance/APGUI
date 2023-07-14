#all imports for basic functions
from ast import Delete, For
from xmlrpc.client import boolean
import customtkinter
import tkinter as tk
import sys					                      
import time
import clr 
import os
import wx
from matplotlib.axis import Axis
import matplotlib.pyplot as plt
from mpl_interactions import ioff, panhandler, zoom_factory
import numpy as np
import pandas as pd
import subprocess
import warnings
import shutil 

warnings.simplefilter("ignore")
clr.AddReference("System.Drawing") 					                       
clr.AddReference("System.Windows.Forms") 	
clr.AddReference("C:\Program Files\Audio Precision\APx500 6.1\API\AudioPrecision.API2") 	# Adding Reference to the APx API
clr.AddReference("C:\Program Files\Audio Precision\APx500 6.1\API\AudioPrecision.API") 
clr.AddReference("System.Collections")

from PIL import Image,ImageTk
from customtkinter.windows.widgets import image
from scipy import misc
from AudioPrecision.API import *
from System.Drawing import Point # Needed for Dialog Boxes
from System.Windows.Forms import Application, Button, Form, Label # Needed for Dialog Boxes
from System.IO import Directory, Path # Used to point to the sample project file
from System import Array
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from textwrap import wrap
from playsound import playsound

customtkinter.set_widget_scaling(1) #adds scaling to the program if you want bigger text it will accomnidate 
customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

class App(customtkinter.CTk):
    def __init__(self):#this will make the GUI and set all the attributes for the program
        super().__init__()
        self.sumbission = ""
        self.modelChoice = ""
        self.seriesChoice = ""
        self.sequence = ""
        self.sequenceTest = "Loudspeaker Production Test"
        self.tests = ["Impedance","PhaseWrapDiff","RMS","THD"]
        self.ticks = [10,20,30,50,100,200,300,500,1000,2000,3000,5000,10000,20000]
        self.finalFilePath = ""
        self.images = []
        self.passedTest = []
        self.failedTest = []
        self.pX = 1
        self.pY = 1
        
        self.models = ["SA","AT","BE","OW","CUBE","QX","QXC","SXC","FXA","ASX","OT","FX","AL","W","OMNI","EXTRA", "NONE"]
        self.modelSA = ["32-4", "32-70",  "42-4","42-70", "63-4","63-70"]
        self.modelAt = ["806", "808","812","816"]
        self.modelBE =["42","43", "52Q", "53Q","62Q","64Q","42-M","43-M","52Q-M","53Q-M","62Q-M","63Q-M","66Q-M","68Q-M","66QBE","68QBE","66QTI","68QTI"]
        self.modelOW =["43","63"]
        self.modelCUBE =["320","330","520","530","620","630","806BE","820","830","1010","1020","1020A"]
        self.modelQX =["8S","10S","620","820"]
        self.modelQXC =["420"]
        self.modelSXC =["550S","Q550S"]
        self.modelFXA =["QX62","62"]
        self.modelFX =["QX62","62"]
        self.modelASX =["66Q","84-4","88Q"]
        self.modelOT = ["62", "63"]
        self.modelAL =["32","33"]
        self.modelW =["42Q","43Q","52Q","53Q"]
        self.modelOMNI =["89","109","129"]
        self.modelEXTRA = []

        #all of the APX attrubites start here 

        self.APx = APx500_Application(APxOperatingMode.SequenceMode)
        self.app = customtkinter.CTk()  # create CTk window like you do with the Tk window
        self.app.geometry("930 x 450")#custom resolution of the GUI.
        self.check_var = customtkinter.StringVar(value="on")

        self.window48Button = customtkinter.CTkButton(master=self.app, text="Window Test 4 - 8", command=self.Window48Button,fg_color="blue")
        self.window48Button.grid(row=1, column= 0, padx = 20, pady = 10)

        self.limit48Button = customtkinter.CTkButton(master=self.app, text= "Limits 4 - 8", command=self.limit48Button,fg_color="blue")
        self.limit48Button.grid(row=2, column= 0, padx = 20, pady = 10)

        self.customButton= customtkinter.CTkButton(master=self.app, text= "Custom", command=self.customButton,fg_color="blue")
        self.customButton.grid(row=3, column= 0, padx = 20, pady = 10)

        self.window70Button= customtkinter.CTkButton(master=self.app, text= "Windows Test 70V", command=self.Window70Button,fg_color="blue")
        self.window70Button.grid(row=4, column= 0, padx = 20, pady = 10)

        self.limt70Button= customtkinter.CTkButton(master=self.app, text= "Limits 70V", command=self.limit70Button,fg_color="blue")
        self.limt70Button.grid(row=5, column= 0, padx = 20, pady = 10)

        self.button6= customtkinter.CTkButton(master=self.app, text= "Passed Test (Right Arrow)", command=self.passButton)
        self.button6.grid(row=6, column=3, padx = 20, pady = 10)

        self.button7= customtkinter.CTkButton(master=self.app, text= "Failed Test (Left Arrow)", command=self.failButton)
        self.button7.grid(row=6, column= 0, padx = 20, pady = 10)

        self.button8= customtkinter.CTkButton(master=self.app, text= "Start Software", command=self.openSoftware,width= 25, height= 25,fg_color="green",text_color="white")
        self.button8.grid(row=0, column= 0, padx = 0, pady = 0)

        self.button9= customtkinter.CTkButton(master=self.app, text= "Submit (Enter)", command=self.testSubmitButton,width= 100, height= 50,fg_color="green")
        self.button9.grid(row=7, column= 5, padx = 20, pady = 10)

        self.button10= customtkinter.CTkButton(master=self.app, text= "Quit (Escape)", command=self.quitbutton,width= 25, height= 25,fg_color="red")
        self.button10.grid(row=0, column= 5, padx = 20, pady = 10)

        self.button11= customtkinter.CTkButton(master=self.app, text= "Start Test (F1)", command=self.startTestButton,width= 25, height= 25,fg_color="green")
        self.button11.grid(row=6, column= 2)

        self.formatButton = customtkinter.CTkButton(master=self.app, text = "Reset Test", command = self.reset, fg_color="green")
        self.formatButton.grid(row= 6, column = 5, padx = 20, pady = 5)
        
        self.user_entry = customtkinter.CTkEntry(master=self.app,placeholder_text= "Enter Serial Number")
        self.user_entry.grid(row=5, column=2)
        
        self.user_entry_custom = customtkinter.CTkEntry(master=self.app,placeholder_text= "Custom Models only")
        self.user_entry_custom.grid(row=2, column=2)
        
        self.user_entry1 = customtkinter.CTkEntry(master=self.app,placeholder_text= "Enter Name Here")
        self.user_entry1.grid(row=1, column=2)

        self.textbox = customtkinter.CTkTextbox(self.app,height= 100,border_color="black")
        self.textbox.grid(row=7, column= 2)

        self.passbox = customtkinter.CTkTextbox(self.app,height= 100,border_color="green")
        self.passbox.grid(row=7, column= 3)

        self.failbox = customtkinter.CTkTextbox(self.app,height= 100,border_color="red")
        self.failbox.grid(row=7, column= 0)   
        
        self.optionmenu_model = customtkinter.StringVar(value="model")          
        self.optionmenu_series = customtkinter.StringVar(value="series")  # set initial value

        self.modelDropDown = customtkinter.CTkOptionMenu(master=self.app,values = self.models, command=self.modelDrop,variable= self.optionmenu_model,fg_color="blue")
        self.modelDropDown.grid(row = 3, column = 2)

        self.seriesDropDown = customtkinter.CTkOptionMenu(master=self.app,variable = self.optionmenu_series,command=self.subModelDrop,fg_color="blue")
        self.seriesDropDown.grid(row = 4, column = 2)     

        self.printPDF = customtkinter.CTkButton(master=self.app, text= "PDF Preview", command=self.createReport)
        self.printPDF.grid( row=2, column= 5)
        
        self.help = customtkinter.CTkButton(master =self.app, text= "Help", command=self.openHelp)
        self.help.grid(row = 1 , column = 5)
        
        self.check_var = customtkinter.StringVar(value="on")
        self.textbox.configure(state="normal")

        self.impedenceButton = customtkinter.CTkButton(master=self.app, text= "Impedance", command=self.showImpedance,width= 25, height= 25,fg_color="blue")
        self.impedenceButton.grid(row=1, column=3)
        self.rmsButton = customtkinter.CTkButton(master=self.app, text= "RMS", command=self.showRMS,width= 25, height= 25,fg_color="blue")
        self.rmsButton.grid(row=2, column=3)
        self.thdButton = customtkinter.CTkButton(master=self.app, text= "THD", command=self.showTHD,width= 25, height= 25,fg_color="blue")
        self.thdButton.grid(row=3, column=3)
        self.PWDButton = customtkinter.CTkButton(master=self.app, text= "PhaseWrapDiff", command=self.showPWD,width= 25, height= 25,fg_color="blue")
        self.PWDButton.grid(row=4, column=3)
        self.ListenButton = customtkinter.CTkButton(master=self.app, text= "Listening Test", command=self.ListenTest,width= 25, height= 25,fg_color="blue")
        self.ListenButton.grid(row=5, column=3)
        
        
        self.app.bind("<Key>", self.key_handler)
        
        if(self.APx.Visible == False):
            self.modelDropDown.configure(state="disabled")
            self.seriesDropDown.configure(state="disabled")
            self.formatButton.configure(state="disabled")
            self.window48Button.configure(state="disabled")
            self.limit48Button.configure(state="disabled")
            self.customButton.configure(state="disabled")
            self.window70Button.configure(state="disabled")
            self.limt70Button.configure(state="disabled")
            self.button6.configure(state="disabled")
            self.button7.configure(state="disabled")
            self.button9.configure(state="disabled")
            self.button11.configure(state="disabled")
            self.impedenceButton.configure(state= "disabled")
            self.rmsButton.configure(state= "disabled")
            self.thdButton.configure(state= "disabled")
            self.PWDButton.configure(state= "disabled")
            self.ListenButton.configure(state= "disabled")
            self.printPDF.configure(state= "disabled")
        self.app.mainloop()
        
    def quitbutton(self):#this is the quit function that ends all process within the program
    
        if(self.APx.Visible == True):
            self.APx.Visible = False
            sys.exit()
        elif(self.APx.Visible == False):
            sys.exit()
        #Button functions for changing the sequence in the APx. 
    def Load_button(self):#this button will load the file format that tests will be done with
        filename = "Davis_AP_template.approjx"
        directory = Directory.GetCurrentDirectory()
        fullpath = Path.Combine(directory, filename)
        self.APx.OpenProject(fullpath)
        #changes the sequence to Window 4- 8
    
    def key_handler(self,event): #this makes key functionality for the buttons and
            if(event.keycode == 39):
                self.passButton()
            if(event.keycode == 37):
                self.failButton()     
            if (event.keycode == 13):
                self.testSubmitButton()
            if (event.keycode == 27):
                self.quitbutton()
            if (event.keycode == 112):
                self.startTestButton()

    
    def Window48Button(self):#button for windows test 4 - 8 
        self.APx.Sequence.Sequences.Activate("Window Test (4-8)")
        self.sequence = "Window Test (4-8)"

        self.limit48Button.configure(fg_color="blue")
        self.customButton.configure(fg_color="blue")
        self.limt70Button.configure(fg_color="blue")
        self.window70Button.configure(fg_color="blue")

        self.window48Button.configure(fg_color="green")
        
        #changes the sequence to Limits 4 8 
    def limit48Button(self):#button for limit test 4 - 8
        self.APx.Sequence.Sequences.Activate("Limits (4-8)")
        self.sequence = "Limits (4-8)"

        self.customButton.configure(fg_color="blue")
        self.limt70Button.configure(fg_color="blue")
        self.window48Button.configure(fg_color="blue")
        self.window70Button.configure(fg_color="blue")

        self.limit48Button.configure(fg_color="green")
        #changes the sequence to Window Test (70V)")
    def Window70Button(self):#button for wondows 70 test
        self.APx.Sequence.Sequences.Activate("Window Test (70V)")
        self.sequence = "Window Test (70V)"

        self.limit48Button.configure(fg_color="blue")
        self.customButton.configure(fg_color="blue")
        self.limt70Button.configure(fg_color="blue")
        self.window48Button.configure(fg_color="blue")

        self.window70Button.configure(fg_color="green")
        #changes the sequence to Limits (70V)
    def limit70Button(self):#button for limit 70 test 
        self.APx.Sequence.Sequences.Activate("Limits (70V)")
        self.sequence = "Limits (70V)"

        self.limit48Button.configure(fg_color="blue")
        self.customButton.configure(fg_color="blue")
        self.window48Button.configure(fg_color="blue")
        self.window70Button.configure(fg_color="blue")
        
        self.limt70Button.configure(fg_color="green")

        #changes the sequence to Custom
    def customButton(self):#button for custom test. 
        self.APx.Sequence.Sequences.Activate("Custom")
        self.sequence = "Custom"
        self.limit48Button.configure(fg_color="blue")
        self.limt70Button.configure(fg_color="blue")
        self.window48Button.configure(fg_color="blue")
        self.window70Button.configure(fg_color="blue")

        self.customButton.configure(fg_color="green")

        #changes the sequence to Window Test (4-8)
        #activates the listening test
    def startTestButton(self):#starts the testing and sequence of calling all the programs.
       # getText = self.textbox.get("0.0","1.end")
       # getfail = self.failbox.get("0.0","1.end")
       # getpass = self.passbox.get("0.0","1.end")
       #if one of the buttons is selected and both of the drop downs are selected
        if(self.sequence == "Custom"):
            if(self.user_entry_custom != "" and self.user_entry.get() != ""):
                self.textbox.delete("0.0","end")
                self.failbox.delete("0.0","end") 
                self.passbox.delete("0.0","end") 
                self.passedTest = []
                self.failedTest = []
                self.button6.configure(state = "normal")
                self.button7.configure(state = "normal")
                self.button9.configure(state = "normal")
                self.impedenceButton.configure(state= "normal")
                self.rmsButton.configure(state= "normal")
                self.thdButton.configure(state= "normal")
                self.PWDButton.configure(state= "normal")
                self.ListenButton.configure(state= "normal")
                self.printPDF.configure(state= "normal")
                self.printTest()
                self.APx.Sequence.Run() 
                getText2 = self.textbox.get("0.0","1.end")
                self.disableButton()
                self.graphData(getText2)
                
        elif(self.sequence != "" and self.optionmenu_model.get() != "model" and  self.optionmenu_series.get() != "series" ):
            if(self.sequence == "Limits (4-8)" or self.sequence == "Limits (70V)"):

                self.APx.Sequence.Run() 
                self.limitTestMaker()
                
            elif((self.sequence == "Window Test (4-8)" or self.sequence == "Window Test (70V)") and self.user_entry.get() != ""):
                self.textbox.delete("0.0","end")
                self.failbox.delete("0.0","end") 
                self.passbox.delete("0.0","end") 
                self.passedTest = []
                self.failedTest = []
                self.button6.configure(state = "normal")
                self.button7.configure(state = "normal")
                self.button9.configure(state = "normal")
                self.impedenceButton.configure(state= "normal")
                self.rmsButton.configure(state= "normal")
                self.thdButton.configure(state= "normal")
                self.PWDButton.configure(state= "normal")
                self.ListenButton.configure(state= "normal")
                self.printPDF.configure(state= "normal")
                self.printTest()
                self.APx.Sequence.Run() 
                getText2 = self.textbox.get("0.0","1.end")
                self.disableButton()
                self.graphData(getText2)
                
    def passButton(self):# this button will pass the failed test
        getText = self.textbox.get("0.0","1.end")   
        if(getText != ""):  
            plt.close()
            self.passbox.insert("0.0",getText+ '\n')
            self.textbox.delete("0.0","1.end") 
            getText = self.textbox.get(2.0,"end")
            self.textbox.delete("0.0","end") 
            self.textbox.insert("1.0",getText)
            getText2 = self.textbox.get("0.0","1.end")
            self.graphData(getText2)
            
        return True
    def failButton(self):# this button will fail the failed test
        getText = self.textbox.get("0.0","1.end") 
        if(getText != ""):
            plt.close()
            self.failbox.insert("0.0",getText + "\n")
            self.textbox.delete("0.0","1.end") 
            getText = self.textbox.get(2.0,"end")
            self.textbox.delete("0.0","end") 
            self.textbox.insert("1.0",getText) 
            getText2 = self.textbox.get("0.0","1.end")
            self.graphData(getText2)
            
        return True 
    def disableButton(self):#this function will disble the beginning buttons to make sure the sure cannot change the sequence 
        self.modelDropDown.configure(state="disabled")
        self.seriesDropDown.configure(state="disabled")   
        self.window48Button.configure(state="disabled")
        self.limit48Button.configure(state="disabled")
        self.customButton.configure(state="disabled")
        self.window70Button.configure(state="disabled")
        self.limt70Button.configure(state="disabled")
        self.user_entry.configure(state="disabled")
        self.user_entry_custom.configure(state="disabled")
    def enableButton(self):#this will enable all the function and buttons before starting the test
        self.modelDropDown.configure(state="normal")
        self.seriesDropDown.configure(state="normal")   
        self.window48Button.configure(state="normal")
        self.limit48Button.configure(state="normal")
        self.customButton.configure(state="normal")
        self.window70Button.configure(state="normal")
        self.limt70Button.configure(state="normal")
        self.user_entry.configure(state="normal")
        self.user_entry_custom.configure(state="normal")
    def openSoftware(self):# this button opens the software and format of the test.
        self.APx.Visible = True
        
        self.modelDropDown.configure(state="normal")
        self.seriesDropDown.configure(state="normal")   
        self.window48Button.configure(state="normal")
        self.limit48Button.configure(state="normal")
        self.customButton.configure(state="normal")
        self.window70Button.configure(state="normal")
        self.limt70Button.configure(state="normal")
        self.user_entry.configure(state="normal")
        self.button11.configure(state="normal")
        self.Load_button()
    def testSubmitButton(self):#sumbits the report and makes a file in the filepath of this model 
        
        if(self.textbox.get("0.0","1.end") == ""):
            self.user_entry.delete("0","end")
            self.textbox.delete("0.0","end")
            self.failbox.delete("0.0","end")
            self.passbox.delete("0.0","end")     
            self.passedTest = []
            self.failedTest = []
            self.enableButton()
    def ListenTest(self):#this will activate the audo for the listening test. I dont know how to route it to the correct audio output but ill find out in testing phase;

        self.APx.Sequence.Sequences.Activate("Listening Test (4-8)")
        self.APx.Sequence.Run()
        self.APx.Sequence.Sequences.Activate(self.sequence)    
        self.removeWhiteSpace("Listening Test")
    
    def modelDrop(self,select):#Drop down menu functon for moldels 
        self.modelChoice = select
        if self.models[0] == self.modelChoice:
            self.seriesDropDown.configure(values = self.modelSA)
        elif self.models[1] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelAt)
        elif self.models[2] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelBE)
        elif self.models[3] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelOW)
        elif self.models[4] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelCUBE)
        elif self.models[5] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelQX)
        elif self.models[6] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelQXC)
        elif self.models[7] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelSXC)
        elif self.models[8] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelFXA)
        elif self.models[9] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelASX)
        elif self.models[10] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelOT)
        elif self.models[11] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelFX)
        elif self.models[12] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelAL)
        elif self.models[13] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelW)
        elif self.models[14] == self.modelChoice:      
            self.seriesDropDown.configure(values = self.modelOMNI)
        elif self.models[15] == self.modelChoice:     
            self.seriesDropDown.configure(values = self.modelEXTRA)  
    def subModelDrop(self,choose):#Drop down menu functon for submodels
        self.finalFilePath = Directory.GetCurrentDirectory()
        self.seriesChoice = choose 
        return choose              
   
    def createReport(self):#creates the sequence report, print failures and passes, exports to a PDF to the same submit filepath
        fileName = 'C:/Users/jlsadmin/Desktop/Sequence_report.pdf'
        documentTitle = 'Report - TEST'
        title = 'JamesLoudSpeaker QA test:'
        testText = "Test:" + self.sequence
        modelText = "Model:"+ self.modelChoice
        seriesText = "Series:" + self.seriesChoice
        serialText = "Serial:" + self.user_entry.get()
        subText = testText + " " + modelText + " " + seriesText + " " + serialText
        passedString = ""
        failString = ""
        textLines = [ 'THD','Phase smooth','Impedence','RMS','Phase Wrap Diff']
        
        image =  Image.open("C:/Users/jlsadmin/Desktop/testStats/pictures/THD.png")
        image1 = Image.open("C:/Users/jlsadmin/Desktop/testStats/pictures/Phase.png")
        image2 = Image.open("C:/Users/jlsadmin/Desktop/testStats/pictures/Impedance.png")
        image3 = Image.open("C:/Users/jlsadmin/Desktop/testStats/pictures/rms.png")
        image4 = Image.open("C:/Users/jlsadmin/Desktop/testStats/pictures/PhaseWrapDiff1.png")
        image5 = Image.open("C:/Users/jlsadmin/Desktop/Hi-rez_JLS-logo-Red-WEB.png")
        
        image.thumbnail ((275, 500), Image.LANCZOS)
        image1.thumbnail((275, 500), Image.LANCZOS)
        image2.thumbnail((275, 500), Image.LANCZOS)
        image3.thumbnail((275, 500), Image.ANTIALIAS)
        image4.thumbnail((275, 500), Image.LANCZOS)
        image5.thumbnail((200, 300), Image.LANCZOS)
        
        imgFold = [image, image1, image2, image3, image4]
        # creating a pdf object
        pdf = canvas.Canvas(fileName)
 
        # setting the title of the document
        pdf.setTitle(documentTitle)
 
        # registering a external font in python
 
        # creating the title by setting it's font
        # and putting it on the canvas
        pdf.setFont('Helvetica', 36)
        pdf.drawCentredString(300, 800, title)
        
        pdf.setFillColorRGB(0, 0, 0)
        pdf.setFont("Helvetica", 14)
        pdf.drawCentredString(300, 770, subText)
    
    
        text = pdf.beginText(40, 680)
        text.setFont("Helvetica", 18)
        text.setFillColor(colors.red)
       
        pdf.drawString(90, 525, textLines[0]) 
        pdf.drawString(320, 525, textLines[1]) 
        pdf.drawString(90, 295, textLines[2]) 
        pdf.drawString(320, 295, textLines[3]) 
        pdf.drawString(90, 65, textLines[4]) 
    
        pdf.drawInlineImage(imgFold[0], 40,540)
        pdf.drawInlineImage(imgFold[1], 290,540)
        pdf.drawInlineImage(imgFold[2], 40,310)
        pdf.drawInlineImage(imgFold[3], 290,310)
        pdf.drawInlineImage(imgFold[4], 40,80)
        pdf.drawInlineImage(image5, 190,10)
        # drawing a images at the
        # specified (x.y) position
        
        whole_text_pass = self.passbox.get("1.0", "end-1c")
        whole_text_failed = self.failbox.get("1.0", "end-1c")
        linesP = whole_text_pass.split("\n")
        linesF = whole_text_failed.split("\n") 
        for line in linesP: 
            self.passedTest.append(line)
        for line in linesF: 
            self.failedTest.append(line)
           
              
        for x in range(len(self.passedTest)):
            passedString = passedString + " " + self.passedTest[x] 
        x = 0
        for x in range(len(self.failedTest)):
            failString = failString + " " + self.failedTest[x] 
            
        text = pdf.beginText(320, 250)
        text.setFillColor(colors.black)
        text.textLine("Passed tests: ")
        text.setFillColor(colors.green)
        x = 0
        while x < (len(self.passedTest)):
            if (x % 2 == 0 and (x + 2 < len(self.passedTest))):
                text.textLine(self.passedTest[x] + ", " + self.passedTest[x+1])
                x += 2
                
            else: 
                 text.textLine(self.passedTest[x])
                 x += 1
        text.setFillColor(colors.black)         
        text.textLine("Failed tests: ")
        text.setFillColor(colors.darkred)
        x = 0
        while x < (len(self.failedTest)):
            if (x % 2 == 0 and (x + 2 < len(self.failedTest))):
                text.textLine(self.failedTest[x] + ", " + self.failedTest[x+1])
                x += 2
            else: 
                 text.textLine(self.failedTest[x])
                 x += 1         
        text.setFillColor(colors.black)      
        text.textLine("Additonal notes:")
        text.textLine("Tested by:" + self.user_entry1.get())
        pdf.drawText(text) 
        pdf.save()
        subprocess.Popen('C:/Users/jlsadmin/Desktop/Sequence_report.pdf',shell=True)
    def printTest(self):
         x = 0
         for x in range(len(self.tests)):
            index = str(x) + '.0'
            testText = self.tests[x] + '\n'
            self.textbox.insert(index, testText)
            text = self.textbox.get("0.0", "end")  # prints the test each time that you run a sequence test in AP 
         self.textbox.insert("end","Listening Test")      
    
    def showImpedance(self):# this will display the impedance of the test
        self.removeWhiteSpace("Impedance")  
        fig, ax = plt.subplots(figsize=(10,7))
        ax.set_title("IMPEDANCE TEST")
        ax.set_ylabel("DBSL1")
        ax.set_xlabel("Frequency")
        ax.set_xlim(25,20002)
        ax.set_ylim(0,25)

        
        x = [ 30,50,100 , 200 ,300, 500 , 1000 ,2000, 3000 , 5000, 10000, 20000]
        ticks = [30,50, 100 , 200 ,300, 500 , "1k" ,"2k", "3k" , "5k", "10k", "20k"]

        ax.set_xscale('log')
        ax.set_xticks(x)
        ax.set_xticklabels(ticks)
        
        filename = "C:/Users/jlsadmin/Desktop/testStats/outputs/Impedance.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/Impedance.xlsx"
        xAxis = pd.read_excel(filename,usecols='A',skiprows=4,)
        yAxis = pd.read_excel(filename,usecols='B',skiprows=4)
        x1Axis = pd.read_excel(filelimit,usecols='A',skiprows=4)
        y1Axis = pd.read_excel(filelimit,usecols='B',skiprows=4)
        x = pd.DataFrame(xAxis).to_numpy()
        x1 = pd.DataFrame(x1Axis).to_numpy()
        y = pd.DataFrame(yAxis).to_numpy()
        y1 = pd.DataFrame(y1Axis).to_numpy()
        x = x.tolist()
        limitX = x1.tolist()
        y = y.tolist()
        limitY = y1.tolist()
        y2 = []
        upperY = []
        lowerY = [] 
        dataX = []
        dataY = [] 
        #ax.yaxis.set_ticks([120, 100 , 80 , 60 , 40 , 20 , 10 , 0 ])
        #ax.plot(x,y)
        pan_handler = panhandler(fig)
        disconnect_zoom = zoom_factory(ax)
        
        for j in range(len(x)):
            dataY.append(y[j][0])
            dataX.append(x[j][0])
        for k in range(len(limitY)):
            
            g = (limitY[k][0] + 3)
            h = (limitY[k][0] - 3)
            upperY.append(g)
            lowerY.append(h)
        
        
        interX = []
        interY = []
        if(self.sequence != "Custom"):
            for q in range(len(dataY) -1 ):
                if ((dataY[q]) >= (upperY[q]) or (dataY[q]) <= (lowerY[q])):
                    interY.append(dataY[q])
                    interX.append(dataX[q])
                
            
            ax.plot(x1,upperY,color= 'red')
            ax.plot(x1,lowerY,color= 'red')
            plt.plot(interX, interY,'ro')
                
        ax.plot(dataX,dataY)
        
        ax.grid(color="grey")
        ax.grid(axis="x", which='minor',color="grey", linestyle="--")
        fig.canvas.mpl_connect('key_press_event',self.on_key)
        plt.savefig("C:/Users/jlsadmin/Desktop/testStats/pictures/Impedance.png")
        plt.autoscale(enable=True, axis= 'y')
        plt.show()           
    def showRMS(self):#shows the RMS graph 
        self.removeWhiteSpace("RMS")  
        fig, ax = plt.subplots(figsize=(10,7))
        ax.set_title("RMS TEST")
        ax.set_ylabel("DBSL1")
        ax.set_xlabel("Frequency")
        ax.set_xlim(50,20002)
        ax.set_ylim(0,100)

        x = [20,30, 50, 100 , 200 ,300, 500 , 1000 ,2000, 3000 , 5000, 10000, 20000]
        ticks = [20,30, 50, 100 , 200 ,300, 500 , "1k" ,"2k", "3k" , "5k", "10k", "20k"]

        ax.set_xscale('log')
        ax.set_xticks(x)
        ax.set_xticklabels(ticks)
        
        filename = "C:/Users/jlsadmin/Desktop/testStats/outputs/rms.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/rmstest.xlsx"

        xAxis = pd.read_excel(filename,usecols='A',skiprows=4,)
        yAxis = pd.read_excel(filename,usecols='B',skiprows=4)
        x1Axis = pd.read_excel(filelimit,usecols='A',skiprows=4)
        y1Axis = pd.read_excel(filelimit,usecols='B',skiprows=4)
        x = pd.DataFrame(xAxis).to_numpy()
        x1 = pd.DataFrame(x1Axis).to_numpy()
        y = pd.DataFrame(yAxis).to_numpy()
        limitY = pd.DataFrame(y1Axis).to_numpy()
        x = x.tolist()
        limitX = x1.tolist()
        y = y.tolist()
        limitY = limitY.tolist()
        upperY = []
        lowerY = [] 
        dataX = []
        dataY = [] 
        #ax.yaxis.set_ticks([100,90,80,70,60,50,40,30,20,10,0])
        pan_handler = panhandler(fig)
        disconnect_zoom = zoom_factory(ax)
        for j in range(len(x)):
            dataY.append(y[j][0])
            dataX.append(x[j][0])
        for k in range(len(limitY)):
            
            g = (limitY[k][0] + 3)
            h = (limitY[k][0] - 3)
            upperY.append(g)
            lowerY.append(h)
        
        
        
                    
        interX = []
        interY = []
        u = 0 #
        if(self.sequence != "Custom"):
            for q in range(len(limitY)):
                if (dataX[q] > 100):
                    if ((dataY[q]) >= (upperY[u])or (dataY[q]) <= (lowerY[u])):
                        interY.append(dataY[q])
                        interX.append(dataX[q])
                    u += 1
            ax.plot(x1,upperY,color= 'red')
            ax.plot(x1,lowerY,color= 'red')
            plt.plot(interX, interY,'ro')
        ax.plot(dataX,dataY)
        
        
        
        ax.grid(color="grey")
        ax.grid(axis="x", which='minor',color="grey", linestyle="--")
        fig.canvas.mpl_connect('key_press_event',self.on_key)

        
        
        plt.autoscale(enable=True, axis= 'y')
        plt.savefig("C:/Users/jlsadmin/Desktop/testStats/pictures/rms.png")
        plt.show()           
    def showPWD(self):#show the PHSAse Wrap difference
        self.removeWhiteSpace("PhaseWrapDiff")  
        fig, ax = plt.subplots(figsize=(10,7))
        ax.set_title("PhaseWrapDiff")
        ax.set_ylabel("DBSL1")
        ax.set_xlabel("Frequency")
        ax.set_xlim(25,20002)
        ax.set_ylim(0,25)

        x = [ 30,37,47,50,100 , 200 ,300, 500 , 1000 ,2000, 3000 , 5000, 10000, 20000]
        ticks = [30,37,47,50, 100 , 200 ,300, 500 , "1k" ,"2k", "3k" , "5k", "10k", "20k"]

        ax.set_xscale('log')
        ax.set_xticks(x)
        ax.set_xticklabels(ticks)
        
        filename = "C:/Users/jlsadmin/Desktop/testStats/outputs/PhaseWrapDiff1.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/PhaseWrapDiff.xlsx"

        xAxis = pd.read_excel(filename,usecols='A',skiprows=4,)
        yAxis = pd.read_excel(filename,usecols='B',skiprows=4)
        x1Axis = pd.read_excel(filelimit,usecols='A',skiprows=4)
        y1Axis = pd.read_excel(filelimit,usecols='B',skiprows=4)
        x = pd.DataFrame(xAxis).to_numpy()
        x1 = pd.DataFrame(x1Axis).to_numpy()
        y = pd.DataFrame(yAxis).to_numpy()
        y1 = pd.DataFrame(y1Axis).to_numpy()
        x = x.tolist()
        x1 = x1.tolist()
        y = y.tolist()
        y1 = y1.tolist()
        y2 = []
        y3 = []
        
        
        ax.yaxis.set_ticks([180,160,140,120,100,80,60,40,20,0,-20,-40,-60,-80,-100,-120,-140,-160,-180])
        for x4 in range(len(y1)): 
            y2.append(160)
        
        pan_handler = panhandler(fig)
        disconnect_zoom = zoom_factory(ax)
              
        
        for x4 in range(len(y1)): 
            y3.append(-160)
        
        

        ax.plot(x,y) 
        ax.grid(color="grey")
        ax.grid(axis="x", which='minor',color="grey", linestyle="--")
        fig.canvas.mpl_connect('key_press_event',self.on_key)
        plt.savefig("C:/Users/jlsadmin/Desktop/testStats/pictures/PhaseWrapDiff1.png")
        
        interX = []
        interY= []
        if(self.sequence != "Custom"):
            ax.plot(x1,y2,color= 'red')
            ax.plot(x1,y3,color= 'red')
            for z in range(len(x)):
                if (y[z][0] == y2[z] or (y[z-1][0] < y2[z] and y[z][0] > y2[z]) or (y[z-1][0] > y2[z] and y[z][0] < y2[z])):
                    interX.append(float(x[z][0]))
                    interY.append(float(y2[z]))
                if (y[z][0] == y3[z] or (y[z-1][0] < y3[z] and y[z][0] > y3[z]) or (y[z-1][0] > y3[z] and y[z][0] < y3[z])):
                    interX.append(float(x[z][0]))
                    interY.append(float(y3[z]))
            plt.plot(interX, interY, 'ro')    
        plt.show()   
    def showTHD(self):#shows the THD rati0
        self.removeWhiteSpace("THD")        
        fig, ax = plt.subplots(figsize=(10,7))
        ax.set_title("THD")
        ax.set_ylabel("DBSL1")
        ax.set_xlabel("Frequency")
        ax.set_xlim(25,20002)
        ax.set_ylim(0,25)

        x = [ 30,50,100 , 200 ,300, 500 , 1000 ,2000, 3000 , 5000, 10000, 20000]
        ticks = [30,50, 100 , 200 ,300, 500 , "1k" ,"2k", "3k" , "5k", "10k", "20k"]
        
        ax.set_xscale('log')
        ax.set_xticks(x)
        ax.set_xticklabels(ticks)
        
        
        filename = "C:/Users/jlsadmin/Desktop/testStats/outputs/THD.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/THD.xlsx"

        xAxis = pd.read_excel(filename,usecols='A',skiprows=4,)
        yAxis = pd.read_excel(filename,usecols='B',skiprows=4)
        x1Axis = pd.read_excel(filelimit,usecols='A',skiprows=4)
        y1Axis = pd.read_excel(filelimit,usecols='B',skiprows=4)
        x = pd.DataFrame(xAxis).to_numpy()
        x1 = pd.DataFrame(x1Axis).to_numpy()
        y = pd.DataFrame(yAxis).to_numpy()
        y1 = pd.DataFrame(y1Axis).to_numpy()
        x = x.tolist()
        x1 = x1.tolist()
        y = y.tolist()
        y1 = y1.tolist()
        y2 = []
        yVal= [100,10,1,.1, .01,.001]
        yTicks = [100,10,1,.1, .01,.001]
        ax.yaxis.set_ticks([100,10,1,.1, .01,.001])
        
        ax.set_yscale('log')
        ax.set_yticks(yVal)
        ax.set_yticklabels(yTicks)
        
        pan_handler = panhandler(fig)
        disconnect_zoom = zoom_factory(ax)
        
        ax.plot(x,y)
        if(self.sequence != "Custom"):
            for x in range(len(y1)):
                g = (y1[x][0] + 6)
                y2.append(g)
            ax.plot(x1,y1,color= 'red')
            ax.plot(x1,y2,color= 'red')
        ax.grid(color="grey")
        ax.grid(axis="x", which='minor',color="grey", linestyle="--")
        fig.canvas.mpl_connect('key_press_event',self.on_key)
        
        plt.autoscale(enable=True, axis= 'y')
        plt.savefig("C:/Users/jlsadmin/Desktop/testStats/pictures/THD.png")
        plt.show()   
  
    def graphData(self,test):# here is the graph testing making of the graphs.   
            if(test == "RMS"):
                self.showRMS()
            elif(test == "Impedance"):
                self.showImpedance()
            elif(test == "PhaseWrapDiff"):
                self.showPWD()
            elif(test == "THD"):
                self.showTHD()
            elif(test == "Listening Test"):
                self.closePlot()                    
    def removeWhiteSpace(self, testName):
        #if Phase name is in pass box or fail box remove it.
        #if phase name is already in the begin
        whole_text = self.textbox.get("1.0", "end-1c")
        x = 0
        lines = whole_text.split("\n")
        for line in lines: 
            if (line.strip() == testName):
                break
            if line.strip() == "":
                    index1 = str(x+1) +'.0'
                    self.textbox.insert(index1,testName)
                    break
            x += 1     
            
        for x in range(1,10): 
            index1 = str(x) +'.0'
            index2 = str(x+1)+ '.0'
            # index3 = str(x+2)+ '.0'
            if(testName in self.passbox.get(index1, index2)):
                self.passbox.delete(index1,index2) 
            if(testName in self.failbox.get(index1, index2)):
                self.failbox.delete(index1, index2)
            
        #BLANK LINE            
    def on_key(self, event):
        if event.key == 'right':
            plt.close()
            self.passButton()
        elif event.key == 'left':
            # Perform left arrow key action
            plt.close()
            self.failButton()
        elif event.key == 'q':
            # Close the window when 'q' key is pressed
            plt.close()   
    def limitTestMaker(self): 
        #get data from test.
        #replace data from test and save to limits
        filename = "C:/Users/jlsadmin/Desktop/testStats/limitsTestExport/Impedance.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/Impedance.xlsx"
        dest = shutil.copyfile(filename, filelimit)

        filename = "C:/Users/jlsadmin/Desktop/testStats/limitsTestExport/Phase.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/Phase.xlsx"
        dest = shutil.copyfile(filename, filelimit)

        filename = "C:/Users/jlsadmin/Desktop/testStats/limitsTestExport/THD.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/THD.xlsx"
        dest = shutil.copyfile(filename, filelimit)
  
        filename = "C:/Users/jlsadmin/Desktop/testStats/limitsTestExport/PhaseSmooth.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/PhaseWrapDiff.xlsx"
        dest = shutil.copyfile(filename, filelimit)
    
        filename = "C:/Users/jlsadmin/Desktop/testStats/limitsTestExport/RMS.xlsx"
        filelimit = "C:/Users/jlsadmin/Desktop/testStats/limits/RMStest.xlsx"
        dest = shutil.copyfile(filename, filelimit)
    
        #test to see if this works.  
    def closePlot(self):
        plt.close()
        self.ListenTest()
    def openHelp(self):
        pdf_path = "C:/Users/jlsadmin/Downloads/APXtesterTutorial.pdf"
        subprocess.Popen([pdf_path], shell=True)
    def reset(self):# this function resets all the test for the user. if they make a mistake
        self.enableButton()
        
        self.user_entry.delete("0","end")
        self.user_entry_custom.delete("0","end")
        self.textbox.delete("0.0","end")
        self.failbox.delete("0.0","end")
        self.passbox.delete("0.0","end")    
if __name__ == "__main__":
    app = App()
    app.mainloop
