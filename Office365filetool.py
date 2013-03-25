#!/usr/bin/python
# -*- coding: utf8 -*-

import wx
import os
import sys
import time
from zipfile import ZipFile

blockedExts = [".ade", ".adp", ".app", ".asa", ".ashx", ".asmx", ".asp", ".bas", ".bat",
               ".cdx", ".cer", ".chm", ".class", ".cmd", ".com", ".config", ".cpl", ".crt",
               ".csh", ".dll", ".exe", ".fxp", ".hlp", ".hta", ".htr", ".htw", ".ida",
               ".idc", ".idq", ".ins", ".isp", ".its", ".jse", ".ksh", ".lnk", ".mad",
               ".maf", ".mag", ".mam", ".maq", ".mar", ".mas", ".mat", ".mau", ".mav",
               ".maw", ".mda", ".mdb", ".mde", ".mdt", ".mdw", ".mdz", ".msc", ".msh",
               ".msh1", ".msh1xml", ".msh2", ".msh2xml", ".mshxml", ".msi", ".msp", ".mst",
               ".ops", ".pcd", ".pif", ".prf", ".prg", ".printer", ".pst", ".reg", ".rem",
               ".scf", ".scr", ".sct", ".shb", ".shs", ".shtm", ".shtml", ".soap", ".stm",
               ".url", ".vb", ".vbe", ".vbs", ".ws", ".wsc", ".wsf", ".wsh"]

blockedSymbols = ["#", "{", "}", "%", "~", "&", "¤", "!", "=", "@", "£", "*", ":", "<", ">", "?", "|", "\""]

blockedPaths = ['c', 'c:', 'c:\\', 'c:/', 'test']


# Function to find things in a list
def find(f, seq):
    found = False
    for item in seq:
        if item == f:
            found = item

    return found


class MainWindow(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, "Office 365 File Analyser", size=(400,400))
        panel = wx.Panel(self, wx.ID_ANY)

        dirbtn = wx.Button(panel, label="Choose the Directory", pos=(25,40), size=(330, 30))
        dirbtn.Bind(wx.EVT_BUTTON, self.chooseDir)

        self.dirchooseText = wx.StaticText(panel, -1, "No directory choosen.", pos=(25,20))

        self.sharepointHelp = wx.StaticText(panel, -1, "Sharepoint URL:", pos=(25, 80))
        self.sharepointUrl = wx.TextCtrl(panel, -1, pos=(25,100), size=(330,25),name=("Kevin"))

        reportbtn = wx.Button(panel, label="Generate Report", pos=(25, 170), size=(330,45))
        reportbtn.Bind(wx.EVT_BUTTON, self.report)

        self.openreport = wx.CheckBox(panel, -1, "Open report when finished", pos=(25, 225))

        docleanupbtn = wx.Button(panel, label="Process the files", pos=(25, 255), size=(330,45))
        docleanupbtn.Bind(wx.EVT_BUTTON, self.docleanup)
    
    def chooseDir(self, event):

        dlg = wx.DirDialog(self, message="Choose the Directory")

        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.dirchooseText.SetLabel(path)
        dlg.Destroy()

    def report(self,event):

        if(self.dirchooseText.GetLabel() == "No directory choosen."):
            wx.MessageBox('You have not choosen any folder!', 'Info', 
        wx.OK | wx.ICON_INFORMATION)


        if(os.path.exists("report.csv")):
            try:
                os.remove("report.csv")
            except IOError as e:
                pass
            
        with open("report.csv", "a") as f:
            f.write("ERRORS\n")

            path = self.dirchooseText.GetLabel()

            for (path, dirs, files) in os.walk(path.encode()):

                # Go through all files          
                for file in files:

                    fileName, fileExtension = os.path.splitext(file)

                    for blockedSymbol in blockedSymbols:

                        # If blockedSymbol is found
                        if file.find(blockedSymbol) != -1:
                            f.write("Symbol;" + path + file + "\n")
                            break
                    
                    for blockedExt in blockedExts:

                        # Looks after if the file extension is in the blocked list    
                        if fileExtension == blockedExt:
                            f.write("File Extension;" + path + "\\" + file + "\n")
                            break
                    

                    # Calculating the length of the final http path
                    # Converting normal spaces to %20 because that is space in the browser
                    filelengthString = os.path.join(path, file)
                    filelengthString = filelengthString.replace(" ", "   ")
                    filelength = len(filelengthString)

                    # Checks for if the final URL is longer than 256 

                    if(filelength - len(self.dirchooseText.GetLabel()) > 256 - len(self.sharepointUrl.GetValue())):
                        f.write("Lenght;" + os.path.join(path, file) + "\n")


            f.close()

        if(self.openreport.GetValue() == True):
            os.startfile("report.csv")

    def docleanup(self, event):
        # Checks for blockedPaths, so you do not accidently zip your Windows folder
        if(find(self.dirchooseText.GetLabel(), blockedPaths) == False):
            
            for (path, dirs, files) in os.walk(self.dirchooseText.GetLabel()):

                # Go through all files          
                for file in files:

                    fileName, fileExtension = os.path.splitext(file)

                    for blockedSymbol in blockedSymbols:

                        # If blockedSymbol is found
                        if file.find(blockedSymbol) != -1:

                            newfilename = file.replace(blockedSymbol,"")

                            # Tries to remove the sign
                            try:
                                os.rename(os.path.join(path, file),os.path.join(path,newfilename))
                                file = newfilename
                                
                            except:
                                print "Kunne ikke omdoebe fil: ", sys.exc_info()
                                newfilename = file.replace(blockedSymbol, "_")
                                os.rename(os.path.join(path, file),os.path.join(path,newfilename))
                                file = newfilename

                    
                    for blockedExt in blockedExts:

                        # Looks after if the file extension is in the blocked list    
                        if fileExtension == blockedExt:
                            newzipfilename = file.replace(blockedExt, blockedExt + ".zip")
                            zippath = os.path.join(path, newzipfilename)
                            
                            # Starts a Zipp file with the same name
                            zf = ZipFile(zippath, mode='w')
                            try:
                                zf.write(os.path.join(path, file), arcname = file)
                            finally:
                                zf.close()
                                os.remove(os.path.join(path, file))

        wx.MessageBox('The program has finished!', 'Info', wx.OK | wx.ICON_INFORMATION)
        
if __name__ == "__main__":
    app = wx.App(False)
    frame = MainWindow()
    frame.Show()
    app.MainLoop()