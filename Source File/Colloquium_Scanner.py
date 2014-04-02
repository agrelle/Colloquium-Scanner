"""  
     NE COLLOQUIUM SCANNER - A program to collect and report student attendance

     This program is free software: you can redistribute it and/or modify
     it under the terms of the GNU General Public License as published by
     the Free Software Foundation, either version 3 of the License, or
     (at your option) any later version.

     This program is distributed in the hope that it will be useful,
     but WITHOUT ANY WARRANTY; without even the implied warranty of
     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     GNU General Public License for more details.

     You should have received a copy of the GNU General Public License
     along with this program.  If not, see <http://www.gnu.org/licenses/>.

  Notes for Use:
 
  - This is a Python program. You must have the Python interpreter installed
    on your machine. If you are using a Windows machine, you can find the appropriate
    version at: <http://www.python.org/> If you are using a Unix machine, it is
    likely that you already have a Python interpreter installed on your machine.

       Appropriate Versions: Python 2.7.x
       Recommended Version:  Python 2.7.3

  - Version restriction is limited to 2.x due to use of wxPython. This program
    extensively uses the wxPython module and as such mandates its installation
    on your local machine. wxPython can be found at <http://wxpython.org/> and
    is available for the latest versions of Python 2.x
 
       wxPython is not available on Python 3.x to my knowledge at this time, but
       if in the future it is, this program MAY or MAY NOT work on it.


"""

# Mandatory system and os dependent path type operation functions
import sys, os
# Sometimes Python2.6 has trouble importing wx.....
if sys.version_info[2] < int(7): sys.path.append('C:\Python26\Lib\site-packages\wx-2.8-msw-unicode')
# Import some Python functions related to the following:
import string, math, wx, time, wx.grid
#from threading import Thread  <--- Old import (depreciated)

class main(wx.Frame):
    def __init__(self, title):
        wx.Frame.__init__(self, None, title="NE Colloquium ID Scanner",
                          size=(640,480),
                          style=wx.DEFAULT_FRAME_STYLE^wx.RESIZE_BORDER^wx.MAXIMIZE_BOX)
        try:
            icon = wx.Icon('Images\\icon.ico', wx.BITMAP_TYPE_ICO,16,16)
            self.SetIcon(icon)
        except: pass
        self.parent = self
        self.status = False
        # Default station ID is 'A'. If you have more than one scanning station,
        # you would ideally identify them differently (example: A, B, C, 1, 2, etc)
        self.id = 'A'
        # Pointer to the file which contains the Comma Separated Values of format:
        #     Student Name , Student ID (newline character)
        self.ID_FILE = None
        # List populated with all students. May contain duplicates without concern
        self.Students = []
        # List of all attendance files. Cannot contain duplicates for accurate answers
        self.ATT_FILE = []
        # List of all ID's in attendance. Will ideally contain duplicates (one entry per
        # attendance instance)
        self.ATT_LIST = []
        # Maximum Colloquiums possible. If set to 0, the maximum student attendance value
        # will be used instead.
        self.max_value = 0
        
        # Create a menu bar with File, Data, and Help elements
        menuBar = wx.MenuBar()
        # File menu contains Station Identification, Program Restart, and Program Exit
        file_menu = wx.Menu()
        self.menuStation = file_menu.Append(wx.ID_ANY, "&Identify Station", "Enter a station ID (Current = " + self.id + " )")
        menuReset = file_menu.Append(wx.ID_REVERT, "&Restart Program", "Reset all Program values")
        menuExit = file_menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Close window and exit program.")
        # Data menu contains Total Number of Colloquiums, Option to Load Student ID file,
        #    Option to Load Attendance File(s), and if both Student ID file and Attendance files
        #    have been loaded, then an option to run a report.
        data_menu = wx.Menu()
        self.menuNumber = data_menu.Append(wx.ID_ANY, "&Total Colloquiums", "Enter Amount of Colloquiums. (Current = " + str(self.max_value) + " )")
        menuLoadPpl = data_menu.Append(wx.ID_ANY, "&Load Students", "Load student ID file")
        menuLoadAttn = data_menu.Append(wx.ID_ANY, "Load Attendance(s)", "Load all attendance files")
        self.menuRun = data_menu.Append(wx.ID_ANY, "Run Report", "Run an attendance report")
        # Help menu contains options to show the License, Help, or About information
        help_menu = wx.Menu()
        menuLicense = help_menu.Append(wx.ID_ANY, "&License", "View License information for this program.")
        menuHelp = help_menu.Append(wx.ID_HELP_CONTENTS, "&Help", "View program Help file.")
        menuAbout= help_menu.Append(wx.ID_ABOUT, "&About"," View information about this program.")
        # Append the three menu items and set the menu bar
        menuBar.Append(file_menu, "&File")
        menuBar.Append(data_menu, "&Data")
        menuBar.Append(help_menu, "&Help")
        self.SetMenuBar(menuBar)
        self.menuRun.Enable(False)

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.OnStation, self.menuStation)
        self.Bind(wx.EVT_MENU, self.OnReset, menuReset)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_MENU, self.OnNumber, self.menuNumber)
        self.Bind(wx.EVT_MENU, self.OnLoadPpl, menuLoadPpl)
        self.Bind(wx.EVT_MENU, self.OnLoadAttn, menuLoadAttn)
        self.Bind(wx.EVT_MENU, self.OnRun, self.menuRun)
        self.Bind(wx.EVT_MENU, self.OnLicense, menuLicense)
        self.Bind(wx.EVT_MENU, self.OnHelp, menuHelp)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

        # Create the status bar
        self.setstatusbar = self.CreateStatusBar()

        # Create the main panel, and make it referenceable
        self.panel = wx.Panel(self)
        self.panel.parent = self
        box = wx.BoxSizer(wx.VERTICAL)

        # Create the panel with information
        panel_info = wx.Panel(self.panel,style=wx.SUNKEN_BORDER)
        sizer_info_v = wx.BoxSizer(wx.VERTICAL)
        sizer_info_h = wx.BoxSizer(wx.HORIZONTAL)

        # Create the panel with the logo
        panel1 = wx.Panel(self.panel)
        box1 = wx.BoxSizer(wx.VERTICAL)
        try:
            image1 = wx.Image("Images\\Main_Frame.jpg", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            wx.StaticBitmap(panel1, -1, image1)
        except: pass
        
        # Create the panel with the button
        panel2 = wx.Panel(panel_info)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        self.buttonp2 = wx.Button(panel2,-1,label="Acquire PUIDs",size=(200,50),style=wx.EXPAND)
        self.buttonp2.Bind(wx.EVT_BUTTON, self.OnAquire)

        # Create a horizontal sizer for panel 2
        sizer_info_h.AddStretchSpacer(1)
        sizer_info_h.Add(panel2, 1, wx.ALL|wx.EXPAND)
        sizer_info_h.AddStretchSpacer(1)

        # Create a panel to act as a color notification panel
        self.panel_color = wx.Panel(panel_info,style=wx.SUNKEN_BORDER)
        self.panel_color.SetBackgroundColour(wx.BLACK)

        # Create a panel to list the status of the program
        panel_status = wx.Panel(panel_info)
        self.status_text = wx.StaticText(panel_status, -1, label="     Status: Passive Mode")

        # Create a panel to list ID Scan Errors
        panel_IDscan = wx.Panel(panel_info)
        self.id_text = wx.StaticText(panel_IDscan, -1, label="ID Scan:  Press Acquire PUIDs")

        # Create a panel to handle text input
        panel_textinput = wx.Panel(panel_info)        
        self.textfield = wx.TextCtrl(panel_textinput,-1,style=wx.TE_PROCESS_ENTER,size=(30,15))
        self.Bind(wx.EVT_TEXT_ENTER,self.OnEnter)

        # Create and set a sizer for the horizontal status panels
        status_sizer = wx.BoxSizer(wx.HORIZONTAL)
        status_sizer.AddStretchSpacer(3)
        status_sizer.Add(self.panel_color, 1)
        status_sizer.Add(panel_status, 3)
        status_sizer.AddStretchSpacer(1)
        status_sizer.Add(panel_IDscan, 4)
        status_sizer.AddStretchSpacer(1)
        status_sizer.Add(panel_textinput, 1)
        status_sizer.AddStretchSpacer(2)

        # Add the information and buttons to the vertical sizer
        sizer_info_v.AddStretchSpacer(2)
        sizer_info_v.Add(sizer_info_h, 3, wx.ALL|wx.EXPAND)
        sizer_info_v.AddStretchSpacer(1)
        sizer_info_v.Add(status_sizer, 1, wx.ALL|wx.EXPAND)
        sizer_info_v.AddStretchSpacer(2)

        # Set the informational panel sizer to the panel
        panel_info.SetSizer(sizer_info_v)

        # Add the panel with the logo and the information
        # panel to the main sizer
        box.Add(panel1)
        box.Add(panel_info, 1, wx.ALL|wx.EXPAND)

        # Set the main sizer in the main panel and draw it
        self.panel.SetSizer(box)
        self.panel.Layout()

        # END OF CREATING THE MAIN WINDOW CONTENTS

    # Determine what happens when you click the program button
    def OnAquire(self,event):
        x = self.parent.buttonp2.GetLabel()
        if x == "Acquire PUIDs":
            self.parent.f = open(''.join(
                ["attendance",'_',self.parent.id,'_','_'.join([str(time.localtime()[i]) for i in range(6)]),'.txt']),
                                 'w')
            self.parent.buttonp2.SetLabel("Stop Acquisition")
            self.parent.id_text.SetLabel("ID Scan:   READY")
            self.parent.status_text.SetLabel("     Status:     ACTIVE")
            self.parent.panel_color.SetBackgroundColour(wx.GREEN)
            self.parent.panel.Refresh()
            self.status = True
            self.parent.textfield.SetFocus()
        else:
            self.parent.panel_color.SetBackgroundColour(wx.BLACK)
            self.parent.panel.Refresh()
            self.parent.buttonp2.SetLabel("Acquire PUIDs")
            self.parent.id_text.SetLabel("ID Scan:  Press Acquire PUIDs")
            self.parent.status_text.SetLabel("     Status: Passive Mode")
            self.parent.panel.Refresh()
            self.status = False
            self.parent.f.close()

    # Lets use an event to poll for PUIDs
    def OnEnter(self,event):
        x = self.parent.textfield.GetValue()
        self.parent.textfield.Clear()
        if len(string.split(x,'=')) == 4: y = string.split(x,'=')[2]
        else: y = False
        if not y:
            self.parent.panel_color.SetBackgroundColour(wx.RED)
            self.parent.id_text.SetLabel("ID Scan: RETRY SCAN")
            self.parent.panel.Refresh()
            self.parent.textfield.SetFocus()
        else:
            self.parent.panel_color.SetBackgroundColour(wx.GREEN)
            self.parent.id_text.SetLabel("ID Scan:   READY")
            self.parent.f.write(''.join([y,'\n']))
            self.parent.panel.Refresh()
            self.parent.textfield.SetFocus()
        
    # Define what happens when you click the "about" menu item
    def OnAbout(self,event):
        text = "NE Colloquium Scanner v1.0\n\nCopyright (C) 2012\nAustin L Grelle\n\nA program to gather, parse,\nsave PUIDs and generate student\nattendance reports for use\nin NE Colloquiums."
        dlg = wx.MessageDialog(self, text, "About NE Colloquium ID Scanner", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

    # Define what happens when you click the "exit" menu item, or click the X
    def OnExit(self, event):
        dlg = wx.MessageDialog(self,
                               "Do you really want to close this application?",
                               "Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            self.Destroy()

    # Define what happens when you select the Station button
    def OnStation(self,event):
        dlg = wx.TextEntryDialog(self,
                                 message = "Suggested Values are A, B, C, etc...",
                                 caption = "Enter Station ID",
                                 defaultValue = self.parent.id,
                                 )
        dlg.ShowModal()
        dlg.Destroy()
        self.parent.id = dlg.GetValue()
        self.parent.menuStation.SetHelp("Enter a station ID (Current = " + self.parent.id + " )")

    # Define what happens when you want to reset the program
    def OnReset(self,event):
        self.parent.ID_FILE = None
        self.parent.ATT_FILE = []
        self.parent.Students = []
        self.parent.ATT_LIST = []
        self.parent.ID = 'A'
        self.status = False
        self.max_value = 0
        self.parent.menuRun.Enable(False)

    # Define what happens when you want to set the total amount of Colloquiums
    def OnNumber(self,event):
        dlg = wx.TextEntryDialog(self,
                                 message = "Entering 0 will result in the highest attendance being used.",
                                 caption = "How many Colloquiums were there?",
                                 defaultValue = str(self.parent.max_value))
        dlg.ShowModal()
        dlg.Destroy()
        self.parent.max_value = int(dlg.GetValue())
        if self.parent.max_value < 0: self.parent.max_value = 0
        self.parent.menuNumber.SetHelp("Enter Amount of Colloquiums. (Current = " + str(self.max_value) + " )")

    # Define what happens when you want to load the Student ID file
    def OnLoadPpl(self,event):
        dlg = wx.FileDialog(self,
                            message = "Load Student ID Data",
                            defaultDir = os.getcwd(),
                            defaultFile = "",
                            wildcard = "Text Files (*.txt)|*.txt",
                            style = wx.OPEN)
        dlg.ShowModal()
        dlg.Destroy()
        self.parent.ID_FILE = dlg.GetPath()
        f = open(self.parent.ID_FILE, 'r')
        students = []
        # Read in all lines from file
        for line in f:
            if string.split(line,'\n') > 1:
                students.append(string.split(line,'\n')[0])
            else: students.append(line)
        # For each file entry, create a student
        # --Each student has a name and ID
        for student in students:
            self.parent.Students.append(
                StudentClass(name = string.split(student,',')[0],
                             ID = string.split(student,',')[1]))
        # Have we loaded IDs and Attendances? If so,
        # the report can be ran, so un-grey the option
        if len(self.parent.ATT_FILE) > 0:
            self.parent.menuRun.Enable(True)

    # Define what happens when you want to load the attendance files
    def OnLoadAttn(self,event):
        dlg = wx.FileDialog(self,
                            message = "Select Attendance File(s)",
                            defaultDir = os.getcwd(),
                            defaultFile = "",
                            wildcard = "Text Files (*.txt)|*.txt",
                            style = wx.OPEN | wx.MULTIPLE)
        dlg.ShowModal()
        dlg.Destroy()
        self.parent.ATT_FILE = dlg.GetPaths()
        # For each attendance file, extract the IDs
        # and add them all to a list
        for each_file in self.parent.ATT_FILE:
            contents = []
            f = open(each_file, 'r')
            for line in f:
                x = string.split(line,'\n')
                if contents.count(int(x[0])) == 0:
                    contents.append(int(x[0]))
            # Store the list in a globally available variable
            self.parent.ATT_LIST = self.parent.ATT_LIST + contents
        # Have we loaded IDs and Attendances? If so,
        # the report can be ran, so un-grey the option
        if self.parent.ID_FILE:
            self.parent.menuRun.Enable(True)

    # If a Student ID file and Attendance File(s) are loaded,
    # then you can run a report. Define what happens when you
    # want to run a report
    def OnRun(self,event):
        # Update the attendance value for each StudentClass
        #  based on the attendance files open
        for each in self.parent.Students:
            each.ATTN = self.parent.ATT_LIST.count(each.ID)
        # Most of the Report functionality is in a separate window
        # Create and show that window
        self.window = ReportWindow(self)
        self.window.Show()

    # Define what happens if you want to view the license information
    def OnLicense(self,event):
        dlg = wx.MessageDialog(self, "To view the license,\nsee license.txt", "License", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

    # Define what happens if you want to see help information
    def OnHelp(self,event):
        dlg = wx.MessageDialog(self, "To view Help,\nplease see Documentation.pdf", "Help", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

# Create a new window which handles the Reporting functionality
class ReportWindow(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title="Attendance Report",
                          size=(424,600),
                          style=wx.DEFAULT_FRAME_STYLE^wx.RESIZE_BORDER^wx.MAXIMIZE_BOX)
        self.parent = parent
        # This panel will house the table of students and attendance values
        self.panel = wx.Panel(self,size=(424,500))
        self.panel.parent = self.parent
        # This panel will house the text-file save button and color indication text
        self.panel2 = wx.Panel(self,size=(424,100))
        self.panel2.parent = self.parent

        # Elements and spacing for the bottom panel with the botton and
        # text indications
        box1 = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        self.button = wx.Button(self.panel2,-1,label="Save as .txt File",size=(100,50),style=wx.EXPAND)
        self.button.Bind(wx.EVT_BUTTON, self.OnSave)
        box2.AddSpacer(25)
        box2.Add(self.button,5)
        box2.AddSpacer(25)
        panel_status = wx.Panel(self.panel2,size=(300,100))
        self.status_text = wx.StaticText(panel_status,
                                         -1,
                                         label="White  =     No Concern\nYellow =     Missed 1 Colloquium\nGreen  =     Record-keeping Problem\nRed     =     Missed >1 Colloquiums")
        box2.Add(panel_status,10)
        box2.AddSpacer(100)
        box1.AddSpacer(10)
        box1.Add(box2)
        box1.AddSpacer(10)
        self.panel2.SetSizer(box1)
        
        # Add the panels to the window and call the Layout
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.panel)
        sizer.Add(self.panel2)
        sizer.AddStretchSpacer(1)
        self.SetSizer(sizer)
        self.Layout()

        # Create the grid within the first panel
        self.grid = wx.grid.Grid(self.panel,-1,size=(414,500))
        self.grid.CreateGrid(0,2)
        self.grid.SetColSize(0,234)
        self.grid.ForceRefresh()
        self.grid.SetColLabelValue(0,"Student Name")
        self.grid.SetColLabelValue(1,"Attendance")
        # Determine what the maximum number of colloquiums
        # possible are. If listed as 0...
        if self.parent.max_value > 0:
            max_value = self.parent.max_value
        # Then find the student with the highest attedance
        # and set the maximum value to him/her
        else:
            max_value = 0
            for each in self.parent.Students:
                if each.ATTN > max_value:
                    max_value = each.ATTN
        # Now, for each student...
        for i in range(len(self.parent.Students)):
            # Add a row to the Grid
            self.grid.AppendRows(1)
            
            # If the student has non-perfect attendance,
            # set their cell colors to the correct indication
            if self.parent.Students[i].ATTN < max_value:
                # Sets the cell to Yellow if the student has only
                #   missed one Colloquium
                if (self.parent.Students[i].ATTN + 1) == max_value:
                    self.grid.SetCellBackgroundColour(i, 0, wx.Colour(255,255,0))
                    self.grid.SetCellBackgroundColour(i, 1, wx.Colour(255,255,0))
                # Sets the cell to Red if the student has missed
                #   more than one Colloquium
                else:
                    self.grid.SetCellBackgroundColour(i, 0, wx.RED)
                    self.grid.SetCellBackgroundColour(i, 1, wx.RED)
            # Sets the cell to Green if the student has swiped his/her
            #  card an amount greater than the total number of
            #  colloquiums. (Indicates a record-keeping problem)
            elif self.parent.Students[i].ATTN > max_value:
                self.grid.SetCellBackgroundColour(i, 0, wx.GREEN)
                self.grid.SetCellBackgroundColour(i, 1, wx.GREEN)

            # Now, fill the first cell with the student's name
            self.grid.SetCellValue(i, 0, self.parent.Students[i].name)
            # Set it to read only so you can't click on it and edit it
            self.grid.SetReadOnly(i,0,True)
            # Fill in the second cell with the student's attendance value
            self.grid.SetCellValue(i, 1, str(self.parent.Students[i].ATTN))
            self.grid.SetReadOnly(i,1,True)

    # If the Save button is clicked, this is what happens
    def OnSave(self,event):
        # First, open a file dialog to determine where to save the file
        dlg = wx.FileDialog(self,
                            message = "Select or Create Save File",
                            defaultDir = os.getcwd(),
                            defaultFile = "",
                            wildcard = "Text Files (*.txt)|*.txt",
                            style = wx.SAVE)
        dlg.ShowModal()
        dlg.Destroy()
        # Its possible the user didn't end up giving a file name,
        # such as if they "cancelled" the window.
        # --> Therefore, check if a filename was given
        if dlg.GetFilename():
            # Determine what the maximum number of colloquiums
            # possible are. If listed as 0...
            if self.parent.max_value > 0:
                max_value = self.parent.max_value
            # Then find the student with the highest attedance
            # and set the maximum value to him/her
            else:
                max_value = 0
                for each in self.parent.Students:
                    if each.ATTN > max_value:
                        max_value = each.ATTN
            # Open the file and write data in a readable format
            f = open(dlg.GetFilename(), 'w')
            f.write("      ----NE COLLOQUIUM REPORT----\n\n")
            f.write("Number of Colloquiums = " + str(self.parent.max_value) + "\n")
            f.write("(If value is zero, then the student with the\n")
            f.write(" most attendance was used for the max value)\n\n")
            f.write("Student ID File Used:  " + str(self.parent.ID_FILE) + "\n\n")
            f.write("Attendance File(s) Used:\n\n")
            for each in self.parent.parent.ATT_FILE:
                f.write("     " + str(each) + "\n")
            f.write("\n")
            f.write("--------------------------------------------------\n")
            f.write("       STUDENT NAME       |     Attendance\n")
            f.write("--------------------------------------------------\n\n")
            for i in range(len(self.parent.Students)):
                name = self.parent.Students[i].name
                value = self.parent.Students[i].ATTN
                if self.parent.Students[i].ATTN < max_value:
                    # Adds a single star (*) after the student's value if it is
                    # one Colloquium less than the maximum
                    if (self.parent.Students[i].ATTN + 1) == max_value:
                        f.write("   " + name + " "*(30 - len(name)) + str(value) + "*" + "\n")
                    # Adds two stars (**) after a student's value if it is
                    # more than one Colloquium less than the maximum
                    else:
                        f.write("   " + name + " "*(30 - len(name)) + str(value) + "**" + "\n")
                # Adds two stars (**) before the student's value if he/she has
                # more than the maximum attendances recorded (indicates a record-keeping
                # problem)
                elif self.parent.Students[i].ATTN > max_value:                    
                        f.write("   " + name + " "*(30 - len(name)) + "**" + str(value) + "\n")
                else: f.write("   " + name + " "*(30 - len(name)) + str(value) + "\n")
            f.close()

# A custom, addressable class with 'name', 'ID', and optional attendance value 'ATTN'
class StudentClass:
    def __init__(self, name, ID, ATTN = 0):
        self.name = name
        self.ID = int(ID)
        self.ATTN = ATTN

# Create the Windows application instance
app = wx.App(redirect=False)
# Create the main frame
frame = main(None)
# Show the main frame
frame.Show()
# Run the application
app.MainLoop()

