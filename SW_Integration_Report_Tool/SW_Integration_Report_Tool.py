
#from PyQt5 import uic
#from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets  import *
#from PyQt5.QtGui  import *
from PyQt5.uic  import  loadUi
from PyQt5.QtCore import QObject, QThread, pyqtSignal , QRunnable, Qt, QThreadPool ,QMutex
from PyQt5 import QtCore, QtGui
#from PyQt5.QtGui import QMovie
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QLineEdit, QMessageBox, QFileDialog
#from matplotlib.backends.backend_qt5agg  import  ( NavigationToolbar2QT  as  NavigationToolbar )
#import  numpy  as  np
#import  random
#import pandas as pd
import sys
import os
from Report_Analysis import*
from datetime import datetime
from SystemFunction_SWC_Report import SWC_SF_Report_Generation_Runnable
from PolarionPlanAPIs import Polarion_Plan_Runnable
import time
#polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
from RTE_API import Task , Task_Flag
import os, shutil

###############################################################################
###################### global variables #######################################

Version={
    "Version": "Version : 4.0.0",
}

Input_Data ={
    "Project":"VW_MEB_Inverter" ,
    "User name": " ",
    "Password": " ",
    "Int_Plan_Doc": " ",
    "Polarion_SW_Plan": " ",
    "RB": "" ,
    "L_Repo": "",
    "Directory": ",",
    "Report_Name": "",
    "SWA_Baseline": "",
    "SF_Mapping": "",
}

Warning_Message ={
    "Warning" : "" ,
}


#########################################################################################
######################### initialization ###############################################

# get current directory
current_directory = os.getcwd()

# Generate Outputs folder
final_directory = os.path.join(current_directory, r'Outputs')
if not os.path.exists(final_directory):
   os.makedirs(final_directory)

# Generate SubFiles folder
final_directory2 = os.path.join(current_directory, r'SubFiles')
if not os.path.exists(final_directory2):
   os.makedirs(final_directory2)

Directory = current_directory.replace("\\", "\\\\")

# delete old output and clear output folders
deleted_directory1 = current_directory+'\SubFiles'
deleted_directory2 = current_directory+'\Outputs'

Delete_Directory_File(deleted_directory1)
Delete_Directory_File(deleted_directory2)



#########################################################################################
# 1. Subclass QRunnable
# main runnable
class Runnable(QRunnable):
    def __init__(self, n):
        super().__init__()
        self.n = n


    def run(self):
        # Get polarion Plan
        infoList1 = []
        infoList1.append(Input_Data["User name"])
        infoList1.append(Input_Data["Password"])
        infoList1.append(Input_Data["Project"])
        infoList1.append(Input_Data["Polarion_SW_Plan"])
        infoList1.append(Input_Data["SWA_Baseline"])

        infoList2 = []
        infoList2.append(Input_Data["User name"])
        infoList2.append(Input_Data["Password"])
        infoList2.append(Input_Data["Project"])
        infoList2.append(Input_Data["SWA_Baseline"])
        infoList2.append(Input_Data["Polarion_SW_Plan"])
        infoList2.append(Input_Data["RB"])
        # Main task
        Main_Task(infoList1, infoList2, Input_Data)
        pass


# main GUI

class MatplotlibWidget(QMainWindow):

    def __init__(self):
        ##print("clicked in init")
        QMainWindow.__init__(self)

        loadUi("Main_Window.ui", self)
        self.show()
        self.UiTimer()
        self.setWindowTitle("SWA Release Tool")

        # creat menue bars
        self.About_Menue = About()
        # tool Bar
        toolbar = QToolBar("Menue tools")
        self.addToolBar(toolbar)
        button_action = QAction("About", self)
        button_action.setStatusTip("About Tool")
        button_action.triggered.connect(self.onMyToolBarButtonClick)
        toolbar.addAction(button_action)
        toolbar2 = QToolBar("Menue tools")

        self.addToolBar(toolbar2)
        button_action2 = QAction("Help", self)
        button_action2.setStatusTip("How to Use tool?")
        button_action2.triggered.connect(self.Help_Menue)
        toolbar2.addAction(button_action2)

        #Get local repo Directory
        self.browse.clicked.connect(self.Get_Directory)

        #Init progress bar
        self.P_Pass.setEchoMode(10)
        self.progressBar.setValue(0)
        self.Version.setText(Version["Version"])
        self.P_Project.currentIndexChanged.connect(self.Select_Proj)

        #self.About.clicked.connect(self.onMyToolBarButtonClick)

        # Run button an execution of main runnable
        self.Run.clicked.connect(self.Script_Runnable)

    # Help Menue Method
    def Help_Menue(self):
        Help_path = os.getcwd()
        file_path = Help_path + "\Help.html"
        os.startfile(file_path)

    # About Menue Method
    def onMyToolBarButtonClick(self):
        self.About_Menue.show()

    # Init Cyclic timer Method -  to refresh main GUI during Running of Runabe and refress program status
    def UiTimer(self):
        # variables
        # count variable
        self.count = 0
        # start flag
        self.start = False
        # creating a timer object
        self.timer = QtCore.QTimer(self)
        # adding action to timer
        self.timer.timeout.connect(self.update_graph)
        # update the timer every tenth second
        self.timer.start(3000)

    # Runnable Initialization method
    def runTasks(self):
        self.Update_GUI_Tasks_names()
        threadCount = QThreadPool.globalInstance().maxThreadCount()
        #self.label.setText(f"Running {threadCount} Threads")
        pool = QThreadPool.globalInstance()
        runnable = Runnable(1)
        # 3. Call start()
        pool.start(runnable)

    # Update Graph method which is called with each timer to Update GUI status
    def update_graph(self):
        # Update current running task
        self.label_10.setText(Task["Name"])

        # Update Progress bar
        self.progressBar.setValue(Task["Progress"])

        # Update Task Complete validation
        self.label___1.setText(Task_Flag["Task1"])
        self.label___2.setText(Task_Flag["Task2"])
        self.label___3.setText(Task_Flag["Task3"])
        self.label___4.setText(Task_Flag["Task4"])
        self.label___5.setText(Task_Flag["Task5"])
        self.label___6.setText(Task_Flag["Task6"])

        # Update warning Message
        self.Warning_Message.setText(Task["Warning_Message"])

    # select current project Method
    def Select_Proj(self):
        Current_Proj= " "

        Current_Cam_Section_Name = str(self.P_Project.currentText())


        if Current_Cam_Section_Name == "VW_MEB_Inverter" :
            Current_Proj="VW_MEB_Inverter"
        elif Current_Cam_Section_Name == "BMW" :
            Current_Proj="obc_35up11kw"
        elif Current_Cam_Section_Name == "PSA":
            Current_Proj="phev_psa_erad_gen2"
        elif Current_Cam_Section_Name == "CEVT":
            Current_Proj="mma_cm1e_dcdc"
        elif Current_Cam_Section_Name == "100-KW":
            Current_Proj="optimus"
        elif Current_Cam_Section_Name == "model_kit":
            Current_Proj="model_kit"
        elif Current_Cam_Section_Name == "DAI":
            Current_Proj="S_DAI_SYS"
        elif Current_Cam_Section_Name == "P2 800V":
            Current_Proj="p2_800v_sic_inv_switching_cell"
        elif Current_Cam_Section_Name == "PSA":
            Current_Proj="phev_psa_erad_gen2"
        Input_Data["Project"] = Current_Proj

    # Get Input from GUI textboxs method
    def Get_Input(self):
        self.Select_Proj()
        Input_Data["User name"]=self.P_User.text()
        Input_Data["Password"] =self.P_Pass.text()

        #Input_Data["Project"] = str(self.P_Project.currentText())
        Input_Data["SWA_Baseline"] = self.SWA_Baseline.text()
        Input_Data["Int_Plan_Doc"] =self.Int_Plan_Document.text()
        Input_Data["Polarion_SW_Plan"] =self.Polarion_Plan.text()
        Input_Data["RB"] =self.Start_RB.text()
        Input_Data["L_Repo"] = self.L_Repo.text()
        Input_Data["Directory"] = Directory
        if self.SF_Mapping.isChecked():
            Input_Data["SF_Mapping"] =1
        else:
            Input_Data["SF_Mapping"] =0

    # Update GUI Task names
    def Update_GUI_Tasks_names(self):
        self.label__1.setText("Create Output report in Subfile Folder")
        self.label__2.setText("Prepare shell script and execute it => Output is Log.txt in subfiles filder")
        self.label__3.setText(
            "Parse Log file and generate Diff logs and pares data into Subfiles/IntegrationReport.xlsx")
        self.label__4.setText(
            "Get polarion plan tasks and WPs and all WIs and verify all Gerrit tickes Ids from polarion and get all there data")
        self.label__5.setText("generate report with system function mapping to SWC from polarion")
        self.label__6.setText("Generate final report")

    # Get Local repo directory Method
    def Get_Directory(self):
        self.dir_path = QFileDialog.getExistingDirectory(self, "Choose Directory", "E:\\")

        self.L_Repo.setText(self.dir_path)
        Input_Data["L_Repo"] = self.L_Repo.text()

    # method to triggered when Run button pressed
    # it is init tasks status
    # and Get Input Data
    # and start the Runnable
    def Script_Runnable(self):
        #print("Getting Input")
        self.label___1.setText("")
        self.label___2.setText("")
        self.label___3.setText("")
        self.label___4.setText("")
        self.label___5.setText("")
        self.label___6.setText("")
        self.Get_Input()
        #print("Finish Input")
        #print("Start Tasks")
        self.runTasks()


# about menue GUI Class
class About(QMainWindow):
    def __init__(self):
        super(About, self).__init__()
        loadUi("About.ui", self)
        self.label_10.setText(Version["Version"])

        # variable
        # Button




# main
if __name__ == '__main__':
    app = QApplication([])
    window  =  MatplotlibWidget ()
    sys.exit(app.exec())

