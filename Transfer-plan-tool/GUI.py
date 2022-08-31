
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
from Integration1 import Runnable_1
from Functional_Lib import*
from datetime import datetime
from SystemFunction_SWC_Report import SWC_SF_Report_Generation_Runnable
from PolarionPlanAPIs import Polarion_Plan_Runnable
import time
#polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")



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
}




Task={
    "Name": "...",
    "Progress": 0,
}

Script={
    "SWREQ_SWC_Bi_Directional": 0,
    "Release_WI": 0,
    "SF_status": 0,
    "SWA_SWREQ_Consistency": 0,
    "Delta_Report": 0,

}

Release_WI={
    "Status": " ",
}

Task_complete_valididty= {
    "Task_1": "..",
    "Task_2": "..",
    "Task_3": "..",
    "Task_4": "..",
    "Task_5": "..",
}

Task_Ignition= {
    "Task_1" :0 ,
    "Task_2" :0 ,
    "Task_3" :0 ,
    "Task_4" :0 ,
    "Task_5": 0,
}

Delta_Message={
    "Messsage" :"Skip this missage by select NO !" ,
    "Question" :"0" ,
    "Flag" :0 ,
    "Window" :0 ,
    "Update_Flag" :0 ,
}

Warning_Message ={
    "Warning" : "" ,
}


#########################################################################################
######################### initialization ###############################################
current_directory = os.getcwd()
final_directory = os.path.join(current_directory, r'Outputs')
if not os.path.exists(final_directory):
   os.makedirs(final_directory)

Directory = current_directory.replace("\\", "\\\\")





#########################################################################################
# 1. Subclass QRunnable
class Runnable(QRunnable):
    def __init__(self, n):
        super().__init__()
        self.n = n


    def run(self):
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        print(Date_Data)
        start_time = time.time()
        Task["Progress"] = 1
        Task["Name"] = "Creat Directory"
        Input_Data["Report_Name"]= Create_Out_Report()
        Task["Progress"] = 3
        Task["Name"] = "Get Gerrit changes"
        Prepare_Shell_Script(Input_Data)
        os.system("Get_Repo.sh")
        Task["Progress"] = 10
        Task["Name"] = "Analyze Git changes"
        Runnable_1(Input_Data)
        Task["Progress"] = 25
        Task["Name"] = "Get polarion Plan Data"
        # Get polarion Plan
        infoList1 =[]
        infoList1.append(Input_Data["User name"])
        infoList1.append(Input_Data["Password"])
        infoList1.append(Input_Data["Project"])
        infoList1.append(Input_Data["Polarion_SW_Plan"])
        infoList1.append(Input_Data["SWA_Baseline"])
        Polarion_Plan_Runnable(infoList1)
        infoList2 = []
        infoList2.append(Input_Data["User name"])
        infoList2.append(Input_Data["Password"])
        infoList2.append(Input_Data["Project"])
        infoList2.append(Input_Data["SWA_Baseline"])
        infoList2.append(Input_Data["Polarion_SW_Plan"])
        infoList2.append(Input_Data["RB"])
        Task["Progress"] = 55
        Task["Name"] = "Get System Function VS SW Componant mapping"
        SWC_SF_Report_Generation_Runnable(infoList2)
        Task["Progress"] = 90
        Task["Name"] = "Generate Final report "
        Report_Generation_Verification(infoList2)
        Task["Progress"] = 100
        Task["Name"] = "Finish ! -'Outputs\Integration_Final_Report.xlsx' report Generated "
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        print(Date_Data)
        print("Execution Time")
        print("--- %s seconds ---" % (time.time() - start_time))



    pass




class MatplotlibWidget(QMainWindow):

    def __init__(self):
        ##print("clicked in init")
        QMainWindow.__init__(self)

        loadUi("Main_Window.ui", self)
        self.show()
        self.UiTimer()
        self.setWindowTitle("SWA Release Tool")


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

        self.browse.clicked.connect(self.Get_Directory)


        self.P_Pass.setEchoMode(10)
        self.progressBar.setValue(0)
        self.Version.setText(Version["Version"])
        self.P_Project.currentIndexChanged.connect(self.Select_Proj)

        #self.About.clicked.connect(self.onMyToolBarButtonClick)

        self.Run.clicked.connect(self.Script_Runnable)

    def Help_Menue(self):
        Help_path = os.getcwd()
        file_path = Help_path + "\Help.html"
        os.startfile(file_path)

    def onMyToolBarButtonClick(self):
        self.Setting1.show()

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

    def runTasks(self):
        threadCount = QThreadPool.globalInstance().maxThreadCount()
        #self.label.setText(f"Running {threadCount} Threads")
        pool = QThreadPool.globalInstance()
        runnable = Runnable(1)
        # 3. Call start()
        pool.start(runnable)

    def update_graph(self):
        #self.Define_Script_To_Run()
        #self.Mandatory_Flags_Updtor()

        self.label_10.setText(Task["Name"])
        self.progressBar.setValue(Task["Progress"])


    def Select_Proj(self):
        Current_Proj= " "
        ##print("Change happened")
        Current_Cam_Section_Name = str(self.P_Project.currentText())
        ##print(Current_Cam_Section_Name)

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
        Input_Data["Project"] = Current_Proj


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
        #print("Inputs ready" ,Input_Data )

        #Input_Data["SWCompPlan"] = Input_Data["Release_Document"]



    def Get_Directory(self):
        self.dir_path = QFileDialog.getExistingDirectory(self, "Choose Directory", "E:\\")
        #print(self.dir_path)
        self.L_Repo.setText(self.dir_path)
        Input_Data["L_Repo"] = self.L_Repo.text()
        #print(self.L_Repo.text())

    def Script_Runnable(self):
        #print("Getting Input")
        self.Get_Input()
        #print("Finish Input")
        #print("Start Tasks")
        self.runTasks()


class About(QMainWindow):
    def __init__(self):
        super(About, self).__init__()
        loadUi("About.ui", self)
        self.label_10.setText(Version["Version"])

        # variable
        # Button





if __name__ == '__main__':
    app = QApplication([])
    window  =  MatplotlibWidget ()
    sys.exit(app.exec())

