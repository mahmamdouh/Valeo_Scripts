
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
#from matplotlib.backends.backend_qt5agg  import  ( NavigationToolbar2QT  as  NavigationToolbar )
#import  numpy  as  np
#import  random
#import pandas as pd
import sys
import os
from REQ_SWC_Traceability import *
from Release_WI import *
from SystemFunction_SWC_Report import *
from Delta_Report import *
from SWA_SWREQ_Consistency import *

# Trial Lib
from Get_Stored_Data_Pandas import Get_Data_Stored_Runnable
from datetime import datetime
import time
#polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
from RTE_Data import Task,Warning_Message,Task_complete_valididty,Task_Ignition


###############################################################################
###################### global variables #######################################
# tool version
Version={
    "Version": "Version : 5.2.1",
}

# GUI Input data from User
Input_Data ={
    "Project":"VW_MEB_Inverter" ,
    "User name": " ",
    "Password": " ",
    "ReqPlan": " ",
    "SWCompPlan": " ",
    "Release_Document": "" ,
    "Old_Release_Document": "",
    "Variant": "variant.KEY: (base\+ base\-)",
}



Script={
    "SWREQ_SWC_Bi_Directional": 0,
    "Release_WI": 0,
    "SF_status": 0,
    "SWA_SWREQ_Consistency": 0,
    "Delta_Report": 0,
    "Terminate": 0,
}

Release_WI={
    "Status": " ",
}



# Delta report Generation Data
Delta_Message={
    "Messsage" :"Skip this missage by select NO !" ,
    "Question" :"0" ,
    "Flag" :0 ,
    "Window" :0 ,
    "Update_Flag" :0 ,
}




#########################################################################################
######################### initialization ###############################################
# Getting current Directory and Generate output folder
current_directory = os.getcwd()
final_directory = os.path.join(current_directory, r'Outputs')
if not os.path.exists(final_directory):
   os.makedirs(final_directory)


#########################################################################################
# main Runnable bassed on user attribute choice
# 1. Subclass QRunnable
class Runnable(QRunnable):
    #Abort_Signal = pyqtSignal()
    def __init__(self, n):
        super().__init__()
        self.n = n
        #Abort_Flag = self.Abort_Signal
        #print(Abort_Flag)
        #if Abort_Flag == 1:
        #    print("Terminate")
        #    self.autoDelete()

    def run(self):
        if Script["SWREQ_SWC_Bi_Directional"] == 1 :
            self.SWC_SWREQ_Bi_Directional()

        if Script["Release_WI"] == 1:
            self.Release_WI("Null","Null")

        if Script["SF_status"] == 1:
            #print("SF is going to Run ")
            self.SF_Status()

        if Script["SWA_SWREQ_Consistency"] == 1:
            self.SWA_SWREQ_Consistency()

        if Script["Delta_Report"] == 1:
            #print("Delta report  is going to Run ")
            self.Delta_Report()

        # Your long-running task goes here ...
    def SWA_SWREQ_Consistency(self):
        Task_Ignition["Task_4"] = 1
        infoList = []
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        infoList.append(Input_Data["ReqPlan"])
        infoList.append(Input_Data["Variant"])
        SWA_SWREQ_Consistency_Runnable(infoList)
        Task_Ignition["Task_4"] = 0
        Task_complete_valididty["Task_4"] = "✓"
        Task["Progress"] = 100
        Task["Name"] = "SWA_SWREQ_Consistency - Finished !"
        Warning_Message["Warning"] = ""

    def Delta_Report(self):
        Task_Ignition["Task_5"] = 1
        infoList = []
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        infoList.append(Input_Data["Old_Release_Document"])
        new = Input_Data["Release_Document"]
        Old = Input_Data["Old_Release_Document"]
        ## Script Algo
        #print("Step 2")
        path = os.getcwd()
        New_file = path + "\Outputs\Traceability_Matrix_Data_" + new + ".xlsx"
        Old_file = path + "\Outputs\Traceability_Matrix_Data_" + Old + ".xlsx"
        #print("Step 3")
        Check_For_Reports(Old_file, New_file)
        #print("Step 3.1")
        #print("Flags",Reports["Old_Release_Flag"],Reports["New_Release_Flag"])
        if Reports["Old_Release_Flag"] == True:
            if Reports["New_Release_Flag"] == True:
                Reports["Start_Flag"] = 1
                #print("Step 3.1.1")
            else:
                #print("Step 3.1.2")
                Delta_Message["Messsage"] = "Traceability_Matrix_Data report for "+ new+ " Not exists !"
                Delta_Message["Question"] = "Do you need to Generate it ?"
                Delta_Message["Window"] =1
                Reports["Start_Flag"] = 0

                while(~(Delta_Message["Update_Flag"])):
                    if Delta_Message["Update_Flag"] == 1:
                        break
                    time.sleep(1)
                    pass
                if Delta_Message["Flag"] == 1:
                    self.Release_WI( "Run", new)
                    Reports["Start_Flag"] = 1

        else:
            Delta_Message["Messsage"] = "Traceability_Matrix_Data report for "+ Old+ " Not exists !"
            Delta_Message["Question"] = "Do you need to Generate it ?"
            Delta_Message["Window"] = 1
            Reports["Start_Flag"] = 0
            #print("waiting for Flag")
            while (~(Delta_Message["Update_Flag"])):
                if Delta_Message["Update_Flag"] == 1 :
                    break
                time.sleep(1)
                pass
            if Delta_Message["Flag"] == 1:
                self.Release_WI("Run", Old)
                Reports["Start_Flag"] = 1
        Task["Progress"] = 3
        Task["Name"] = "Delta_Report - Started !"
        if Reports["Start_Flag"] == 1:
            Task["Progress"] = 10
            Task["Name"] = "Getting Baselinnes Data "
            Old_Release_Data = Get_Release_WIs(Old_file)
            New_Release_Data = Get_Release_WIs(New_file)

            Task["Progress"] = 30
            Task["Name"] = "Compare Baselines"
            Add_WIs, Deletec_WIs = Compare_SWA_Baseline(New_Release_Data, Old_Release_Data)
            #print("Step pass")
            Task["Progress"] = 90
            Task["Name"] = "Generate Report "
            #print("Step pass 2")
            workbook1, worksheet, worksheet1, worksheet2, worksheet3 = Delta_Generate_Report(infoList, Delta_Kpi_Added,Delta_Kpi_Removed,path)
            #print("Step pass 2")
            workbook1,worksheet2, worksheet3 = Generate_Report_Data(workbook1, Add_WIs, Deletec_WIs, worksheet2, worksheet3,infoList)
            #print("Step pass 55555")
            while True:
                try:
                    workbook1.close()
                    Warning_Message["Warning"] = ""
                    # os.startfile(output_workbook)
                    break
                except:
                    Warning_Message["Warning"]= "File Is already opened ! , please close Excel file "
                    pass




        ## End algo
        Task["Progress"] = 100
        Task["Name"] = "Delta_Report - Finished !"
        Task_Ignition["Task_5"] = 0
        Task_complete_valididty["Task_5"] = "✓"
        Warning_Message["Warning"] = ""

    def Release_WI(self,Mode,Release):
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("Release_WI - Start time ",Date_Data )

        Task_Ignition["Task_2"] = 1
        infoList=[]
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        infoList.append(Input_Data["ReqPlan"])
        infoList.append(Input_Data["Variant"])
        if  Mode != "Null":
            infoList = []
            infoList.append(Input_Data["User name"])
            infoList.append(Input_Data["Password"])
            infoList.append(Input_Data["Project"])
            infoList.append(Release)
            infoList.append(Input_Data["ReqPlan"])

        Release_WI_Runnable(infoList)

    def SWC_SWREQ_Bi_Directional(self):

        Task_Ignition["Task_1"] = 1
        infoList = []
        Task["Name"] = "Get Input Data"
        Task["Progress"] = 1
        # Get user input data.
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["ReqPlan"])
        infoList.append(Input_Data["Release_Document"])
        infoList.append(Input_Data["Variant"])
        REQ_SWC_Traceability_Runnable(infoList)
        # #print(Input_Data)

    def SF_Status(self):
        Task_Ignition["Task_3"] = 1
        Task["Name"] = "Get inputs "
        Task["Progress"] = 1
        infoList = []
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        SystemFunction_SWC_Report_Runnable(infoList)



class MatplotlibWidget(QMainWindow):
    #Abort_Signal = pyqtSignal()

    def __init__(self):
        ##print("clicked in init")
        QMainWindow.__init__(self)

        loadUi("Main_Window.ui", self)
        self.show()
        self.UiTimer()
        self.setWindowTitle("SWA Release Tool")
        self.Setting1 = About()
        self.delta1 = Delta()

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


        self.P_Pass.setEchoMode(10)
        self.progressBar.setValue(0)
        self.Version.setText(Version["Version"])
        self.P_Project.currentIndexChanged.connect(self.Select_Proj)

        #self.About.clicked.connect(self.onMyToolBarButtonClick)

        self.Run.clicked.connect(self.Script_Runnable)
        #self.Abort.clicked.connect(self.Abort_Button_Action)


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
        pool = QThreadPool.globalInstance()
        runnable = Runnable(1)
        # 3. Call start()
        pool.start(runnable)


    def update_graph(self):
        self.Warning_Message.setText(Warning_Message["Warning"])
        self.Define_Script_To_Run()
        self.Mandatory_Flags_Updtor()
        self.label_10.setText(Task["Name"])

        # check Delta Window
        if Delta_Message["Window"] == 1 :
            Delta_Message["Window"] =0
            time.sleep(1)
            self.delta1.show()


        self.progressBar.setValue(Task["Progress"])

        self.movie = QMovie(Task["Animated_GIF"])
        if Task_complete_valididty["Task_1"] != "✓" and Task_Ignition["Task_1"] == 1:
            self.SCRIPT_1_C_M.setMovie(self.movie)
            self.movie.start()
        elif Task_complete_valididty["Task_1"] == "✓":
            self.SCRIPT_1_C_M.setText(Task_complete_valididty[ "Task_1"])

        if Task_complete_valididty["Task_2"] != "✓" and Task_Ignition["Task_2"] == 1:
            self.SCRIPT_2_C_M.setMovie(self.movie)
            self.movie.start()
        elif Task_complete_valididty["Task_2"] == "✓":
            self.SCRIPT_2_C_M.setText(Task_complete_valididty[ "Task_2"])

        if Task_complete_valididty["Task_3"] != "✓" and Task_Ignition["Task_3"] == 1:
            self.SCRIPT_3_C_M.setMovie(self.movie)
            self.movie.start()
        elif Task_complete_valididty["Task_3"] == "✓":
            self.SCRIPT_3_C_M.setText(Task_complete_valididty["Task_3"])


        if Task_complete_valididty["Task_4"] != "✓" and Task_Ignition["Task_4"] == 1:
            self.SCRIPT_4_C_M.setMovie(self.movie)
            self.movie.start()
        elif Task_complete_valididty["Task_4"] == "✓":
            self.SCRIPT_4_C_M.setText(Task_complete_valididty[ "Task_4"])

        if Task_complete_valididty["Task_5"] != "✓" and Task_Ignition["Task_5"] == 1:
            self.SCRIPT_5_C_M.setMovie(self.movie)
            self.movie.start()
        elif Task_complete_valididty["Task_5"] == "✓":
            self.SCRIPT_5_C_M.setText(Task_complete_valididty[ "Task_5"])

        # Update Warning Message
        if Task_Ignition["Task_1"] == 1:
            Warning_Message["Warning"] = Release_WI_Warning_Message["Message"]
        if Task_Ignition["Task_2"] == 1:
            #print("Warning Message:", Warning_Message["Warning"],REQ_SWC_Traceability_Warning_Message["Message"])
            Warning_Message["Warning"] = REQ_SWC_Traceability_Warning_Message["Message"]
        if Task_Ignition["Task_3"] == 1:
            Warning_Message["Warning"] = SystemFunction_SWC_Report_Warning_Message["Message"]
        if Task_Ignition["Task_4"] == 1:
            pass
            #Warning_Message["Warning"] = Release_WI_Warning_Message["Message"]
        if Task_Ignition["Task_5"] == 1:
            Warning_Message["Warning"] = Delta_Report_Warning_Message["Message"]


    def Mandatory_Flags_Updtor(self):
        if Script["SWREQ_SWC_Bi_Directional"] == 1 or Script["SWA_SWREQ_Consistency"] == 1 or Script[
            "Release_WI"] == 1 or Script["SF_status"] == 1 or Script["Delta_Report"] == 1:
            self.User_Mark.setText("*")
            self.pass_Mark.setText("*")
            self.Project_Mark.setText("*")
            self.New_Release_Mark.setText("*")
        elif Script["SWREQ_SWC_Bi_Directional"] == 0 and Script["SWA_SWREQ_Consistency"] == 0 and Script[
            "Release_WI"] == 0 and Script["SF_status"] == 0 and Script["Delta_Report"] == 0:
            self.User_Mark.setText("")
            self.pass_Mark.setText("")
            self.Project_Mark.setText("")
            self.New_Release_Mark.setText("")
            self.ReqFolder_Mark.setText("")
            self.Old_Release_Mark.setText("")

        if Script["SWREQ_SWC_Bi_Directional"] == 1 or Script["SWA_SWREQ_Consistency"] == 1:
            self.ReqFolder_Mark.setText("*")
        else:
            self.ReqFolder_Mark.setText("")

        if Script["Delta_Report"] == 1:
            self.Old_Release_Mark.setText("*")
        else:
            self.Old_Release_Mark.setText("")

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
        infoList = []
        Input_Data["User name"]=self.P_User.text()
        Input_Data["Password"] =self.P_Pass.text()
        #Input_Data["Project"] = str(self.P_Project.currentText())
        Input_Data["ReqPlan"] =self.REQ_Path.text()
        Input_Data["Old_Release_Document"] =self.Old_Release_Path.text()
        Input_Data["Release_Document"] =self.Release_Folder.text()
        #Input_Data["SWCompPlan"] = Input_Data["Release_Document"]
        if self.Variant_Not.isChecked():
            if self.Base_P_Variant.isChecked():
                Input_Data["Variant"] = "NOT variant.KEY:base\+"
            if self.Base_M_Variant.isChecked():
                Input_Data["Variant"] = "NOT variant.KEY:base\-"
            if (self.Base_P_Variant.isChecked() and self.Base_M_Variant.isChecked() ):
                Input_Data["Variant"] = "NOT variant.KEY: (base\+ base\-)"
        else :
            if self.Base_P_Variant.isChecked():
                Input_Data["Variant"] = "variant.KEY:base\+"
            if self.Base_M_Variant.isChecked():
                Input_Data["Variant"] = "variant.KEY:base\-"
            if (self.Base_P_Variant.isChecked() and self.Base_M_Variant.isChecked() ):
                Input_Data["Variant"] = "variant.KEY: (base\+ base\-)"



    def Script_Runnable(self):
        Task_complete_valididty["Task_1"] = ""
        Task_complete_valididty["Task_2"] = ""
        Task_complete_valididty["Task_3"] = ""
        Task_complete_valididty["Task_4"] = ""
        Task_complete_valididty["Task_5"] = ""

        self.SCRIPT_1_C_M.setText("")
        self.SCRIPT_2_C_M.setText("")
        self.SCRIPT_3_C_M.setText("")
        self.SCRIPT_4_C_M.setText("")
        self.SCRIPT_5_C_M.setText("")
        Task["Animated_GIF"] = "activity.gif"
        self.Define_Script_To_Run()
        self.Get_Input()
        self.runTasks()
    # for future Implementation
    def Abort_Button_Action(self):
        Task["Animated_GIF"] = "AnimatedStop.gif"
        Task["Name"] = "Stop !"
        self.Abort_Scripts_flags()
        self.Abort_Signal.emit()
        #self.runTasks()

    # for future Implementation
    def Abort_Scripts_flags(self):
        Script["SWREQ_SWC_Bi_Directional"] = 0
        Script["Release_WI"] = 0
        Script["SF_status"] = 0
        Script["Delta_Report"] = 0
        Script["SWA_SWREQ_Consistency"] = 0

    def Define_Script_To_Run(self):
        if self.SWREQ_SWC.isChecked():
            Script["SWREQ_SWC_Bi_Directional"] = 1
        else:
            Script["SWREQ_SWC_Bi_Directional"] = 0

        if self.WI_Report.isChecked():
            Script["Release_WI"] = 1
        else:
            Script["Release_WI"] = 0

        if self.SF_SWC.isChecked():
            Script["SF_status"] = 1
        else:
            Script["SF_status"] = 0

        if self.Delta_Report.isChecked():
            Script["Delta_Report"] = 1
        else:
            Script["Delta_Report"] = 0

        if self.SWA_SWREQ.isChecked():
            Script["SWA_SWREQ_Consistency"] = 1
        else:
            Script["SWA_SWREQ_Consistency"] = 0

class About(QMainWindow):
    def __init__(self):
        super(About, self).__init__()
        loadUi("About.ui", self)
        self.label_10.setText(Version["Version"])

        # variable
        # Button

class Delta(QMainWindow):
    def __init__(self):
        super(Delta, self).__init__()
        loadUi("Delta_GUI.ui", self)

        self.Detailes.clicked.connect(self.Show_Details)
        self.Yes.clicked.connect(self.SET_Yes_Message)
        self.No.clicked.connect(self.SET_NO_Message)

        # variable
        # Button
    def Show_Details(self):
        #print("Details")
        #print("message",Delta_Message["Messsage"])
        #print("Question",Delta_Message["Question"])
        self.Message.setText(str(Delta_Message["Messsage"]))
        self.Question.setText(str(Delta_Message["Question"]))


    def SET_Yes_Message(self):
        Delta_Message["Flag"] = 1
        Delta_Message["Update_Flag"] = 1
        #print("Yes Pressed " ,Delta_Message["Flag"],Delta_Message["Update_Flag"] )
        self.close()

    def SET_NO_Message(self):
         Delta_Message["Flag"] = 0
         Delta_Message["Update_Flag"] = 1
         #print("No Pressed ", Delta_Message["Flag"], Delta_Message["Update_Flag"])
         self.close()


if __name__ == '__main__':
	app = QApplication([])
	window  =  MatplotlibWidget ()
	sys.exit(app.exec())

