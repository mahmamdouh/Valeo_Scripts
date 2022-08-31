
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



###############################################################################
###################### global variables #######################################


Version={
    "Version": "Version : 4.0.0",
}

Input_Data ={
    "Project":"VW_MEB_Inverter" ,
    "User name": " ",
    "Password": " ",
    "ReqPlan": " ",
    "SWCompPlan": " ",
    "Release_Document": "" ,
    "Old_Release_Document": "",
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
#########################################################################################
######################### initialization ###############################################
current_directory = os.getcwd()
final_directory = os.path.join(current_directory, r'Outputs')
if not os.path.exists(final_directory):
   os.makedirs(final_directory)


#########################################################################################
# 1. Subclass QRunnable
class Runnable(QRunnable):
    def __init__(self, n):
        super().__init__()
        self.n = n
        my_signal = pyqtSignal()


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
        global SWA_Missed_IDs_With_Document
        global SWREQ_Missed_IDs_With_Document
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("SWA_SWREQ_Consistency - Start time ", Date_Data)

        Task_Ignition["Task_4"] = 1
        infoList = []
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        infoList.append(Input_Data["ReqPlan"])

        ####################################### Script ###############################################3
        # Prepare Data Set from Polarion Baselines Folders
        SWA_Array, SWAREQ_Array = Get_Data(infoList)
        # Generate Temp Report
        #Temp_File_Name = Generate_Temp_Report(infoList, SWA_Array, SWAREQ_Array)

        #SWA_Array, SWAREQ_Array = Get_Data_Stored_Runnable("\Outputs\SWA_SWREQ_Consistency_Temp_24_11_1_21_p321.xlsx")

        # Create Report Sheet
        workbook1, worksheet, worksheet1, worksheet2, worksheet3 = Generate_Report(infoList)

        # Generate Overall Report
        SWA_SWREQ_Total_KPI1, SWA_ocument_Not_have_match_name = Compare_Document_Names(SWA_Array, SWAREQ_Array, "SWA")
        SWA_SWREQ_Total_KPI1 = remove_Duplicates(SWA_SWREQ_Total_KPI1)
        workbook1, worksheet1 = Generate_SWA_Report(SWA_SWREQ_Total_KPI1, workbook1, worksheet1)
        SWA_SWREQ_Total_KPI = []

        SWA_SWREQ_Total_KPI2, SWRQ_ocument_Not_have_match_name = Compare_Document_Names(SWAREQ_Array, SWA_Array, "REQ")
        SWA_SWREQ_Total_KPI2 = remove_Duplicates(SWA_SWREQ_Total_KPI2)
        workbook1, worksheet2 = Generate_SWREQ_Report(SWA_SWREQ_Total_KPI2, workbook1, worksheet2)

        # find Miss-Consistency match Documents Names
        SWA_Document_Not_have_match_name, SWREQ_Document_Not_have_match_name = Find_Inconsistency_Document_Names(
            SWA_Array, SWAREQ_Array)
        SWA_Document_Not_have_match_name = remove_Duplicates(SWA_Document_Not_have_match_name)
        SWREQ_Document_Not_have_match_name = remove_Duplicates(SWREQ_Document_Not_have_match_name)

        # find Missed WIS  with Documents Names
        SWA_Missed_IDs_With_Document = remove_Duplicates(SWA_Missed_IDs_With_Document)
        SWREQ_Missed_IDs_With_Document = remove_Duplicates(SWREQ_Missed_IDs_With_Document)

        #############################################################33
        ### Mised IDS aLgo1
        NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array = Compare_All_Document_Inconsistency_Missed_Ids(
            SWA_Missed_IDs_With_Document, SWAREQ_Array, "SWA", "Algo1_1")
        Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array, "Algo1_1")

        index = 1
        NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array = Compare_All_Document_Inconsistency_Missed_Ids(
            SWREQ_Missed_IDs_With_Document, SWA_Array, "SWREQ", "Algo1_2")
        Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array, "Algo1_2")

        workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_IDs_With_Document_Report_Split1, workbook1,
                                                                     worksheet3, index)
        workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_IDs_With_Document_Report_Split2, workbook1,
                                                                     worksheet3, index)
        ### Mised Document aLgo2
        NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array = Compare_All_Document_Not_Founded_In_Other_Baseline(
            SWA_Document_Not_have_match_name, SWA_Array, SWAREQ_Array, "SWA", "Algo2_1")
        Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array, "Algo2_1")

        NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array = Compare_All_Document_Not_Founded_In_Other_Baseline(
            SWREQ_Document_Not_have_match_name, SWAREQ_Array, SWA_Array, "SWREQ", "Algo2_2")
        Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array, "Algo2_2")
        # Missed_Document_Report = remove_Duplicates(Missed_Document_Report)
        workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_Document_Report_Report_Split1, workbook1,
                                                                     worksheet3,
                                                                     index)
        workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_Document_Report_Report_Split2, workbook1,
                                                                     worksheet3,
                                                                     index)

        workbook1.close()

        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("SWA_SWREQ_Consistency - Stop time ", Date_Data)
        Task_Ignition["Task_4"] = 0
        Task_complete_valididty["Task_4"] = "✓"
        Task["Progress"] = 100
        Task["Name"] = "SWA_SWREQ_Consistency - Finished !"
        pass

    def Delta_Report(self):
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("Delta_Report - Start time ", Date_Data)
        Task_Ignition["Task_5"] = 1
        infoList = []
        #print("Step 1")
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        infoList.append(Input_Data["Old_Release_Document"])
        #print("Step 1")
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
            #print("Step 3.1.3")
            Delta_Message["Messsage"] = "Traceability_Matrix_Data report for "+ Old+ " Not exists !"
            Delta_Message["Question"] = "Do you need to Generate it ?"
            Delta_Message["Window"] = 1
            Reports["Start_Flag"] = 0
            #print("waiting for Flag")
            while (~(Delta_Message["Update_Flag"])):
                if Delta_Message["Update_Flag"] == 1 :
                    break
                time.sleep(1)
                #print("waiting for Flag",Delta_Message["Update_Flag"])
                pass
            #print("Flag response come with ",Delta_Message["Flag"])
            if Delta_Message["Flag"] == 1:
                #print("Step 3.1.4")
                #print("Generating report for ", Old)
                self.Release_WI("Run", Old)
                Reports["Start_Flag"] = 1
        #print("Step 4")
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
            workbook1.close()
            #print("Step pass 666666666")

        ## End algo
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("Delta_Report - Stop time ", Date_Data)
        Task["Progress"] = 100
        Task["Name"] = "Delta_Report - Finished !"
        Task_Ignition["Task_5"] = 0
        Task_complete_valididty["Task_5"] = "✓"
        #print(Task_Ignition["Task_5"],Task_complete_valididty["Task_5"])

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
        if  Mode != "Null":
            infoList = []
            infoList.append(Input_Data["User name"])
            infoList.append(Input_Data["Password"])
            infoList.append(Input_Data["Project"])
            infoList.append(Release)
            infoList.append(Input_Data["ReqPlan"])
        #print(infoList)
        Task["Progress"] = 0
        # Create Script Output folder.
        Task["Name"]="Create Script Output folder"
        output_path = create_output_directory()
        #print("Task 1")
        # Connect to Polarion.
        Task["Name"] = "Connecting to Polarion"
        Task["Progress"] = 10
        folder_content = get_folder_data(infoList)
        # For each document get workitems IDs and data.
        Task["Name"] = "For each document get workitems IDs and data"
        Task["Progress"] = 30
        #print("Task 2")
        docs_content = get_work_items_ids(folder_content,infoList)
        #print(infoList)
        Task["Progress"] = 40
        #print("Task 3")
        docs_content_detail = get_work_items_data(docs_content,infoList)
        Task["Name"] = "Create Excel sheet and write data"
        #print("Task 3.1")
        # Create Excel sheet and write data.
        Task["Progress"] = 90
        RWI_data_write(docs_content_detail, folder_content,infoList,output_path)
        #print("Task 4")
        Task["Progress"] =100
        Task["Name"] = "Release WI - Finished !"
        Task_Ignition["Task_2"] = 0
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("Release_WI - Stop time ", Date_Data)
        #print("Task 5")
        Task_complete_valididty["Task_2"] = "✓"

    def SWC_SWREQ_Bi_Directional(self):
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("SWC_SWREQ_Bi_Directional - Start time ", Date_Data)
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
        # #print(Input_Data)

        Task["Name"] = "Create / Open Excel file"
        Task["Progress"] = 2
        # Create / Open Excel file.
        workbook1, worksheet1, worksheet2, worksheet3,worksheet4 = excel_open(infoList)
        Task["Name"] = "Get polarion data 1st direction"
        Task["Progress"] = 3
        # Get polarion data 1st direction.
        #print("step 1")
        workitems_list, polarion_object = polarion_query_REQvsSWComp(infoList)
        #KPI["SWREQNum"] = len(workitems_list)
        Task["Name"] = "Analyze the data then write it in Excel file 1st direction"
        Task["Progress"] = 40
        #print("step 2")
        # Analyze the data then write it in Excel file 1st direction.
        workbook1 = data_analysis_REQvsSWComp(workitems_list, workbook1, worksheet1, worksheet3,worksheet4, infoList,
                                              polarion_object)
        Task["Name"] = "Get polarion data 2nd direction"
        Task["Progress"] = 50
        # Get polarion data 2nd direction.
        #print("step 3")
        workitems_list = polarion_query_SWCompvsREQ(infoList, polarion_object)
        KPI["SWCNum"] = len(workitems_list)
        Task["Name"] = "Analyze the data then write it in Excel file 2nd direction"
        Task["Progress"] = 80
        #print("step 4")
        # Analyze the data then write it in Excel file 2nd direction.
        workbook1 = data_analysis_SWCompvsREQ(workitems_list, workbook1, worksheet2, worksheet3, infoList,
                                              polarion_object)
        #print("step 5")
        polarion_object.disconnect()
        Task["Name"] = "Create KPI sheet "
        Task["Progress"] = 99
        #print("step 6")
        data = [
            ['Total SWREQ', 'Covered SWREQ', 'Total DWI', 'Covered DWI', 'Total SWC', 'Covered SWC'],
            [0, 0, 0, 0, 0, 0],
        ]
        #print("step 7")
        data[1][0] = KPI["Covered_SWREQ"]
        data[1][1] = KPI["SWREQNum"] - KPI["Covered_SWREQ"]
        data[1][2] = KPI["SWRDINum"]
        data[1][3] = KPI["SWRDINum"] - KPI["Covered_DIAG"]
        data[1][4] = KPI["Covered_SWC"]
        data[1][5] = KPI["SWCNum"] - KPI["Covered_SWC"]

        #print("step 8")
        workbook1 = Pi_Chart_creation(workbook1, worksheet3, data,infoList)
        #print("step 9")
        BAD_IDS(workbook1, Missed_SWRq, Missed_SWC,Missed_DWI, worksheet3, infoList)
        #print("step 10")
        Task["Name"] = "Finished !"
        Task["Progress"] = 100
        Task_Ignition["Task_1"] = 0
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("SWC_SWREQ_Bi_Directional - Stop time ", Date_Data)
        #print("step 11")
        Task_complete_valididty["Task_1"] = "✓"

    def SF_Status(self):
        now = datetime.now()  # current date and time
        Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
        #print("SF_Status - Start time ", Date_Data)
        Task_Ignition["Task_3"] = 1
        Task["Name"] = "Get inputs "
        Task["Progress"] = 1
        infoList = []
        infoList.append(Input_Data["User name"])
        infoList.append(Input_Data["Password"])
        infoList.append(Input_Data["Project"])
        infoList.append(Input_Data["Release_Document"])
        Task["Name"] = "Create / Open Excel file"
        Task["Progress"] = 2
        #print("Step 1")
        output_path = create_output_directory()
        Task["Name"] = "Create Report File"
        Task["Progress"] = 3
        #print("Step 2")
        folder_content = get_folder_data(infoList)
        #print("Step 3")
        output_workbook = Create_Report(output_path,infoList)
        Task["Name"] = "Get Document Information"
        Task["Progress"] = 10
        #print("Step 4")
        docs_content = SF_get_work_items_ids(folder_content, infoList, output_workbook)
        #print("Step 5")
        SF_Array = prepare_SF_array(folder_content, docs_content)
        Task["Name"] = "analyse Data ..."
        Task["Progress"] = 85
        #print("Step 6")
        docs_content_detail , SF_Array ,workbook,worksheet3 = get_work_items_Details_data(docs_content, infoList, SF_Array, output_workbook)
        SWC_SF_Array_Mapp = Separaate_SWC_SF(SF_Array)
        Write_SWC_VS_SF_Report(SWC_SF_Array_Mapp, workbook, worksheet3)
        #print("Step 7")
        Task["Name"] = "Finished !"
        Task["Progress"] = 100
        Task_Ignition["Task_3"] = 0
        now2 = datetime.now()  # current date and time
        Date_Data2 = now2.strftime("%m/%d/%Y, %H:%M:%S")
        #print("SF_Status - Stop time ", Date_Data2)
        Task_complete_valididty["Task_3"] = "✓"



class MatplotlibWidget(QMainWindow):

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


    def Help_Menue(self):
        #print("Hello How can Help you ")
    # method for widgets
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
        ##print(Task["Name"])
        ##print(Task["Name"])
        self.Define_Script_To_Run()
        self.Mandatory_Flags_Updtor()
        self.label_10.setText(Task["Name"])

        # check Delta Window
        if Delta_Message["Window"] == 1 :
            Delta_Message["Window"] =0
            time.sleep(1)
            self.delta1.show()


        self.progressBar.setValue(Task["Progress"])

        self.movie = QMovie("activity.gif")
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



    def Script_Runnable(self):
        self.Define_Script_To_Run()
        self.Get_Input()
        self.runTasks()

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

