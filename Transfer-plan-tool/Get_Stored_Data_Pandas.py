import pandas as pd
import os
import xlsxwriter
from datetime import datetime
import os.path



def Get_Release_WIs(file1):
    Release1_Data={
     "SWA_Data": 0,
     "SWREQ_Data": 0,
    }
    # SW interface
    df = pd.read_excel(file1, 'Sheet1') # can also index sheet by name or fetch all sheets
    DataList= df['commit'].tolist()
    df= df.fillna(0)
    Data = df.to_dict()
    Release1_Data["SWA_Data"] = Data
    return Data


def Get_Polarion_Plan_Data(file1):
    Release1_Data={
     "Parent_Data": 0,
     "Chield_Data": 0,
    }
    # SW interface
    df = pd.read_excel(file1, 'Sheet1') # can also index sheet by name or fetch all sheets
    #print(df)
    #DataList= df['commit'].tolist()
    df= df.fillna(0)
    Data = df.to_dict()
    Release1_Data["SWA_Data"] = Data
    return Data


def Get_Integration_Report_Data(file1):
    # SW interface
    df = pd.read_excel(file1, 'Sheet1')  # can also index sheet by name or fetch all sheets
    Gerrit_Tickets= df['Tag'].tolist()
    df = df.fillna(0)
    Data = df.to_dict()
    return Data ,Gerrit_Tickets


def Get_SWC_System_Function_Data(file1):
    # SW interface
    df = pd.read_excel(file1, 'SWC VS SF')  # can also index sheet by name or fetch all sheets
    SWC_List= df['Software Component'].tolist()
    df = df.fillna(0)
    Data = df.to_dict()
    return Data ,SWC_List

