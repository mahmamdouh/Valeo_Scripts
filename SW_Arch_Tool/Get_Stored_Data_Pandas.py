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
    df = pd.read_excel(file1, 'SWA_Data') # can also index sheet by name or fetch all sheets
    DataList= df['IDs'].tolist()
    Data = df.to_dict()
    Release1_Data["SWA_Data"] = Data
    ##print("Data Structure")
    ##print(df)
    #trial code
    ##print(Release1_Data["SWA_Data"])




    # HSI Element
    df = pd.read_excel(file1, 'SWREQ_Data') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["SWREQ_Data"] =Data
    ##print(Release1_Data["SWREQ_Data"])

    return Release1_Data

def filter_ID_array(ID_Array):
    for elements in ID_Array:
        if elements =="":
            ID_Array.remove(elements)
    return ID_Array

def Get_Data_Stored_Runnable(New_file):
    path = os.getcwd()
    New_file = path + New_file
    data_structure = Get_Release_WIs(New_file)
    SWA_Array = []
    SWREQ_Array = []
    for elements in data_structure:
        array_Bak = []
        for element in elements:
            for sub_element in data_structure[elements]["Name"]:
                ##print(data_structure[elements]["Name"][sub_element])
                ##print(data_structure[elements]["Short_Name"][sub_element])
                ##print(data_structure[elements]["IDs"][sub_element])
                SF_Dict = {
                    "Name": "",
                    "Short_Name": "",
                    "Status": "",
                    "IDs": [],
                    "ID_Str": "",
                    "Ids_Number": 0,
                }

                SF_Dict["Name"] = data_structure[elements]["Name"][sub_element]
                SF_Dict["Short_Name"] = data_structure[elements]["Short_Name"][sub_element]
                SF_Dict["Status"] = data_structure[elements]["Status"][sub_element]
                ID_String = str(data_structure[elements]["IDs"][sub_element])
                ID_list = ID_String.split(",")
                ID_list =filter_ID_array(ID_list)
                SF_Dict["IDs"] = ID_list
                SF_Dict["ID_Str"] = data_structure[elements]["ID_Str"][sub_element]
                SF_Dict["Ids_Number"] = data_structure[elements]["Ids_Number"][sub_element]
                ##print(data_structure[elements]["Name"][element])
                ##print(data_structure[elements]["Short_Name"][sub_element])
                ##print(data_structure[elements]["Status"][sub_element])
                ##print(data_structure[elements]["IDs"][sub_element])
                ##print(data_structure[elements]["ID_Str"][sub_element])
                ##print(data_structure[elements]["Ids_Number"][sub_element])
                array_Bak.append(SF_Dict)
        if elements == "SWA_Data":
            SWA_Array = array_Bak
        elif elements == "SWREQ_Data":
            SWREQ_Array = array_Bak

    return SWA_Array,SWREQ_Array



'''
SWA_Array,SWREQ_Array=Get_Data_Stored_Runnable("\Outputs\SWA_SWREQ_Consistency_Temp_24_11_1_21_p321.xlsx")
#print("Finished")

'''