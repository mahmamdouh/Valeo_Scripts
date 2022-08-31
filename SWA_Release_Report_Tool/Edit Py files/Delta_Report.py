
import pandas as pd
import os
import xlsxwriter
from datetime import datetime
import os.path

#############
#Global variable
Delta_Kpi_Added={
        "SW interface": 0,
        "HSI Element": 0,
        "Static view": 0,
        "Dynamic view": 0,
        "Runnables": 0,
        "SWC": 0,
        "DWI": 0,
        "REQ": 0,
        "Generic": 0,
}

Delta_Kpi_Removed={
        "SW interface": 0,
        "HSI Element": 0,
        "Static view": 0,
        "Dynamic view": 0,
        "Runnables": 0,
        "SWC": 0,
        "DWI": 0,
        "REQ": 0,
        "Generic": 0,
}

Reports={
        "Old_Release_Name": 0,
        "New_Release_Name": 0,
        "Old_Release_Flag": 0,
        "New_Release_Flag": 0,
        "Start_Flag": 0,
}



# Get Release 1 Data

def Get_Release_WIs(file1):
    Release1_Data={
        "SW interface": 0,
        "HSI Element": 0,
        "Static view": 0,
        "Dynamic view": 0,
        "Runnables": 0,
        "SWC": 0,
        "DWI": 0,
        "REQ": 0,
        "Generic": 0,
    }
    # SW interface
    df = pd.read_excel(file1, 'SW Interfaces') # can also index sheet by name or fetch all sheets
    DataList= df['IDs'].tolist()
    Data = df.to_dict()
    Release1_Data["SW interface"] = Data
    ##print("Data Structure")
    ##print(df)
    #trial code




    # HSI Element
    df = pd.read_excel(file1, 'HSI Elements') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["HSI Element"] =Data


    # Static view
    df = pd.read_excel(file1, 'Static Views') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["Static view"] = Data


    # Dynamic view
    df = pd.read_excel(file1, 'Dynamic Views') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["Dynamic view"] = Data


    # runnables
    df = pd.read_excel(file1, 'Runnables') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["Runnables"] = Data


    # SWC
    df = pd.read_excel(file1, 'swComponent') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["SWC"] = Data


    # Req
    df = pd.read_excel(file1, 'softwareRequirement') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["REQ"] = Data


    # DWI
    df = pd.read_excel(file1, 'diagnostic') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["DWI"] = Data


    # Generic
    df = pd.read_excel(file1, 'Generic') # can also index sheet by name or fetch all sheets
    Data = df.to_dict()
    Release1_Data["Generic"] = Data

    return Release1_Data


def Compare_SWA_Baseline(Data_New , Data_Old):
    WIs = ["SW interface","HSI Element","Static view","Dynamic view","Runnables","SWC","DWI","REQ","Generic"]
    New_WIs = []
    Old_WIs = []
    #print("Start Comparing ....")
    #Compare New WIs added

    for element in WIs:
        New_WIs.append(Compare_By_Item_Direction1(Data_New , Data_Old , element))
        ##print("Finishing of Comparing Added WI for ",element )
    #print("Get added WIs")
    # Compare Deleted WIs
    for element in WIs:
        Old_WIs.append(Compare_By_Item_Direction2(Data_Old,Data_New,element))
        ##print("Finishing of Comparing Deleted WI for ", element)
    #print("Get deleted WIs")
    return  New_WIs,Old_WIs

def Compare_By_Item_Direction1(Data_New ,Data_Old , Type):
    ADD_WI_Array = []
    Old_Release_Ids =[]
    new_Release_Ids = []

    for Items in Data_Old[Type]["IDs"]:
        Old_Release_Ids.append(Data_Old[Type]["IDs"][Items])

    for Items in Data_New[Type]["IDs"]:
        new_Release_Ids.append(Data_New[Type]["IDs"][Items])

    for element in Data_New[Type]["IDs"]:
        ##print("Element to compare:",element)
        ADD_WI_DICT = {
            "IDs": "",
            "Title": "",
            "Type": "",
            "Architecture Type": "",
            "System Function": "",
        }


        if Data_New[Type]["IDs"][element] in Old_Release_Ids :
            pass
        else:
            ADD_WI_DICT["IDs"]=Data_New[Type]["IDs"][element]
            ADD_WI_DICT["Title"]=Data_New[Type]["Title"][element]
            ADD_WI_DICT["Type"]=Data_New[Type]["Type"][element]
            ADD_WI_DICT["Architecture Type"]=Data_New[Type]["Architecture Type"][element]
            ADD_WI_DICT["System Function"]=Data_New[Type]["System Function"][element]
            if ADD_WI_DICT in ADD_WI_Array:
                pass
            else:
                 ADD_WI_Array.append(ADD_WI_DICT)
                 Update_KPI(Type, "Add")

    return ADD_WI_Array

def Compare_By_Item_Direction2(Data_New ,Data_Old , Type):
    ADD_WI_Array = []
    Old_Release_Ids =[]
    new_Release_Ids = []

    for Items in Data_Old[Type]["IDs"]:
        Old_Release_Ids.append(Data_Old[Type]["IDs"][Items])

    for Items in Data_New[Type]["IDs"]:
        new_Release_Ids.append(Data_New[Type]["IDs"][Items])


    for element in Data_New[Type]["IDs"]:
        ##print("Element to compare:",element)
        ADD_WI_DICT = {
            "IDs": "",
            "Title": "",
            "Type": "",
            "Architecture Type": "",
            "System Function": "",
        }


        if Data_New[Type]["IDs"][element] in Old_Release_Ids :
            pass
        else:
            ADD_WI_DICT["IDs"]=Data_New[Type]["IDs"][element]
            ADD_WI_DICT["Title"]=Data_New[Type]["Title"][element]
            ADD_WI_DICT["Type"]=Data_New[Type]["Type"][element]
            ADD_WI_DICT["Architecture Type"]=Data_New[Type]["Architecture Type"][element]
            ADD_WI_DICT["System Function"]=Data_New[Type]["System Function"][element]
            if ADD_WI_DICT in ADD_WI_Array:
                pass
            else:
                 ADD_WI_Array.append(ADD_WI_DICT)
                 Update_KPI(Type, "Removed")
    #print("Pass here 2.3")
    return ADD_WI_Array

def Delta_Generate_Report(info_list,Data1,Data2,Path):
    # Create / Open Excel file.+
    #print("Step here Pass gedan 0")
    File_Name = Path +"\Outputs\Delta_Report_" + info_list[3] +"_VS_"+ info_list[4]+ ".xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet= workbook1.add_worksheet("Execution details")
    worksheet1 = workbook1.add_worksheet("KPI")
    worksheet2 = workbook1.add_worksheet("New Work items")
    worksheet3 = workbook1.add_worksheet("removed Work items")
    #print("Step here Pass gedan 1")
    worksheet2.autofilter('A1:E5000')
    worksheet3.autofilter('A1:E5000')
    cell_format = workbook1.add_format()
    cell_format.set_bold()
    #print("Step here Pass gedan 2")
    # Add columns title.
    worksheet2.write("A1", 'ID',cell_format)
    worksheet2.write("B1", 'Title',cell_format)
    worksheet2.write("C1", 'Type',cell_format)
    worksheet2.write("D1", 'Architecture Type',cell_format)
    worksheet2.write("E1", 'System function',cell_format)
    #print("Step here Pass gedan 3")
    worksheet3.write("A1", 'ID',cell_format)
    worksheet3.write("B1", 'Title',cell_format)
    worksheet3.write("C1", 'Type',cell_format)
    worksheet3.write("D1", 'Architecture Type',cell_format)
    worksheet3.write("E1", 'System function',cell_format)


    # KPI Header
    worksheet1.write("B1", 'SW interface',cell_format)
    worksheet1.write("C1", 'HSI Element',cell_format)
    worksheet1.write("D1", 'Static view',cell_format)
    worksheet1.write("E1", 'Dynamic view',cell_format)
    worksheet1.write("F1", 'Runnables',cell_format)
    worksheet1.write("G1", 'SWC',cell_format)
    worksheet1.write("H1", 'DWI',cell_format)
    worksheet1.write("I1", 'REQ',cell_format)
    worksheet1.write("J1", 'Generic',cell_format)

    worksheet1.write("A2", 'Total Added',cell_format)
    worksheet1.write("A3", 'Total Removed',cell_format)

    #print("Step here Pass gedan 4")
    # Execution Header
    now = datetime.now()  # current date and time
    Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
    #print("Step here Pass gedan 5")
    worksheet.write('A1', "Exection Date",cell_format)
    worksheet.write('B1', Date_Data)
    worksheet.write('A2', "Generated By ",cell_format)
    worksheet.write('B2', str(info_list[0]))
    worksheet.write('A3', "SWA Baseline1 ",cell_format)
    worksheet.write('B3', str(info_list[3]))
    worksheet.write('A4', "SWA Baseline2 ",cell_format)
    worksheet.write('B4', str(info_list[4]))

    worksheet1.write("B2", str(Data1['SW interface']))
    worksheet1.write("C2", str(Data1['HSI Element']))
    worksheet1.write("D2", str(Data1['Static view']))
    worksheet1.write("E2", str(Data1['Dynamic view']))
    worksheet1.write("F2", str(Data1['Runnables']))
    worksheet1.write("G2", str(Data1['SWC']))
    worksheet1.write("H2", str(Data1['DWI']))
    worksheet1.write("I2", str(Data1['REQ']))
    worksheet1.write("J2", str(Data1['Generic']))

    worksheet1.write("B3", str(Data2['SW interface']))
    worksheet1.write("C3", str(Data2['HSI Element']))
    worksheet1.write("D3", str(Data2['Static view']))
    worksheet1.write("E3", str(Data2['Dynamic view']))
    worksheet1.write("F3", str(Data2['Runnables']))
    worksheet1.write("G3", str(Data2['SWC']))
    worksheet1.write("H3", str(Data2['DWI']))
    worksheet1.write("I3", str(Data2['REQ']))
    worksheet1.write("J3", str(Data2['Generic']))
    #print("Step here Pass gedan")

    return workbook1,worksheet, worksheet1, worksheet2 ,worksheet3

def Get_Input_Data():
    infoList = []

    # Get data from Input_Data_File.txt.
    with open("Polarion.txt") as infoFile:
        for line in infoFile:
            # Skip empty lines.
            if len(line) < 5:
                continue
            # Record non-empty lines.
            infoList.append(line[(line.find(':') + 1):].rstrip('\n'))

    infoList[2] = infoList[2].lower()

    if infoList[2] == "100kw":
        infoList[2] = "optimus"
    elif infoList[2] == "model kit" or infoList[2] == "modelkit" or infoList[2] == "model_kit" or infoList[2] == "model-kit":
        infoList[2] = "model_kit"
    elif infoList[2] == "vw_meb_inverter":
        infoList[2] = "VW_MEB_Inverter"
    elif infoList[2] == "vw_meb_inverter_base_minus":
        infoList[2] = "VW_MEB_Inverter_Base_Minus"

    return infoList


def Update_KPI(Type,Action):
    if Action == "Add":
        Delta_Kpi_Added[Type]=Delta_Kpi_Added[Type]+1
    elif Action == "Removed":
        Delta_Kpi_Removed[Type] = Delta_Kpi_Removed[Type] + 1


def Delta_Report_Data_Write(workbook1,worksheet,index,Data,infoList):
    Id = str(Data[0])
    ##print("Pass here please 1")
    a = 'A' + str(index)
    b = 'B' + str(index)
    c = 'C' + str(index)
    d = 'D' + str(index)
    e = 'E' + str(index)

    url = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + infoList[2] + "/workitem?id=" + Id
    URL = str(url)

    # Data write.
    worksheet.write_url(a, URL, string=Id)
    worksheet.write(b, str(Data[1]))
    worksheet.write(c, str(Data[2]))
    worksheet.write(d, str(Data[3]))
    worksheet.write(e, str(Data[4]))

    return workbook1,worksheet




def Generate_Report_Data(workbook1,Add_WIs,Deletec_WIs, worksheet2,worksheet3,infoList):
    index = 1
    #print("Function pass 1")
    for element in Add_WIs:
        for subelement in element:
            # #print("Added ", subelement)
            Data1 = []
            Data1.append(str(subelement["IDs"]))
            Data1.append(str(subelement["Title"]))
            Data1.append(str(subelement["Type"]))
            Data1.append(str(subelement["Architecture Type"]))
            Data1.append(str(subelement["System Function"]))
            index = index + 1
            ##print("Function pass 11")
            workbook1,worksheet2 = Delta_Report_Data_Write(workbook1, worksheet2, index, Data1,infoList)


    index = 1
    for element in Deletec_WIs:
        for subelement in element:
            Data1 = []
            Data1.append(str(subelement["IDs"]))
            Data1.append(str(subelement["Title"]))
            Data1.append(str(subelement["Type"]))
            Data1.append(str(subelement["Architecture Type"]))
            Data1.append(str(subelement["System Function"]))
            index = index + 1
            ##print("Deleted WIs : ",subelement["IDs"],subelement["System Function"])
            workbook1,worksheet3 = Delta_Report_Data_Write(workbook1, worksheet3, index, Data1,infoList)
    #print("Function Back 1")
    return workbook1,worksheet2,worksheet3

def Check_For_Reports(Old_file_Path , New_file_Path):
    Reports["Old_Release_Flag"] = os.path.exists(Old_file_Path)
    Reports["New_Release_Flag"] = os.path.exists(New_file_Path)


'''

if __name__ == '__main__':
    new= "24_11_1_19_p330"
    Old= "24_11_1_21_p321"
    infoList = Get_Input_Data()
    path = os.getcwd()
    New_file = path + "\Outputs\Traceability_Matrix_Data_"+new +".xlsx"
    Old_file = path + "\Outputs\Traceability_Matrix_Data_"+Old +".xlsx"




    Old_Release_Data = Get_Release_WIs(Old_file)
    New_Release_Data = Get_Release_WIs(New_file)
    #for element in Old_Release_Data:
     #   #print("Old data : ",Old_Release_Data[element]["IDs"])

    Add_WIs , Deletec_WIs = Compare_SWA_Baseline(New_Release_Data, Old_Release_Data)

    workbook1, worksheet, worksheet1, worksheet2, worksheet3 = Delta_Generate_Report(infoList,Delta_Kpi_Added, Delta_Kpi_Removed,path)

    workbook1,worksheet2,worksheet3= Generate_Report_Data(workbook1,Add_WIs, Deletec_WIs, worksheet2, worksheet3,infoList)
    workbook1.close()


'''






