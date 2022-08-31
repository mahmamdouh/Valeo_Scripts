import connectors.polarion_connector as c
import xlsxwriter
import re
import http.client
import requests
import os
from Get_Stored_Data_Pandas import Get_Data_Stored_Runnable
from datetime import datetime

SWA_SWREQ_Total_KPI =[]
Missed_IDs_With_Document=[]









SWA_Doc_List =[]
SWREQ_Doc_List =[]
SWA_SWREQ_Consistency_Warning_Message ={
    "Message" : ""
}

def create_output_directory():
    output_path = os.getcwd() + "\Outputs"

    # Check existence of Output folder and create it if not.
    if not os.path.exists(output_path):
        os.makedirs(output_path)
        ###print(output_path + ' : created')

    return output_path

def SWA_SWREQ_Create_Report(output_path):
    output_workbook = output_path + "\SWA_SWREQ_Consistency.xlsx"

    return output_workbook

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

def SWA_SWREQ_polarion_query_REQvsSWComp(infoList):
    ###print("Connecting to Polarion...")
    NOTConnected = True

    # Set polarion query.
    sql = "SQL:(select WI.C_PK from MODULE M inner join REL_MODULE_WORKITEM RMW ON RMW.FK_URI_MODULE = M.C_URI inner " \
          "join WORKITEM WI on WI.C_URI = RMW.FK_URI_WORKITEM where (M.C_LOCATION like '%" + infoList[3] + "%' ))"

    if infoList[2] == "VW_MEB_Inverter":
        query = "project.id:" + infoList[2] + " AND type: softwareRequirement AND " + sql + " AND NOT status: obsolete " \
                                                            "AND variant.KEY: (base\+ base\-)"
    else:
        query = "project.id:" + infoList[2] + " AND type: softwareRequirement AND " + sql + " AND NOT status: obsolete "

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get workitems.
            workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
                query, "priority", ["id", "title", "type", "linkedWorkItemsDerived"])
            ###print("Receiving Requirements data...")
            NOTConnected = False
        except Exception as e:
            ###print(str(e))
            ###print("Retry...")
            pass
        except:
            pass

    return workitems_list, polarion_object

def SWA_SWREQ_excel_open():
    # Create / Open Excel file.
    workbook1 = xlsxwriter.Workbook('Outputs\Req_SWC_Bi_Directional_Report.xlsx')

    worksheet1 = workbook1.add_worksheet("SWREQ ")
    worksheet2 = workbook1.add_worksheet("SWA")

    # Add columns title.
    #worksheet1.write("A1", 'Req ID')


    return workbook1, worksheet1, worksheet2

def get_SWA_folder_data(info_list):
    ##print("Connecting to Polarion...")
    not_connected = True

    # Connect to Polarion database
    while not_connected:
        try:
            # Connect using the username and password
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion")
            polarion_object.connect(str(info_list[0]), str(info_list[1]))

            folder_content = ''

            folder_content = polarion_object.tracker_webservice.service.getModules(info_list[2], info_list[3])
            not_connected = False
        except BaseException as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except Exception as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except http.client.HTTPException as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except http.client.RemoteDisconnected as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except requests.exceptions.ConnectionError as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except ConnectionError as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except:
            pass

    ###print("Folder content downloaded")

    # Disconnect Polarion.
    polarion_object.disconnect()
    ##print("I am here now ")
    return folder_content

def get_SWRq_folder_data(info_list):
    ##print("Connecting to Polarion...")
    not_connected = True

    # Connect to Polarion database
    while not_connected:
        try:
            # Connect using the username and password
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion")
            polarion_object.connect(str(info_list[0]), str(info_list[1]))

            folder_content = ''

            folder_content = polarion_object.tracker_webservice.service.getModules(info_list[2], info_list[4])
            not_connected = False
        except BaseException as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except Exception as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except http.client.HTTPException as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except http.client.RemoteDisconnected as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except requests.exceptions.ConnectionError as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except ConnectionError as e:
            # ##print connection error message
            ###print(str(e))
            ##print("Retry...")
            pass
        except:
            pass

    ###print("Folder content downloaded")

    # Disconnect Polarion.
    polarion_object.disconnect()
    ##print("I am here now ")
    return folder_content

def SWA_SWREQ_prepare_SWA_SF_array(folder_content,docs_content):
    SF_Array = []
    SF_Dict = {
        "Name": "",
        "Status": "",
        "IDs": [],
    }
    for element in folder_content:
        SF_Dict = {
            "Name": "",
            "Short_Name": "",
            "Status": "",
            "IDs": [],
            "ID_Str": "",
            "Ids_Number": 0,
        }
        SF_Dict["Name"] = str(element.title)
        SF_Dict["Status"] = str(element.status.id)
        SWA_DOC_Tile = re.split("\s", SF_Dict["Name"], 1)
        SF_Dict["Short_Name"] =SWA_DOC_Tile[0]
        SF_Array.append(SF_Dict)

    for pair in docs_content:
        for element in SF_Array:
            if pair == element["Name"]:
                element["IDs"] = docs_content[pair]
            else:
                pass
    return SF_Array

def SWA_SWREQ_prepare_SWREQ_SF_array(folder_content,docs_content):
    SF_Array = []
    SF_Dict = {
        "Name": "",
        "Status": "",
        "IDs": [],
    }
    for element in folder_content:
        SF_Dict = {
            "Name": "",
            "Short_Name": "",
            "Status": "",
            "IDs": [],
            "ID_Str": "",
            "Ids_Number":0,
        }
        SF_Dict["Name"] = str(element.title)
        SF_Dict["Status"] = str(element.status.id)
        SWA_DOC_Tile = re.split("\s", SF_Dict["Name"], 1)
        SWA_DOC_Tile2 = re.split("\s", SWA_DOC_Tile[1], 1)
        SF_Dict["Short_Name"] =SWA_DOC_Tile2[0]
        SF_Array.append(SF_Dict)

    for pair in docs_content:
        for element in SF_Array:
            if pair == element["Name"]:
                element["IDs"] = docs_content[pair]
            else:
                pass
    return SF_Array

def SWA_SWREQ_get_work_items_ids(folder_content,info_list,output_workbook):
    docs_title_list = []
    docs_content = {}
    index =2

    # Get documents title in list and make dictionary for IDs list for each document.
    for element in folder_content:
        docs_title_list.append(element.title)
        docs_content[element.title] = element.homePageContent
        #wb=Write_SF_DATA(info_list, wb, system_Function, index, str(element.title), str(element.status.id),output_workbook)
        #index = index + 1

    # Loop on each document and get it's IDs, then save it in the dictionary.
    for element in docs_title_list:
        home_page_content = docs_content[element].content

        res = []
        ids_list = []

        # Search for IDs in the page data.
        for elements in re.finditer("params=id=", home_page_content):
            # ##print(home_page_content[elements.end():elements.end() + 16])
            if info_list[2] == "optimus":
                res.append(home_page_content[elements.end():elements.end() + 11])
            elif info_list[2] == "model_kit":
                res.append(home_page_content[elements.end():elements.end() + 15])
            elif info_list[2] == "mma_cm1e_dcdc":
                res.append(home_page_content[elements.end():elements.end() + 19])
            elif info_list[2] == "VW_MEB_Inverter":
                res.append(home_page_content[elements.end():elements.end() + 17])
            elif info_list[2] == "VW_MEB_Inverter_Base_Minus":
                res.append(home_page_content[elements.end():elements.end() + 23])
            elif info_list[2] == "PMA1":
                res.append(home_page_content[elements.end():elements.end() + 11])
            elif info_list[2] == "me_s230":
                res.append(home_page_content[elements.end():elements.end() + 11])

        # Filter duplicated IDs.
        for i in range(len(res)):
            # Check if the string ends with digits (number) and return None if not.
            for j in range(3):
                last_character = re.search(r'\d+$', res[i])
                if last_character is None:
                    res[i] = res[i][:-1]
                j += 1
            if res[i] not in ids_list:
                ids_list.append(res[i])

        # Save the IDs list in the dictionary.
        docs_content[element] = ids_list

    return docs_content

def SWA_SWREQ_get_work_items_Details_data(docs_content,info_list,SF_Array,output_workbook):
    docs_content_detail = {}
    docs_content_detail2 = {}
    ##print("Connecting to Polarion...")
    # Connect using the username and password
    polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion")
    polarion_object.connect(str(info_list[0]), str(info_list[1]))
    Document_SF ={}
    Doc_Title =""
    Doc_Status=""
    index=1
    # Get workitem details for each document.
    for elements in SF_Array:
        ###print("Name :", elements["Name"])
        ###print("Status :", elements["Status"])
        ###print("IDs :", elements["IDs"])
        ids_list = elements["IDs"]
        ###print("Doc element :",docs_content)
        #Status = docs_content[element.status]
        SW_COmp =[]
        #docs_content_detail2[element]
        not_connected = True

        # Connect to Polarion database
        while not_connected:
            try:
                ###print("trying Except ")
                # Clear the list data to avoid any corrupted data.
                work_item_list = []
                ids_group = '('

                # Get WorkItems data.
                for i in range(len(ids_list)):
                    ids_group += (ids_list[i] + ' ')
                ids_group = ids_group[:-1]
                ids_group += ')'

                query = str("project.id:" + info_list[2] + " AND id:" + ids_group)

                work_item_list.append(polarion_object.tracker_webservice.service.queryWorkItems(
                    query, "priority", ["id", "title", "type"]))
                not_connected = False
            except ConnectionError as e:
                # ##print connection error message
                ###print("Error Except 1")
                ###print(str(e))
                ##print("Retry...")
                pass
            except Exception as e:
                # ##print connection error message
                ###print("Error Except 2")
                ###print(str(e))
                ##print("Retry...")
                pass
            except BaseException as e:
                ###print("Error Except 3")
                # ##print connection error message
                ###print(str(e))
                ##print("Retry...")
                pass
            except http.client.HTTPException as e:
                ###print("Error Except 4")
                # ##print connection error message
                ###print(str(e))
                ##print("Retry...")
                pass
            except http.client.RemoteDisconnected as e:
                # ##print connection error message
                ###print("Error Except 5")
                ###print(str(e))
                ##print("Retry...")
                pass
            except requests.exceptions.ConnectionError as e:
                # ##print connection error message
                ###print("Error Except 6")
                ###print(str(e))
                ##print("Retry...")
                pass
            except:
                ##print("Retry...")
                pass
        ###print("SF Done => ",elements["Name"])

        for element in work_item_list:
            for sub_element in element:
                if sub_element.type.id == "softwareRequirement":
                    ID = sub_element.id
                    SW_COmp.append(ID)
        elements["IDs"] = SW_COmp
        elements["Ids_Number"] = len(elements["IDs"])

    polarion_object.disconnect()
    return SF_Array

def Get_Data(infoList):

    # Create / Open Excel file.
    output_path = create_output_directory()
    output_workbook = SWA_SWREQ_Create_Report(output_path)
    ##print("Step 1")
    workbook1, worksheet1, worksheet2 = SWA_SWREQ_excel_open()
    # ##print("Step 2")
    ##print(infoList)
    # workitems_list, polarion_object = SWA_SWREQ_polarion_query_REQvsSWComp(infoList)
    ##print("Step 3")
    SWA_folder_content = get_SWA_folder_data(infoList)
    ##print("Step 4")
    SWAREQ_folder_content = get_SWRq_folder_data(infoList)
    ##print("Step 5")
    SWA_docs_content = SWA_SWREQ_get_work_items_ids(SWA_folder_content, infoList, output_workbook)
    ##print("Step 6")
    SWAREQ_docs_content = SWA_SWREQ_get_work_items_ids(SWAREQ_folder_content, infoList, output_workbook)
    ##print("Step 7")
    SWA_Array = SWA_SWREQ_prepare_SWA_SF_array(SWA_folder_content, SWA_docs_content)
    ##print("Step 8")
    SWAREQ_Array = SWA_SWREQ_prepare_SWREQ_SF_array(SWAREQ_folder_content, SWAREQ_docs_content)
    ##print("Step 9")
    SWA_Array = SWA_SWREQ_get_work_items_Details_data(SWA_folder_content, infoList, SWA_Array, output_workbook)
    ##print("Step 10")
    SWAREQ_Array = SWA_SWREQ_get_work_items_Details_data(SWAREQ_folder_content, infoList, SWAREQ_Array, output_workbook)
    ##print("Step 11")
    ##print("End")

    return SWA_Array, SWAREQ_Array

def Compare_Document_Names(SWA_A , SWREQ_A ,Type,SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document):
    "Short_Name"
    Document_Not_have_match_name=[]
    for SWA_items in SWA_A:
        SWA_Document_Title = SWA_items["Short_Name"]
        for SWRQ_items in SWREQ_A:
            SWRQ_Document_Title = SWRQ_items["Short_Name"]
            if (SWA_Document_Title == SWRQ_Document_Title):
                ###print("Founded :",SWRQ_items["Name"] ,SWA_items["Name"] )
                miss_Match_FLag_Rtn,SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document=Compare_SWA_SWAREQ_Document_Ids(SWA_items, SWRQ_items ,Type,SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document)

        else:
            pass
    return SWA_SWREQ_Total_KPI , Document_Not_have_match_name , SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document


def Find_Inconsistency_Document_Names(SWA_A , SWREQ_A):
    "Short_Name"
    SWA_Document_Not_have_match_name=[]
    SWREQ_Document_Not_have_match_name = []
    SWA_Documents=[]
    SWREQ_Documents = []
    # prepare Lists
    for SWA_items in SWA_A:
        SWA_Documents.append(SWA_items["Short_Name"])

    for SWRQ_items in SWREQ_A:
        SWREQ_Documents.append(SWRQ_items["Short_Name"])

    # compare SWA List
    for SWA_items in SWA_A:
        if SWA_items["Short_Name"] in SWREQ_Documents:
            pass
        else:
            SWA_Document_Not_have_match_name.append(SWA_items["Name"])

    # compare SWREQ List
    for SWRQ_items in SWREQ_A:
        if SWRQ_items["Short_Name"] in SWA_Documents:
            pass
        else:
            SWREQ_Document_Not_have_match_name.append(SWRQ_items["Name"])

    return SWA_Document_Not_have_match_name,SWREQ_Document_Not_have_match_name

def Compare_All_Document_Inconsistency_Missed_Ids(SWA_A , SWREQ_A,Document,Algo,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2):
    Found_Flag = 0
    Found_Flag_array = []
    NOT_Found_Flag_array = []
    NOT_Found_Flag_array_Data = []
    Inconsistency_Flag = 0
    for SWA_items in SWA_A:
        ###print("=================================")
        for SWRQ_items in SWREQ_A:
            SWA_SWREQ_KPI = {
                "SWA_DOC": "",
                "SWREQ_DOC": "",
                "Message": 0,
                "Missing_Ids": [],
            }
            SWA_SWREQ_KPI_Founded = {
                "SWA_DOC": "",
                "SWREQ_DOC": "",
                "Message": 0,
                "Missing_Ids": "",
            }
            SWA_SWREQ_KPI_NOT_Founded = {
                "SWA_DOC": "",
                "SWREQ_DOC": "",
                "Message": 0,
                "Missing_Ids": "",
            }
            ID_list = SWRQ_items["IDs"]
            if len(ID_list) == 0:
                pass
            else:
                for Id_Element in SWA_items["IDs"]:
                    if Id_Element in SWRQ_items["IDs"]:
                        SWA_SWREQ_KPI_Founded["SWA_DOC"] = SWA_items["Name"]
                        SWA_SWREQ_KPI_Founded["SWREQ_DOC"] = SWRQ_items["Name"]
                        SWA_SWREQ_KPI_Founded["Message"] = Document + " Wrong Reference"
                        SWA_SWREQ_KPI_Founded["Missing_Ids"]=Id_Element
                        if Id_Element not in Found_Flag_array:
                            ###print("Find Element",Id_Element )
                            if Algo == "Algo1_1":
                                Missed_IDs_With_Document_Report_Split1.append(SWA_SWREQ_KPI_Founded)
                            elif Algo == "Algo1_2":
                                Missed_IDs_With_Document_Report_Split2.append(SWA_SWREQ_KPI_Founded)
                            Found_Flag_array.append(Id_Element)
                    else:
                        Found_Flag == 1
                        if Id_Element not in NOT_Found_Flag_array:
                            NOT_Found_Flag_array.append(Id_Element)
                        SWA_SWREQ_KPI["SWA_DOC"] = SWA_items["Name"]
                        SWA_SWREQ_KPI["SWREQ_DOC"] = ""
                        SWA_SWREQ_KPI["Message"] = Document + " WIS not mentioned anywhere"
                        SWA_SWREQ_KPI["Missing_Ids"] = Id_Element
                        if SWA_SWREQ_KPI not in NOT_Found_Flag_array_Data:
                            NOT_Found_Flag_array_Data.append(SWA_SWREQ_KPI)

    return NOT_Found_Flag_array,NOT_Found_Flag_array_Data,Found_Flag_array,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2

def Filter_Total_Missing_Ids(NOT_Found_Flag_array,NOT_Found_Flag_array_Data,Found_Flag_array,Current_Algo,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2):
    for items in NOT_Found_Flag_array:
        if items not in Found_Flag_array:
            ###print("Missing Elements",items)
            for elements in NOT_Found_Flag_array_Data:
                if items == elements["Missing_Ids"]:
                    if Current_Algo == "Algo1_1":
                        if elements not in Missed_IDs_With_Document_Report_Split1:
                            Missed_IDs_With_Document_Report_Split1.append(elements)
                    elif Current_Algo == "Algo1_1":
                        if elements not in Missed_IDs_With_Document_Report_Split2:
                            Missed_IDs_With_Document_Report_Split2.append(elements)
                    elif Current_Algo == "Algo2_1":
                        if elements not in Missed_Document_Report_Report_Split1:
                           Missed_Document_Report_Report_Split1.append(elements)
                    elif Current_Algo == "Algo2_2":
                        if elements not in Missed_Document_Report_Report_Split2:
                            Missed_Document_Report_Report_Split2.append(elements)
    return Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2


def Compare_SWA_SWAREQ_Document_Ids(SWA_IDs ,SWRQ_IDs ,Type,SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document):
    miss_Match_FLag_Rtn = 0
    Temp_arr =[]
    SF_Dict = {
        "Name": "",
        "IDs": [],
    }
    SWA_SWREQ_KPI = {
        "SWA_DOC": "",
        "SWREQ_DOC": "",
        "Match_Ids": 0,
        "Not_Match_Ids": 0,
        "Total_Ids": 0,
        "Other_DOC_Ids":0,
        "IDs": [],
        "Missing_Ids": [],
        "Miss_Match_Status": 0,
    }
    Total_Ids =0
    SWA_SWREQ_KPI["SWA_DOC"]=SWA_IDs["Name"]
    SWA_SWREQ_KPI["SWREQ_DOC"]=SWRQ_IDs["Name"]
    for element in SWA_IDs["IDs"]:
        Total_Ids = Total_Ids +1
        if element in SWRQ_IDs["IDs"]:
            SWA_SWREQ_KPI["Match_Ids"] +=1
        else:
            miss_Match_FLag_Rtn =1
            SWA_SWREQ_KPI["Miss_Match_Status"] =1
            SWA_SWREQ_KPI["Not_Match_Ids"]+=1
            SWA_SWREQ_KPI["Missing_Ids"].append(element)
            Temp_arr.append(element)
            if Type == "SWA":
                SF_Dict["IDs"]=Temp_arr
                SF_Dict["Name"] = SWA_IDs["Name"]
                SWA_Missed_IDs_With_Document.append(SF_Dict)
            elif Type == "REQ":
                ###print("##print Missed SWREQ:",SWA_IDs["Name"])
                ###print("Lost Ids:",Temp_arr)
                SF_Dict["IDs"] = Temp_arr
                SF_Dict["Name"]=SWA_IDs["Name"]
                SWREQ_Missed_IDs_With_Document.append(SF_Dict)
    SWA_SWREQ_KPI["Total_Ids"] = Total_Ids
    SWA_SWREQ_KPI["IDs"] = Total_Ids
    SWA_SWREQ_Total_KPI.append(SWA_SWREQ_KPI)
    SWA_SWREQ_KPI["Other_DOC_Ids"] = len(SWRQ_IDs["IDs"])
    return miss_Match_FLag_Rtn,SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document

def Generate_Report(info_list):
    # Create / Open Excel file.
    File_Name = "Outputs\SWA_SWREQ_Consistency_" + info_list[3] + ".xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet = workbook1.add_worksheet("Execution Details")
    worksheet1 = workbook1.add_worksheet("SWA VS SWREQ Overall Compare")
    worksheet2 = workbook1.add_worksheet("SWREQ VS SWA Overall Compare")
    worksheet3 = workbook1.add_worksheet("Inconsistency Report")


    worksheet1.autofilter('A1:D5000')
    worksheet2.autofilter('A1:D5000')
    worksheet3.autofilter('A1:D5000')

    cell_format = workbook1.add_format()
    cell_format.set_bold()

    # Headers.
    worksheet1.write("A1", 'SW ARCH Document', cell_format)
    worksheet1.write("B1", 'SW REQ Document', cell_format)
    worksheet1.write("C1", 'Consistency', cell_format)
    worksheet1.write("D1", 'Numbers', cell_format)
    worksheet1.write("E1", 'Missing IDs', cell_format)

    worksheet2.write("A1", 'SW REQ Document', cell_format)
    worksheet2.write("B1", 'SW Arch Document', cell_format)
    worksheet2.write("C1", 'Consistency', cell_format)
    worksheet2.write("D1", 'Numbers', cell_format)
    worksheet2.write("E1", 'Missing IDs', cell_format)

    worksheet3.write("A1", 'SW REQ Document', cell_format)
    worksheet3.write("B1", 'SW Arch Document', cell_format)
    worksheet3.write("C1", 'Consistency Result', cell_format)
    worksheet3.write("D1", 'Missing IDs', cell_format)

    now = datetime.now()  # current date and time
    Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
    # #print("step 2")
    worksheet.write('A1', "Exection Date", cell_format)
    worksheet.write('B1', Date_Data)
    worksheet.write('A2', "Generated By ", cell_format)
    worksheet.write('B2', str(info_list[0]))
    worksheet.write('A3', "SWREQ Baseline ", cell_format)
    worksheet.write('B3', str(info_list[4]))
    worksheet.write('A4', "SWA Baseline ", cell_format)
    worksheet.write('B4', str(info_list[3]))
    worksheet.write('A5', "Variant", cell_format)
    worksheet.write('B5', str(info_list[5]))

    return workbook1,worksheet,worksheet1,worksheet2,worksheet3

def Generate_Temp_Report(info_list,SWA_Array, SWAREQ_Array):
    # Create / Open Excel file.
    output_path = os.getcwd()
    File_Name = output_path +"\Outputs\SWA_SWREQ_Consistency_Temp_" + info_list[3] + ".xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet = workbook1.add_worksheet("SWA_Data")
    worksheet1 = workbook1.add_worksheet("SWREQ_Data")

    # Headers.
    cell_format = workbook1.add_format()
    cell_format.set_bold()
    worksheet.write("A1", 'Name', cell_format)
    worksheet.write("B1", 'Short_Name', cell_format)
    worksheet.write("C1", 'Status', cell_format)
    worksheet.write("D1", 'IDs', cell_format)
    worksheet.write("E1", 'ID_Str', cell_format)
    worksheet.write("F1", 'Ids_Number', cell_format)

    worksheet1.write("A1", 'Name', cell_format)
    worksheet1.write("B1", 'Short_Name', cell_format)
    worksheet1.write("C1", 'Status', cell_format)
    worksheet1.write("D1", 'IDs', cell_format)
    worksheet1.write("E1", 'ID_Str', cell_format)
    worksheet1.write("F1", 'Ids_Number', cell_format)

    Index = 1

    for Elements in SWA_Array:
        ##print("Data To ##print" ,Elements["Name"] )
        Index += 1
        a = 'A' + str(Index)
        b = 'B' + str(Index)
        c = 'C' + str(Index)
        d = 'D' + str(Index)
        e = 'E' + str(Index)
        f = 'F' + str(Index)
        ##print("Index A",a)
        ID_String =""
        for element in Elements["IDs"]:
            ID_String = ID_String +","+element
        worksheet.write(a,Elements["Name"])
        worksheet.write(b, Elements["Short_Name"])
        worksheet.write(c, Elements["Status"])
        worksheet.write(d, ID_String)
        worksheet.write(e, "-")
        worksheet.write(f, Elements["Ids_Number"])
    # SWAREQ_Array

    Index = 1
    for Elements in SWAREQ_Array:
        Index += 1
        a = 'A' + str(Index)
        b = 'B' + str(Index)
        c = 'C' + str(Index)
        d = 'D' + str(Index)
        e = 'E' + str(Index)
        f = 'F' + str(Index)
        ID_String =""
        for element in Elements["IDs"]:
            ID_String = ID_String +","+element
        worksheet1.write(a,Elements["Name"])
        worksheet1.write(b, Elements["Short_Name"])
        worksheet1.write(c, Elements["Status"])
        worksheet1.write(d, ID_String)
        worksheet1.write(e, "-")
        worksheet1.write(f, Elements["Ids_Number"])

    workbook1.close()
    return File_Name

def Data_Write(workbook,worksheet,Data,index):
    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})
    format2 = workbook.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100'})


    a = 'A' + str(index)
    b = 'B' + str(index)
    c = 'C' + str(index)
    d = 'D' + str(index)
    e = 'E' + str(index)

    worksheet.write(a, str(Data[0]))
    worksheet.write(b, str(Data[1]))
    if Data[2] == "Match" :
        worksheet.write(c, str(Data[2]),format2)
    elif Data[2] == "Miss Match" :
        worksheet.write(c, str(Data[2]),format1)
    else:
        worksheet.write(c, str(Data[2]))
    worksheet.write(d, str(Data[3]))
    worksheet.write(e, str(Data[4]))

    return workbook,worksheet

def ARRAY_TO_String_Comma_Separated(Array):
    Rtn_Value =""
    for elements in Array:
        Rtn_Value = Rtn_Value + "," + str(elements)
    return Rtn_Value

def remove_Duplicates(Array_L):
    res = []
    for i in Array_L:
        if i not in res:
            res.append(i)
    return res


def Compare_All_Document_Not_Founded_In_Other_Baseline(Document_Short_name ,SWA_A , SWREQ_A,Document,Algo,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2):
    Found_Flag = 0
    Current_Comparing=""
    Found_Flag_array = []
    NOT_Found_Flag_array = []
    NOT_Found_Flag_array_Data = []
    Inconsistency_Flag = 0
    Document_List=[]
    SWA_Document_List_Compared = []

    SWA_Document_List__Ready_Compared = []
    for SWA_items in Document_Short_name:
        Document_List.append(SWA_items)
    for SWA_items in SWA_A:
        if SWA_items["Name"] in Document_List:
            SWA_Document_List__Ready_Compared.append(SWA_items)
    SWA_Document_List__Ready_Compared = remove_Duplicates(SWA_Document_List__Ready_Compared)
    for SWA_items in SWA_Document_List__Ready_Compared:
        SWREQ_Document_List_Compared =[]
        Current_Comparing=SWA_items["Name"]
        for SWRQ_items in SWREQ_A:
            if SWRQ_items["Name"] in SWREQ_Document_List_Compared:
                pass
            else:
                if SWRQ_items not in SWREQ_Document_List_Compared:
                    ###print("Start Comparing :", SWA_items["Name"], " With ", SWRQ_items["Name"])
                    SWREQ_Document_List_Compared.append(SWRQ_items["Name"])
                    SWA_SWREQ_KPI = {
                        "SWA_DOC": "",
                        "SWREQ_DOC": "",
                        "Message": 0,
                        "Missing_Ids": [],
                    }
                    SWA_SWREQ_KPI_Founded = {
                        "SWA_DOC": "",
                        "SWREQ_DOC": "",
                        "Message": 0,
                        "Missing_Ids": "",
                    }
                    SWA_SWREQ_KPI_NOT_Founded = {
                        "SWA_DOC": "",
                        "SWREQ_DOC": "",
                        "Message": 0,
                        "Missing_Ids": "",
                    }
                    ID_list = SWRQ_items["IDs"]
                    if len(ID_list) == 0:
                        pass
                    else:
                        for Id_Element in SWA_items["IDs"]:
                            if Id_Element in SWRQ_items["IDs"]:
                                SWA_SWREQ_KPI_Founded["SWA_DOC"] = SWA_items["Name"]
                                SWA_SWREQ_KPI_Founded["SWREQ_DOC"] = SWRQ_items["Name"]
                                SWA_SWREQ_KPI_Founded["Message"] = Document + " Wrong Reference"
                                SWA_SWREQ_KPI_Founded["Missing_Ids"]=Id_Element
                                ###print("Find Element ",Id_Element ," Wich was in :",SWA_items["Name"] ,"Founded in :", SWRQ_items["Name"])
                                if Id_Element not in Found_Flag_array:
                                    ###print("Find Element",Id_Element )
                                    if Algo == "Algo2_1":
                                        Missed_Document_Report_Report_Split1.append(SWA_SWREQ_KPI_Founded)
                                    elif Algo == "Algo2_2":
                                        Missed_Document_Report_Report_Split2.append(SWA_SWREQ_KPI_Founded)
                                    Found_Flag_array.append(Id_Element)
                            else:
                                Found_Flag == 1
                                if Id_Element not in NOT_Found_Flag_array:
                                    NOT_Found_Flag_array.append(Id_Element)
                                SWA_SWREQ_KPI["SWA_DOC"] = SWA_items["Name"]
                                SWA_SWREQ_KPI["SWREQ_DOC"] = ""
                                SWA_SWREQ_KPI["Message"] = Document + " WIS not mentioned anywhere"
                                SWA_SWREQ_KPI["Missing_Ids"] = Id_Element
                                if SWA_SWREQ_KPI not in NOT_Found_Flag_array_Data:
                                    NOT_Found_Flag_array_Data.append(SWA_SWREQ_KPI)
        ###print("===========================================")
    return NOT_Found_Flag_array,NOT_Found_Flag_array_Data,Found_Flag_array,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2

def Generate_SWA_Report(SWA_SWREQ_Total_KPI1,workbook1, worksheet1):
    index = 1
    for items in SWA_SWREQ_Total_KPI1:
        Data = []
        index += 1
        ###print("Arch_Document ", items["SWA_DOC"])
        ###print("REQ_Document ", items["SWREQ_DOC"])
        ###print("Match ", items["Match_Ids"])
        ###print("Missmatch ", items["Not_Match_Ids"])
        ###print("Total ", items["Total_Ids"])
        Data.append(items["SWA_DOC"])
        Data.append(items["SWREQ_DOC"])
        if items["Match_Ids"] == items["IDs"]:
            Data.append("Match")
            cc = str(items["IDs"]) + "/" + str(items["IDs"])
            Data.append(cc)
            Data.append("")
        else:
            Data.append("Miss Match")
            SWA_Total_Ids = str(items["Total_Ids"])
            Missing_Ids = items["Total_Ids"] - items["Not_Match_Ids"]
            DD = SWA_Total_Ids + "/" + str(items["Other_DOC_Ids"])
            Data.append(DD)
            missing_IDs = ARRAY_TO_String_Comma_Separated(items["Missing_Ids"])
            Data.append(missing_IDs)
        ###print("Data TO store : ", Data)
        workbook1, worksheet1 = Data_Write(workbook1, worksheet1, Data, index)
    return workbook1, worksheet1

def Generate_SWREQ_Report(SWA_SWREQ_Total_KPI2,workbook1, worksheet2):
    index = 1
    for items in SWA_SWREQ_Total_KPI2:
        Data = []
        index += 1
        ###print("Arch_Document ", items["SWA_DOC"])
        ###print("REQ_Document ", items["SWREQ_DOC"])
        ###print("Match ", items["Match_Ids"])
        ###print("Missmatch ", items["Not_Match_Ids"])
        Data.append(str(items["SWA_DOC"]))
        Data.append(str(items["SWREQ_DOC"]))
        if items["Match_Ids"] == items["IDs"]:
            Data.append("Match")
            CC = str(items["IDs"]) + "/" + str(items["IDs"])
            Data.append(CC)
            Data.append("")
        else:
            Data.append("Miss Match")
            SWREQ_Total_Ids = str(items["Total_Ids"])
            Missing_Ids = items["Total_Ids"] - items["Not_Match_Ids"]
            CC = SWREQ_Total_Ids + "/" + str(items["Other_DOC_Ids"])
            Data.append(CC)
            missing_IDs = ARRAY_TO_String_Comma_Separated(items["Missing_Ids"])
            Data.append(missing_IDs)
        ###print("Data TO store : ", Data)
        workbook1, worksheet2 = Data_Write(workbook1, worksheet2, Data, index)
    return workbook1, worksheet2

def Check_Document_Incocnsistency(SWA_IDs ,SWRQ_IDs ):
    miss_Match_FLag_Rtn = 0
    for element in SWA_IDs["IDs"]:
        if element in SWRQ_IDs["IDs"]:
            pass
        else:
            miss_Match_FLag_Rtn =1
    return miss_Match_FLag_Rtn ,




def Generate_Inconsistency_Report(SWA_SWREQ_Total_KPI2,workbook1, worksheet3, index):
    for items in SWA_SWREQ_Total_KPI2:
        ###print("Current Indix:",index)
        Data = []
        index += 1
        SWA_Document = ""
        SWREQ_Document = ""
        rtn = Document_Name_Check_Arch_REQ(items["SWREQ_DOC"])
        if rtn == "SWA":
            SWA_Document = items["SWREQ_DOC"]
        else:
            SWREQ_Document = items["SWREQ_DOC"]

        rtn = Document_Name_Check_Arch_REQ(items["SWA_DOC"])
        if rtn == "SWA":
            SWA_Document = items["SWA_DOC"]
        else:
            SWREQ_Document = items["SWA_DOC"]

        Data.append(str(SWREQ_Document))
        Data.append(str(SWA_Document))
        Data.append(str(items["Message"]))
        Data.append(str(items["Missing_Ids"]))
        Data.append("")
        workbook1, worksheet3 = Data_Write(workbook1, worksheet3, Data, index)
    return workbook1, worksheet3 , index

def Document_Variant_Check(Input_Path ,variant,Document_Name ):
    Base_P_Key = "P"
    Base_M_Key = "Q"

    Split_Key = re.split("_", Input_Path)
    x = len(Split_Key)

    num = re.sub(r'\D', "", Split_Key[x-1])

    if variant == "variant.KEY:base\+":
        #print("Variant founded ")
        Search_Key = Base_P_Key+num
        if re.search(Search_Key, Document_Name):
            return "Base+"
    elif variant == "variant.KEY:base\-":
        Search_Key = Base_M_Key + num
        if re.search(Search_Key, Document_Name):
            return "Base-"
    else:
        return 0

def Filter_Document_Variant(Input_Path ,variant,SWRQ_Arr,infoList):
    SWREQ_Array = []
    #print("Variant is ",variant)
    for element in SWRQ_Arr:
        rtn = Document_Variant_Check(infoList[4], variant, element["Name"])
        #print(rtn)
        if rtn == "Base+" :
            SWREQ_Array.append(element)
        elif rtn == "Base-" :
            SWREQ_Array.append(element)
    #print("Function pass ")
    return SWREQ_Array

def Document_Name_Check_Arch_REQ(Document):
    if Document in SWA_Doc_List:
        return "SWA"
    elif Document in SWREQ_Doc_List:
        return "SWREQ"

def Document_Name_List(SWA_Arr  , SWRQ_Arr):
    Arch_Documents =[]
    REQ_Documents = []
    for elements in SWA_Arr:
        Arch_Documents.append(elements["Name"])
    for elements in SWRQ_Arr:
        REQ_Documents.append(elements["Name"])
    return Arch_Documents , REQ_Documents

def SWA_SWREQ_Consistency_Runnable(infoList):
    Missed_Document_Report_Report_Split1 = []
    Missed_Document_Report_Report_Split2 = []

    Missed_IDs_With_Document_Report_Split1 = []
    Missed_IDs_With_Document_Report_Split2 = []

    SWA_Missed_IDs_With_Document = []
    SWREQ_Missed_IDs_With_Document = []

    #print(infoList)
    #variant = "variant.KEY:base\+"
    #infoList = Get_Input_Data()
    ##print(infoList)
    #print("Step 1 .")
    SWA_Array, SWAREQ_Array = Get_Data(infoList)
    #print("Data Generated")
    #Temp_File_Name =Generate_Temp_Report(infoList, SWA_Array, SWAREQ_Array)
    #print("Data Stored")
    #SWA_Array, SWAREQ_Array = Get_Data_Stored_Runnable("\Outputs\SWA_SWREQ_Consistency_Temp_24_11_1_31_p331.xlsx")
    #print("Step 2 .")
    # Filter Document By name
    SWAREQ_Array2 = Filter_Document_Variant(infoList[4],infoList[5], SWAREQ_Array,infoList)
    SWAREQ_Array = SWAREQ_Array2
    #print("Step 3 .")
    #print("Step 1 validation",SWAREQ_Array2)
    SWA_Doc_List , SWREQ_Doc_List = Document_Name_List(SWA_Array, SWAREQ_Array)
    #print("Step 2 validation", SWA_Doc_List , SWREQ_Doc_List)
    workbook1, worksheet,worksheet1, worksheet2, worksheet3 = Generate_Report(infoList)
    #print("Step 4 .")
    ##print("(1)=============================================================================")
    SWA_SWREQ_Total_KPI1 ,SWA_ocument_Not_have_match_name , SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document=Compare_Document_Names( SWA_Array, SWAREQ_Array,"SWA" , SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document)
    SWA_SWREQ_Total_KPI1=remove_Duplicates(SWA_SWREQ_Total_KPI1)
    workbook1, worksheet1 = Generate_SWA_Report(SWA_SWREQ_Total_KPI1, workbook1, worksheet1)
    SWA_SWREQ_Total_KPI = []
    #print("Step 5 .")
    ##print("(2)=============================================================================")
    SWA_SWREQ_Total_KPI2 ,SWRQ_ocument_Not_have_match_name,SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document = Compare_Document_Names(SWAREQ_Array,SWA_Array,"REQ",SWA_Missed_IDs_With_Document,SWREQ_Missed_IDs_With_Document)
    SWA_SWREQ_Total_KPI2 = remove_Duplicates(SWA_SWREQ_Total_KPI2)
    workbook1, worksheet2 = Generate_SWREQ_Report(SWA_SWREQ_Total_KPI2, workbook1, worksheet2)
    ##print("(3)=============================================================================")
    #print("Step 6 .")
    # find Miss-Consistency match Documents Names
    SWA_Document_Not_have_match_name, SWREQ_Document_Not_have_match_name = Find_Inconsistency_Document_Names(SWA_Array,SWAREQ_Array)
    ##print("(4)")
    SWA_Document_Not_have_match_name = remove_Duplicates(SWA_Document_Not_have_match_name)
    SWREQ_Document_Not_have_match_name = remove_Duplicates(SWREQ_Document_Not_have_match_name)
    #print("Step 7 .")
    # find Missed WIS  with Documents Names
    SWA_Missed_IDs_With_Document = remove_Duplicates(SWA_Missed_IDs_With_Document)
    SWREQ_Missed_IDs_With_Document = remove_Duplicates(SWREQ_Missed_IDs_With_Document)
    ##print("(5)")
    #print("Step 8 .")
    #############################################################33
    ### Mised IDS aLgo
    NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2=Compare_All_Document_Inconsistency_Missed_Ids(SWA_Missed_IDs_With_Document, SWAREQ_Array, "SWA","Algo1_1",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2=Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,"Algo1_1",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    #print("Step 9 .")
    ##print("(6)")
    index=1
    ##print("(7) save Data ")
    NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2 = Compare_All_Document_Inconsistency_Missed_Ids(SWREQ_Missed_IDs_With_Document, SWA_Array, "SWREQ","Algo1_2",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    ##print("(8)")
    ###print("Current Indix ===========:", index)
    #print("Step 10 .")
    Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2=Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,"Algo1_2",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    #print("Step 11 .")
    ##print("(9)")
    Missed_IDs_With_Document_Report_Split1 = remove_Duplicates(Missed_IDs_With_Document_Report_Split1)
    Missed_IDs_With_Document_Report_Split2 = remove_Duplicates(Missed_IDs_With_Document_Report_Split2)
    #print("Step 13 .")
    workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_IDs_With_Document_Report_Split1, workbook1, worksheet3,index)
    workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_IDs_With_Document_Report_Split2, workbook1, worksheet3,index)
    #############################################################33
    #print("Step 14 .")
    ### Mised Document aLgo
    NOT_Found_Flag_array,NOT_Found_Flag_array_Data,Found_Flag_array,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2= Compare_All_Document_Not_Founded_In_Other_Baseline(SWA_Document_Not_have_match_name,SWA_Array, SWAREQ_Array, "SWA","Algo2_1",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2=Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,"Algo2_1",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    #print("Step 16 .")
    NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2 = Compare_All_Document_Not_Founded_In_Other_Baseline(
        SWREQ_Document_Not_have_match_name,SWAREQ_Array, SWA_Array, "SWREQ","Algo2_2",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    #print("Step 17 .")
    Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2=Filter_Total_Missing_Ids(NOT_Found_Flag_array, NOT_Found_Flag_array_Data, Found_Flag_array,"Algo2_2",Missed_Document_Report_Report_Split1,Missed_Document_Report_Report_Split2,Missed_IDs_With_Document_Report_Split1,Missed_IDs_With_Document_Report_Split2)
    #print("Step 18 .")
    Missed_Document_Report_Report_Split1 = remove_Duplicates(Missed_Document_Report_Report_Split1)
    Missed_Document_Report_Report_Split2 = remove_Duplicates(Missed_Document_Report_Report_Split2)
    #print("Step 19 .")



    workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_Document_Report_Report_Split1, workbook1, worksheet3,
                                                                 index)
    #print("Step 20 .")
    workbook1, worksheet3, index = Generate_Inconsistency_Report(Missed_Document_Report_Report_Split2, workbook1,
                                                                 worksheet3,
                                                                 index)
    #print("Step 21 .")
    while True:
        try:
            workbook1.close()
            # os.startfile(output_workbook)
            break
        except:
            SWA_SWREQ_Consistency_Warning_Message["Message"] = "File Is already opened ! , please close Excel file "
            pass
    #print("Step 22 .")








