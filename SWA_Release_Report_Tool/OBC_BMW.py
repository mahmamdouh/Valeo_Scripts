'''
/* ********************************************************************** */
/* Sourcefile:      REQ_SWC_Traceability.py                               */
/*                                                                        */
/* Department:      VSeA-Cairo                                            */
/*                                                                        */
/* Author :      Ahmed SOLIMAN-MAHMOUD                                    */
/* Mail :        ahmed.soliman-mahmoud@valeo.com                          */
/* Author :      Mahmoud Elmohtady                                        */
/* Mail :        mahmoud.elmohtady@valeo.com                              */
/*                                                                        */
/* Version :        2.0.0                                                 */
/* ********************************************************************** */
/* Copyright (C) Valeo 2022                                               */
/* All Rights Reserved.  Confidential                                     */
/* ********************************************************************** */
'''

import connectors.polarion_connector as c
import xlsxwriter
import re
from Pichart import Pi_Chart_creation
from datetime import datetime



def polarion_query_SW_interface(infoList):
    #print("Connecting to Polarion...")
    NOTConnected = True

    # Set polarion query.
    sql = "SQL:(select WI.C_PK from MODULE M inner join REL_MODULE_WORKITEM RMW ON RMW.FK_URI_MODULE = M.C_URI inner " \
          "join WORKITEM WI on WI.C_URI = RMW.FK_URI_WORKITEM where (M.C_LOCATION like '%23_01_02_gdd%' ))"

    query = "project.id:obc_35up11kw AND type:sw_interface AND " + sql + " AND status: released "

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get workitems.
            workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
                query, "priority", ["id", "title", "type","linkedWorkItems" ,"linkedWorkItemsDerived" ,"customFields.lower_limit.KEY","customFields.upper_limit.KEY","customFields.update_constraint.KEY","customFields.coded_type.KEY","customFields.init_value.KEY"])
            #print("Receiving Requirements data...")
            NOTConnected = False
        except Exception as e:
            print(str(e))
            #print("Retry...")
            pass
        except:
            pass

    return workitems_list, polarion_object



def Get_Input_Data():
    infoList = []

    # Get data from Input_Data_File.txt.
    with open("Input\Polarion.txt") as infoFile:
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


def data_analysis_SWInterface(workitems_list,workbook1, worksheet):
    index = 1
    index1 = 1
    page_indix = 2
    #x=10
    sql = "SQL:(select WI.C_PK from MODULE M inner join REL_MODULE_WORKITEM RMW ON RMW.FK_URI_MODULE = M.C_URI inner " \
          "join WORKITEM WI on WI.C_URI = RMW.FK_URI_WORKITEM where (M.C_LOCATION like '%23_01_02_gdd%' ))"

    for i in range(len(workitems_list)):
        Data = []
        # Check if the workitem type is SW requirement.
        if workitems_list[i].type.id == "sw_interface":
            index += 1
            Interface_ID = workitems_list[i].id
            Interface_Title = workitems_list[i].title
            Data.append(Interface_ID)
            Data.append(Interface_Title)
            #print("Interface ", Interface_ID , Interface_Title)
            lwiIDs_P = []
            lwiIDs_U = []
            try:
                # Search in the LWI elements for realize elements, get it's type and save it's ID.
                for j in range(len(workitems_list[i].linkedWorkItems)):
                    #print("Role in WIs",str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].role.id))
                    role = str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].role.id)
                    if str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].role.id) == 'provide':
                        Query = sql + "AND id:OBC_35U11K\-"
                        workitemID = str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].workItemURI).split("}")[-1]
                        #workitemData = workitem_type(workitemID, infoList, polarion_object)
                        #print(str(workitemID))
                        ID_Number = re.split("-", str(workitemID))
                        Query = Query+ID_Number[1]
                        #print(ID_Number[1])
                        SWC_List = polarion_object.tracker_webservice.service.queryWorkItems(
                            Query, "priority", ["id", "title", "type"])
                        #print("WI List" ,SWC_List[0].title)
                        lwiIDs_P.append(str(SWC_List[0].title))
                for j in range(len(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem)):
                    try:
                        #print("Linked WIs Role",str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].role.id))
                        if str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].role.id) == "use":
                            Query = sql + "AND id:OBC_35U11K\-"
                            #print("Is used BY " ,workitemID )
                            workitemID = str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].workItemURI).split("}")[-1]
                            #workitemData = workitem_type(workitemID, infoList,polarion_object)
                            ID_Number = re.split("-", str(workitemID))
                            Query = Query + ID_Number[1]
                            # print(ID_Number[1])
                            SWC_List = polarion_object.tracker_webservice.service.queryWorkItems(
                                Query, "priority", ["id", "title", "type"])
                            #print("WI List", SWC_List[0].title)
                            lwiIDs_U.append(str(SWC_List[0].title))

                    except:
                        pass
            except:
                pass

            # Add ',' after each ID.
            lwiData_P = ""
            if len(lwiIDs_P) > 0:
                for x in range(len(lwiIDs_P)):
                    lwiData_P = lwiData_P + str(lwiIDs_P[x])
                    # Check if the last ID is added to remove the ','.
                    if x + 1 < len(lwiIDs_P):
                        lwiData_P = lwiData_P + "; "
            lwiData_U = ""
            if len(lwiIDs_U) > 0:
                for x in range(len(lwiIDs_U)):
                    lwiData_U = lwiData_U + str(lwiIDs_U[x])
                    # Check if the last ID is added to remove the ','.
                    if x + 1 < len(lwiIDs_U):
                        lwiData_U = lwiData_U + "; "

            Data.append(lwiData_P)
            Data.append(lwiData_U)
        ### get all WI attributes

        #if workitems_list[i].customFields.Custom[0].key ==
        Lower_Limit = workitems_list[i].customFields.Custom[0].value
        Upper_Limit = workitems_list[i].customFields.Custom[1].value
        Update_Constraint = workitems_list[i].customFields.Custom[2].value
        Code_Type = workitems_list[i].customFields.Custom[3].value
        Init_Value = ""
        #if workitems_list[i].customFields.Custom[0].key ==
        try :
            Init_Value = workitems_list[i].customFields.Custom[4].value.content
        except:
            Init_Value = ""
            #print("Interface : " , reqID,reqTitle , lwiData_P,lwiData_U)
        Data.append(Code_Type)
        Data.append(Init_Value)
        Data.append(Lower_Limit)
        Data.append(Upper_Limit)
        workbook1, worksheet= Data_Weite(workbook1, worksheet, Data, page_indix)
        page_indix = page_indix + 1

    return


def excel_open(info_list):
    # Create / Open Excel file.
    File_Name = "SubFiles\Sw interfaces.xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet= workbook1.add_worksheet("Sw interfaces")
    # Add columns title.
    worksheet.write("A1", 'ID')
    worksheet.write("B1", 'Title')
    worksheet.write("C1", 'Provided by SWC')
    worksheet.write("D1", 'Used by SWC ')
    worksheet.write("E1", 'Coded Type')
    worksheet.write("F1", 'Init Value')
    worksheet.write('G1', "Lower Limit")
    worksheet.write('H1', 'Upper Limit')

    return workbook1, worksheet

def Data_Weite(WB , Sheet , Data ,index):
    a = 'A' + str(index)
    b = 'B' + str(index)
    c = 'C' + str(index)
    d = 'D' + str(index)
    e = 'E' + str(index)
    f = 'F' + str(index)
    g = 'G' + str(index)
    h = 'H' + str(index)

    Sheet.write(a,Data[0])
    Sheet.write(b,Data[1])
    Sheet.write(c,Data[2])
    Sheet.write(d,Data[3])
    Sheet.write(e,Data[4])
    Sheet.write(f,Data[5])
    Sheet.write(g,Data[6])
    Sheet.write(h,Data[7])

    return WB , Sheet



infoList = Get_Input_Data()
workbook1, worksheet = excel_open(infoList)
# Create / Open Excel file.
workitems_list, polarion_object = polarion_query_SW_interface(infoList)
data_analysis_SWInterface(workitems_list,workbook1, worksheet)
print("Data Finished")
workbook1.close()