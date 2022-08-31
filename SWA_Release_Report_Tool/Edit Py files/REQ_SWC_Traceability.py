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
from Pichart import Pi_Chart_creation
from datetime import datetime

SWC_DIAG_IDS = {}
KPI ={
    "SWREQNum" : 0,
    "SWRDINum" : 0,
    "SWCNum" : 0,
    "Covered_SWREQ": 0,
    "Covered_DIAG": 0,
    "Covered_SWC": 0,
}

Missed_SWRq=[]
Missed_SWC=[]
Missed_DWI=[]
def data_analysis_REQvsSWComp(workitems_list, workbook, worksheet1,worksheet3,worksheet4, infoList,polarion_object):
    index = 1
    index1 = 1
    #x=10
    for i in range(len(workitems_list)):
        SWC_Flag =0
        #x=x+1

        # Check if the workitem type is SW requirement.
        if workitems_list[i].type.id == "softwareRequirement":
            KPI["SWREQNum"] = KPI["SWREQNum"] +1
            index += 1
            reqID = workitems_list[i].id
            reqTitle = workitems_list[i].title
            lwiIDs = []
            try:
                # Search in the LWI elements for realize elements, get it's type and save it's ID.
                for j in range(len(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem)):
                    try:
                        if str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].role.id) == "realize":
                            workitemID = str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].workItemURI).split("}")[-1]
                            workitemData = workitem_type(workitemID, infoList,polarion_object)
                            if workitemData.type.id == "swComponent":
                                lwiIDs.append(str(workitemID))
                                SWC_Flag = 1
                    except:
                        pass
                if SWC_Flag:
                    KPI["Covered_SWREQ"] = KPI["Covered_SWREQ"] +1
                else:
                    Missed_SWRq.append(workitems_list[i].id)
                    #worksheet3 =KPI_data_write_SWRQ_LostIds(workbook,workitems_list[i], x, worksheet3, infoList)
            except:
                pass

            # Add ',' after each ID.
            lwiData = ""
            if len(lwiIDs) > 0:
                for x in range(len(lwiIDs)):
                    lwiData = lwiData + str(lwiIDs[x])
                    # Check if the last ID is added to remove the ','.
                    if x + 1 < len(lwiIDs):
                        lwiData = lwiData + ", "


            data_write(workbook,reqID, reqTitle, lwiData, index, worksheet1,infoList,"Null","Null")
        if workitems_list[i].type.id == "diagnostic":
            ##print("WOrkitem : ", workitems_list[i].id)
            ##print("Type : ", workitems_list[i].type.id)
            KPI["SWRDINum"] = KPI["SWRDINum"] + 1
            index1 += 1
            reqID = workitems_list[i].id
            reqTitle = workitems_list[i].title
            DIAGlwiIDs = []
            try:
                # Search in the LWI elements for realize elements, get it's type and save it's ID.
                for j in range(len(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem)):
                    try:
                        if str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].role.id) == "realize":
                            workitemID = \
                            str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].workItemURI).split("}")[-1]
                            workitemData = workitem_type(workitemID, infoList, polarion_object)
                            if workitemData.type.id == "swComponent":
                                DIAGlwiIDs.append(str(workitemID))
                                SWC_Flag = 1
                    except:
                        pass
                ##print("DIag WI ",reqID , reqTitle )
                ##print("Linked DIag WI ",DIAGlwiIDs )
                if SWC_Flag:
                    KPI["Covered_DIAG"] = KPI["Covered_DIAG"] + 1
                else:
                    Missed_DWI.append(workitems_list[i].id)
                    # worksheet3 =KPI_data_write_SWRQ_LostIds(workbook,workitems_list[i], x, worksheet3, infoList)
            except:
                pass

            # Add ',' after each ID.
            lwiData = ""
            if len(DIAGlwiIDs) > 0:
                for x in range(len(DIAGlwiIDs)):
                    lwiData = lwiData + str(DIAGlwiIDs[x])
                    # Check if the last ID is added to remove the ','.
                    if x + 1 < len(DIAGlwiIDs):
                        lwiData = lwiData + ", "
            ##print("DIAG : " ,reqID )
            data_write(workbook, reqID, reqTitle, lwiData, index1, worksheet4, infoList, "Null", "Null")
    #workbook.save('Outputs\Req_SWC_Bi_Directional_Report.xlsx')
    # Clear data.
    workitems_list = []
    return workbook
    #workbook.close()




def data_analysis_SWCompvsREQ(workitems_list, workbook, worksheet2,worksheet3, infoList,polarion_object):
    index = 1
    #x=10
    ##print("Start Data analysis ")
    for i in range(len(workitems_list)):
        SWR_Flag = 0
        #x=x+1
        # Check if the workitem type is SW Component.
        if workitems_list[i].type.id == "swComponent" :
            index += 1
            swCompID = workitems_list[i].id
            swCompTitle = workitems_list[i].title
            lwiIDs = []
            DIAG_WI_Ids =[]
            try:
                # Search in the LWI elements for realize elements, get it's type and save it's ID.
                for j in range(len(workitems_list[i].linkedWorkItems.LinkedWorkItem)):
                    try:
                        if str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].role.id) == "realize":
                            workitemID = str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].workItemURI).split("}")[-1]
                            workitemData = workitem_type(workitemID, infoList,polarion_object)
                            if workitemData.type.id == "softwareRequirement":
                                lwiIDs.append(str(workitemID))
                                SWR_Flag = 1
                            elif workitemData.type.id == "diagnostic":
                                DIAG_WI_Ids.append(str(workitemID))
                                SWC_DIAG_IDS[workitems_list[i].id] = workitemID
                            #else:
                                ##print("Missed Relizes :", workitemData.type.id)
                                ##print("Linked Relizes :", workitemID)
                    except:
                        pass
                if SWR_Flag:
                    KPI["Covered_SWC"] = KPI["Covered_SWC"] +1
                else:
                    Missed_SWC.append(workitems_list[i].id)
                    #worksheet3 = KPI_data_write_SWC_LostIds(workbook,workitems_list[i], x, worksheet3, infoList)
            except:
                pass

            # Add ',' after each ID.
            lwiData = ""
            if len(lwiIDs) > 0:
                for x in range(len(lwiIDs)):
                    lwiData = lwiData + str(lwiIDs[x])
                    # Check if the last ID is added to remove the ','.
                    if x + 1 < len(lwiIDs):
                        lwiData = lwiData + ", "

            # Add ',' after each ID.
            DIAGData = ""
            if len(DIAG_WI_Ids) > 0:
                for x in range(len(DIAG_WI_Ids)):
                    DIAGData = DIAGData + str(DIAG_WI_Ids[x])
                    # Check if the last ID is added to remove the ','.
                    if x + 1 < len(DIAG_WI_Ids):
                        DIAGData = DIAGData + ", "
            # Write the requirement ID and LWI IDs.
            data_write(workbook,swCompID, swCompTitle, lwiData, index, worksheet2,infoList,DIAGData,"DIAG")
    ##print("finalize Start Data analysis ")
    return workbook
    #workbook.close()


def workitem_type(id, infoList,polarion_object):
    # New query based on workitems ID to get the type.
    query = str("project.id:" + infoList[2] + " AND id:" + id)

    type = polarion_object.tracker_webservice.service.queryWorkItems(query, "priority", ["type", "status"])

    return type[0]


def data_write(workbook1,data1, data2, data3, index, worksheet,infoList,DIAG_WI_Ids,Type):
    # Identify data index.
    a = 'A' + str(index)
    b = 'B' + str(index)
    c = 'C' + str(index)
    d = 'D' + str(index)

    url = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + infoList[2] + "/workitem?id=" + data1

    # Data write.
    worksheet.write_url(a, url, string=data1)
    worksheet.write(b, data2)
    worksheet.write(c, data3)
    if Type == "DIAG":
        worksheet.write(d, DIAG_WI_Ids)
    #workbook1.save('Outputs\Req_SWC_Bi_Directional_Report.xlsx')

def KPI_data_write(workbook1,data1, data2, data3, data4, worksheet3,infoList):
    # data1 => total covered REq
    # data2 => total covered SWC
    # data3 => Not covered req
    # data4 => Not covered SWC
    # data5 => Not covered SWC IDs
    # data6 => Not covered SWC IDS
    # Identify data index.
    a1 = 'A1'
    a2 = 'A2'
    a3 = 'A3'
    a4 = 'A4'

    b1 = 'B1'
    b2 = 'B2'
    b3 = 'B3'
    b4 = 'B4'
    ##print("Data #print")
    ##print("Data #print",data1, data2, data3, data4)
    worksheet3.write(a1, "Total Req Number")
    worksheet3.write(a2, "Not covered Req Number")
    worksheet3.write(a3, "Total SWC Number")
    worksheet3.write(a4, "Not covered SWC Number")

    worksheet3.write(b1, data1)
    worksheet3.write(b2, data2)
    worksheet3.write(b3, data3)
    worksheet3.write(b4, data4)
    return workbook1

def KPI_data_write_SWC_LostIds(workbook1,data6, index, worksheet,infoList):
    # data1 => total covered REq
    # data2 => total covered SWC
    # data3 => Not covered req
    # data4 => Not covered SWC
    # data5 => Not covered SWC IDs
    # data6 => Not covered SWC IDS
    bb = 'E' + str(index)
    dd = 'D' + str(index)
    url2 = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + infoList[2] + "/workitem?id=" + data6


    # Data write.
    worksheet.write_url(bb, url2, string=data6)
    #if data6 in SWC_DIAG_IDS:
    #    worksheet.write_url(dd, url2, string=SWC_DIAG_IDS[data6])
    return worksheet

def KPI_data_write_SWRQ_LostIds(workbook1,data5, index, worksheet,infoList):
    # data1 => total covered REq
    # data2 => total covered SWC
    # data3 => Not covered req
    # data4 => Not covered SWC
    # data5 => Not covered SWC IDs
    # data6 => Not covered SWC IDS
    aa = 'A' + str(index)
    url2 = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + infoList[2] + "/workitem?id=" + data5

    # Data write.
    worksheet.write_url(aa, url2, string=data5)
    return worksheet

def KPI_data_write_DWI_LostIds(workbook1,data5, index, worksheet,infoList):
    # data1 => total covered REq
    # data2 => total covered SWC
    # data3 => Not covered req
    # data4 => Not covered SWC
    # data5 => Not covered SWC IDs
    # data6 => Not covered SWC IDS
    aa = 'C' + str(index)
    url2 = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + infoList[2] + "/workitem?id=" + data5

    # Data write.
    worksheet.write_url(aa, url2, string=data5)
    return worksheet

def polarion_query_REQvsSWComp(infoList):
    ##print("Connecting to Polarion...")
    NOTConnected = True

    # Set polarion query.
    sql = "SQL:(select WI.C_PK from MODULE M inner join REL_MODULE_WORKITEM RMW ON RMW.FK_URI_MODULE = M.C_URI inner " \
          "join WORKITEM WI on WI.C_URI = RMW.FK_URI_WORKITEM where (M.C_LOCATION like '%" + infoList[3] + "%' ))"

    if infoList[2] == "VW_MEB_Inverter":
        query = "project.id:" + infoList[2] + " AND type:(diagnostic softwareRequirement) AND " + sql + " AND NOT status: obsolete " \
                                                            "AND variant.KEY: (base\+ base\-)"
    else:
        query = "project.id:" + infoList[2] + " AND type:(diagnostic softwareRequirement) AND " + sql + " AND NOT status: obsolete "

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get workitems.
            workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
                query, "priority", ["id", "title", "type", "linkedWorkItemsDerived"])
            ##print("Receiving Requirements data...")
            NOTConnected = False
        except Exception as e:
            #print(str(e))
            ##print("Retry...")
            pass

    return workitems_list, polarion_object


def polarion_query_DWIvsSWComp(infoList):
    ##print("Connecting to Polarion...")
    NOTConnected = True

    # Set polarion query.
    sql = "SQL:(select WI.C_PK from MODULE M inner join REL_MODULE_WORKITEM RMW ON RMW.FK_URI_MODULE = M.C_URI inner " \
          "join WORKITEM WI on WI.C_URI = RMW.FK_URI_WORKITEM where (M.C_LOCATION like '%" + infoList[3] + "%' ))"

    if infoList[2] == "VW_MEB_Inverter":
        query = "project.id:" + infoList[2] + " AND type:diagnostic AND " + sql + " AND NOT status: obsolete " \
                                                            "AND variant.KEY: (base\+ base\-)"
    else:
        query = "project.id:" + infoList[2] + " AND type:diagnostic AND " + sql + " AND NOT status: obsolete "

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get workitems.
            workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
                query, "priority", ["id", "title", "type", "linkedWorkItemsDerived"])
            ##print("Receiving Requirements data...")
            NOTConnected = False
        except Exception as e:
            #print(str(e))
            ##print("Retry...")
            pass

    return workitems_list, polarion_object



def polarion_query_SWCompvsREQ(infoList,polarion_object):
    # Set polarion query.
    sql = "SQL:(select WI.C_PK from MODULE M inner join REL_MODULE_WORKITEM RMW ON RMW.FK_URI_MODULE = M.C_URI inner " \
          "join WORKITEM WI on WI.C_URI = RMW.FK_URI_WORKITEM where (M.C_LOCATION like '%" + infoList[4] + "%' ))"

    query = "project.id:" + infoList[2] + " AND type: swComponent AND " + sql + " AND variant.KEY: (base\+ base\-)"

    ##print("Receiving SW Components data...")
    workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
        query, "priority", ["id", "title", "type", "linkedWorkItems"])

    return workitems_list


def excel_open(info_list):
    # Create / Open Excel file.
    File_Name = "Outputs\Req_SWC_Bi_Directional_Report_"+info_list[4]+".xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet= workbook1.add_worksheet("Execution details")
    worksheet3 = workbook1.add_worksheet("KPI")
    worksheet1 = workbook1.add_worksheet("Req Vs SWC")
    worksheet4 = workbook1.add_worksheet("DWI Vs SWC")
    worksheet2 = workbook1.add_worksheet("SWC Vs REq")
    worksheet1.autofilter('A1:D1000')
    worksheet4.autofilter('A1:D1000')
    worksheet2.autofilter('A1:D1000')


    # Add columns title.
    worksheet1.write("A1", 'Req ID')
    worksheet1.write("B1", 'Title')
    worksheet1.write("C1", 'SWC')

    worksheet4.write("A1", 'DIAG ID')
    worksheet4.write("B1", 'Title')
    worksheet4.write("C1", 'SWC')

    worksheet2.write("A1", 'SWC ID')
    worksheet2.write("B1", 'Title')
    worksheet2.write("C1", 'Req')
    worksheet2.write("D1", 'Diagnostic')

    worksheet3.write("A19", 'Not Covered SWREQ')
    worksheet3.write("B19", 'Justification')
    worksheet3.write("C19", 'Not Covered DWI')
    worksheet3.write("D19", 'Justification')
    worksheet3.write("E19", 'Not Covered SWc')
    worksheet3.write("F19", 'Justification')
    #print("step 1")
    cell_format = workbook1.add_format()
    cell_format.set_bold()
    now = datetime.now()  # current date and time
    Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
    #print("step 2")
    worksheet.write('A1', "Exection Date",cell_format)
    worksheet.write('B1', Date_Data)
    worksheet.write('A2', "Generated By ",cell_format)
    worksheet.write('B2', str(info_list[0]))
    worksheet.write('A3', "SWA Baseline ",cell_format)
    worksheet.write('B3', str(info_list[4]))
    worksheet.write('A4', "SWREQ Baseline ",cell_format)
    worksheet.write('B4', str(info_list[3]))
    #print("step 3")
    return workbook1, worksheet1, worksheet2 ,worksheet3,worksheet4


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


def #print_script_version():
    # Script Version Number.
    return("Version : 3.2.1 ")

    ##print('Version', script_version, script_build_number)


def BAD_IDS(workbook1,SWRQ_List , SWC_List ,DWI_List,worksheet ,infoList):
    #print("Start")
    for i in range(len(SWRQ_List)):
        ##print(str(SWRQ_List[i]))
        worksheet= KPI_data_write_SWRQ_LostIds(workbook1, str(SWRQ_List[i]), (i+20), worksheet, infoList)
    for i in range(len(SWC_List)):
        ##print(str(SWC_List[i]))
        worksheet= KPI_data_write_SWC_LostIds(workbook1, str(SWC_List[i]), (i+20), worksheet, infoList)
    for i in range(len(DWI_List)):
        worksheet=KPI_data_write_DWI_LostIds(workbook1, str(SWC_List[i]), (i+20), worksheet, infoList)

    workbook1.close()
    #print("Finish")
    return 0





'''
# #print Script Version.
#print_script_version()
# Get user input data.
infoList = Get_Input_Data()
# Create / Open Excel file.
workbook1, worksheet1, worksheet2 ,worksheet3,excel_open = excel_open(infoList)
# Get polarion data 1st direction.

workitems_list, polarion_object = polarion_query_REQvsSWComp(infoList)
workitems_list2, polarion_object = polarion_query_DWIvsSWComp(infoList)

KPI["SWREQNum"]=len(workitems_list)
# Analyze the data then write it in Excel file 1st direction.
workbook1= data_analysis_REQvsSWComp(workitems_list, workbook1, worksheet1,worksheet3, infoList,polarion_object)
# Get polarion data 2nd direction.
workitems_list = polarion_query_SWCompvsREQ(infoList,polarion_object)
KPI["SWCNum"]=len(workitems_list)
# Analyze the data then write it in Excel file 2nd direction.
workbook1= data_analysis_SWCompvsREQ(workitems_list, workbook1, worksheet2,worksheet3, infoList,polarion_object)
polarion_object.disconnect()

#workbook1 = KPI_data_write(workbook1,KPI["SWREQNum"], KPI["Covered_SWREQ"], KPI["SWCNum"], KPI["Covered_SWC"], worksheet3,infoList)
#Missed_SWRq=['VWMEB-Inv-87299','VWMEB-Inv-87551','VWMEB-Inv-87557','VWMEB-Inv-118774','VWMEB-Inv-85305','VWMEB-Inv-87420']
#Missed_SWC=['VWMEB-Inv-87299','VWMEB-Inv-87551','VWMEB-Inv-87557','VWMEB-Inv-118774','VWMEB-Inv-85305','VWMEB-Inv-87420']
data = [
        ['Total SWREQ', 'Covered SWREQ', 'Total DWI' ,'Covered DWI', 'Total SWC' ,'Covered SWC'],
        [0, 0, 0 , 0, 0 , 0],
    ]
data[1][0]=KPI["Covered_SWREQ"]
data[1][1]=KPI["SWREQNum"]-KPI["Covered_SWREQ"]
data[1][2]=************8
data[1][3]=************8
data[1][4]=KPI["Covered_SWC"]
data[1][5]=KPI["SWCNum"]-KPI["Covered_SWC"]

workbook1= Pi_Chart_creation(workbook1,worksheet3,data)
BAD_IDS(workbook1,Missed_SWRq , Missed_SWC ,worksheet3 ,infoList)

'''