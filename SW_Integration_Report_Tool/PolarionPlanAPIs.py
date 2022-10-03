import connectors.polarion_connector as c
from tqdm import tqdm
import xlsxwriter
from RTE_API import Log_File


def data_analysis(workitems_list, workbook, worksheet, infoList,polarion_object):
    index = 2
    for i in tqdm(range(len(workitems_list)), desc="Loading..."):
        LWI = []
        counter = 0
        # Get workitem type.
        type = workitem_type(workitems_list[i].id, infoList,polarion_object)
        if type.type.id == 'workpackage' or type.type.id == 'defect' or type.type.id == 'issue' or\
                type.type.id == 'changeRequest' or type.type.id == 'softwareRequirement':
            try:
                # Get LWIs ID, type and status.
                for j in range(len(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem)):
                    lwiID = str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].workItemURI).split("}")[-1]
                    lwiData = workitem_type(lwiID, infoList,polarion_object)
                    if lwiData.type.id == 'task':
                        counter = counter + 1
                        if str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].role.id) == 'implements':
                            type = ''
                            LWI.append(type + '' +
                                       str(workitems_list[i].linkedWorkItemsDerived.LinkedWorkItem[j].workItemURI).split("}")[-1])
                            LWI.append(str(lwiData.type.id))
                            LWI.append(str(lwiData.status.id))
                            LWI.append(str(lwiData.severity.id))
                            LWI.append(str(lwiData.priority.id))
                            try:
                                LWI.append(str(lwiData.resolution.id))
                            except:
                                LWI.append(str(""))
                                pass
                            try:
                                LWI.append(str(lwiData.customFields.Custom.value.id))
                            except:
                                LWI.append(str(""))
                                pass



            except:
                pass
            # try:
            #     for j in range(len(workitems_list[i].linkedWorkItems.LinkedWorkItem)):
            #         LWI = str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].workItemURI).split("}")[-1]
            #         type = workitem_type(LWI, infoList)
            #         if type.id == 'task':
            #             LWI = LWI + str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].role.id) + ': '
            #             LWI = LWI + str(workitems_list[i].linkedWorkItems.LinkedWorkItem[j].workItemURI).split("}")[-1]
            #             if j != max(range(len(workitems_list[i].linkedWorkItems.LinkedWorkItem))):
            #                 LWI = LWI + ', '
            # except:
            #     counter = counter + 1
            #     pass

            index,workbook,worksheet = data_write(workitems_list, LWI, counter, i, index, workbook,worksheet)

    workbook.close()


def data_write(workitems_list, LWI, counter, i, index, workbook,worksheet):
    resolution = ""
    custumFields = ""
    if counter > 0:
        # Identify data index.
        a = 'A' + str(index)
        b = 'B' + str(index)
        c = 'C' + str(index)
        d = 'D' + str(index)
        e = 'E' + str(index)
        f = 'F' + str(index)
        g = 'G' + str(index)
        h = 'H' + str(index)
        # Data write.
        url = 'https://vseapolarion.vnet.valeo.com/polarion/redirect/project/optimus/workitem?id=' + workitems_list[i].id
        worksheet.write_url(a, str(url), string=str(workitems_list[i].id))
        worksheet.write(b, workitems_list[i].title)
        worksheet.write(c, workitems_list[i].type.id)
        worksheet.write(d, workitems_list[i].severity.id)
        worksheet.write(e, workitems_list[i].priority.id)
        worksheet.write(f, workitems_list[i].status.id)

        try :
            worksheet.write(g, workitems_list[i].resolution.id)
            resolution = workitems_list[i].resolution.id
        except:
            print("Can not find Resolution of : ",workitems_list[i].id)
            pass
        try :
            worksheet.write(h, workitems_list[i].customFields.Custom.value.id)
            custumFields = workitems_list[i].customFields.Custom.value.id
        except:
            print("Can not find custumFields of : ", workitems_list[i].id)
            pass



        z = 0
        while z < (len(LWI)/7):
            ii = 'I' + str(index)
            j = 'J' + str(index)
            k = 'K' + str(index)
            l = 'L' + str(index)
            m = 'M' + str(index)
            n = 'N' + str(index)
            o = 'O' + str(index)
            worksheet.write(ii, LWI[z*7])
            worksheet.write(j, LWI[(z*7)+1])
            worksheet.write(k, LWI[(z*7)+2])
            worksheet.write(l, LWI[(z * 7) + 3])
            worksheet.write(m, LWI[(z * 7) + 4])
            worksheet.write(n, LWI[(z * 7) + 5])
            worksheet.write(o, LWI[(z * 7) + 6])
            index = index + 1
            z = z + 1
        if z > 1:
            print("Input Data" , url, workitems_list[i].id, workitems_list[i].title,workitems_list[i].type.id,workitems_list[i].severity.id,workitems_list[i].priority.id, workitems_list[i].status.id, resolution, custumFields)
            workbook,worksheet=cell_merge(index-z, index-1, url, str(workitems_list[i].id), workitems_list[i].title,workitems_list[i].type.id,workitems_list[i].severity.id,workitems_list[i].priority.id, workitems_list[i].status.id, resolution, custumFields,workbook,worksheet)


    return index,workbook,worksheet


def cell_merge(x, y, url, id, title, type,severity,priority,status,resolution,customFields,workbook,worksheet):
    # Create a format to use in the merged range.
    merge_format = workbook.add_format({
        'valign': 'vcenter'})
    merge_format1 = workbook.add_format({
        'font_color': 'blue',
        'underline': 1,
        'valign': 'vcenter'})

    a = 'A' + str(x) + ':A' + str(y)
    b = 'B' + str(x) + ':B' + str(y)
    c = 'C' + str(x) + ':C' + str(y)
    d = 'D' + str(x) + ':D' + str(y)
    e = 'E' + str(x) + ':E' + str(y)
    f = 'F' + str(x) + ':F' + str(y)
    g = 'G' + str(x) + ':G' + str(y)
    h = 'H' + str(x) + ':H' + str(y)

    # Merge cells.
    worksheet.merge_range(a, url, merge_format)
    worksheet.merge_range(b, title, merge_format)
    worksheet.merge_range(c, type, merge_format)
    worksheet.merge_range(d, severity, merge_format)
    worksheet.merge_range(e, priority, merge_format)
    worksheet.merge_range(f, status, merge_format)
    worksheet.merge_range(g, resolution, merge_format)
    worksheet.merge_range(h, customFields, merge_format)


    # Edit ID cells Hyperlink and format.
    a = 'A' + str(x)
    worksheet.write_url(a, str(url), merge_format1, string=str(id))
    return workbook,worksheet


def workitem_type(id, infoList,polarion_object):
    # New query based on workitems ID to get the type.
    query = str("project.id:" + infoList[2] + " AND id:" + id)

    type = polarion_object.tracker_webservice.service.queryWorkItems(query, "priority", ["type", "status" ,"severity" , "priority" , "resolution" , "customFields.subProjectDiscipline.KEY"])

    return type[0]


def polarion_query(infoList):
    print("Connecting to Polarion...")
    NOTConnected = True

    query = "project.id:" + infoList[2] + " AND PLAN:(" + infoList[2] + "/" + infoList[3] + ")"

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get workitems.
            workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
                query, "priority", ["id", "title", "status", "type", "linkedWorkItems", "linkedWorkItemsDerived" ,"severity" , "priority" , "resolution" , "customFields.subProjectDiscipline.KEY" ] )
            NOTConnected = False
        except Exception as e:
            #print(str(e))
            pass
        
    return workitems_list, polarion_object


def excel_open():
    # Create / Open Excel file.
    workbook = xlsxwriter.Workbook('Outputs\Polarion_Plan.xlsx')
    worksheet = workbook.add_worksheet()

    # Add columns title.
    worksheet.write("A1", 'ID')
    worksheet.write("B1", 'Title')
    worksheet.write("C1", 'Type')
    worksheet.write("D1", 'Severity')
    worksheet.write("E1", 'Priority')
    worksheet.write("F1", 'Status')
    worksheet.write("G1", 'Resolution')

    worksheet.write("H1", 'Linked Work Items')
    worksheet.write("I1", 'LWI:Type')
    worksheet.write("J1", 'LWI:Status')
    worksheet.write("K1", 'LWI:Severity')
    worksheet.write("L1", 'LWI:Priority')
    worksheet.write("M1", 'LWI:Status')
    worksheet.write("N1", 'LWI:Resolution')

    return workbook, worksheet


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

    if infoList[2] == "100kW" or infoList[2] == "100kw" or infoList[2] == "100KW" or infoList[2] == "100Kw":
        infoList[2] = "optimus"

    return infoList


def print_script_version():
    # Script Version Number.
    script_version = '1.0.0'
    script_build_number = '(1)'

    print('Version', script_version, script_build_number)






def Connect_to_polarion(infoList):
    print("Connecting to Polarion...")
    NOTConnected = True

    query = "project.id:" + infoList[2] + " AND PLAN:(" + infoList[2] + "/" + infoList[3] + ")"

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get workitems.
            NOTConnected = False
        except Exception as e:
            # print(str(e))
            pass
    return polarion_object


def Gerrit_workitem_type(id, infoList,polarion_object):
    # New query based on workitems ID to get the type.
    query = str("project.id:" + infoList[2] + " AND id:" + id)
    #print("Getting for ID :", id)
    try:
        type = polarion_object.tracker_webservice.service.queryWorkItems(query, "priority", ["type", "status" ,"severity" , "priority" , "resolution" , "customFields.subProjectDiscipline.KEY"])
        #print("Data Founded")
        return type[0]
    except:
        return "Invalid"



def Get_Gerrit_WI_Data(Task_List , infoList):
    #print("Pass 1")
    WI_Status_array =[]

    polarion_object =Connect_to_polarion(infoList)
    #print("Pass 2")
    for id in Task_List:
        WI_Status = {
            "ID": "",
            "Data": []
        }
        if id== 0 :
            pass
        else:
            Data = Gerrit_workitem_type(id, infoList, polarion_object)
            if Data == "Invalid" :
                #print("NOT ----- Valid Data of :", id)
                WI_Status["ID"] = id
                WI_Status["Data"] = "Not project Data"
                WI_Status_array.append(WI_Status)
            else:
                #print("Valid Data of :" , id)
                WI_Status["ID"] = id
                WI_Status["Data"] = Data
                WI_Status_array.append(WI_Status)
    #print("Pass 3")
    return WI_Status_array


def Polarion_Plan_Runnable(infoList,Tags_List):
    # Print Script Version.
    #print_script_version()
    # Get user input data.
    #infoList = Get_Input_Data()
    #print(infoList)
    # Create / Open Excel file.
    #####workbook, worksheet = excel_open()
    # Get polarion data.
    ####workitems_list, polarion_object = polarion_query(infoList)
    # Analize the data then write it in Excel file.
    ####data_analysis(workitems_list, workbook, worksheet, infoList,polarion_object)
    # Disconnet polarion.
    ####polarion_object.disconnect()

    #print("Gerrit IDs" ,Tags_List )

    WI_Status_array = Get_Gerrit_WI_Data(Tags_List, infoList)
    #print("Pass 6" , WI_Status_array)
    Polarion_Ticket_Data =[]
    Gerrit_Tickets_Info =[]
    for elements in WI_Status_array:
        Tickets_Info ={
            "ID": "" ,
            "priority": "" ,
            "severity": "" ,
            "status": "" ,
            "resolution": "" ,
            "FGL": "" ,
        }
        Tickets_Info["ID"]=elements["ID"]
        try:
            Tickets_Info["status"] = elements["Data"].status.id
            Tickets_Info["priority"] = elements["Data"].priority.id
            Tickets_Info["severity"] = elements["Data"].severity.id
            Tickets_Info["resolution"] = elements["Data"].resolution.id

            Tickets_Info["FGL"]=elements["Data"].customFields['Custom'][0]['value']['EnumOptionId'][0]['id']
        except :
            Tickets_Info["FGL"] = ""
        Polarion_Ticket_Data.append(Tickets_Info)
    #print("Pass 7")
    return Polarion_Ticket_Data


#infoList=['melmohta','Aliya$1994','optimus','B03_00']
#Task_List = ['100kW-61978', '100kW-60413', '100kW-59749', '100kW-42315', '100kW-52768', '100kW-59439', '100kW-61752', '100kW-61719', '100kW-61525', 'DEVOPS-2068']
#Polarion_Ticket_Data =Get_Gerrit_WI_Data(Task_List , infoList)
#print(Polarion_Ticket_Data)
#Polarion_Plan_Runnable(infoList)

