import connectors.polarion_connector as c
from tqdm import tqdm
import xlsxwriter


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
    if counter > 0:
        # Identify data index.
        a = 'A' + str(index)
        b = 'B' + str(index)
        c = 'C' + str(index)
        d = 'D' + str(index)

        # Data write.
        url = 'https://vseapolarion.vnet.valeo.com/polarion/redirect/project/optimus/workitem?id=' + workitems_list[i].id
        worksheet.write_url(a, str(url), string=str(workitems_list[i].id))
        worksheet.write(b, workitems_list[i].title)
        worksheet.write(c, workitems_list[i].type.id)
        worksheet.write(d, workitems_list[i].status.id)
        j = 0
        while j < (len(LWI)/3):
            e = 'E' + str(index)
            f = 'F' + str(index)
            g = 'G' + str(index)
            worksheet.write(e, LWI[j*3])
            worksheet.write(f, LWI[(j*3)+1])
            worksheet.write(g, LWI[(j*3)+2])
            index = index + 1
            j = j + 1
        if j > 1:
            workbook,worksheet=cell_merge(index-j, index-1, url, workitems_list[i].id, workitems_list[i].title,
                       workitems_list[i].type.id, workitems_list[i].status.id,workbook,worksheet)

    return index,workbook,worksheet


def cell_merge(x, y, url, id, title, type, status,workbook,worksheet):
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

    # Merge cells.
    worksheet.merge_range(a, url, merge_format)
    worksheet.merge_range(b, title, merge_format)
    worksheet.merge_range(c, type, merge_format)
    worksheet.merge_range(d, status, merge_format)

    # Edit ID cells Hyperlink and format.
    a = 'A' + str(x)
    worksheet.write_url(a, str(url), merge_format1, string=str(id))
    return workbook,worksheet


def workitem_type(id, infoList,polarion_object):
    # New query based on workitems ID to get the type.
    query = str("project.id:" + infoList[2] + " AND id:" + id)

    type = polarion_object.tracker_webservice.service.queryWorkItems(query, "priority", ["type", "status"])

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
                query, "priority", ["id", "title", "status", "type", "linkedWorkItems", "linkedWorkItemsDerived"])
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
    worksheet.write("D1", 'Status')
    worksheet.write("E1", 'Linked Work Items')
    worksheet.write("F1", 'LWI:Type')
    worksheet.write("G1", 'LWI:Status')

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


def Polarion_Plan_Runnable(infoList):
    # Print Script Version.
    #print_script_version()
    # Get user input data.
    #infoList = Get_Input_Data()
    #print(infoList)
    # Create / Open Excel file.
    workbook, worksheet = excel_open()
    # Get polarion data.
    workitems_list, polarion_object = polarion_query(infoList)
    # Analize the data then write it in Excel file.
    data_analysis(workitems_list, workbook, worksheet, infoList,polarion_object)
    # Disconnet polarion.
    polarion_object.disconnect()

#infoList=['melmohta','Aliya$1994','optimus','B03_00']
#Polarion_Plan_Runnable(infoList)