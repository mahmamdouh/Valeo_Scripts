import script.handlers.connectors.polarion_connector as c
import pandas as pd
import xlsxwriter
import re


dataList = []
dataList.append('')
dataList.append('')
dataList.append('')

def description_extraction(description, counter, worksheet):
    table = pd.read_html(description, header=0)
    pass

    indexList = description.split("<br/>\n")
    dataList[2] = indexList
    data_write(dataList, counter, worksheet)

    # for elements in re.finditer("<br/>", description):
    #     indexList.append(elements.start())
    #     indexList.append(elements.end())
    # print(indexList)
    # print(".................")
    # i = 2
    # print(len(indexList))
    # try:
    #     if indexList[0] != 0:
    #         print(description[:indexList[0]].replace('\n', ''))
    # except:
    #     pass
    # while i <= len(indexList)-1:
    #     print(i)
    #     print(indexList[i])
    #     print(description[indexList[i-1]:indexList[i]].replace('\n', ''))
    #     i = i + 2
    #     #if elements[0].start() != 0:
    #     #    print(description[:elements[0].start()])
    #     #print(description[elements.start()])
    #
    # indexList = []
    # indexList_1 = []
    # for elements in re.finditer("<li>", description):
    #     indexList.append(elements.end())
    # print(indexList)
    # for elements in re.finditer("</li>", description):
    #     indexList_1.append(elements.start())
    # print(indexList_1)
    # print(".................")
    # i = 0
    # print(len(indexList), " ", len(indexList_1))
    # while i <= len(indexList)-1:
    #     print(i)
    #     print(indexList[i])
    #     print(indexList_1[i])
    #     print(description[indexList[i]:indexList_1[i]].replace('\n', ''))
    #     i = i + 1


def data_write(dataList, counter, worksheet):
    # Identify data index.
    a = 'A' + str(counter)
    b = 'B' + str(counter)
    c = 'C' + str(counter)

    cData = str(dataList[2]).replace("'", "").replace("[", "").replace("]", "")

    # Data write.
    worksheet.write(a, dataList[0])
    worksheet.write(b, dataList[1])
    worksheet.write(c, cData)


def get_workItems_data(idsList, worksheet):
    counter = 1

    for id in idsList:
        workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
            "id:" + id, "priority", ["id", "title", "status", "type", "description", "linkedWorkItems"])
        # print(workitems_list[0].id)
        # print(workitems_list[0].title)
        # print(workitems_list[0].status.id)
        # print(workitems_list[0].type.id)
        # print(workitems_list[0].description)

        if workitems_list[0].type.id == 'information':
            dataList[1] = workitems_list[0].title
            for elements in workitems_list[0].linkedWorkItems.LinkedWorkItem:
                if elements.role.id == 'parent':
                    LWIid = str(elements.workItemURI).split("}")[-1]
                    #print(LWIid)
                    workitem_title(LWIid)
            if workitems_list[0].description != None:
                description_extraction(workitems_list[0].description.content, counter, worksheet)
            counter = counter + 1


def workitem_title(id):
    # New query based on workitems ID to get the title.
    query = str("id:" + id)

    title = polarion_object.tracker_webservice.service.queryWorkItems(query, "priority", ["title"])

    dataList[0] = title[0].title


def data_analysis(homePageContent, infoList, worksheet):
    res = []
    ids_list = []

    # Search for IDs in the page data.
    for elements in re.finditer("params=id=", homePageContent):
        # print(home_page_content[elements.end():elements.end() + 16])
        if infoList[2] == "optimus":
            res.append(homePageContent[elements.end():elements.end() + 11])
        elif infoList[2] == "model_kit":
            res.append(homePageContent[elements.end():elements.end() + 15])
        elif infoList[2] == "VW_MEB_Inverter":
            res.append(homePageContent[elements.end():elements.end() + 16])
        elif infoList[2] == "VW_MEB_Inverter_Base_Minus":
            res.append(homePageContent[elements.end():elements.end() + 23])

    # Filter duplicated IDs.
    for i in range(len(res)):
        # Check if the string ends with digits (number) and return None if not.
        last_character = re.search(r'\d+$', res[i])
        if last_character is None:
            res[i] = res[i][:-1]
        if res[i] not in ids_list:
            ids_list.append(res[i])

    get_workItems_data(ids_list, worksheet)


def polarion_query(infoList):
    print("Connecting to Polarion...")
    NOTConnected = True

    while NOTConnected:
        try:
            # Connect to polarion with credentials.
            polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
            polarion_object.connect(infoList[0], infoList[1])
            # Get document data.
            homepage = polarion_object.tracker_webservice.service.getModuleByLocation(infoList[2], infoList[6])
            homePageContent = homepage.homePageContent.content
            NOTConnected = False
        except Exception as e:
            print(str(e))
            pass

    return homePageContent, polarion_object


def excel_open():
    # Create / Open Excel file.
    workbook = xlsxwriter.Workbook('Outputs\Output_1.xlsx')
    worksheet1 = workbook.add_worksheet('CR Scope')
    worksheet2 = workbook.add_worksheet('Defect Scope')
    worksheet3 = workbook.add_worksheet('BSW scope')
    worksheet4 = workbook.add_worksheet('SSW scope')
    worksheet5 = workbook.add_worksheet('LLSW scope')
    worksheet6 = workbook.add_worksheet('DRCO scope')

    worksheet = [worksheet1, worksheet2, worksheet3, worksheet4, worksheet5, worksheet6]

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

    # Remove the empty space at the beginning of each element.
    i = 0
    for element in infoList:
        try:
            if element[0] == " ":
                element = element.replace(" ", "")
                infoList[i] = element
        except:
            pass
        i += 1

    # Rename project based on the project name in the database.
    if infoList[2] == "100kW" or infoList[2] == "100kw" or infoList[2] == "100KW" or infoList[2] == "100Kw":
        infoList[2] = "optimus"
    elif infoList[2] == "model kit" or infoList[2] == "Model Kit" or infoList[2] == "Model kit" or \
            infoList[2] == "model Kit" or infoList[2] == "modelkit" or infoList[2] == "ModelKit" or \
            infoList[2] == "Modelkit" or infoList[2] == "modelKit":
        infoList[2] = "model_kit"
    elif "VW_MEB" in infoList[2] and "Minus" in infoList[2]:
        infoList[2] = "VW_MEB_Inverter_Base_Minus"
    elif "VW_MEB" in infoList[2]:
        infoList[2] = "VW_MEB_Inverter"

    # Edit arch document name.
    infoList[6] = infoList[6].replace("%20", " ")

    return infoList


def print_script_version():
    # Script Version Number.
    script_version = '1.0.1'
    script_build_number = '(2)'

    print('Version', script_version, script_build_number)


if __name__ == '__main__':


    polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion/")
    polarion_object.connect('asolima2', 'Asolima2228')

    workitems_list = polarion_object.tracker_webservice.service.queryWorkItems(
        "id:" + "100kW-58022", "priority", ["id", "title", "status", "type", "description", "linkedWorkItems"])

    indexList = []
    for elements in re.finditer("<br/>", workitems_list[0].description.content):
        indexList.append(elements.start())
        indexList.append(elements.end())
        print(indexList)

    # Print Script Version.
    print_script_version()
    # Get user input data.
    infoList = Get_Input_Data()
    # Create / Open Excel file.
    workbook, worksheet = excel_open()
    # Get polarion data.
    homePageContent, polarion_object = polarion_query(infoList)
    # Analyze the data then write it in Excel file.
    data_analysis(homePageContent, infoList, worksheet)
    # Save Excel file.
    workbook.close()
    # Disconnect polarion.
    polarion_object.disconnect()
    print("Successfully downloaded")
