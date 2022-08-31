import connectors.polarion_connector as c
from openpyxl import Workbook
import os
import re
import http.client
import requests
from xlsxwriter.workbook import Workbook
from datetime import datetime
#from GUI import Task
from openpyxl.chart import BarChart, Reference, Series

################### Global varible


SF_SWC={}

SF_Status={}
SystemFunction_SWC_Report_Warning_Message ={
    "Message" : ""
}
def get_input_data():
    info_list = []

    # Get data from Input_Data_File.txt
    with open("Input_Data_File.txt") as InfoFile:
        for line in InfoFile:
            info_list.append(line[(line.find(':') + 1):].rstrip('\n'))

    i = 0
    for element in info_list:
        info_list[i] = element.replace(" ", "")
        i += 1
    #print(info_list)
    return info_list


def create_output_directory():
    output_path = os.getcwd() + "\Outputs"

    # Check existence of Output folder and create it if not.
    if not os.path.exists(output_path):
        os.makedirs(output_path)
        ##print(output_path + ' : created')

    return output_path


def get_folder_data(info_list):
    #print("Connecting to Polarion...")
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
            # #print connection error message
            #print(str(e))
            #print("Retry...")
            pass
        except Exception as e:
            # #print connection error message
            #print(str(e))
            #print("Retry...")
            pass
        except http.client.HTTPException as e:
            # #print connection error message
            #print(str(e))
            #print("Retry...")
            pass
        except http.client.RemoteDisconnected as e:
            # #print connection error message
            #print(str(e))
            #print("Retry...")
            pass
        except requests.exceptions.ConnectionError as e:
            # #print connection error message
            #print(str(e))
            #print("Retry...")
            pass
        except ConnectionError as e:
            # #print connection error message
            #print(str(e))
            #print("Retry...")
            pass
        except:
            #print("Retry...")
            pass

    ##print("Folder content downloaded")

    # Disconnect Polarion.
    polarion_object.disconnect()
    #print("I am here now ")
    return folder_content


def SF_get_work_items_ids(folder_content,info_list,output_workbook):
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
            # #print(home_page_content[elements.end():elements.end() + 16])
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


def get_work_items_Details_data(docs_content,info_list,SF_Array,output_workbook):
    docs_content_detail = {}
    docs_content_detail2 = {}
    #print("Connecting to Polarion...")
    # Connect using the username and password
    polarion_object = c.Polarion("https://vseapolarion.vnet.valeo.com/polarion")
    polarion_object.connect(str(info_list[0]), str(info_list[1]))
    Document_SF ={}
    Doc_Title =""
    Doc_Status=""
    index=1
    # Get workitem details for each document.
    for elements in SF_Array:
        ##print("Name :", elements["Name"])
        ##print("Status :", elements["Status"])
        ##print("IDs :", elements["IDs"])
        ids_list = elements["IDs"]
        ##print("Doc element :",docs_content)
        #Status = docs_content[element.status]
        SW_COmp =[]
        #docs_content_detail2[element]
        not_connected = True

        # Connect to Polarion database
        while not_connected:
            try:
                ##print("trying Except ")
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
                # #print connection error message
                ##print("Error Except 1")
                ##print(str(e))
                #print("Retry...")
                pass
            except Exception as e:
                # #print connection error message
                ##print("Error Except 2")
                ##print(str(e))
                #print("Retry...")
                pass
            except BaseException as e:
                ##print("Error Except 3")
                # #print connection error message
                ##print(str(e))
                #print("Retry...")
                pass
            except http.client.HTTPException as e:
                ##print("Error Except 4")
                # #print connection error message
                ##print(str(e))
                #print("Retry...")
                pass
            except http.client.RemoteDisconnected as e:
                # #print connection error message
                ##print("Error Except 5")
                ##print(str(e))
                #print("Retry...")
                pass
            except requests.exceptions.ConnectionError as e:
                # #print connection error message
                ##print("Error Except 6")
                ##print(str(e))
                #print("Retry...")
                pass
            except:
                #print("Retry...")
                pass

        ID_Str =""
        for element in work_item_list:
            for sub_element in element:
                if sub_element.type.id == "swComponent":
                    ID = sub_element.id + " " + sub_element.title + " \n"
                    ID_Str = ID_Str + str(ID)
                    SW_COmp.append(ID)
                    ##print("ID :",ID)
        elements["IDs"] = SW_COmp
        elements["ID_Str"] = ID_Str
        ID_Str = ""

        #Task["Name"] = "Document" + str(element)+ "' data downloaded successfully"
        ##print("Document '" + element + "' data downloaded successfully")

    # Create an new Excel file and add a worksheet.
    #workbook =
    #print("Pass here ")
    workbook = Workbook(output_workbook)

    worksheet4 = workbook.add_worksheet('Execution Details')
    worksheet2 = workbook.add_worksheet('KPI')
    worksheet = workbook.add_worksheet('SF report')
    worksheet3 = workbook.add_worksheet('SWC VS SF')
    worksheet.autofilter('A1:D5000')
    worksheet3.autofilter('A1:D5000')
    # Widen the first column to make the text clearer.
    worksheet.set_column('C:C', 20)
    # Add a cell format with text wrap on.
    cell_format = workbook.add_format({'text_wrap': True})
    worksheet.write('A1', "System function")
    worksheet.write('B1', "Status")
    worksheet.write('C1', "SWC")
    indix =2

    Released = 0
    non_released = 0
    Total_Docs = len(SF_Array)
    for elements in SF_Array:
        cloumA = 'A' + str(indix)
        cloumB = 'B' + str(indix)
        cloumC = 'C' + str(indix)
        indix = indix +1
        # Write a wrapped string to a cell.
        ##print(cloumA,cloumB)
        Document_Link =link_generate(info_list,elements["Name"], "document")
        worksheet.write(cloumA,Document_Link, cell_format)
        worksheet.write(cloumB,elements["Status"], cell_format)
        worksheet.write(cloumC, elements["ID_Str"], cell_format)
        if elements["Status"] == "released":
            Released = Released +1
        else:
            non_released = non_released + 1
        ##print("Name :", elements["Name"])
        ##print("Status :", elements["Status"])
        ##print("IDs :", elements["IDs"])
    #print("Pass here 1")
    cell_format = workbook.add_format()
    cell_format.set_bold()
    #print("Pass here 2")
    worksheet2.write('A1',"Total SF",cell_format)
    worksheet2.write('A2', "Released SF",cell_format)
    worksheet2.write('A3', "Non Released SF",cell_format)
    worksheet2.write('B1', Total_Docs)
    worksheet2.write('B2', Released)
    worksheet2.write('B3', non_released)
    #print("Pass here 3")
    '''
    now = datetime.now()  # current date and time
    Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
    worksheet4.write('A1', "Exection Date",cell_format)
    worksheet4.write('B1', Date_Data)
    worksheet4.write('A2', "Generated By ",cell_format)
    worksheet4.write('B2', str(info_list[0]))
    worksheet4.write('A3', "SWA Baseline",cell_format)
    worksheet4.write('B3', str(info_list[3]))
    '''
    #print("Pass here 4")
    # Disconnect Polarion.
    polarion_object.disconnect()
    #workbook.close()
    #print("Pass here 5")
    return docs_content_detail,SF_Array,workbook,worksheet3

def Create_Report(output_path,infoList):
    File_Name = "\System_Function_Report.xlsx"
    output_workbook = output_path + File_Name

    return output_workbook



def link_generate(info_list, id, link_type):
    if link_type == "workitem":
        link = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + info_list[2] + "/workitem?id=" + id
        HyperLink = "=HYPERLINK(\"" + link + "\",\"" + id + "\")"
    elif link_type == "testrun":
        link = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + info_list[2] + "/testrun?id=" + id
        HyperLink = "=HYPERLINK(\"" + link + "\",\"" + id + "\")"
    elif link_type == "document":
        link = "https://vseapolarion.vnet.valeo.com/polarion/#/project/" + info_list[2] + "/wiki/" + id
        HyperLink= "=HYPERLINK(\"" + link + "\",\""+id+"\")"

    return HyperLink


def get_kpi(kpi, type, status):
    for element in kpi:
        if element == type:
            kpi[type][0]["total"] +=1
            if status == "released":
                kpi[type][0]["released"] += 1
            else:
                kpi[type][0]["not released"] += 1

    return kpi


def prepare_SF_array(folder_content,docs_content):
    SF_Array = []
    SF_Dict = {
        "Name": "",
        "Status": "",
        "IDs": [],
    }
    for element in folder_content:
        SF_Dict = {
            "Name": "",
            "Status": "",
            "IDs": [],
            "ID_Str": "",
        }
        SF_Dict["Name"] = str(element.title)
        SF_Dict["Status"] = str(element.status.id)
        SF_Array.append(SF_Dict)

    for pair in docs_content:
        for element in SF_Array:
            if pair == element["Name"]:
                element["IDs"] = docs_content[pair]
        else:
            pass
    return SF_Array

def Separaate_SWC_SF(SF_Array):
    SWC_ID_Array=[]
    SWC_SF_Array_Mapp =[]

    for elements in SF_Array:
        for element in elements["IDs"]:
            ##print("Element : ", element)
            ##print("array of Element : ", SWC_ID_Array)
            if element in SWC_ID_Array:
                pass
            else:
                SWC_ID_Array.append(element)

    for elements in SWC_ID_Array:
        SF_SWC = {
            "ID_Name": "",
            "SWC_ID_Mapp": [],
            "SWC_SF_Mapp": "",
        }
        # elements => SWC IDs
        for element in SF_Array:
            #element => Dictionary of name of SF and SWC IDS in it
            for sub_elements in element["IDs"]:
                if sub_elements == elements:
                    if element in SF_SWC["SWC_ID_Mapp"]:
                        pass
                    else:
                        SF_SWC["SWC_ID_Mapp"].append(element["Name"])
                        SF_SWC["SWC_SF_Mapp"] = SF_SWC["SWC_SF_Mapp"] + element["Name"] + "\n"

            else :
                pass
        SF_SWC["ID_Name"] = elements
        SWC_SF_Array_Mapp.append(SF_SWC)
    return SWC_SF_Array_Mapp

def Write_SWC_VS_SF_Report(SWC_SF_Array_Mapp,workbook,worksheet3):
    worksheet3.set_column('B:B', 20)
    # Add a cell format with text wrap on.
    cell_format = workbook.add_format({'text_wrap': True})
    worksheet3.write('A1', "Software Component")
    worksheet3.write('B1', "System function")
    indix = 2
    Strenth = len(SWC_SF_Array_Mapp)
    for i in range(Strenth):
        ##print("Name:" , SWC_SF_Array_Mapp["ID_Name"])
        ##print("SF_ Name:",SWC_SF_Array_Mapp["SWC_SF_Mapp"] )
        SWC_ID = SWC_SF_Array_Mapp[i]["ID_Name"]
        System_Functions =SWC_SF_Array_Mapp[i]["SWC_SF_Mapp"]
        #for subelement in SWC_SF_Array_Mapp["SWC_ID_Mapp"]:
        #    #print("===",subelement)
        #    SF = str(subelement)+ " \n"
        #    System_Functions = System_Functions +SF
        cloumA = 'A' + str(indix)
        cloumB = 'B' + str(indix)
        indix = indix + 1
        worksheet3.write(cloumA, SWC_ID, cell_format)
        worksheet3.write(cloumB, System_Functions, cell_format)
    while True:
        try:
            workbook.close()
            #os.startfile(output_workbook)
            break
        except:
            SystemFunction_SWC_Report_Warning_Message["Message"] = "File Is already opened ! , please close Excel file "
            pass



def SWC_SF_Report_Generation_Runnable(info_list):
    #info_list = get_input_data()
    output_path = create_output_directory()
    folder_content = get_folder_data(info_list)
    output_workbook= Create_Report(output_path,info_list)
    docs_content = SF_get_work_items_ids(folder_content,info_list,output_workbook)
    SF_Array= prepare_SF_array(folder_content,docs_content)
    docs_content_detail , SF_Array ,workbook,worksheet3 = get_work_items_Details_data(docs_content,info_list,SF_Array,output_workbook)
    SWC_SF_Array_Mapp= Separaate_SWC_SF(SF_Array )
    Write_SWC_VS_SF_Report(SWC_SF_Array_Mapp,workbook,worksheet3)

