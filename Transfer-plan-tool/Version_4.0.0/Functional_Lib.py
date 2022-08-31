import xlsxwriter
from Get_Stored_Data_Pandas import Get_Polarion_Plan_Data,Get_Integration_Report_Data,Get_SWC_System_Function_Data
import re
import difflib
from datetime import datetime
def Prepare_Shell_Script(Input_Data):
    #print("Start of generating file ...")
    #print("Input Data :",Input_Data)
    file1 = open("Get_Repo.sh", "w")
    local_repo = Input_Data["L_Repo"].replace("/", "\\\\")
    L = ["cd " + local_repo, "\ngit logreview --name-only --decorate  " + Input_Data[
        "RB"] + "..HEAD  >"+Input_Data["Directory"]+"\\\\SubFiles\\\\log.txt"]
    file1.writelines(L)
    file1.close()
#"\ngit pull --rebase",


def Create_Out_Report():
    #print("Create")
    File_Name = "SubFiles\Integration_Report.xlsx"
    output_workbook = File_Name
    workbook1 = xlsxwriter.Workbook(output_workbook)

    # Create Excel document

    Sheet1 = workbook1.add_worksheet('Sheet1')


    format1 = workbook1.add_format({'bg_color': '#c7f1ff',
                                   'font_color': '#9C0006'})
    format1.set_bold()

    Sheet1.write("A1", "commit", format1)
    Sheet1.write("B1", "Author", format1)
    Sheet1.write("C1", "Change ID", format1)
    Sheet1.write("D1", "Tag", format1)
    Sheet1.write("E1", "SWCs", format1)
    Sheet1.write("F1", "Functional affected SWCS", format1)
    Sheet1.write("G1", "Affected System Function", format1)
    #print("Create 1")
    while True:
        try:
            workbook1.close()
            #os.startfile(output_workbook)
            break
        except:
            print("File Is already opened ! , please close Excel file ")
            pass

    #print("Create 2")
    return "SubFiles\Integration_Report.xlsx"

def Get_Polarion_Plan_IDs(File):
    Polarion_Origin_Plan =[]
    Polarion_Chield_Plan = []
    Polarion_Data = Get_Polarion_Plan_Data(File)
    #print(Polarion_Data["ID"])
    for element in Polarion_Data["ID"]:
        if not (Polarion_Data["ID"][element] == 0):
            #print(Polarion_Data["ID"][element])
            Polarion_Origin_Plan.append(Polarion_Data["ID"][element])

    for element in Polarion_Data["Linked Work Items"]:
        if not (Polarion_Data["Linked Work Items"][element] == 0):
            #print("LWIS ",Polarion_Data["Linked Work Items"][element])
            Polarion_Chield_Plan.append(Polarion_Data["Linked Work Items"][element])
    return Polarion_Origin_Plan,Polarion_Chield_Plan


def Compare_SWC_SystemFunction_List(Integrated_SWC , SWC_SF_List,Polarion_Origin_Plan , Polarion_Chield_Plan):
    #commit	Author	Change ID	Tag	SWCs	Functional affected SWCS	Affected System Function Polarion Plan Verification
    Out_Put_Report_Array=[]
    for element in Integrated_SWC['commit']:
        Out_Report_Data = {
            "commit": "",
            "Author": "",
            "Change ID": "",
            "Gerrit Ticket": "",
            "SWCs": "",
            "Functional affected SWCS": "",
            "Affected System Function": "",
            "Polarion Plan Verification": "",
        }
        if Integrated_SWC['Functional affected SWCS'][element] ==0:
            #Fill Data To print
            Out_Report_Data["commit"]=Integrated_SWC['commit'][element]
            Out_Report_Data["Author"]=Integrated_SWC['Author'][element]
            Out_Report_Data["Change ID"]=Integrated_SWC['Change ID'][element]
            Out_Report_Data["Gerrit Ticket"]=Integrated_SWC['Tag'][element]
            Out_Report_Data["SWCs"]=Integrated_SWC['SWCs'][element]
            Out_Report_Data["Functional affected SWCS"]=Integrated_SWC['Functional affected SWCS'][element]
            Plan_Verification =Compare_Polarion_Plan(Polarion_Origin_Plan, Polarion_Chield_Plan, Integrated_SWC['Tag'][element])
            Out_Report_Data["Polarion Plan Verification"] = Plan_Verification
            #print(Out_Report_Data)
        else:
            #Fill Data
            Out_Report_Data["commit"] = Integrated_SWC['commit'][element]
            Out_Report_Data["Author"] = Integrated_SWC['Author'][element]
            Out_Report_Data["Change ID"] = Integrated_SWC['Change ID'][element]
            Out_Report_Data["Gerrit Ticket"] = Integrated_SWC['Tag'][element]
            Out_Report_Data["SWCs"] = Integrated_SWC['SWCs'][element]
            Out_Report_Data["Functional affected SWCS"] = Integrated_SWC['Functional affected SWCS'][element]
            Plan_Verification = Compare_Polarion_Plan(Polarion_Origin_Plan, Polarion_Chield_Plan,
                                                      Integrated_SWC['Tag'][element])
            Out_Report_Data["Polarion Plan Verification"] = Plan_Verification
            #Do Compare
            Integrated_SWC_List =Integrated_SWC['Functional affected SWCS'][element].split(",")
            for SWC_Element in Integrated_SWC_List:
                Affected_SF_List =[]
                for elements in SWC_SF_List['Software Component']:
                    SWC_New_List = []
                    SWC_New_List.append(SWC_SF_List['Software Component'][elements].split())
                    for SWC_New_List_Elements in SWC_New_List :
                        for Sub_Elements in SWC_New_List_Elements:
                            temp = difflib.SequenceMatcher(None, SWC_Element, Sub_Elements)
                            if (temp.ratio() > 0.9 ):
                                #print('Compare Between : ',SWC_Element, Sub_Elements)
                                #print('Similarity Score: ', temp.ratio())
                                #print("SWC is :",SWC_Element )
                                #print("System Function Is:", SWC_SF_List['System function'][elements])
                                Affected_SF_List.append(SWC_SF_List['System function'][elements])
                    Affected_SF_List =remove_Duplicates (Affected_SF_List)
                    Out_Report_Data["Affected System Function"] = Affected_SF_List

        Out_Put_Report_Array.append(Out_Report_Data)
    return Out_Put_Report_Array
            #print(Integrated_SWC['Functional affected SWCS'][element])

        #System function
    #temp = difflib.SequenceMatcher(None, string1, string2)

    #print(temp.get_matching_blocks())
    #print('Similarity Score: ', temp.ratio())


def remove_Duplicates(Array_L):
    res = []
    for i in Array_L:
        if i not in res:
            res.append(i)
    return res

def Compare_Polarion_Plan(Polarion_Origin_Plan , Polarion_Chield_Plan , Ticket):
    rtn_Val = ""
    if Ticket in Polarion_Origin_Plan :
        rtn_Val= "✓"
    elif Ticket in Polarion_Chield_Plan :
        rtn_Val= "✓"
    else :
        rtn_Val= "Not Planned"
    return rtn_Val




def Generate_Integration_Report_Report(info_list,Out_Report_Data):
    # Create / Open Excel file.
    File_Name = "Outputs\Integration_Final_Report.xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet = workbook1.add_worksheet("Intagration_Report")
    worksheet2 = workbook1.add_worksheet("Execution Details")



    worksheet.autofilter('A1:D50000')

    cell_format = workbook1.add_format()
    cell_format.set_bold()

    cell_format2 = workbook1.add_format({'text_wrap': True})

    # Headers.
    # commit	Author	Change ID	Tag	SWCs	Functional affected SWCS	Affected System Function Polarion Plan Verification
    format1 = workbook1.add_format({'bg_color': '#124173',
                                    'font_color': '#ffffff'})
    format1.set_bold()

    Failed_format = workbook1.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})
    Pass_format = workbook1.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100'})

    worksheet.write("A1", 'commit', format1)
    worksheet.write("B1", 'Author', format1)
    worksheet.write("C1", 'Change ID', format1)
    worksheet.write("D1", 'Gerrit Ticket', format1)
    worksheet.write("E1", 'SWCs', format1)
    worksheet.write("F1", 'Functional affected SWCS', format1)
    worksheet.write("G1", 'Affected System Function', format1)
    worksheet.write("H1", 'Polarion Plan Verification', format1)

    now = datetime.now()  # current date and time
    Date_Data = now.strftime("%m/%d/%Y, %H:%M:%S")
    # print("Step 2")
    worksheet2.write('A1', "Exection Date", cell_format)
    worksheet2.write('B1', Date_Data)
    worksheet2.write('A2', "Generated By ", cell_format)
    worksheet2.write('B2', str(info_list[0]))
    worksheet2.write('A3', "SWA Baseline ", cell_format)
    worksheet2.write('B3', str(info_list[4]))
    worksheet2.write('A4', "Polarion Plan", cell_format)
    worksheet2.write('B4', str(info_list[3]))
    worksheet2.write('A5', "Starting RB", cell_format)
    worksheet2.write('B5', str(info_list[5]))

    indix =2
    for Elements in Out_Report_Data:
        aa = 'A' + str(indix)
        bb = 'B' + str(indix)
        cc = 'C' + str(indix)
        dd = 'D' + str(indix)
        ee = 'E' + str(indix)
        ff = 'F' + str(indix)
        gg = 'G' + str(indix)
        hh = 'H' + str(indix)
        worksheet.write(aa, Elements['commit'])
        worksheet.write(bb, Elements['Author'])
        worksheet.write(cc, Elements['Change ID'])
        worksheet.write(dd, Elements['Gerrit Ticket'])
        worksheet.write(ee, Elements['SWCs'])
        worksheet.write(ff, Elements['Functional affected SWCS'])
        worksheet.write(gg, change_List_To_string(Elements['Affected System Function']) ,cell_format2)
        if Elements['Polarion Plan Verification'] == "Not Planned":
            worksheet.write(hh, Elements['Polarion Plan Verification'], Failed_format)
        elif Elements['Polarion Plan Verification'] == "✓":
            worksheet.write(hh, Elements['Polarion Plan Verification'], Pass_format)
        else :
            worksheet.write(hh, Elements['Polarion Plan Verification'])

        indix = indix +1
    while True:
        try:
            workbook1.close()
            #os.startfile(output_workbook)
            break
        except:
            print ("File 'Integration_Final_Report' Is already opened ! , please close Excel file ")
            pass

def change_List_To_string(List):
    List_Str =""
    for elements in List:
        List_Str = List_Str + " " + elements
    return List_Str

def Report_Generation_Verification(infoList):
    Polarion_Origin_Plan,Polarion_Chield_Plan=Get_Polarion_Plan_IDs("Outputs\Polarion_Plan.xlsx")
    Integrated_Report, Gerrit_Tickets = Get_Integration_Report_Data("SubFiles\Integration_Report.xlsx")
    SWC_SF_List, SWC_List = Get_SWC_System_Function_Data("Outputs\System_Function_Report.xlsx")
    Out_Put_Report_Array =Compare_SWC_SystemFunction_List(Integrated_Report, SWC_SF_List,Polarion_Origin_Plan, Polarion_Chield_Plan)
    infoList = ['melmohta', 'Aliya$1994', 'optimus', 'B03_00','Mahmoud' , "E;mohtady"]
    Generate_Integration_Report_Report(infoList,Out_Put_Report_Array)