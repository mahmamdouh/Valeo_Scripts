import xlsxwriter
from Get_Stored_Data_Pandas import Get_Polarion_Plan_Data,Get_Integration_Report_Data,Get_SWC_System_Function_Data
import re
import difflib
from datetime import datetime
import os
from RTE_API import Task , Verification_Plan_Flag ,Task_Flag
from Gerrit_Changes_Analysis import Gerrit_Changes_Analysis_Runnable
from PolarionPlanAPIs import Polarion_Plan_Runnable
import time
from SystemFunction_SWC_Report import SWC_SF_Report_Generation_Runnable
import os
import os, shutil
from SystemFunction_SWC_Report import create_output_directory,Create_Report
#from xlsxwriter.workbook import Workbook
#from openpyxl import Workbook


def Delete_Directory_File(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))



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
    Sheet1.write("H1", "Date", format1)
    Sheet1.write("I1", "Rolling Build", format1)
    Sheet1.write("J1", "Commit message", format1)
    Sheet1.write("K1", "File List", format1)
    Sheet1.write("L1", "ParFiles", format1)
    Sheet1.write("M1", "A2lFiles", format1)
    Sheet1.write("N1", "SourceFiles", format1)
    Sheet1.write("O1", "HeaderFiles", format1)
    Sheet1.write("P1", "RTRT", format1)






    #print("Create 1")
    while True:
        try:
            workbook1.close()
            Task["Warning_Message"] = ""
            #os.startfile(output_workbook)
            break
        except:
            print("File Is already opened ! , please close Excel file ")
            Task["Warning_Message"] = "[Integration_Report.xlsx] file Is already opened ! , please close Excel file "
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


def Compare_SWC_SystemFunction_List(Integrated_SWC , SWC_SF_List ):
    #commit	Author	Change ID	Tag	SWCs	Functional affected SWCS	Affected System Function Polarion Plan Verification
    Out_Put_Report_Array=[]
    for element in Integrated_SWC['commit']:
        #print("Parsing result from report :" , element)
        Out_Report_Data = {
            "commit": "",
            "Author": "",
            "Change ID": "",
            "Polarion Ticket": "",
            "SWCs": "",
            "Functional affected SWCS": "",
            "Affected System Function": "",
            "Date": "",
            "Rolling Build": "",
            "Commit Message": "",
            "file List": "",
            "ParFiles": "",
            'A2lFiles': "",
            "SourceFiles": "",
            "HeaderFiles": "",
            "RTRT": "",
            "Polarion Plan Verification": "",
        }
        if Integrated_SWC['Functional affected SWCS'][element] ==0:
            #Fill Data To print
            Out_Report_Data["commit"]=Integrated_SWC['commit'][element]
            Out_Report_Data["Author"]=Integrated_SWC['Author'][element]
            Out_Report_Data["Change ID"]=Integrated_SWC['Change ID'][element]
            Out_Report_Data["Polarion Ticket"]=Integrated_SWC['Tag'][element]
            Out_Report_Data["SWCs"]=Integrated_SWC['SWCs'][element]
            Out_Report_Data["Functional affected SWCS"]=Integrated_SWC['Functional affected SWCS'][element]
            Out_Report_Data["Date"] = Integrated_SWC['Date'][element]
            Out_Report_Data["Rolling Build"] = Integrated_SWC['Rolling Build'][element]
            Out_Report_Data["Commit Message"] = Integrated_SWC['Commit message'][element]
            Out_Report_Data["file List"] = Integrated_SWC['File List'][element]
            Out_Report_Data["ParFiles"] = Integrated_SWC['ParFiles'][element]
            Out_Report_Data["A2lFiles"] = Integrated_SWC['A2lFiles'][element]
            Out_Report_Data["SourceFiles"] = Integrated_SWC['SourceFiles'][element]
            Out_Report_Data["HeaderFiles"] = Integrated_SWC['HeaderFiles'][element]
            Out_Report_Data["RTRT"] = Integrated_SWC['RTRT'][element]
            #Plan_Verification =Compare_Polarion_Plan(Polarion_Origin_Plan, Polarion_Chield_Plan, Integrated_SWC['Tag'][element])
            #Out_Report_Data["Polarion Plan Verification"] = Plan_Verification
            #print(Out_Report_Data)
        else:
            #Fill Data
            Out_Report_Data["commit"] = Integrated_SWC['commit'][element]
            Out_Report_Data["Author"] = Integrated_SWC['Author'][element]
            Out_Report_Data["Change ID"] = Integrated_SWC['Change ID'][element]
            Out_Report_Data["Polarion Ticket"] = Integrated_SWC['Tag'][element]
            Out_Report_Data["SWCs"] = Integrated_SWC['SWCs'][element]
            Out_Report_Data["Date"] = Integrated_SWC['Date'][element]
            Out_Report_Data["Rolling Build"] = Integrated_SWC['Rolling Build'][element]
            Out_Report_Data["Commit Message"] = Integrated_SWC['Commit message'][element]
            Out_Report_Data["file List"] = Integrated_SWC['File List'][element]
            Out_Report_Data["ParFiles"] = Integrated_SWC['ParFiles'][element]
            Out_Report_Data["A2lFiles"] = Integrated_SWC['A2lFiles'][element]
            Out_Report_Data["SourceFiles"] = Integrated_SWC['SourceFiles'][element]
            Out_Report_Data["HeaderFiles"] = Integrated_SWC['HeaderFiles'][element]
            Out_Report_Data["RTRT"] = Integrated_SWC['RTRT'][element]
            Out_Report_Data["Functional affected SWCS"] = Integrated_SWC['Functional affected SWCS'][element]
            #Plan_Verification = Compare_Polarion_Plan(Polarion_Origin_Plan, Polarion_Chield_Plan,
            #                                          Integrated_SWC['Tag'][element])
            #Out_Report_Data["Polarion Plan Verification"] = Plan_Verification
            #Do Compare
            Integrated_SWC_List =Integrated_SWC['Functional affected SWCS'][element].split(",")
            for SWC_Element in Integrated_SWC_List:
                Affected_SF_List =[]
                try :
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
                except:
                    Out_Report_Data["Affected System Function"] = ""
                    pass

        Out_Put_Report_Array.append(Out_Report_Data)
    return Out_Put_Report_Array
            #print(Integrated_SWC['Functional affected SWCS'][element])

        #System function
    #temp = difflib.SequenceMatcher(None, string1, string2)

    #print(temp.get_matching_blocks())
    #print('Similarity Score: ', temp.ratio())

def Create_Temp_OutPut():
    #print("Create")
    File_Name = "Outputs\System_Function_Report.xlsx"
    output_workbook = File_Name
    workbook1 = xlsxwriter.Workbook(output_workbook)

    while True:
        try:
            workbook1.close()
            Task["Warning_Message"] = ""
            #os.startfile(output_workbook)
            break
        except:
            print("File Is already opened ! , please close Excel file ")
            Task["Warning_Message"] = "[System_Function_Report.xlsx] file Is already opened ! , please close Excel file "
            pass

    #print("Create 2")
    return "SubFiles\Integration_Report.xlsx"

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

def Generate_KPI(Out_Report_Data):
    KPI_Data = []
    Num_Of_Commits = 0
    Commits_Have_Functional_Changes = 0
    Commits_Have_RTRT_Changes = 0
    DRCO_Commits = 0
    LLSW_Commits = 0
    BSW_Commits = 0
    SSW_Commits = 0
    Integration_Commits = 0
    Others_Commits = 0
    Orphan_Commits = 0
    # Collect Data
    Num_Of_Commits = len(Out_Report_Data)
    for elements in Out_Report_Data :
        if elements['Functional affected SWCS'] != "0":
            Commits_Have_Functional_Changes = Commits_Have_Functional_Changes +1

        if elements['RTRT'] != "-":
            Commits_Have_RTRT_Changes = Commits_Have_RTRT_Changes +1

        if elements['FGL'] != "":
            LLSW_result = re.search("LLSW", elements['FGL'])
            BSW_result = re.search("BSW", elements['FGL'])
            DRCO_result = re.search("DRCO", elements['FGL'])
            SSW_result = re.search("Integration", elements['FGL'])
            Integration_result = re.search("LLSW", elements['FGL'])
            isw_result = re.search("isw", elements['FGL'])

            if LLSW_result:
                LLSW_Commits = LLSW_Commits +1
            elif BSW_result:
                BSW_Commits = BSW_Commits + 1
            elif DRCO_result:
                DRCO_Commits = DRCO_Commits + 1
            elif SSW_result:
                SSW_Commits = SSW_Commits + 1
            elif Integration_result:
                Integration_Commits = Integration_Commits + 1
            elif isw_result:
                Integration_Commits = Integration_Commits + 1
            else:
                Orphan_Commits = Orphan_Commits+1
        else:
            Others_Commits = Others_Commits +1
    KPI_Data = [Num_Of_Commits,Commits_Have_Functional_Changes,Commits_Have_RTRT_Changes,DRCO_Commits,LLSW_Commits,BSW_Commits,SSW_Commits,Integration_Commits,Others_Commits,Orphan_Commits]

    return KPI_Data

def Write_KPI_Data_To_Sheet(workbook1 ,worksheet0 ):
    #######################################################################
    #
    # Create a new bar chart.
    #
    chart1 = workbook1.add_chart({'type': 'bar'})

    # Configure the first series.
    chart1.add_series({
        'name': '=Summary_KPI!$A$1',
        'categories': '=Summary_KPI!$A$2:$A$11',
        'values': '=Summary_KPI!$B$2:$B$11',
    })

    # Configure a second series. Note use of alternative syntax to define ranges.
    chart1.add_series({
        'name': ['Summary_KPI', 0, 1],
        'categories': ['Summary_KPI', 1, 0, 11, 0],
        'values': ['Summary_KPI', 1, 2, 11, 2],
    })

    # Add a chart title and some axis labels.
    chart1.set_title({'name': 'Summary_KPI'})
    chart1.set_x_axis({'name': 'Numbers'})
    chart1.set_y_axis({'name': 'Attribute'})

    # Set an Excel chart style.
    chart1.set_style(11)

    # Insert the chart into the worksheet (with an offset).
    worksheet0.insert_chart('D2', chart1, {'x_offset': 25, 'y_offset': 10})




    return workbook1 ,worksheet0


def Generate_Integration_Report_Report(info_list,Out_Report_Data,Polarion_IDs_Data,Verification_Plan_Flag):
    # Create / Open Excel file.

    File_Name = "Outputs\Integration_Final_Report.xlsx"
    workbook1 = xlsxwriter.Workbook(File_Name)
    worksheet0 = workbook1.add_worksheet("Summary_KPI")
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
    worksheet.write("D1", 'Polarion Ticket', format1)
    worksheet.write("E1", 'priority', format1)
    worksheet.write("F1", 'severity', format1)
    worksheet.write("G1", 'status', format1)
    worksheet.write("H1", 'resolution', format1)
    worksheet.write("I1", 'Feature Group', format1)
    worksheet.write("J1", 'SWCs', format1)
    worksheet.write("K1", 'Functional affected SWCS', format1)
    worksheet.write("L1", 'Affected System Function', format1)
    worksheet.write("M1", 'Date', format1)
    worksheet.write("N1", 'Rolling Build	', format1)
    worksheet.write("O1", 'Commit Message	', format1)
    worksheet.write("P1", 'file List', format1)
    worksheet.write("Q1", 'ParFiles', format1)
    worksheet.write("R1", 'A2lFiles', format1)
    worksheet.write("S1", 'SourceFiles', format1)
    worksheet.write("T1", 'HeaderFiles', format1)
    worksheet.write("U1", 'RTRT', format1)
    worksheet.write("V1", 'Polarion Plan Verification', format1)


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
    #print("report Passing 2")
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
        ii = 'I' + str(indix)
        jj = 'J' + str(indix)
        kk = 'K' + str(indix)
        ll = 'L' + str(indix)
        mm = 'M' + str(indix)
        nn = 'N' + str(indix)
        oo = 'O' + str(indix)
        pp = 'P' + str(indix)
        qq = 'Q' + str(indix)
        rr = 'R' + str(indix)
        ss = 'S' + str(indix)
        tt = 'T' + str(indix)
        uu = 'U' + str(indix)
        vv = 'V' + str(indix)


        worksheet.write(aa, Elements['commit'])
        worksheet.write(bb, Elements['Author'])
        worksheet.write(cc, Elements['Change ID'])
        worksheet.write(dd, Elements['Polarion Ticket'])
        ##################################
        ID_Check_Flag = 0



        Polarion_ID_Data , ID_Check_Flag = Get_Ticket_Status(Elements['Polarion Ticket'] , Polarion_IDs_Data)
        Elements["FGL"] = Polarion_ID_Data['FGL']
        #print(Elements)
        #print("report Passing 2.2")
        if ID_Check_Flag == 1 :
            worksheet.write(ee, Polarion_ID_Data['priority'])
            worksheet.write(ff, Polarion_ID_Data['severity'])
            worksheet.write(gg, Polarion_ID_Data['status'])
            worksheet.write(hh, Polarion_ID_Data['resolution'])
            worksheet.write(ii, Polarion_ID_Data['FGL'])


        ###########################33
        worksheet.write(jj, Elements['SWCs'])
        worksheet.write(kk, Elements['Functional affected SWCS'])
        worksheet.write(ll, change_List_To_string(Elements['Affected System Function']), cell_format2)
        worksheet.write(mm, Elements['Date'])
        worksheet.write(nn, Elements['Rolling Build'])
        worksheet.write(oo, Elements['Commit Message'])
        worksheet.write(pp, Elements['file List'])
        worksheet.write(qq, Elements['ParFiles'])
        worksheet.write(rr, Elements['A2lFiles'])
        worksheet.write(ss, Elements['SourceFiles'])
        worksheet.write(tt, Elements['HeaderFiles'])
        worksheet.write(uu, Elements['RTRT'])
        if Verification_Plan_Flag == 1:
            worksheet.write(vv, Elements['Polarion Plan Verification'])



        if Elements['Polarion Plan Verification'] == "Not Planned":
            worksheet.write(qq, Elements['Polarion Plan Verification'], Failed_format)
        elif Elements['Polarion Plan Verification'] == "✓":
            worksheet.write(qq, Elements['Polarion Plan Verification'], Pass_format)
        else :
            worksheet.write(qq, Elements['Polarion Plan Verification'])

        indix = indix +1

    # KPI Data WRite
    KPI_Data = Generate_KPI(Out_Report_Data)
    #headers
    worksheet0.write("A2", 'Num_Of_Commits', format1)
    worksheet0.write("A3", 'Commits_Have_Functional_Changes', format1)
    worksheet0.write("A4", 'Commits_Have_RTRT_Changes', format1)
    worksheet0.write("A5", 'DRCO_Commits', format1)
    worksheet0.write("A6", 'LLSW_Commits', format1)
    worksheet0.write("A7", 'BSW_Commits', format1)
    worksheet0.write("A8", 'SSW_Commits', format1)
    worksheet0.write("A9", 'Integration_Commits', format1)
    worksheet0.write("A10", 'Others_Commits', format1)
    worksheet0.write("A11", 'Orphan_Commits', format1)

    # Data
    worksheet0.write("B2", KPI_Data[0])
    worksheet0.write("B3", KPI_Data[1])
    worksheet0.write("B4", KPI_Data[2])
    worksheet0.write("B5", KPI_Data[3])
    worksheet0.write("B6", KPI_Data[4])
    worksheet0.write("B7", KPI_Data[5])
    worksheet0.write("B8", KPI_Data[6])
    worksheet0.write("B9", KPI_Data[7])
    worksheet0.write("B10", KPI_Data[8])
    worksheet0.write("B11", KPI_Data[9])
    workbook1,worksheet0 = Write_KPI_Data_To_Sheet(workbook1, worksheet0)
    #workbook1.insert_chart('D2', Cahrt, {'x_offset': 25, 'y_offset': 10})



    #print("report Passing 3")
    while True:
        try:
            workbook1.close()
            Task["Warning_Message"] = ""
            #os.startfile(output_workbook)
            break
        except:
            print ("File 'Integration_Final_Report' Is already opened ! , please close Excel file ")
            Task["Warning_Message"] = "[Integration_Final_Report.xlsx] file Is already opened ! , please close Excel file "
            pass

def change_List_To_string(List):
    List_Str =""
    for elements in List:
        List_Str = List_Str + " " + elements
    return List_Str

def Report_Generation_Verification(infoList,Polarion_IDs_Data,Verification_Plan_Flag,Input_Data):
    #Polarion_Origin_Plan,Polarion_Chield_Plan=Get_Polarion_Plan_IDs("Outputs\Polarion_Plan.xlsx")
    SWC_SF_List =[]
    Integrated_Report, Gerrit_Tickets = Get_Integration_Report_Data("SubFiles\Integration_Report.xlsx")

    if Input_Data["SF_Mapping"] == 1:
        SWC_SF_List, SWC_List = Get_SWC_System_Function_Data("Outputs\System_Function_Report.xlsx")
    else:
        SWC_SF_List=0

    Out_Put_Report_Array =Compare_SWC_SystemFunction_List(Integrated_Report, SWC_SF_List)

    Generate_Integration_Report_Report(infoList,Out_Put_Report_Array,Polarion_IDs_Data,Verification_Plan_Flag)



def Get_Ticket_Status(ID , Polarion_IDs_Data):
    Rtn_ID = {}
    rtn_Flg = 0
    for sub_Elements in Polarion_IDs_Data :
        if str(ID) == str(sub_Elements["ID"]) :
            Rtn_ID = sub_Elements
            rtn_Flg = 1
            return Rtn_ID , rtn_Flg
    if rtn_Flg == 0:
        return 0 , 0

def Main_Task(infoList1, infoList2 , Input_Data):
    #print("SF Flag:",Input_Data["SF_Mapping"])
    print(Input_Data)
    Task["Progress"] = 1
    Task["Name"] = "Creat Directory"

    # Create Output report in Subfile Folder
    print("Create Output report in Subfile Folder")
    time.sleep(1)
    Input_Data["Report_Name"] = Create_Out_Report()

    Task_Flag["Task1"] ="✓"
    Task["Progress"] = 3
    Task["Name"] = "Prepare shell script and execute it => Output is Log.txt in subfiles filder"

    # Prepare shell script and execute it => Output is Log.txt in subfiles filder
    print("Prepare shell script and execute it => Output is Log.txt in subfiles filder")
    time.sleep(1)
    Prepare_Shell_Script(Input_Data)
    os.system("Get_Repo.sh")
    Task_Flag["Task2"] = "✓"
    Task["Progress"] = 10
    Task["Name"] = "Parse Log file and generate Diff logs and pares data into Subfiles/IntegrationReport.xlsx"

    # Parse Log file and generate Diff logs and pares data into Subfiles/IntegrationReport.xlsx
    print("Parse Log file and generate Diff logs and pares data into Subfiles/IntegrationReport.xlsx")
    time.sleep(1)
    Polarion_Tags_List = Gerrit_Changes_Analysis_Runnable(Input_Data)
    Task_Flag["Task3"] = "✓"
    Task["Progress"] = 25
    Task["Name"] = "Get polarion plan tasks and WPs and all WIs and verify all Gerrit tickes Ids from polarion and get all there data"

    # Get polarion plan tasks and WPs and all WIs and verify all Gerrit tickes Ids from polarion and get all there data
    print("Get polarion plan tasks and WPs and all WIs and verify all Gerrit tickes Ids from polarion and get all there data")
    time.sleep(1)

    Polarion_IDs_Data = Polarion_Plan_Runnable(infoList1, Polarion_Tags_List)

    Task_Flag["Task4"] = "✓"
    #print(Polarion_IDs_Data)
    Task["Progress"] = 55
    Task["Name"] = "generate report with system function mapping to SWC from polarion"

    # generate report with system function mapping to SWC from polarion
    print("generate report with system function mapping to SWC from polarion")
    time.sleep(1)
    print(Input_Data["SF_Mapping"])
    if Input_Data["SF_Mapping"] == 1:
        SWC_SF_Report_Generation_Runnable(infoList2)
    else:
        print("Generate Empty File")
        Create_Temp_OutPut()

    Task_Flag["Task5"] = "✓"
    Task["Progress"] = 90
    Task["Name"] = "Generate Final report "

    # Generate final report
    print("Generate final report")
    time.sleep(1)
    Report_Generation_Verification(infoList2, Polarion_IDs_Data, Verification_Plan_Flag[ "Flag"],Input_Data)
    Task_Flag["Task6"] = "✓"
    Task["Progress"] = 100
    Task["Name"] = "Finish ! -'Outputs\Integration_Final_Report.xlsx' report Generated "
    #file = "C:\\Documents\\file.txt"
    #os.startfile(file)

'''
infoList1 =['melmohta', 'Aliya$1994', 'p2_800v_sic_inv_switching_cell', '', '23_03_03_01_01_a2_02']
infoList2 =['melmohta', 'Aliya$1994', 'p2_800v_sic_inv_switching_cell', '23_03_03_01_01_a2_02', '', 'P2-800V_SIC-MASTER-0001-20220603']
Input_Data ={
    'Project': 'p2_800v_sic_inv_switching_cell',
    'User name': 'melmohta',
    'Password': 'Aliya$1994',
    'Int_Plan_Doc': '',
    'Polarion_SW_Plan': '',
    'RB': 'P2-800V_SIC-MASTER-0001-20220603',
    'L_Repo': 'C:\\\\P2-800V-Gen5\\\\proj5742_inv_gen5',
    'Directory': 'C:\\\\Users\\\\melmohta\\\\Desktop\\\\Integrator_Role\\\\Scripts\\\\Transfer_Plan_Script\\\\Vsersion_2',
    'Report_Name': '',
    'SWA_Baseline': '23_03_03_01_01_a2_02',
    'SF_Mapping' : "" ,
}
Main_Task(infoList1, infoList2 , Input_Data)


'''

