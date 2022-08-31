import re

Input_Path = "27_05_18_p330_and_q330"
variant= "variant.KEY:base\-"
Documents = ["Inverter ActeSafeSt 01_SWRS_P330","Inverter ActeSwtgSig 01_SWRS_Q330","Inverter DchaDcLink 01_SWRS_P330_Q330","Inverter DetmnCooltFlow 01_SWRS_P330_Q330","Inverter DetmnCstrReqVal 01_SWRS_P330","Inverter DetmnCstrReqVal 01_SWRS_Q330"]

#variant.KEY:base\-

def Document_Variant_Check(Input_Path ,variant,Document_Name ):
    Base_P_Key = "P"
    Base_M_Key = "Q"

    Split_Key = re.split("_", Input_Path)
    x = len(Split_Key)

    num = re.sub(r'\D', "", Split_Key[x-1])

    if variant == "variant.KEY:base\+":
        Search_Key = Base_P_Key+num
        if re.search(Search_Key, Document_Name):
            return "Base+"
    elif variant == "variant.KEY:base\-":
        Search_Key = Base_M_Key + num
        if re.search(Search_Key, Document_Name):
            return "Base-"
    else:
        return 0


