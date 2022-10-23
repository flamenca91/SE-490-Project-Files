import docx
import re
import xlwings

docRelation = {"HRD":("HRS"), "HRS":("PRS"), "PRS":("URS","RISK"), "HTR":("HTP"), "HTP":("HRD", "HRS"), \
               "SDS":("BOLUS","ACE","AID"), "ACE":("PRS", "TBV", "DER"), "BOULUS":("PRS"), "AID":("PRS","DER"), \
               "SVAL":("BOLUS", "ACE", "AID"), "SVATR":("SVAL"), "UT":("UNIT"), "INS": ("UNIT")}      # to be created by the GUI

docFile = {"HRD":"HDS_new_pump.docx", "HRS":"HRS_new_pump.docx", "HTP":"HTP_new_pump.docx", "HTR":"HTR_new_pump.docx", \
           "PRS":"PRS_new_pump.docx", "RISK":"RiskAnalysis_Pump.docx", "SDS":"SDS_New_pump_x04.docx", "ACE":"SRS_ACE_Pump_X01.docx", \
           "BOLUS":"SRS_BolusCalc_Pump_X04.docx", "SRS":"SRS_DosingAlgorithm_X03.docx", "SVAL":"SVaP_new_pump.docx", \
           "SVATR":"SVaTR_new_pump.docx", "UT":"SVeTR_new_pump.docx", "URS":"URS_new_pump.docx"}

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    fullText = [ele for ele in fullText if ele.strip()]
    return fullText


#def main():
txtLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/HDS_new_pump.docx')
index = 0
ind = []

for t in txtLst:
    if re.search('.*:HRD:', t):
        ind.append(index)
        tt=t
        y = re.findall('\S*:HRD:\S*', t)
        z = re.findall('\S*:HRS:\S*', t)
        tt = tt.replace(y[0], '')
        tt = tt.replace(z[0], '')
        tt = tt.strip()

        print(tt)
        #print(t)
        print(y)
        print(z)
    index = index + 1

print(ind)
#print(txtLst)


