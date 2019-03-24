'''
Author: Ethan Fowler EHD
Date: Mar 23 2019
Description: Used to take a .pst file that has a folder with all the process fail emails that are sent and automatically plug them into the Process Fail Report. 

Future Plans:
- Write the report itself
- GUI for user usablity
- Read from inbox directly
'''
##To read pst
import pypff

##To write excel file
import pandas as pd
from pandas import ExcelWriter



def parse_email(email):
    temp = {"Date":None,"Ticket":None,"Agent":None,"Reason":None, "Kickback":None}
    
    ##Get the reason and the ticket number
    if "IM" not in email.subject.split(":")[1]:
        temp["Reason"] = email.subject.split(":")[1][1:]
        temp["Ticket"] = "N/A"
    else:
        reason = email.subject.split(":")[1]
        reason = reason.replace("(", "")
        reason = reason.replace(")", "")
        temp["Reason"] = reason.split("IM")[0][1:-1]
        temp["Ticket"] = "IM" + reason.split("IM")[1]

    ##Get the agent and the date
    for x in email.transport_headers.split("\n"):
        if "To" in x:
            agent = x.split(":")[1].strip()
            agent = agent.replace("<","")
            agent = agent.replace(">", "")
            temp["Agent"] = agent
        elif "Date" in x:
            temp['Date'] = x.split(":")[1:-2][0][1:]

    ##Get kickbacks
    for y in str(email.plain_text_body).split("\\n"):
        if "Kickback" in y:
            temp["Kickback"] = y.split(":")[1].strip("\\r").upper()[1:]
        else:
            temp["Kickback"] = "N"

    return temp


def get_process_fails(pstlocation):
    process_fail_list = []
    pstFile = open(pstlocation, "rb")

    pff_file = pypff.file()

    pff_file.open_file_object(pstFile)
    for i,x in enumerate(pff_file.get_root_folder().sub_items):
        for y in x.sub_items:
            if y.name == "Process Failures":
                for z in y.sub_items:
                    if z.subject.split(":")[0] != "Process Failure":
                        continue
                    
                    process_fail = parse_email(z)

                    process_fail_list.append(process_fail)
    return process_fail_list
                
def write_report(process_fails, report):
   
    writer = ExcelWriter(report)
    data = {"Date":[], "Ticket":[], "Agent":[], "Reason":[], "Kickback":[]}
    for row in process_fails:
        for key, val in row.items():
            data[key].append(val)
    
    df = pd.DataFrame(data)
    df.to_excel(writer, 'Sheet 1', index=False)
    writer.save()
    

if __name__ == "__main__":
    ##input pst location
    process_fails = get_process_fails("ProcessFails.pst")

    write_report(process_fails, "Report.xlsx")