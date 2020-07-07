import fitz
import xlsxwriter

pdf_document = "hamdenarrestlog.pdf"
doc = fitz.open(pdf_document)
page=doc.pageCount
workbook = xlsxwriter.Workbook('Output.xlsx')
row = 0
column = 0

worksheet = workbook.add_worksheet()
for i in range(page):
            page1 = doc.loadPage(i)
            page1text = page1.getText("text")
            page=page1text
            page=page.split(" Remarks:")
            for j in range(len(page)-1):

                try:
                    race=page[j].split("Name:")[0]
                    race1=str.splitlines(race)
                    for i in race1:
                        if race1[0] == '':
                            race_of_arrested = race1[11]
                        elif race1[0] == 'HAMDEN POLICE DEPARTMENT':
                             race_of_arrested = race1[18]
                        else:
                            race_of_arrested = "Race not Found"
                except: race_of_arrested = "error"

                try:
                    sex=page[j].split("Name:")[0]
                    sex1=str.splitlines(sex)
                    for i in sex1:
                        if sex1[0] == '':
                            sex_of_arrested = sex1[12]
                        elif sex1[0] == 'HAMDEN POLICE DEPARTMENT':
                             sex_of_arrested = sex1[19]
                        else:
                            sex_of_arrested = "Sex not Found"
                except: sex="error"
                try:
                    dob=page[j].split("D.O.B.")[1]
                    dob1=(str.splitlines(dob))
                    date_of_birth = dob1[1]
                except:
                    dob="error"
                try:
                    incidentNo=page[j].split(" Arresting Officer:")[0].split("\n")[-2]
                    if(incidentNo[0].isdigit()):
                        incidentNo = page[j].split(" Arresting Officer:")[0].split("\n")[-2]
                    else:
                        incidentNo =""
                except:
                    incidentNo=""
                try:
                    arrestingOfficer=page[j].split(" Arresting Officer:")[0].split("\n")[-3]
                except:
                    arrestingOfficer=""
                try:
                    whereArrested=page[j].split("Where Arrested:")[1].split("Court Date:")[0].strip()
                except:
                    whereArrested=""
                try:
                    dateArrested=page[j].split("Arrested:")[1].strip(" Where")
                    dateTimeArrestedList=(str.splitlines(dateArrested))
                    dateA: str=dateTimeArrestedList[1]
                    timeA=dateTimeArrestedList[2]
                except:
                    dateArrested=""
                charge=""
                try:
                    a=page[j].split("Description")[1].split("\n")
                    a = filter(None, a)
                    a = list(filter(None, a))
                    for i in range(len(a)):
                        if(a[i][0].isdigit()):
                            if(len(a[i])>4):
                                charge=charge+a[i]+"|"
                except:
                    charge=""
                l=charge
                content = [incidentNo,dateA,timeA,arrestingOfficer,whereArrested,date_of_birth,race_of_arrested,sex_of_arrested,charge]
                for item in content:
                    # write operation perform
                    worksheet.write(row, column, item)
                    # incrementing the value of row by one
                    # with each iteratons.
                    column += 1
                row+=1
                column=0

workbook.close()
