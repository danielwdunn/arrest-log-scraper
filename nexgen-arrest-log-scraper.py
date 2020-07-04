import fitz
import xlsxwriter

pdf_document = "arrestlog.pdf"
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
                content = [incidentNo,arrestingOfficer,whereArrested,charge]
                for item in content:
                    # write operation perform
                    worksheet.write(row, column, item)
                    # incrementing the value of row by one
                    # with each iteratons.
                    column += 1
                row+=1
                column=0

workbook.close()
