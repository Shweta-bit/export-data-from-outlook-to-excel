import win32com.client
import os
import xlsxwriter

outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")
workbook = xlsxwriter.Workbook("AlertMAnagerReport.xlsx")
worksheet = workbook.add_worksheet()

contentCol = ["RecievedTime", "Alertname", "namespace", "Job", "message"]

Col = 0

for item in contentCol:
	worksheet.write(0,Col,item)
	Col +=1
	
inbox = outlook.Getdefaultfolder(6)

message = inbox.items

message.sort("[RecievedTime]", True)

row = 1

try:
	for message in messages:
		alertname = " "
		namespace = " "
		job = " "
		messageAm = "NA || "
		
		if str(message.SenderName)== "alertmanager@fareportal.com" and (message.UnRead==True or message.UnRead==False):
			body = str(message.body)
			recievedTime = str(message.RecievedTime)
			senderName = str(message.SenderName)
			messageLines = body.splitlines()
			
			for line in messageLines:
				x = line.split()
				if x[0] == "alertname":
					alertname += x[1]+" || "
				if x[0] == "namespace":
					namespace += x[1]+" || "
				if x[0] == "job":
					job += x[1]+" || "
				if x[0] == "message":
					message += x[1]+" || "
					
			
			worksheet.write(row, 0, recievedTime)
			worksheet.write(row,1, alertname)
			worksheet.write(row, 2 , namespace)
			worksheet.write(row, 3, job)
			worksheet.write(row, 4 , message)
			
			row +=1
except:
	print("encountered an exception")
				