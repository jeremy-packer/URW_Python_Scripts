import os
import re
import sys
import pytz
import random
import datetime
import win32com.client

#this creates an outlook object and connects to current outlook session
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#grabs the sales inbox
inbox = outlook.Folders("URW Sales").Folders("Inbox")
#grabs the retailers email folder
retailer_inbox = inbox.Folders("Retailer Emails")

##############MAKE SURE TO MANUALLY CHANGE MONTH AND YEAR IF APPLICABLE FOR REPORTING################	
FolderPath = "Z:\\MARKETING\\ED&A\\Sales and Traffic Analysis\\Monthly Sales Reporting\\2021\\02 - February\\Retailers"
emails = inbox.Items

#Loops through each email
for email in emails:

	#Aeropostale
	if email.SenderEmailAddress == "salesreporting=aeropostale.com@lucernex.com":
	#if email.SenderEmailAddress[-11:] == "nautica.com":
		print("Aeropostale")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Aeropostale\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Aeropostale\\Email - " + str(counter) + ".pdf" , 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Aeropostale\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Aeropostale\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Aeropostale"))
	
	#Aldo	
	if email.SenderEmailAddress == "realestate=aldogroup.com@lucernex.com":
		#generic counter
		print("Aldo")
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Aldo\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Aldo\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Aldo\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Aldo\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Aldo"))
	
	#Apple
	if email.Subject.strip()[0:5] == "Apple":
		#generic counter
		print("Apple")
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Apple\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Apple\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Apple\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Apple\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Apple"))	

	#Ascena
	#if email.SenderEmailAddress == "Jyoti.Bandal@ascenaretail.com":
	#if email.SenderEmailAddress == "ShailendraKumar.Vaishya@ascenaretail.com":
	if email.Subject.strip()[0:17] == "[LxRetail] Ascena":
		print("Ascena")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Ascena\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Ascena\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Ascena\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Ascena\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Ascena"))	
	
	#Zumiez and Bare Escentuals have no distinguishing indicators so make sure to comment out whatever email hasn't been recieved
	#Bare Escentuals	
	####if email.SenderEmailAddress == "salesreporting@bareescentuals.com" or email.SenderEmailAddress.strip()[-12:] == "shiseido.com":
	#if email.SenderEmailAddress == "SalesReport@CoStarREManager.com":	
	#	print("Bare Escentuals")
	#	#generic counter
	#	counter = 1
	#	#checks to see if the file name exists. if the file name does exist increment counter by 1. 
	#	while os.path.exists(FolderPath + "\\Bare Escentuals\\Email - " + str(counter) + ".pdf"):
	#		counter = counter + 1
	#	#run once the while loop stops and determines that a file with email# doesn't exist
	#	email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Bare Escentuals\\Email - " + str(counter) + ".pdf", 17)
	#	#checks to see if attachment type is something we want
	#	counter2 = 1
	#	for att in email.Attachments:
	#		hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
	#		if hold:
	#			while os.path.exists(FolderPath + "\\Bare Escentuals\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
	#				counter2 = counter2 + 1
	#			att.SaveAsFile(FolderPath + "\\Bare Escentuals\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
	#	email.Move(retailer_inbox.Folders("Bare Escentuals"))
	
	#Ben Bridge	
	if email.SenderEmailAddress == "Lori.Hamamoto@BenBridge.com":
		print("Ben Bridge")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Ben Bridge\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Ben Bridge\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Ben Bridge\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Ben Bridge\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Ben Bridge"))
		
	#Brighton Collectibles
	if email.SenderEmailAddress == "icampas@brighton.com":
		print("Brighton Collectibles")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Brighton Collectibles\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Brighton Collectibles\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Brighton Collectibles\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Brighton Collectibles\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Brighton Collectibles"))
		
	#Build A Bear
	if email.SenderEmailAddress == "donotreply=buildabear.com@lucernex.com":
		print("Build A Bear")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Build A Bear\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Build A Bear\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Build A Bear\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Build A Bear\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Build A Bear"))

	#Claires
	if email.SenderEmailAddress == "HEATHER.RUIS@Claires.com":
		print("Claires")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Claires\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Claires\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Claires\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Claires\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Claires"))
	
	#Coffee Bean
	#if email.SenderEmailAddress.strip()[-14:] == "coffeebean.com":
	#if email.SenderEmailAddress.strip() == "norman.galido@jws.com.ph":
	if email.SenderEmailAddress.strip() == "marites.castor@jws.com.ph":
		print("coffee bean")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Coffee Bean\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Coffee Bean\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Coffee Bean\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Coffee Bean\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Coffee Bean"))

	#Cotton On
	if email.SenderEmailAddress.strip() == "noreply@leaseeagle.com":
		print("Cotton On")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Cotton On\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Cotton On\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Cotton On\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Cotton On\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Cotton On"))

	#Disney
	if email.SenderEmailAddress.strip()[-10:] == "disney.com":
		print("Disney")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Disney\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Disney\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Disney\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Disney\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Disney"))
	
	#Eddie Bauer
	if (re.search("Eddie Bauer", email.Body) != None or re.search("Eddie Bauer", email.Subject) != None):
		print("Eddie Bauer")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Eddie Bauer\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Eddie Bauer\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Eddie Bauer\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Eddie Bauer\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Eddie Bauer"))	
		
	#Express
	if re.search("EXPRESS FASHION", email.Body) != None or re.search("EXPRESS, LLC", email.Body) != None :
	#if email.Subject.strip() == "[LxRetail] Express Sales Report":
		print("Express")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Express\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Express\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Express\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Express\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Express"))	
		
	#Finish Line
	if email.SenderEmailAddress == "LeaseAdmin=finishline.com@lucernex.com":
		print("Finish Line")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Finish Line\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Finish Line\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Finish Line\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Finish Line\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Finish Line"))	


	#Footlocker
	if email.SenderEmailAddress == "donotreply=footlocker.lucernex.com@lucernex.com":
		print("Footlocker")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Footlocker\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Footlocker\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Footlocker\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Footlocker\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Footlocker"))			
		
	#Forever 21
	#if email.SenderEmailAddress == "f21sales21@gmail.com":
	if email.SenderEmailAddress == "ella.n@forever21.com":
		print("Forever 21")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Forever 21\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Forever 21\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Forever 21\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Forever 21\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Forever 21"))		
		
	#Gap
	if email.SenderEmailAddress.strip()[-7:] == "gap.com":
		print("Gap")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Gap\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Gap\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Gap\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Gap\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Gap"))	
		
	#Genesco
	if email.SenderEmailAddress == "GenescoDoNotReply@lucernex.com":
	#if email.SenderEmailAddress == "CERTIFIEDSALES@genesco.com":
		print("Genesco")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Genesco\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Genesco\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Genesco\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Genesco\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Genesco"))
		
	#GNC
	if email.SenderEmailAddress == "GNCDoNotReply@Tangoanalytics.com":
		print("GNC")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\GNC\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\GNC\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\GNC\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\GNC\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("GNC"))
		
	#Guess
	if email.SenderEmailAddress.strip()[-9:] == "guess.com":
		print("Guess")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Guess\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Guess\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Guess\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Guess\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Guess"))
		

	#H&M
	if email.SenderEmailAddress.strip()[-9:] == "US@hm.com":
		print("H&M")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\H&M\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\H&M\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\H&M\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\H&M\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("H&M"))		
		
	#Helzberg	
	if email.Subject.strip()[0:19] == "[LxRetail] Helzberg":
		print("Helzberg")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Helzberg\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Helzberg\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Helzberg\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Helzberg\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Helzberg"))	
		
	#Hot Topic	
	if email.Subject.strip()[0:20] == "[LxRetail] Hot Topic" or email.SenderEmailAddress.strip()[-12:0] == "hottopic.com" or email.SenderEmailAddress.strip()[0:17] == "tdstoresalesrepor":
		print("Hot Topic")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Hot Topic\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Hot Topic\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Hot Topic\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Hot Topic\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Hot Topic"))	
	
	#J. Crew
	#if email.SenderEmailAddress.strip()[-9:] == "JCrew.Com": 
	if email.SenderEmailAddress.strip() == "JCrewDoNotReply@Tangoanalytics.com":
		print("J Crew")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\J. Crew\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\J. Crew\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\J. Crew\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\J. Crew\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("J. Crew"))	
		
	#J. Jill
	if email.SenderEmailAddress == "donotreply@lucernex.com" and (email.Subject[0:18] == "[LxRetail] J. Jill" or re.search("JILL", email.Body) != None) : 
		print("J. Jill")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\J. Jill\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\J. Jill\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\J. Jill\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\J. Jill\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("J. Jill"))		
		
	#Jos A Bank
	if email.SenderEmailAddress == "ladministration@jos-a-bank.com": 
		print("Jos A Bank")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Jos A Bank\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Jos A Bank\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Jos A Bank\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Jos A Bank\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Jos A Bank"))

	#Kay Jewelers
	#if email.SenderEmailAddress == "Steven.Sarich@signetjewelers.com":
	#if email.SenderEmailAddress == "Sue.Major@signetjewelers.com": 
	if (re.search("Kay #", email.Body) != None):
		print("Kay Jewelers")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Kay Jewelers\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Kay Jewelers\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Kay Jewelers\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Kay Jewelers\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Kay Jewelers"))	
		
	#L Brands
	if email.SenderEmailAddress.strip()[-6:] == "lb.com": 
		print("L Brands")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\L Brands\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\L Brands\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\L Brands\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\L Brands\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("L Brands"))		
	
	#Lids
	#if email.SenderEmailAddress.strip() == "leasing@lids.com":
	#if email.SenderEmailAddress.strip() == "Shannon.Frazier@lids.com":
	if email.SenderEmailAddress.strip() == "donotreply=lids.com@lucernex.com":
		print("Lids")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Lids\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Lids\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Lids\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Lids\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Lids"))

	
	#L'Occitane
	if email.SenderEmailAddress.strip()[-13:] == "loccitane.com": 
		print("L Occitane")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\L'Occitane\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\L'Occitane\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\L'Occitane\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\L'Occitane\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("L'Occitane"))		
		
	#L'Oreal
	if email.SenderEmailAddress.strip()[-10:] == "loreal.com": 
		print("L'Oreal")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\L'Oreal\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\L'Oreal\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\L'Oreal\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\L'Oreal\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("L'Oreal"))			
			
	#Lovesac
	if (re.search("SAC ACQUISITION LLC", email.Body) != None or re.search("Lovesac", email.Body) != None):
	#if email.Subject[0:18] == "[LxRetail] Lovesac": 
	#if email.Subject == "[LxRetail] Gross Sales Information":
		print("Lovesac")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Lovesac\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Lovesac\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Lovesac\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Lovesac\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Lovesac"))
	
	#Lucky Brand
	#if email.SenderEmailAddress == "Leaseadminsales@costarremail.com":
	if (re.search("Lucky Brand", email.Body) != None or re.search("LUCKY BRAND", email.Body) != None):
		print("Lucky Brand")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Lucky Brand\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Lucky Brand\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Lucky Brand\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Lucky Brand\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Lucky Brand"))
		
	#Lululemon
	if email.SenderEmailAddress.strip()[-13:] == "lululemon.com": 
		print("Lululemon")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Lululemon\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Lululemon\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Lululemon\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Lululemon\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Lululemon"))	
		
	#Luxottica
	if email.SenderEmailAddress.strip()[-19:] == "luxotticaretail.com": 
		print("Luxottica")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Luxottica\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Luxottica\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Luxottica\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Luxottica\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Luxottica"))	
		
	#Men's Wearhouse
	if email.SenderEmailAddress == "leaseadmin@tailoredbrands.com": 
		print("Men's Wearhouse")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Men's Wearhouse\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Men's Wearhouse\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Men's Wearhouse\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Men's Wearhouse\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Men's Wearhouse"))	

	#New York & Company
	##if email.SenderEmailAddress == "donotreply@lucernex.com":
	#if (re.search("New York", email.Body) != None or re.search("[LxRetail] New York", email.Subject) != None):
	##if email.Subject.strip()[0:19] == "[LxRetail] New York": 
	#	print("New York & Company")
	#	#generic counter
	#	counter = 1
	#	#checks to see if the file name exists. if the file name does exist increment counter by 1. 
	#	while os.path.exists(FolderPath + "\\New York & Company\\Email - " + str(counter) + ".pdf"):
	#		counter = counter + 1
	#	#run once the while loop stops and determines that a file with email# doesn't exist
	#	email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\New York & Company\\Email - " + str(counter) + ".pdf", 17)
	#	#checks to see if attachment type is something we want
	#	counter2 = 1
	#	for att in email.Attachments:
	#		hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
	#		if hold:
	#			while os.path.exists(FolderPath + "\\New York & Company\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
	#				counter2 = counter2 + 1
	#			att.SaveAsFile(FolderPath + "\\New York & Company\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
	#	email.Move(retailer_inbox.Folders("New York & Company"))			
		
	#PacSun
	#if (re.search("Pacific Sunwear", email.Body) != None or re.search("Pacific Sunwear", email.Subject) != None):
	##if email.SenderEmailAddress == "rramsey@consultasg.com":
	if (re.search("Pacific Sunwear", email.Body) != None or email.SenderEmailAddress.strip()[-18:] == "pacificsunwear.com"):
		print("PacSun")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\PacSun\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\PacSun\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\PacSun\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\PacSun\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("PacSun"))		
		
	#Papyrus
	#if email.SenderEmailAddress == "landlordinv=srgretail.com@proleaseweb.com":
	if email.SenderEmailAddress == "landlordinv@srgretail.com":
		print("Papyrus")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Papyrus\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Papyrus\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Papyrus\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Papyrus\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Papyrus"))	
		
	#Restoration Hardware
	if email.SenderEmailAddress == "SalesReportingRestorationHardware@CoStarREManager.com":
		print("Restoration Hardware")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Restoration Hardware\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Restoration Hardware\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Restoration Hardware\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Restoration Hardware\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Restoration Hardware"))	
		
	#See's
	if email.SenderEmailAddress[-8:] == "sees.com":
		print("See's")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\See's\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\See's\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\See's\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\See's\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("See's"))	
		
	#Sephora
	#if email.SenderEmailAddress == "Lease_Compliance@sephora.com":
	if email.SenderEmailAddress == "sephora_mgmt@accruent.com":
		print("Sephora")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Sephora\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Sephora\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Sephora\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Sephora\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Sephora"))	


	#Signet Jewelers
	if email.SenderEmailAddress == "Steven.Sarich@signetjewelers.com" and (re.search("Kay #", email.Body) == None):
		print("Signet Jewelers")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Signet Jewelers\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Signet Jewelers\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Signet Jewelers\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Signet Jewelers\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Signet Jewelers"))			
		
	#Skechers
	if email.SenderEmailAddress == "skxfinance@skechers.com":
		print("Skechers")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Skechers\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Skechers\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Skechers\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Skechers\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Skechers"))	
		
	#Spencer Gifts
	if email.SenderEmailAddress == "SpencerGifts-Salereporting@costarremail.com":
		print("Spencer Gifts")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Spencer Gifts\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Spencer Gifts\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Spencer Gifts\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Spencer Gifts\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Spencer Gifts"))				

	#Starbucks
	if email.SenderEmailAddress == "noreply.salesreporting@starbucks.com":
		print("Starbucks")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Starbucks\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Starbucks\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Starbucks\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Starbucks\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Starbucks"))		
		
	#Swarovski
	if email.SenderEmailAddress.strip()[-13:] == "swarovski.com":
		print("Swarovski")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Swarovski\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Swarovski\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Swarovski\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Swarovski\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Swarovski"))		
		
	#Tapestry
	if email.SenderEmailAddress.strip() == "salesnotification@tapestry.com":
		print("Tapestry")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Tapestry\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Tapestry\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Tapestry\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Tapestry\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Tapestry"))
		
	##The Children's Place
	#if email.SenderEmailAddress == "support@CoStarREManager.com":
	#if email.SenderEmailAddress == "SalesReport@CoStarREManager.com":
	if (re.search("Children's Place", email.Body) != None):
		print("The Children's Place")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\The Children's Place\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\The Children's Place\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\The Children's Place\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\The Children's Place\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("The Children's Place"))		
		
	#The Walking Company
	#if email.SenderEmailAddress.strip()[-21:] == "thewalkingcompany.com":
	if email.SenderEmailAddress.strip() == "EstelaT@walkingco.com":
		print("The Walking Company")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\The Walking Company\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\The Walking Company\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\The Walking Company\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\The Walking Company\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("The Walking Company"))		

	#VF Corporation
	if email.SenderEmailAddress == "Steven_Frommell@vfc.com":
		print("VF Corporation")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\VF Corporation\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\VF Corporation\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\VF Corporation\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\VF Corporation\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("VF Corporation"))	
		
	#Yankee Candle
	if email.SenderEmailAddress[-12:] == "newellco.com":
		print("Yankee Candle")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Yankee Candle\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Yankee Candle\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Yankee Candle\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Yankee Candle\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Yankee Candle"))		
		
	#Zara
	if email.SenderEmailAddress == "VeskoDr@inditex.com":
		print("Zara")
		#generic counter
		counter = 1
		#checks to see if the file name exists. if the file name does exist increment counter by 1. 
		while os.path.exists(FolderPath + "\\Zara\\Email - " + str(counter) + ".pdf"):
			counter = counter + 1
		#run once the while loop stops and determines that a file with email# doesn't exist
		email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Zara\\Email - " + str(counter) + ".pdf", 17)
		#checks to see if attachment type is something we want
		counter2 = 1
		for att in email.Attachments:
			hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
			if hold:
				while os.path.exists(FolderPath + "\\Zara\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
					counter2 = counter2 + 1
				att.SaveAsFile(FolderPath + "\\Zara\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
		email.Move(retailer_inbox.Folders("Zara"))			
		
	#Zumiez and Bare Escentuals have no distinguishing indicators so make sure to comment out whatever email hasn't been recieved
	#Zumiez
	##if email.SenderEmailAddress == "leaseaccounting@zumiez.com":
	#if email.SenderEmailAddress == "SalesReport@CoStarREManager.com":
	#	print("Zumiez")
	#	#generic counter
	#	counter = 1
	#	#checks to see if the file name exists. if the file name does exist increment counter by 1. 
	#	while os.path.exists(FolderPath + "\\Zumiez\\Email - " + str(counter) + ".pdf"):
	#		counter = counter + 1
	#	#run once the while loop stops and determines that a file with email# doesn't exist
	#	email.GetInspector.WordEditor.ExportAsFixedFormat(FolderPath + "\\Zumiez\\Email - " + str(counter) + ".pdf", 17)
	#	#checks to see if attachment type is something we want
	#	counter2 = 1
	#	for att in email.Attachments:
	#		hold = (att.DisplayName.strip()[-3:] != "jpg" and att.DisplayName.strip()[-3:] !=  "png" and att.DisplayName.strip()[-3:] !=  "peg" and att.DisplayName.strip()[-3:] !=  "gif")		
	#		if hold:
	#			while os.path.exists(FolderPath + "\\Zumiez\\" + str(counter) + "." + str(counter2) + " - " + att.FileName):
	#				counter2 = counter2 + 1
	#			att.SaveAsFile(FolderPath + "\\Zumiez\\" + str(counter) + "." + str(counter2) + " - " + att.FileName)
	#	email.Move(retailer_inbox.Folders("Zumiez"))		
		