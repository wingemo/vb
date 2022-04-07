'------------------------------------------'
' Philip Wingemo 2022-04-05
' FX Solutions 
' Close currency account
'------------------------------------------'

Public obj_file, file 
Public excel_path  ' Path to customer list
Public mailbox ' Mailbox to load emails
Public p_id ' ID for user
Public folder_year ' Mailbox folder year
Public folder_emails ' Folder with emails
Public email_error ' Email for failed closing
Public email_sucess  ' Email for sucess closing
Public email_fxtrade ' Email customer Fx Trade
Public flag_atlas, flag_menyb, subject

Function importMail()
	'Import mail from outlook'
	'----------------------------------'
	
	Dim objOutlook 
	Dim objOutlookMsg
	Dim olMailItem
	Dim olFolderInbox, iTotalMails
	
	olFolderInbox = 6 : sSubject = ""
	
	Set objOutlook = CreateObject("Outlook.Application")
	Set objNamespace = objOutlook.GetNamespace("MAPI")
	
	Set objOwner = objNamespace.CreateRecipient(mailbox)

	objOwner.Resolve
	
	'Create reference to Inbox Folder
	Set oInbox = objNamespace.GetSharedDefaultFolder(objOwner, olFolderInbox) 
	
	'Find all items in the Inbox Folder
	Set folder = oInbox.Folders(folder_year).Folders(folder_emails)
	Set oAllMails = folder.Items
	
	'Find out properties of the mail item
	importMail = oAllMails(1).Body

End Function

Function countMail
	Dim objOutlook 
	Dim objOutlookMsg
	Dim olMailItem
	Dim olFolderInbox, iTotalMails
	
	olFolderInbox = 6 : sSubject = ""
	
	Set objOutlook = CreateObject("Outlook.Application")
	Set objNamespace = objOutlook.GetNamespace("MAPI")
	
	Set objOwner = objNamespace.CreateRecipient(mailbox)

	objOwner.Resolve
	
	'Create reference to Inbox Folder
	Set oInbox = objNamespace.GetSharedDefaultFolder(objOwner, olFolderInbox) 
	
	'Find all items in the Inbox Folder
	Set folder = oInbox.Folders(folder_year).Folders(folder_emails)
	Set oAllMails = folder.Items
	
	countMail = 0
	
	For Each obj In oAllMails 
    		countMail = 1 + countMail 
      Next
	
End Function

Function extractNumbers(mail) 
	'Extract numbers from mail'
	'----------------------------------'
       Dim  folder, tmp, id, pos

       Set regex = CreateObject("vbscript.regexp")
	regex.Global = True
	regex.Pattern = "^\D+"
	mail = regex.Replace(mail, Empty)
	regex.Pattern = "\D+"
	mail = regex.Replace(mail, "-")
	extractNumbers = Split(mail, "-") 
End Function

Function vktoclosAtlas(numbers)
	'MainFrame delete account'
	'----------------------------------'

       Dim tmp
  
	TeWindow("TeWindow").TeScreen("P901PWO ST").TeField("Account ID").Set "h"
	TeWindow("TeWindow").TeScreen("P901PWO ST").TeField("Cust").Set numbers(3)
	TeWindow("TeWindow").TeScreen("P901PWO ST").TeField("Cust").SetCursorPos
	TeWindow("TeWindow").TeScreen("P901PWO ST").SendKey TE_ENTER
	'TeWindow("TeWindow").TeScreen("P901PWO ST").Sync
	TeWindow("TeWindow").TeScreen("P901PWO ST").SendKey TE_ENTER
	TeWindow("TeWindow").TeScreen("P901PWO ST").Sync
	vktoclosAtlas = TeWindow("TeWindow").TeScreen("P901PWO ST").GetText(23, 2, 24, 41)
	pos = InStr(tmp, "Visual")
	If  pos > 0 Then
		TeWindow("TeWindow").TeScreen("P901PWO ST").SendKey TE_ENTER
	End If 
	TeWindow("TeWindow").TeScreen("P901PWO ST").TeField("P901PWO(protected)").SetCursorPos
	TeWindow("TeWindow").TeScreen("P901PWO ST").SendKey TE_PF10
	'TeWindow("TeWindow").TeScreen("P901PWO ST").Sync
	
	flag_atlas = "X"
	flag_menyb = "X"
	
End Function 

Function searchExcel(numbers)
	'Search orgnr in Excel'
	'----------------------------------'

	'Create Excel Object
	Set Myexcel = createobject("excel.application")
	'Open the Excel File
	Set WB = Myexcel.Workbooks.Open(excel_path,true,true)
	'Get the Control on Specific Sheet
	Set searchExcel = WB.Worksheets("Klara kunder").Range("B:B"). Find(numbers(3))
	'Quit from Excel Application
	Myexcel.DisplayAlerts = False
	WB.Close
	Myexcel.Quit
	Set Mysheet=Nothing
	Set WB=Nothing
	Set Myexcel=Nothing
End Function 

Sub sendMail(message)
	'Send mail through outlook'
	'----------------------------------'

	Set objOutlook = CreateObject("Outlook.Application") 
	Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
	objOutlookMsg.To = "_FB.LCI.FX Solutions & Support" 
	' Body of the message
	objOutlookMsg.Subject = subject	
	objOutlookMsg.Body = message
	' Send the message
	objOutlookMsg.Send
End Sub 

Sub sendMailAttachment(message)
	Set objOutlook = CreateObject("Outlook.Application") 
	Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
	objOutlookMsg.To = "_FB.LCI.FX Solutions & Support" 
	' Body of the message
	objOutlookMsg.Subject = "Output logs - Bobby" 	
	objOutlookMsg.Attachments.Add "L:\CRM Retail\CRM Retail\Customer Support Retail\Customer Support Retail\Customer Support Retail - LUS-FXtrade\logs.txt"
	objOutlookMsg.Attachments.Add "L:\CRM Retail\CRM Retail\Customer Support Retail\Customer Support Retail\Customer Support Retail - LUS-FXtrade\config.txt"
	objOutlookMsg.Attachments.Add "L:\CRM Retail\CRM Retail\Customer Support Retail\Customer Support Retail\Customer Support Retail - LUS-FXtrade\MailTemplate.txt"
	objOutlookMsg.Body = message
	' Send the message
	objOutlookMsg.Send
End Sub

Function readMailTemplate() 
	Set objStream = CreateObject("ADODB.Stream")
	objStream.CharSet = "utf-8"
	objStream.Open
	objStream.LoadFromFile("L:\CRM Retail\CRM Retail\Customer Support Retail\Customer Support Retail\Customer Support Retail - LUS-FXtrade\MailTemplate.txt")
	readMailTemplate = objStream.ReadText()
End Function 
	
Sub initialization
	'Assignment of initial values'
	'----------------------------------'
	Dim initial_values(8) 
	Dim temp, value 

	Set FSO = CreateObject("Scripting.FileSystemObject")
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set oFile2 = FSO.OpenTextFile("L:\CRM Retail\CRM Retail\Customer Support Retail\Customer Support Retail\Customer Support Retail - LUS-FXtrade\config.txt", ForReading, True)
	temp = 0
	Do Until oFile2.AtEndOfStream = True
		value = oFile2.ReadLine
		split_string = Split(value, "=") 
		initial_values(temp) = split_string(1)
		temp = temp + 1
	Loop 
	oFile2.Close
	
	excel_path = initial_values(0)
	mailbox =  initial_values(1)
	p_id  = initial_values(2)
       folder_year = initial_values(3)
       folder_emails  = initial_values(4)
	
	Set obj_file =CreateObject("Scripting.FileSystemObject")
       Set file=obj_file.CreateTextFile("L:\CRM Retail\CRM Retail\Customer Support Retail\Customer Support Retail\Customer Support Retail - LUS-FXtrade\logs.txt", True)
	
End Sub

Sub logs(message)
	file.WriteLine message
End Sub

Sub moveMail(map)
	
	Dim objOutlook 
	Dim objOutlookMsg
	Dim olMailItem
	Dim olFolderInbox, iTotalMails
	
	olFolderInbox = 6 : sSubject = ""
	
	Set objOutlook = CreateObject("Outlook.Application")
	Set objNamespace = objOutlook.GetNamespace("MAPI")
	
	Set objOwner = objNamespace.CreateRecipient(mailbox)

	objOwner.Resolve
	
	'Create reference to Inbox Folder
	Set oInbox = objNamespace.GetSharedDefaultFolder(objOwner, olFolderInbox) 
	
	'Find all items in the Inbox Folder
	Set folder = oInbox.Folders(folder_year).Folders(folder_emails)
	Set oAllMails = folder.Items(1) 
	Set department = oInbox.Folders(folder_year).Folders(map) 
	oAllMails.Move department 

End Sub

Sub openAtlas()
SystemUtil.Run("C:\Users\P901PWO\Desktop\Atlas.zws")	
Window("Windows Security").Type "p901pwo"
Window("Windows Security").Type  micReturn
TeWindow("TeWindow").TeScreen("screen21070").TeField("Användarid").Set "p901pwo_"
TeWindow("TeWindow").TeScreen("screen21070").TeField("Lösenord").SetSecure "624e881db5d460d054c83e2a0b5066ca075e07bba65b"
TeWindow("TeWindow").TeScreen("screen21070").TeField("Funktion").Set "VKTOCLOS"
TeWindow("TeWindow").TeScreen("screen21070").TeField("Subbranch").SetCursorPos
TeWindow("TeWindow").TeScreen("screen21070").SendKey TE_ENTER
TeWindow("TeWindow").TeScreen("screen21070").Sync
End Sub 

Sub openMenyB
SystemUtil.Run("C:\Users\Public\Documents\Zephyr\PASSPORT PC TO HOST\Meny-B sso.zws")	
Window("Windows Security").Type "p901pwo"
Window("Windows Security").Type  micReturn 
wait 5
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "t"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "ims"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "7"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "lus590"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "010101"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "510"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "        "
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
End Sub

Sub markMenyB(account)
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type account
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type micShiftDwn +  micTab  + micShiftUp
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "a"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "AVSLUTAD PUA P901PWO"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type "J"
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micF3
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type " "
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micReturn
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab
Window("Swedbank PassTicket").WinObject("C:\Users\Public\Documents\Zeph").Type  micTab 
End Sub

Sub main() 
	' Sekvens av operationer '
	'----------------------------------'

	Dim mail, numbers, findObj
	Dim folder, message, count
	openAtlas()
	openMenyB()
	count = countMail()
	
	For a = 1 To count
	On Error Resume Next
		logs("------------------------------------------------")
		logs(time)
		mail =  importMail()
		numbers = extractNumbers(mail)
		logs(numbers(3))
		
		Set findObj = searchExcel(numbers)
		
		If Not findObj is Nothing Then
			logs("FX-Trade")
			moveMail(". Prio 2 Fx Trade")
		Else
			message = vktoclosAtlas(numbers)
			pos = InStr(message, "Visual")
			If  pos = 0 Then
				logs("Status: " + message +" Map: . Prio 2 Error")
				moveMail(". Prio 2 Error")
			Else
				logs("Status: avslutat" + " Atlas: " + flag_atlas + " MenyB: " + flag_menyb +" Map: . Avslutade valutakonton")
				markMenyB(numbers(3))
				subject = numbers(3) + " " + numbers(1)
				sendmail(readMailTemplate())
				moveMail(". Avslutade valutakonton")
			End If
		End If
	Next
	file.Close
	sendMailAttachment("Output logs / Bobby")
End Sub

'Assignment of initial values'	
initialization()

'The entry point'
main()
