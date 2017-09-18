'main.vbs
'This script is mainly for the main control for file transfering and logging
'created by Gordon
'June and July

Dim wshshell 
Set wshshell = WScript.CreateObject ("WSCript.shell")

Dim ProfileRunTime
ProfileRunTime  =  0

ForWriting = 2
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Testing script
'Dim location
'location = "%comspec% /C C:\""Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client""\sftpg3.exe -B C:\ftp_batch\test.bat"
'wshshell.run location, 1, True

'List declaration
Dim tempFolderN
Dim ListJafterdownload 
Set ListJafterdownload =  CreateObject("System.Collections.ArrayList")
ListJafterdownload.Add("")


'testing
Set ListJafterdownload2 =  CreateObject("System.Collections.ArrayList")
ListJafterdownload2.Add("")

Dim JdownFolder : set JdownFolder = objFSO.GetFolder("C:\\ftp_fromCiti\")


For Each objFileForDownload In JdownFolder.Files 
	
	If objFSO.GetExtensionName(objFileForDownload.Name) = "TXT" Or  objFSO.GetExtensionName(objFileForDownload.Name) = "txt" Then
		
		tempFolderN = objFileForDownload 
		ListJafterdownload2.Add(tempFolderN)
		
	End If
	
Next

location = "%comspec% /K C:\""Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client""\sftpg3.exe -B C:\ftp_batch\downloadFromCiti.bat > C:\ftp_vbs\report.log"
wshshell.run location, 1, True


Set objFileToReadScript = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\ftp_vbs\report.log",1)

	Dim listForCitiTrans
	Set listForCitiTrans = CreateObject("System.Collections.ArrayList")
	listForCitiTrans.Add ("Cannot Null") 
	Dim ReadEnd : ReadEnd = 0 
	
	Dim listForCitiTransName
	Set listForCitiTransName = CreateObject("System.Collections.ArrayList")
	listForCitiTransName.Add("")
	
	
	do while not objFileToReadScript.AtEndOfStream 
		listForCitiTrans.Add(objFileToReadScript.ReadLine())
		ReadEnd = ReadEnd + 1 
	loop

Dim valInit
Dim ValueCheck : ValueCheck = 0 	
	
For i = 0 To listForCitiTrans.Count - 2

If(Instr(listForCitiTrans(i),"mget")<>0 ) Then 
If(i+ValueCheck > 0 And i+ValueCheck <listForCitiTrans.Count ) Then
Do Until Instr(listForCitiTrans(i + ValueCheck),"bye")<>0


Dim successflag

If  Instr(listForCitiTrans(i + ValueCheck) , "100%") Then
successflag = "Successful transfered"

Else
successflag = "transfered fail"

End If

Dim stopsign : stopsign = Instr(listForCitiTrans(i + ValueCheck),"|")
valInit = LEFT(listForCitiTrans(i + ValueCheck), stopsign) 
Dim finalkey : finalkey = valInit + " " + successflag
listForCitiTransName.Add(finalkey)
ValueCheck = ValueCheck + 1 
Loop
End If 
End If

Next


For Each objFileForDownload2 In JdownFolder.Files 
	
	If objFSO.GetExtensionName(objFileForDownload2.Name) = "TXT" Or  objFSO.GetExtensionName(objFileForDownload2.Name) = "txt" Then
		
		tempFolderN = objFileForDownload 
		ListJafterdownload.Add(tempFolderN)
		
	End If
	
Next


Dim toCitiConfol
Dim ListToCitiUpload
Set ListToCitiUpload = CreateObject("System.Collections.ArrayList")
ListToCitiUpload.Add("")
Dim justUpCiti : set justUpCiti = objFSO.GetFolder("C:\\ftp_toCiti\")

Dim from06Confol
Dim Listfrom06
Set Listfrom06 = CreateObject("System.Collections.ArrayList")
Listfrom06.Add("")
Dim justfrom06 : set justfrom06 = objFSO.GetFolder("C:\\ftp_toCiti\")

Dim To06Confol
Dim ListTo06
Set ListTo06 = CreateObject("System.Collections.ArrayList")
Dim justto06 : set justto06 = objFSO.GetFolder("C:\\ftp_fromCiti\")

Dim QAlist
Set QAlist = CreateObject("System.Collections.ArrayList")
QAlist.Add ("") 

Dim Deblist
Set Deblist = CreateObject("System.Collections.ArrayList")
Deblist.Add ("") 

Dim fromCitiList
Set  fromCitiList = CreateObject("System.Collections.ArrayList")
fromCitiList.Add ("") 

Dim toCitiList
Set  toCitiList = CreateObject("System.Collections.ArrayList")
toCitiList.Add ("") 



'Variable declaration
Dim number 
Dim QAmax  
QAmax = 0 
Dim debMax
debMax = 0
Dim tF
Dim ONum   
Dim QATime
Dim DebTime
Dim QAhasMatches
Dim DebhasMatches
Dim FPath
Dim UsrRightLogin
UsrRightLogin = True 
Dim deleteSuccess
deleteSuccess =  0 
Dim updateSuccess 
redim Stack(-1) 'This is data structure called stack
FPath = "C:\\ftp_fromCiti\" 
Dim QAPath
QAPath = "C:\ftp_fromCiti\IPP_AIA_HK_DOWNLOAD_QA_FILE.txt"
Dim debPath
debPath = "C:\ftp_fromCiti\IPP_AIA_HK_DOWNLOAD_DEBIT_FILE.txt"
Dim CitiUATIPCheck : CitiUATIPCheck = 0
Dim zerosixCheck :zerosixCheck = 0 
Dim QAFileExists : QAFileExists = false
Dim DebFileExists : DebFileExists = false
Dim DebCorrectFile 
DebCorrectFile = 0 


'precheck the file whehther exists things 
Dim tempFilList

For Each objF2 In objFSO.GetFolder("C:\\ftp_toCiti\").Files
	
	tempFilList = objF2
	toCitiList.Add(tempFilList)			   
	
Next

For Each objF3 In objFSO.GetFolder("C:\\ftp_fromCiti\").Files
	
	tempFilList = objF3
	fromCitiList.Add(tempFilList)			   
	
Next


Dim temppath
DocAction objFSO.GetFolder(FPath)
Sub DocAction(objFolder)
	Dim objFile
'Renaming file funtion created by gordon on 6/7
	For Each objFile In objFolder.Files
		
		FileN = ""
		FileN = ObjFSO.GetFileName(objFile.Path)
		
'Section A : ignore the naming with english letter at the end by RegEx
		
		number = RIGHT(FileN,12)   
		number = LEFT(number,8)
		
		ONum = RIGHT(tF,12)
		ONum = LEFT(ONum,8)
		
		Set re = New RegExp
		re.Pattern = "[a-z]"
		re.IgnoreCase = True
		re.Global = True
		DebhasMatches = re.Test(RIGHT(number,1))
		QAhasMatches = re.Test(RIGHT(number,1))
		
'section A.2 If find case that not match two case 
		
		If Instr(FileN,"IPP_AIA_HK_DOWNLOAD_DEBIT") = 0 And Instr(FileN,"IPP_AIA_HK_DOWNLOAD_QA") = 0 Then
			temppath = objFile.Path
			Deblist.Add(temppath + "is not either IPP_AIA_HK_DOWNLOAD_DEBIT Or IPP_AIA_HK_DOWNLOAD_QA,so it deleted.") 
			objFSO.DeleteFile(objFile.Path) 
			
		End If	  
			
'Section B :  Find the most updatest file referenced as document ending 's date 
			
			If Instr(FileN,"IPP_AIA_HK_DOWNLOAD_QA") <> 0  Then  'for QA Case
				
				If QAhasmatches = true Then
					temppath = objFile.Path 
					QAlist.Add(temppath + " deleted because of File name is not end with date but string.") 
					objFSO.DeleteFile(objFile.Path)
					QATime = QATime + 1
					QAFileExists = true
					
				ElseIf QAmax>number Then
					temppath = objFile.Path
					QAlist.Add(temppath + " deleted because it 's not the updatest.")
					objFSO.DeleteFile(objFile.Path) 
					QAFileExists = true
					
				ElseIf number>ONum And QATime > 0 Then
					temppath = tF
					QAlist.Add( tF +" deleted because it 's not the updatest.")
					objFSO.DeleteFile(tF)
					tF = FPath+ObjFSO.GetFileName(objFile.Path)
					QAFileExists = true
					
				ElseIf number>QAmax Then
					QAmax = number 
					tF = FPath+ObjFSO.GetFileName(objFile.Path)
					QATime = QATime + 1 
					
				Else
					QAlist.Add("No File Name Checking for QA")
					
				End If
			End If 	
			
			
			If Instr(FileN,"IPP_AIA_HK_DOWNLOAD_DEBIT")<> 0 Then   'for Debit Case
				
				
				If Debhasmatches = true Then
					temppath = objFile.Path
					Deblist.Add(temppath+ " deleted because of File name is not end with date but string.")
					objFSO.DeleteFile(objFile.Path)
					debTime = debTime + 1
					DebFileExists = true
					
				ElseIf debMax>number Then
					temppath = objFile.Path
					Deblist.Add(temppath+ " deleted because it 's not the updatest.")
					objFSO.DeleteFile(objFile.Path) 
					DebFileExists = true
					
				ElseIf number>ONum And debTime > 0 Then
					temppath = tF
					Deblist.Add(tF+ " deleted because it 's not the updatest.")
					objFSO.DeleteFile(tF)
					tF = FPath + ObjFSO.GetFileName(objFile.Path)
					DebFileExists = true
					
				ElseIf number>debMax Then
					debMax = number 
					tF = FPath +ObjFSO.GetFileName(objFile.Path)
					debTime = debTime + 1 
					
				Else
					Deblist.Add ("No File Name Checking for Debit")
					
				End If
				
			End If 
			
			
			
			
		Next
		
'Section C : Changed the ending of QA txt 's naming into standard form as Client request [C:\\ftp_fromCiti\IPP_AIA_HK_DOWNLOAD_QA.txt]
		
		Dim objF
		For Each objF In objFolder.Files
			
			FileN = ""
			FileN = ObjFSO.GetFileName(objF.Path)
			
			If Instr(FileN,"IPP_AIA_HK_DOWNLOAD_QA") <> 0  Then
				
				QAlist.Add(objF.Path + " had already renamed into client 's standard ")            
				
				If objFSO.FileExists(QAPath) Then
					objFSO.DeleteFile(QAPath)
				End If 
				objFSO.MoveFile objF.Path, QAPath
			End If
			
			If Instr(FileN,"IPP_AIA_HK_DOWNLOAD_DEBIT") <> 0  Then                
				Deblist.Add(objF.Path + " is the most updatest file and fulfilled the standard from client.")
			End If
		Next
		
	End Sub
	
	
'UPLOAD TO hkgabqwfis01 (QA & Debit)
	location = "%comspec% /K C:\""Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client""\sftpg3.exe --verbose -B C:\ftp_batch\uploadToH06.bat > C:\ftp_vbs\report2.log "
	wshshell.run location, 1, True 
	
'self log check from system by Gordon 07282016
	
	Set objFileToReadScript = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\ftp_vbs\report.log",1)
	
	Dim listFor06Trans
	Set listFor06Trans = CreateObject("System.Collections.ArrayList")
	listFor06Trans.Add ("") 
	
	Dim listFor06TransName
	Set listFor06TransName = CreateObject("System.Collections.ArrayList")
	listFor06TransName.Add("")
	
	do while not objFileToReadScript.AtEndOfStream 
		listFor06Trans.Add(objFileToReadScript.ReadLine())
	loop

Dim valInit4
Dim ValueCheck4 : ValueCheck4 = 0 	
	
For i = 1 To listFor06Trans.Count - 1 

If(Instr(listFor06Trans(i),"put")<>0 ) Then 
If( (i+ValueCheck4) > 0 And (i+ValueCheck4) < (listFor06Trans.Count-1) ) Then

Do Until Instr(listFor06Trans(i + ValueCheck4),"Quit Now")<>0 

' Dim successflag4
' If Instr(listFor06Trans(i + ValueCheck4) , "100%") <> 0 And Instr(listFor06Trans(i + ValueCheck4 + 1 ) , "skipped") = 0 Then
' successflag4 = "Success transfered"
' Else
' successflag4 = "transfered fail"
' End If

WScript.Echo listFor06Trans(i)

Dim stopsign4 : stopsign4 = Instr(listFor06Trans(i + ValueCheck4),"|")
valInit4 = LEFT(listFor06Trans(i + ValueCheck4), stopsign4) 
Dim finalkey4: finalkey4 = valInit4 

WScript.Echo finalkey4

listFor06TransName.Add(finalkey4)
ValueCheck4 = ValueCheck4 + 1 

Loop
 
End If 
End If

Next



 For Each objFileTo06 In justto06.Files 
	
	If objFSO.GetExtensionName(objFileTo06.Name) = "TXT" Or objFSO.GetExtensionName(objFileTo06.Name) = "txt" Then
		
		tempFolderN = objFileTo06 
		ListTo06.Add(tempFolderN)
		
	End If
	
 Next

	
	
	
'DOWNLOAD FROM hkgabqwfis01 (Policy)
	location = "%comspec% /K C:\""Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client""\sftpg3.exe --verbose -B C:\ftp_batch\policy_downloadFromH06.bat > C:\ftp_vbs\report.log"
	wshshell.run location, 1, True
	
	'self log check from system by Gordon 07282016
	
Set objFileToReadScript = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\ftp_vbs\report.log",1)
	
	Dim listFor06Trans2
	Set listFor06Trans2 = CreateObject("System.Collections.ArrayList")
	listFor06Trans2.Add ("") 
	
	Dim listFor06TransName2
	Set listFor06TransName2 = CreateObject("System.Collections.ArrayList")
	'listFor06TransName2.Add("")
	
	do while not objFileToReadScript.AtEndOfStream 
		listFor06Trans2.Add(objFileToReadScript.ReadLine())
	loop

	
Dim valInit3
Dim ValueCheck3 : ValueCheck3 = 0 	
	
For i = 1 To listFor06Trans2.Count - 1 

If(Instr(listFor06Trans2(i),"mget")<>0 ) Then 

If( (i+ValueCheck3) > 0 And (i+ValueCheck3) < (listFor06Trans2.Count-1) ) Then
Do Until Instr(listFor06Trans2(i + ValueCheck3),"Quit Now")<>0 And Instr(listFor06Trans2(i + ValueCheck3),"Echo")<>0

Dim stopsign3 : stopsign3 = Instr(listFor06Trans2(i + ValueCheck3),".TXT")
valInit3 = LEFT(listFor06Trans2(i + ValueCheck3), stopsign3) 

Dim finalkey3: finalkey3 = valInit3
listFor06TransName2.Add(finalkey3)
ValueCheck3 = ValueCheck3 + 1 

Loop
 
End If 
End If

Next

	For Each objFileFor06 In justfrom06.Files 
	
	If objFSO.GetExtensionName(objFileFor06.Name) = "TXT" Or objFSO.GetExtensionName(objFileFor06.Name) = "txt" Then
		
		tempFolderN = objFileFor06 
		Listfrom06.Add(tempFolderN)
		
	End If
	
    Next
	
	Dim listuseforsuccess06
	Set listuseforsuccess06 = CreateObject("System.Collections.ArrayList")
	Dim listuseforfail06 : set listuseforfail06 = CreateObject("System.Collections.ArrayList")
	
	listuseforsuccess06.Add("")
	listuseforfail06.Add("")

		Dim sizeQ2 
		Dim size1 : size1 = listFor06TransName2.Count - 1 
		Dim CheckGotten1
		Dim variableTrans1
		Dim word2
					
		For j = 1 To listFor06TransName2.Count -  1 
		word2  = "C:\\ftp_toCiti\"+listFor06TransName2(j) + "TXT"
		
		If Objfso.FileExists(word2) Then
		listuseforsuccess06.Add(Objfso.getFileName(word2) + " | transfered success")
		from06 = from06 + 1
		Else
		listuseforfail06.Add( word2 +" | transfer failure")
        End If	
		
		Next
					


	
'THAT'S FOR UPLOADING TO CITI (Policy)

location = "%comspec% /K C:\""Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client""\sftpg3.exe --verbose -B C:\ftp_batch\policy_uploadToCiti.bat > C:\ftp_vbs\report.log"


'self log check from system by Gordon 07282016

wshshell.run location, 1, True
	
	Set objFileToReadScript = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\ftp_vbs\report.log",1)
	
	Dim listForCitiTrans2
	Set listForCitiTrans2 = CreateObject("System.Collections.ArrayList")
	listForCitiTrans2.Add ("") 
	
	Dim listForCitiTransName2
	Set listForCitiTransName2 = CreateObject("System.Collections.ArrayList")
	listForCitiTransName2.Add("")
	
	do while not objFileToReadScript.AtEndOfStream 
		listForCitiTrans2.Add(objFileToReadScript.ReadLine())
	loop

	
Dim valInit2
Dim ValueCheck2 : ValueCheck2 = 0 	
	
For i = 1 To listForCitiTrans2.Count - 1 

If(Instr(listForCitiTrans2(i),"mput")<>0 ) Then 

If( (i+ValueCheck2) > 0 And (i+ValueCheck2) < (listForCitiTrans2.Count-1) ) Then
Do Until Instr(listForCitiTrans2(i + ValueCheck2),"Quit Now")<>0 And Instr(listForCitiTrans2(i + ValueCheck2),"Echo")<>0

Dim successflag2

If Instr(listForCitiTrans2(i + ValueCheck2) , "100%") <> 0 And Instr(listForCitiTrans2(i + ValueCheck2 + 1 ) , "skipped") = 0 Then
successflag2 = "Successful transfered"
Else
successflag2 = "transfered fail"
End If

Dim stopsign2 : stopsign2 = Instr(listForCitiTrans2(i + ValueCheck2),"|")
valInit2 = LEFT(listForCitiTrans2(i + ValueCheck2), stopsign2) 

Dim finalkey2: finalkey2 = valInit2 + " " + successflag2

listForCitiTransName2.Add(finalkey2)
ValueCheck2 = ValueCheck2 + 1 

Loop
 
End If 
End If

Next

For Each objFileForUpload In justUpCiti.Files 
	
If objFSO.GetExtensionName(objFileForUpload.Name) = "TXT" Or objFSO.GetExtensionName(objFileForUpload.Name) =  "txt" Then
		
tempFolderN = objFileForUpload 
ListToCitiUpload.Add( ObjFSO.GetFileName(tempFolderN) )
		
End If
	
 Next

 
'Logging of FTP Transfer System
'created by Gordon 7/12
'because of the server script may not totally fulfill the requirement for system
	
	Dim fsoForWriting 
	fsoForWriting = 8
	
	Dim objTime
	Set objTime = CreateObject("Scripting.FileSystemObject")
	
	
'Open the text file
	Dim objTextStream
	Set objTextStream = objTime.OpenTextFile("C:\\ftp_vbs\clientlog.txt", fsoForWriting, True)
	objTextStream.WriteLine "*********************************************"
	
	
'Log Merging by Gordon 07132016
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\ftp_vbs\event.txt",1)
	Dim strLine
	Dim listime
	Dim list
	Set list = CreateObject("System.Collections.ArrayList")
	list.Add ("Cannot Null") 
	
	
	Dim CopyFilelist
	Set CopyFilelist = CreateObject("System.Collections.ArrayList")
	CopyFilelist.Add ("Cannot Null") 
	
	
	do while not objFileToRead.AtEndOfStream 
		list.Add(objFileToRead.ReadLine())
		listime = listime + 1 
	loop
	
	Dim updateno
	Dim signal 
	nwerrorsignal = 0
	
	Dim loginName 
	loginName = "" 
	
	objTextStream.WriteLine "main.vbs run on  : " + FormatDateTime(now, 1) + " " + FormatDateTime(now, 4)
	
	Dim CitiNo : CitiNo = 0
	Dim habNo : habNo = 0 
	
	Dim Debexists
	Dim numoffile
	Dim QACorrectFile
	Dim DebFile
	Dim existsfinishedFlag
	existsfinishedFlag = 1 
	Dim stat 
	Dim copiedfinishString
	Dim deletedfinishedString
	Dim neatestPortID
	Dim trackProblemScript
	Dim signalExec
	Dim failName 
	Dim ENFlag
	ENFlag = true
	Dim ENFlag2 
	ENFlag2 = true
	Dim scriptName 
	Dim CheckEntry 
	Dim CheckStandNo
	CheckStandNo = 0 
	Dim AlgoName : AlgoName =""
	Dim FailLogName : FailLogName =""
	Dim ExistName : ExistName =""
	
	Dim globbinged : globbinged = false
	
	Dim requiredlistName : requiredlistName = "" 
	
	Dim from06
	Dim tocitibank


	Dim QuickCheckfor06 : QuickCheckfor06 = true
	Dim QuickCheck : QuickCheck = true
	
	For updateno = listime To 0 Step -1
		
		If Instr(list(updateno),"WHOLE FINISHED") <> 0 Then
			existsfinishedFlag = existsfinishedFlag + 1
		End If
		
'Detect Login Problem And do file transfer issue 
		
		
	If Instr(list(updateno),"B-3") <> 0 And existsfinishedFlag = 1  Then 'check success
			For AlgoNum =  10 To 0 Step -1 
				If( updateno - AlgoNum  < listime And updateno - AlgoNum  > 0 ) then 
				If Instr(list(updateno - AlgoNum ),"open") <> 0 And existsfinishedFlag = 1 And Instr(list(updateno),"bye") = 0 And Instr(list(updateno),"quit") = 0  And ( Instr(list(updateno - AlgoNum ),"CitiUATIP56S") <> 0  Or Instr(list(updateno - AlgoNum ),"hkgab") <> 0 )Then
					requiredlistName = list(updateno - AlgoNum) 
					
					Dim AlgoPos : AlgoPos = list(updateno - AlgoNum)
					AlgoName = Right(AlgoPos,(Len(AlgoPos)-Instr(AlgoPos,"username:"))+1 )
					Dim ReAlgo : ReAlgo = ((Len(AlgoPos)-Instr(AlgoPos,"username:"))+1) 
					
					AlgoName = LEFT(AlgoName,(Instr(AlgoName,"Program")-3))
					
					scriptName = Right(list(updateno - AlgoNum),Len(list(updateno- AlgoNum ))-Instr(list(updateno - AlgoNum ),"open") - 3 )
					scriptName = LEFT(scriptName,Len(scriptName) - 1 ) 
					
					If Instr(requiredlistName,"CitiUAT") <> 0 Then 
						
						CitiUATIPCheck = CitiUATIPCheck + 1

						For Serveri = 0 To 10 
						
						If(updateno + Serveri < listime And updateno + Serveri > 0 ) Then
						If Instr(list(updateno + Serveri ),"Server's hostkey not accepted") Then
						push Stack,"The server login fail due to hostkey not matched"
						QuickCheck = false
						End If 
						End If
						
						
						Next
						
						If CitiUATIPCheck = 1 And QuickCheck = true Then
							
							push Stack,""
							push Stack,""
							push Stack,""
							push Stack,"End "+scriptName 
							
				    'List checking as verbose appoarch by Gordon 07282016
							
				    Dim sizeQ 
					Dim size : size = ListToCitiUpload.Count - 1 
					
					Dim CheckGotten
					Dim variableTrans
					
					
					If Objfso.FileExists("C:\\ftp_toCiti\success\") Then
					ObjFSO.deleteFolder "C:\\ftp_toCiti\success\"
					End If 
					
					If Objfso.FileExists("C:\\ftp_toCiti\failure") Then
					ObjFSO.deleteFolder "C:\\ftp_toCiti\failure"
					End If 
					
					If not Objfso.FileExists("C:\\ftp_toCiti\success\") Then
                    Objfso.CreateFolder "C:\\ftp_toCiti\success\"
                    End If

                    If not Objfso.FileExists("C:\\ftp_toCiti\failure") Then
                    Objfso.CreateFolder "C:\\ftp_toCiti\failure"
                    End If

					For j = 0 To listForCitiTransName2.Count -  1 
					
					
					If Instr (listForCitiTransName2(j) ,"TXT" )<> 0 Then
					sizeQ = (Instr(listForCitiTransName2(j),"TXT")) 
					variableTrans = LEFT(listForCitiTransName2(j),sizeQ) 
				    Dim word : word  = "C:\\ftp_toCiti\"+variableTrans + "XT"
					
					If Objfso.FileExists(word) Then
                    ObjFSO.MoveFile word,"C:\\ftp_toCiti\success\"
                    End If
					
					End If 
					
					Next
					
					Dim folderForSuccess 
                    set folderForSuccess = Objfso.GetFolder("C:\\ftp_toCiti\success\")

                    Dim folder2ForSuccess
                    set folder2ForSuccess = Objfso.GetFolder("C:\\ftp_toCiti\failure\")
					
					Dim folder3ForSuccess
                    set folder3ForSuccess = Objfso.GetFolder("C:\\ftp_toCiti\")

                    Dim CountFile 
                    Dim overall :overall = 0 
					
					For Each num2 In folder3ForSuccess.Files
					overall = overall + 1 
                    Next					
					
                    For Each num In folderForSuccess.Files
                    CountFile = CountFile + 1
					toCitibank = toCitibank + 1
                    push Stack,Objfso.getFileName(num)+ " | transfered success"
                    Next
					
                    If overall > 0 Then
                    ObjFso.MoveFile "C:\\ftp_toCiti\*.txt","C:\\ftp_toCiti\failure\"
                    End If


                    For Each num in folder2ForSuccess.Files
                    push Stack,Objfso.getFileName(num)+ " | transfered failure"
                    Next

                    Dim strFolder
                    dtmValue = Now()
                    strFolder = "C:\\backup\"& Month(dtmValue) & "_" & Day(dtmValue) & "_" & Hour(dtmValue) &"_" & Second(dtmValue)&"_toCiti_backup"
                    objFSO.CreateFolder strFolder

					
                    objFSO.MoveFolder "C:\\ftp_toCiti\*",strFolder
					
					

				' For Each objFileStar In justfrom06.Files 
					   ' Set re = New RegExp
		                ' re.Pattern = "[a-z]"
		                ' re.IgnoreCase = True
		                ' re.Global = True
					   '''''''''''''''''''''''''''''''''''''
					   
					   ' For j = 0 To listForCitiTransName2.Count -  1 
					  
					   
					   ' If Instr(listForCitiTransName2(j),objFileStar) <>  0  Then
					     ' WScript.Echo " I success" + objFileStar
					   	 ' push Stack,objFileStar + "| Success transfered"
                         ' tocitibank = tocitibank + 1 
						 ' CheckGotten = True
						 ' Exit For 
					   ' End If 
					   
					    ' j=j+1
						
					' Next
					  
					  ' If CheckGotten = false Then 	
					  ' push Stack, objFileStar + "| failed to transfer"
					  ' End If
	
                ' Next
					
					   ' For i = 0 To size
					   
                            ' CheckGotten = false
							
						' For j = 0 To listForCitiTransName2.Count -1
						
							'WScript.Echo listForCitiTransName2(j)
							' Set re = New RegExp
		                    ' re.Pattern = "[a-z]"
		                    ' re.IgnoreCase = True
		                    ' re.Global = True
							
							 ' Dim hasMatchesforCiti2 : hasMatchesforCiti2 = re.Test(LEFT(listForCitiTransName2(j),1))
							 ' hasMatchesforCiti2 = true
							' If hasMatchesforCiti2 = True Then
								
						'If listForCitiTransName2(j) <>Null Then
								' sizeQ = (Instr(listForCitiTransName2(j),"TXT")) 
								

								' If ListToCitiUpload(i) <> Null then
						        ' variableTrans = LEFT(listForCitiTransName2(j),sizeQ) 
								' variableTrans = LEFT(listForCitiTransName2(j),Len(variableTrans) -  1 ) 
								' End If 
								
			                    ' Dim quickvarfast : quickvarfast = variableTrans
								
							    ' If Instr(ListToCitiUpload(i),quickvarfast) <>  0  Then
								' WScript.Echo ListToCitiUpload(i)
								' push Stack,ListToCitiUpload(i) + "| Success transfered"
                                ' tocitibank = tocitibank + 1 
								' CheckGotten = True
						        ' Exit For
								

								' End If
								
								' j = j + 1 
								
                        ' Next 	
						
								' If CheckGotten = false Then 	
								' push Stack, ListToCitiUpload(i) + "| failed to transfer"
								' End If
								' i = i + 1 
	   
                      ' Next
							
							' For j = 1 To ListToCitiUpload.Count - 1 
		                    ' push Stack,ObjFSO.GetFileName(ListToCitiUpload(j)) & " transfered to CitiUATIP16 from ToCiti success. "  
							' tocitibank = tocitibank + 1 
                            ' Next
							
							If tocitibank = 0 Then
								push Stack,"No file transfer to CitiUATIP56 from toCiti"
							End If 
							
							push Stack,"Remote locate in "+scriptName+" for file transfer from toCiti To CITIUATIP56"
					
					
					
						    ElseIf CitiUATIPCheck = 2 And QuickCheck = true Then
							
							push Stack,""
							push Stack,""
							push Stack,""
							push Stack,"End "+scriptName 
							
							'checking whether the value match with file transfered
							
							Set re = New RegExp
		                    re.Pattern = "[a-z]"
		                    re.IgnoreCase = True
		                    re.Global = True

                             For i = 0 To listForCitiTransName.Count - 1 

                             If Instr(objFSO.GetExtensionName(listForCitiTransName(i)),"TXT") <> 0 Then
							 
							   Dim hasMatchesforCiti : hasMatchesforCiti = re.Test(LEFT(listForCitiTransName(i),1))
							   
							   If hasMatchesforCiti = True Then
							   push Stack,listForCitiTransName(i)
                               CitiNo = CitiNo + 1 
							   End If 
							   
							 End If 
							 
							Next
                            
							' For i = 1 To ListJafterdownload.Count - 1
								' push Stack,ObjFSO.GetFileName(ListJafterdownload(i)) &"transfered from CitiUATIP16 to FromCiti success."
								' CitiNo = CitiNo + 1 
							' Next

							
							If CitiNo = 0 Then
								push Stack,"No file in the fromCitiFolder which transfered from fromCiti"
							End If 
							
							push Stack,"Remote locate in "+scriptName 
						End If 
						
					End If 
					
					If Instr(requiredlistName,"hkgabqwfis01")<>0 Then
					
						zerosixCheck = zerosixCheck + 1
						WScript.Echo zerosixCheck
						
						
						For i = 0 To 10 
						
						If Instr(list(updateno + i ),"Server's hostkey not accepted") Then
						push Stack,"The server login fail due to hostkey not matched"
						QuickCheckfor06 = false
						End If 
						
						Next
						
						If zerosixCheck = 1 And QuickCheckfor06 = true Then
							
							push Stack,""
							push Stack,""
							push Stack,""
							push Stack,"End "+scriptName
							
							Set re = New RegExp
		                    re.Pattern = "[a-z]"
		                    re.IgnoreCase = True
		                    re.Global = True
							
					For i = 0 To listuseforsuccess06.Count -1 
                     push Stack,listuseforsuccess06(i)
					Next
					
					For j = 0 To listuseforfail06.Count - 1 
					 push Stack,listuseforfail06(j)
					Next
					
					
							' If hasmatchesfor06 = True Then 
							' push Stack,listFor06TransName2(q)
                            ' from06 = from06 + 1     							
							' End If
							
							' Next

							' For q = 1 To Listfrom06.Count - 1
								' push Stack,ObjFSO.GetFileName(Listfrom06(q)) & " transfered from hkgabqwfis01 To ToCiti success."
								' from06 = from06 + 1 
							' Next
							
				    If from06 = 0 Then
					push Stack,"No file in the ToCiti transfered from H06"
					End If 
							
							
					push Stack,"Remote locate in "+scriptName 
							

							
					ElseIf zerosixCheck = 2 And QuickCheckfor06 = true Then
							push Stack,""
							push Stack,""
							push Stack,""
							push Stack,"End "+scriptName
								
							' For g = 0 To listFor06TransName.Count - 1 
							   
							   ' Dim hasMatchesfor062 : hasMatchesfor062 = re.Test(LEFT(listFor06TransName(g),1))
							   ' If hasMatchesfor062 = True Then
							   

							   ' push Stack,listFor06TransName(g)
							   ' habNo = habNo + 1 
							   ' End If 
							   
							' Next
		
					'Dim sizeQ 
					'Dim size : size = ListToCitiUpload.Count - 1 
					
					'Dim CheckGotten
					'Dim variableTrans
					
					
					If Objfso.FileExists("C:\\ftp_fromCiti\success\") Then
					ObjFSO.deleteFolder "C:\\ftp_fromCiti\success\"
					End If 
					
					If Objfso.FileExists("C:\\ftp_fromCiti\failure") Then
					ObjFSO.deleteFolder "C:\\ftp_fromCiti\failure"
					End If 
					
					If not Objfso.FileExists("C:\\ftp_fromCiti\success\") Then
                    Objfso.CreateFolder "C:\\ftp_fromCiti\success\"
                    End If

                    If not Objfso.FileExists("C:\\ftp_fromCiti\failure") Then
                    Objfso.CreateFolder "C:\\ftp_fromCiti\failure"
                    End If

					For j = 0 To listFor06TransName.Count -  1 
					
					If Instr (listFor06TransName(j) ,"TXT" )<> 0 Then
					sizeQ = (Instr(listFor06TransName(j),"TXT")) 
					variableTrans = LEFT(listFor06TransName(j),sizeQ) 
				    Dim word3 : word3  = "C:\\ftp_toCiti\"+variableTrans + "XT"
					
					If Objfso.FileExists(word3) Then
					WScript.Echo word3
                    ObjFSO.MoveFile word3,"C:\\ftp_toCiti\success\"
                    End If
					
					End If 
					
					Next
					
					Dim folderForSuccess06 
                    set folderForSuccess06 = Objfso.GetFolder("C:\\ftp_fromCiti\success\")

                    Dim folder2ForSuccess06
                    set folder2ForSuccess06 = Objfso.GetFolder("C:\\ftp_fromCiti\failure\")
					
					Dim folder3ForSuccess06
                    set folder3ForSuccess06 = Objfso.GetFolder("C:\\ftp_fromCiti\")

                    Dim CountFile2 
                    Dim overall2 :overall2 = 0 
					
					For Each num2 In folder3ForSuccess.Files
					overall2 = overall2 + 1 
                    Next					
					
                    For Each num In folderForSuccess06.Files
                    CountFile2 = CountFile2 + 1
					toCitibank = toCitibank + 1
					habNo = habNo + 1  
                    push Stack,Objfso.getFileName(num)+ " | transfered success"
                    Next
					
                    If overall2 > 0 Then
                    ObjFso.MoveFile "C:\\ftp_toCiti\*.txt","C:\\ftp_fromCiti\Citi\failure\"
                    End If

                    For Each num in folder2ForSuccess06.Files
                    push Stack,Objfso.getFileName(num)+ " | transfered failure"
                    Next

                    Dim strFolder3
                    dtmValue = Now()
                    strFolder3 = "C:\\backup\"& Month(dtmValue) & "_" & Day(dtmValue) & "_" & Hour(dtmValue) &"_" & Second(dtmValue)&"_fromCiti_backup"
                    objFSO.CreateFolder strFolder3
                    objFSO.MoveFolder "C:\\ftp_fromCiti\*",strFolder3
							
							' For g = 0 To ListTo06.Count - 1 
								' push Stack,ObjFSO.GetFileName(ListTo06(g))&" transfered from fromCiti to hkgabqwfis01 success."
								' habNo = habNo + 1 
							' Next
														
							If habNo = 0 Then
								push Stack,"No file transfered to hkgabqwfis01 from fromCiti"
							End If 
							
							push Stack,"Remote locate in "+scriptName 
							
							End If			 
							
						End If
						
						push Stack, AlgoName&" had already login in "& scriptName 
						
					End If 
				End If
			Next
			
		End If 
		
		
'Broker client disconnect CASE
		If Instr(list(updateno),"Broker_client_disconnect") <> 0 And existsfinishedFlag = 1 Then
		
			For Numb = 0 To 10
				If( updateno - Numb  < listime And updateno - Numb > 0 ) then 'check NULL and error file case
				If Instr(list(updateno-Numb),"open") <> 0  And existsfinishedFlag = 1  And Instr(list(updateno-Numb),"CitiUATIP56S") = 0 And Instr(list(updateno-Numb),"hkgabqwfis01") = 0 Then 
					FailLogName = Right(list(updateno - Numb ),Len(list(updateno - Numb ))-Instr(list(updateno - Numb ),"username:")+1 )
					FailLogName = LEFT(FailLogName,(Len(list(updateno- Numb ))-Instr(list(updateno - Numb ),"username:")) - Instr(list(updateno - Numb ),"Program"))
					scriptName = Right(list(updateno - Numb ),Len(list(updateno - Numb ))-Instr(list(updateno - Numb ),"open") - 3 )
					scriptName = LEFT(scriptName,Len(scriptName) - 1) 
					
					If Instr(list(updateno-Numb),"hkgabqwfis01") = 0 And Instr(list(updateno-Numb),"CitiUATIP56S") = 0 Then
						push Stack, FailLogName&" fail to login in "& scriptName +  " because of the wrong profile Name."
						
					Else
						push Stack, FailLogName&" fail to login in "& scriptName +  " because of user name wrong or they don 't have right to access" 
						
					End If
					
					
					If Instr(list(updateno-Numb),"hkgabqwfis01") <>0  Or Instr(list(updateno-Numb),"CitiUATIP56S")<>0  Then 
						
						push Stack, FailLogName&" fail to login in "& scriptName +  "because of Profile " + scriptName + " does not exists."
						
						If Instr(list(updateno-Numb),"hkgabqwfis01")<>0 Then
							zerosixCheck = zerosixCheck + 1 
						End If
						
						If Instr(list(updateno-Numb),"CitiUATIP16")<>0 Then
							CitiUATIPCheck = CitiUATIPCheck + 1
						End If
						
					End If
					
					CheckEntry = scriptName
				End If
			End If
		Next
	End If
	
	
	
	If Instr(list(updateno),"File Downloaded") <> 0 And existsfinishedFlag = 1 Then
		
		If CheckStandNo = 0 And existsfinishedFlag = 1  Then  
			
			Dim objFTemp : objFTemp = ""
			
			For Each objF In objFSO.GetFolder(FPath).Files
				
                'file checking section
				
				If Instr(objF,QAPath) Then
					QACorrectFile = QACorrectFile + 1
					QAnumOfFile = QAnumOfFile + 1 
				End If
				
				
				Dim rightcheck
				If Instr(objF,"ftp_fromCiti\IPP_AIA_HK_DOWNLOAD_DEBIT_FILE") Then
					rightcheck= RIGHT(objF,7)   
					rightcheck= LEFT(rightcheck,2)
					If RIGHT(rightcheck,1) >= 0 And RIGHT(rightcheck,1) <= 9 Then 
						DebCorrectFile = DebCorrectFile + 1
					End If
					
				End If
				
				WScript.Echo objF 
				WScript.Echo QACorrectFile
				
				numOfFile = numOfFile + 1 
				
			Next
			
			push Stack,"%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
			
			push Stack,"" 
			
			
			If numOfFile = 0 Then  
				push Stack,"Both QA and Debit file does not exists or both files are not fulfill standard thurs the wrong standard file would automatic deleted"
			End If   
			If numofFile > 0 Then    
				If QACorrectFile = 1 Then 
					push Stack,"QA File matches the requirement as assigned"
				ElseIf QACorrectFile = 0 Then
					push Stack,"There are not exists any QA File Or the file not fulfill standard thurs the wrong standard would automatic deleted"
				End If
				
				If DebCorrectFile = 1 Then 
					push Stack,"Debit File matches the requirement as assigned"
				ElseIf DebCorrectFile = 0 Then
					push Stack,"There are not exists any Deb File Or the file not fulfill standard thurs the wrong standard would automatic deleted"
				End If
				
			End If
			
			push Stack,"Conclusion : " 
			push Stack,"" 
			
			For QA = 0 To QAlist.Count - 1 
				push Stack, ObjFSO.GetFileName(QAlist(QA))
			Next 

			For Deb = 0 To Deblist.Count - 1 
				push Stack, ObjFSO.GetFileName(Deblist(Deb))
			Next  
			
			
			push Stack,"Here are the checking log : " 
			push Stack,"Checking the file whether fulfill standard ... "
			
			
			push Stack,"%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
			
			CheckStandNo = CheckStandNo + 1 
			
		End If
		
	End If
	
	
'That 's the case for determine broken client disconnect
	If Instr(list(updateno),"Broker_client_disconnect") <> 0 And existsfinishedFlag = 1 Then 
		
		If Instr(list(updateno - 1),"3017") <> 0 And Instr(list(updateno - 1),"ssh-tectia-configurator") <> 0 And existsfinishedFlag = 1 Then
			push Stack,"No file had copied."
		End If 
		
	End If
	
	If Instr(list(updateno),"Globing pattern") <> 0 And existsfinishedFlag = 1 Then
		globbinged = true
	End If
	
'Str Copy File finished Case
	
	If Instr(list(updateno),"Sftc_copy_file_finished") <> 0 And Instr(list(updateno),"SUCCESS") <> 0 And existsfinishedFlag = 1 Then 
		updateSuccess = true
	End If
	
	
Next


objFileToRead.Close

For i = 0 To getUBound(Stack)
	objTextStream.WriteLine pop(Stack)
Next 


If updateSuccess = false Then
	objTextStream.WriteLine"No file had copied."
End If

wshshell.run "%comspec% /K C:\""Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client""\sftpg3.exe -B C:\ftp_vbs\finalCleanUp.bat", 1, True


'Appoarch 1 : Backup  
Dim strFolderCiti,strFolder2 


dtmValue = Now()

strFolderCiti = "C:\\"& Month(dtmValue) & "_" & Day(dtmValue) & "_" & Hour(dtmValue) &"_" & Minute(dtmValue)&"_toCiti_backup"
strFolder2 = "C:\\"& Month(dtmValue) & "_" & Day(dtmValue) & "_" & Hour(dtmValue) &"_" & Minute(dtmValue)&"_fromCiti_backup"

Dim folder 
set folder = objFSO.GetFolder("C:\\ftp_toCiti\")

Dim ToCitiCount : ToCitiCount =  0 


Dim downFolder : set downFolder = objFSO.GetFolder("C:\\ftp_fromCiti\")

Dim FromCitiCount2 : FromCitiCount2 = 0 

For Each objFileForCiti In folder.Files 
	
	If objFSO.GetExtensionName(objFileForCiti.Name) = "TXT" Or objFSO.GetExtensionName(objFileForCiti.Name) = "txt" Or objFSO.GetExtensionName(objFileForCiti.Name) = "old" Then
		
		ToCitiCount = ToCitiCount + 1 
		WScript.Echo ToCitiCount
	End If
	
Next


For Each objFile2ForOther In downFolder.Files 
	
	If objFSO.GetExtensionName(objFile2ForOther.Name) = "TXT" Or objFSO.GetExtensionName(objFile2ForOther.Name) = "txt" Or objFSO.GetExtensionName(objFile2ForOther.Name) = "old" Then
		
		FromCitiCount2 = FromCitiCount2 + 1 
		
	End If
	
Next

If ToCitiCount<>0 Then
	objFSO.CreateFolder strFolderCiti
	objFSO.MoveFile "C:\\ftp_toCiti\*.TXT" , strFolderCiti
	set newfilepos = objFSO.GetFolder(strFolderCiti)
	newfilepos.Move ("C:\\update\")
	
ElseIf ToCitiCount = 0  Then
End If


If FromCitiCount2<>0 Then
	objFSO.CreateFolder strFolder2
	objFSO.MoveFile "C:\\ftp_fromCiti\*.TXT" , strFolder2
	set newfilepos2 = objFSO.GetFolder(strFolder2)
	newfilepos2.Move ("C:\\update\")
	
ElseIf FromCitiCount2 = 0 Then 
End If

'Appoarch 2 : delete
'DeleteFileWriting
'by Gordon 07182016
'Issue : because file would deleted so that I cannot put the signal at the end of file.
'***remember to add back ldelete * in the bat otherwise nothing would happen***

' Dim deleteflag
' deleteflag = 1 
' Dim deleteSu 
' deleteSu = 0

' Dim deleteRun
'deleteFile Section because of most file are keep tracked until successflag assigned
' Dim objDelete
' Set objDelete = CreateObject("Scripting.FileSystemObject")
' Dim deletelist
' Set deletelist = CreateObject("System.Collections.ArrayList")
' deletelist.Add ("Cannot Null") 


'Open the text file
' Set objTextStream = objDelete.OpenTextFile("C:\\ftp_vbs\clientlog.txt", fsoForWriting, True)
' Set objDelete = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\ftp_vbs\event.txt",1)

' do while not objDelete.AtEndOfStream 
' deletelist.Add(objDelete.ReadLine())
' deleteRun = deleteRun + 1 
' loop


' For updateno = deleteRun To 0 Step -1
' If Instr(deletelist(updateno),"Delete End") <> 0 Then
' deleteflag = deleteflag + 1
' End If

' If Instr(deletelist(updateno),"Sftc_delete_file") <> 0 And Instr(deletelist(updateno),"SUCCESS") <> 0 And deleteflag = 2 Then 
' stat = RIGHT(deletelist(updateno),Len(deletelist(updateno)) - Instr(deletelist(updateno),"File")+3)
' objTextStream.WriteLine"this file deleted : "&stat
' deleteSu = deleteSu + 1 
' End If

' Next 
'objDelete.Close

If deleteSu =  0 Then
	objTextStream.WriteLine"The file had already backuped and all backup file transfer to C:\\update\"
End If 


objTextStream.WriteLine "*********************************************"
objTextStream.WriteLine""
objTextStream.WriteLine""
objTextStream.WriteLine""
objTextStream.WriteLine""
objTextStream.Close


'====================================================================
'Stack Implementation
'Implementated by Gordon 07142016
'referenced from Francis de la Cerna 
'link : https://gallery.technet.microsoft.com/scriptcenter/c05af93f-1213-4238-9c96-6218141bf66d#content
'====================================================================

'==================================================================== 
'  getUbound(arr) 
'  returns the Ubound of arr, -1 if arr has no elements 
'  USED BY ALL OTHER SUBS AND FUNCTIONS IN THIS SCRIPT 
'==================================================================== 
function getUbound(arr) 
	dim uba 
	uba = -1 
	on error resume next 
	uba = ubound(arr) 
	getUbound = uba 
end function 

'==================================================================== 
'push arr, var 
'  adds var to arr 
'==================================================================== 
sub push(arr, var) 
	dim uba 
	uba = getUBound(arr) 
	redim preserve arr(uba+1) 
	arr(uba+1) = var 
end sub 

'==================================================================== 
'pop(arr) 
'  returns the last element in arr and removes it from arr 
'  returns NULL if arr has no elements 
'==================================================================== 
function pop(arr) 
	dim uba, var 
	uba = getUbound(arr) 
	if uba < 0 then 
		var = NULL 
	else 
		var = arr(uba) 
		redim preserve arr(uba-1) 
	end if 
	pop = var 
end function 

'==================================================================== 
'top(arr) 
'  returns the last element in arr but does not remove it from arr 
'  returns NULL if arr has no elements 
'==================================================================== 
function top(arr) 
	dim uba, var 
	uba = getUbound(arr) 
	if uba < 0 then 
		var = NULL 
	else 
		var = arr(uba) 
	end if 
	top = var 
end function 

'==================================================================== 
'pushObj arr, obj 
'  adds obj to arr 
'==================================================================== 
sub pushObj(arr, obj) 
	dim uba 
	uba = getUBound(arr) 
	redim preserve arr(uba+1) 
	Set arr(uba+1) = obj 
end sub 

'==================================================================== 
'popObj(arr) 
'  returns the last element in arr and removes it from arr 
'  returns Nothing if arr has no elements 
'==================================================================== 
function popObj(arr) 
	dim uba, obj 
	uba = getUbound(arr) 
	if uba < 0 then 
		set obj = Nothing 
	else 
		set obj = arr(uba) 
		redim preserve arr(uba-1) 
	end if 
	set popObj = obj 
end function 

'==================================================================== 
'topObj(arr) 
'  returns the last element in arr but does not remove it from arr 
'  returns Nothing if arr has no elements 
'==================================================================== 
function topObj(arr) 
	dim uba, obj 
	uba = getUbound(arr) 
	if uba < 0 then 
		set obj = Nothing 
	else 
		set obj = arr(uba) 
	end if 
	set topObj = obj 
end function 





