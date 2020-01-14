# UFTFunctions
UFTFunctions

Main Core function :

' Declare variables and constants
Dim PRINT_LOG_LEVEL, RESULTS_LOG_LEVEL, objWscript
Public Const micDebug = -1 ' new mic Value used in ReportMessage function.
Set objWscript = CreateObject("Wscript.Shell")
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************

Public Function ReportMessage (iMic,sTitle, sDesc, iPrintLogLevel, iResultsLogLevel)
	' iMic :  The error level of the message [ micDebug(-1), micDone(2), micWarning(3), micPass(0), micFail(1) ]
	' sTitle:  The title or summary of the message.
	' sDesc: The details or description of the message.
	' iPrintLogLevel:  The Pring Log Level.  This value overrides the global value set at run time.
	' iResultsLogLevel:  The Results  Log Level.  This value overrides the global value set at run time.

	If PRINT_LOG_LEVEL = "" Then
		PRINT_LOG_LEVEL = 0
	End If

	If RESULTS_LOG_LEVEL = "" Then
		RESULTS_LOG_LEVEL = 0
	End If
   
   ' If iPrintLogLevel is specified it overrides the global  Log Level.
   ' Otherwise the Log Level is set to the Global  value
	If iPrintLogLevel = "" Then
		iPrintLogLevel = PRINT_LOG_LEVEL
	End If
   ' If iResultsLogLevel is specified it overrides the global  Log Level.
   ' Otherwise the Log Level is set to the Global  value	
	If iResultsLogLevel = "" Then
		iResultsLogLevel = RESULTS_LOG_LEVEL
	End If
	If not iPrintLogLevel = "noBrowser" Then
		If Browser("Browser").Exist Then
			urlString = Browser("Browser").Page("Page").GetROProperty("url")
		else
			urlString = "none"
		End If
	else
		urlString = "none"
	End If
    '  Assign numerical value to Event Status and set event specific variables.
	genericDesc = sDesc & VbCrLf & _
									"*****************************   Checkpoint Details *****************************" & VbCrLf & _
									 "      Execution Time:  " & Now & VbCrLf & _
									 "      Test Name:  " & Environment("TestName") & VbCrLf & _
									"      Test Iteration:  " & Environment("TestIteration") & VbCrLf & _
									"      Action Name:  " & Environment("ActionName") & VbCrLf & _
									"      Action Iteration:  " & Environment("ActionIteration") & VbCrLf & _
									"      Local Host Name:  " & Environment("LocalHostName") & VbCrLf & _
									"      Operating System:  " & Environment("OS") & VbCrLf & _
									"      Operating System Version :  " & Environment("OSVersion") & VbCrLf & _
									 "      Url :  " & urlString & VbCrLf & _
								    "********************************************************************************"
	Select Case iMic
			Case micDebug
					iEventStatus = 0
					iMic = 2
					sDesc = genericDesc
			Case micDone
					iEventStatus = 1
					sDesc = genericDesc
			Case micWarning
					iEventStatus = 2
					sDesc = genericDesc
			Case micPass
					iEventStatus = 3
					sDesc = genericDesc
			Case micFail
					If not iPrintLogLevel = "noBrowser" Then
						Call take_ScreenShot(ssLocation)
					End If
					iEventStatus = 4
					sDesc = sDesc & VbCrLf & _
									"*****************************   ERROR DETAILS *****************************" & VbCrLf & _
									 "      Error Time:  " & Now & VbCrLf & _
									 "      Test Name:  " & Environment("TestName") & VbCrLf & _
									"      Test Iteration:  " & Environment("TestIteration") & VbCrLf & _
									"      Action Name:  " & Environment("ActionName") & VbCrLf & _
									"      Action Iteration:  " & Environment("ActionIteration") & VbCrLf & _
									"      Local Host Name:  " & Environment("LocalHostName") & VbCrLf & _
									"      Operating System:  " & Environment("OS") & VbCrLf & _
									"      Operating System Version :  " & Environment("OSVersion") & VbCrLf & _
									 "      Url :  " & urlString & VbCrLf & _
								    "********************************************************************************" & VbCrLf & _
									ssLocation
    End Select

	' Sends messages with event statuses greater than the Log Level to the QTP Print Log
	If iEventStatus >= iPrintLogLevel Then
			Print sTitle & ":  " & sDesc
	End If

	' Sends messages with event statuses greater than the Log level to the QTP Results Log
    If iEventStatus >= iResultsLogLevel Then
			Reporter.ReportEvent iMic,  sTitle, sDesc
	End If
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function take_ScreenShot(ssLocation)
   'Declare windows script host network object
	Dim WshNetwork
	'Create the object
	Set WshNetwork = CreateObject("WScript.Network") 


		'Create a date string in the format mm-dd-yyyy
		Call createDate(today,today2,today3)
		'Get the test name
		tName = Environment("TestName")
		'Adjust the test name to contain only valid characters for a folder name
		tName = Replace(tName,"/","")
		'Add the date and the test name together to create the folder name
		fName =  tName & "_" & today3
		'Set the output folder to a variable
		outFolder = "C:\QA\Automation\eVenue\Automated Functional Testing\Scripts\QTP Bug Screenshots\" & fName
		'Create the fso object
		Set fso = CreateObject("Scripting.FileSystemObject")
		'Check if the output folder exists, if not, create it
		fExist = fso.FolderExists(outFolder)
		If not fExist = True Then
			fso.CreateFolder (outFolder)
		End If
		'Declare the datestamp variable
		datestamp = Now() 
		'Create a output file name based on the test name and current time
		ssName = Environment("TestName") & "_"&datestamp
		'Replace characters in test name to create valid file name
		ssName = Replace(ssName,"/","") 
		ssName = Replace(ssName,":","")
		'Create a variable containing the full file path
		ssLocation = "C:\QA\Automation\eVenue\Automated Functional Testing\Scripts\QTP Bug Screenshots\" & fName & "\" & ssName & ".png"
		'Create and image capture object 
		set ImageCap = CreateObject("SNAGIT.ImageCapture") 
		'Configure capture object to capture a window
		ImageCap.Input = 1
		'Configure capture object to save to a file
		ImageCap.Output = 2
		'Configure the capture object to save file as a specified name
		ImageCap.OutputImageFile.FileNamingMethod = 1
		'Configure name of output file
		ImageCap.OutputImageFile.FileName = ssName
		'Set location to save output file
		ImageCap.OutputImageFile.Directory = outFolder
		'Set auto scroll setting to scroll up/down and left/right
		ImageCap.AutoScrollOptions.AutoScrollMethod = 1
		'Set the capture delay to 1 second to bring the window to the foreground
		ImageCap.AutoScrollOptions.Delay = 1
		'Configure capture object to capture at the specified coordinates
		ImageCap.InputWindowOptions.SelectionMethod = 3
		'Get the browser's height/width, then set the x,y capture point in the center
		bHeight = Browser("Browser").GetROProperty("height")
		bWidth = Browser("Browser").GetROProperty("width")
		ImageCap.InputWindowOptions.XPos = bWidth/2
		ImageCap.InputWindowOptions.YPos = bHeight/2
		'Invoke the capture sequence
		ImageCap.Capture 
		'Wait until capture is completed
		Do Until ImageCap.IsCaptureDone 
			Wait 0,500
		Loop 
	'End If
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function createDate(today,today2,today3)
dateYear = Year(Date)
dateMonth = Month(Now)
dateDay = Day(Date)
If dateMonth = "10" or dateMonth = "11" or dateMonth = "12"Then
	else
	dateMonth = "0" & dateMonth
End If
dateDay2 = dateDay/10
dateResult= dateDay2 < 1
If dateResult = "True" Then
	dateDay = "0" & dateDay
End If
today = dateMonth & "/" & dateDay & "/" & dateYear
today2 = dateYear & dateMonth & dateDay
today3 = dateMonth & "-" & dateDay & "-" & dateYear
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function createDate2(today) 'This function does not add zeros to single digits
	dateYear = Year(Date)
	dateMonth = Month(Now)
	dateDay = Day(Date)
	today = dateMonth & "/" & dateDay & "/" & dateYear
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function DateFormatter(outDate)	'This function yields today's date into this format:  January 19, 2010
	dd = DatePart("D", Date)
	mm = DatePart("M", Date)
	yyyy = DatePart("YYYY", Date)
	
	mth = MonthName(mm)
	
	outDate = mth & " " & dd & ", " & yyyy
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function disabledCheck(dObject)
   disabled = 1
	Do until disabled = "0"
		 disabled = Browser("Browser").Page("Page").WebList(dObject).GetROProperty("disabled")
	Loop
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function stopWatch(cStartTime,cEndTime,elapsed,errorTimeout,timeout)
   cEndTime = Time
   elapsed = cEndTime - cStartTime
   elapsed = Round(elapsed,0)
	timeout = elapsed > errorTimeout
End Function 
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Public Function StringToAscii(str)
	Dim result, x
	StringToAscii = ""
	If Len(str)=0 Then Exit Function
	If Len(str)=1 Then
		result = Asc(Mid(str, 1, 1))
		StringToAscii = Left("000", 3-Len(CStr(result))) & CStr(result)
		Exit Function
	End If
	result = ""
	For x=1 To Len(str)
		result = result & StringToAscii(Mid(str, x, 1))
	Next
	StringToAscii = result
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
'Option Explicit
Dim objFileSystem,objResultFile,strFileName,strVal1,strVal2,timestamp
	Sub OpenResultsExcel(strInputFileName)
      	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
		Set objResultFile = objFileSystem.OpenTextFile(strFileName, 2, True)
		objResultFile.WriteLine "strVal1,strVal2,timestamp"
	End Sub
    

Sub WriteDataToExcel(strVal1,strVal2,timestamp)
	objResultFile.WriteLine strVal1&","&strVal2&","&timestamp
End Sub
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Function xlread_cell(xlpath, xlsheet, xlrow, xlcol)
	'   Read the value from a cell with in an excel file.
	'i/p : xlpath, xlsheet, xlrow, xlcol
	'o/p : cell value
	Dim myxlapp, myxlsheet
	Set myxlapp = createobject("Excel.Application")
	myxlapp.workbooks.open xlpath ' Open that XLApp in this new created object
	
	Set myxlsheet = myxlapp.activeworkbook.worksheets(xlsheet)
'	print " cell value is " & myxlsheet.cells(xlrow, xlcol)
	
	xlread_cell = myxlsheet.cells(xlrow, xlcol)

	myxlapp.activeworkbook.close ' Close all opened workbooks.
	myxlapp.application.quit ' Close the Excel App.
	
	Set myxlapp = nothing ' Release the memory held for the object
	Set myxlsheet = nothing
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************
Function xlwrite_cell(xlpath, xlsheet, xlrow, xlcol, xldata)
	'   Read the value from a cell with in an excel file.
	'i/p : xlpath, xlsheet, xlrow, xlcol, xldata
	'o/p : cell value
	Dim myxlapp, myxlsheet
	Set myxlapp = createobject("Excel.Application")
	myxlapp.workbooks.open xlpath ' Open that XLApp in this new created object
	
	Set myxlsheet = myxlapp.activeworkbook.worksheets(xlsheet)
    	
	myxlsheet.cells(xlrow, xlcol) = xldata

	myxlapp.activeworkbook.save ' Save the data before we close it.
	myxlapp.activeworkbook.close ' Close all opened workbooks.
	myxlapp.application.quit ' Close the Excel App.
	
	Set myxlapp = nothing ' Release the memory held for the object
	Set myxlsheet = nothing
End Function
'********************************************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************************************

Key word function:

'*************************************************'launch Portal App*******************************************************************************************************************	
Function StartApp(sValidURL)
	   'Close any existing IE browser windows
	   systemutil.CloseProcessByName "iexplore.exe"   
      SystemUtil.Run "iexplore.exe", sValidURL

	If Browser("Browser").Page("Page").Exist(5) then
     wait 1
   ' strHwnd = Browser("Browser").object.HWND
   ' Window("hwnd:=" & strHwnd.HWND).Maximize
    
	      startApp = "Pass"
	 else
		  startApp = "Fail"
	End If
    If Browser("Browser").Page("Page").WebElement("AccountSetting_lable").Exist(0) Then
	     Browser("Browser").Page("Page").WebElement("AccountSetting_lable").Click
	     Browser("Browser").Page("Page").WebElement("Logout_Link").Click    
        
    End If
End Function

'******************************************************__End__*************************************************************************************************************************

'*************************************************'Log-In Page Testing*****************************************************************************************************************
Function TypeUserName(sTestData)
If Browser("Browser").Page("Page").WebEdit("FBEmail_InputField").Exist(2) Then
		Browser("Browser").Page("Page").WebEdit("FBEmail_InputField").Set sTestData
		vUsername = Browser("Browser").Page("Page").WebEdit("FBEmail_InputField").GetROProperty("value")
	 If vUsername <> "" then
	      TypeUserName = "Pass"
	 else
		  TypeUserName = "Fail"
	End If
	End If
End Function

Function TypePassword(sTestData)
		Browser("Browser").Page("Page").WebEdit("FBPassword_InputFiled").Set sTestData
		vUsername = Browser("Browser").Page("Page").WebEdit("FBPassword_InputFiled").GetROProperty("value")
	 If vUsername <> "" then
	      TypePassword = "Pass"
	 else
		  TypePassword = "Fail"
	End If
End Function


Function SubmitCredentials()
	If Browser("Browser").Page("Page").Webbutton("FBLogin_btm").Exist(10) Then
		Browser("Browser").Page("Page").Webbutton("FBLogin_btm").Click
	    SubmitCredentials = "Pass"
	    Else
	    SubmitCredentials = "Fail"
	End If
End Function

Function VerifySignInPositive()

	   If Browser("Browser").Page("Page").Link("Home_Text").Exist(2) Then
	      VerifySignInPositive = "Pass"
	 else
		  VerifySignInPositive = "Fail"
	   'fnSendEmailFromOutlook
	End If
End Function


Function VerifySignInNegitive()

	 If Browser("Browser").Page("Page").Link("Home_Text").Exist(1) Then

	      VerifySignInNegitive = "Fail"
	 else
		  VerifySignInNegitive = "Pass"
				  'fnSendEmailFromOutlook ' Send an Email to <name and or cc to name>
	End If
End Function


Function SignOff()
     If Browser("Browser").Page("Page").WebElement("AccountSetting_lable").Exist(3) Then
	    Browser("Browser").Page("Page").WebElement("AccountSetting_lable").Click
	    Browser("Browser").Page("Page").WebElement("Logout_Link").Click
	      SignOff = "Pass"
	      else
	      SignOff = "Fail"
	    End If
End Function

Function VerifySignOff()
     	If Browser("Browser").Page("Page").WebEdit("FBEmail_InputField").Exist(3) Then
	      VerifySignOff = "Pass"
	      else
	      VerifySignOff = "Fail"
	    End If
End Function

Function ValidateTC3Portal_Lable()
     If Browser("Browser").Page("Page").WebElement("TC3_Portal_Lable_WebElement").Exist(50) then 
		Call ReportMessage(micPass, "A 'TC3 Portal' Lable is displaying on Log-in page.", "Test Pass: ""","","")
	      ValidateTC3Portal_Lable = "Pass" 
	 else
		  ValidateTC3Portal_Lable = "Fail"
		Call ReportMessage(micFail, "A 'TC3 Portal' Lable is not displaying on Log-in page. 'TC3 Portal' lable should display top of User-Name Field on Log-In page.","Test Failed:","","")
	End If
End Function

Function ValidateUserName_Lable()
     If Browser("Browser").Page("Page").WebElement("User Name_Lable_WebElement").Exist(50) then 
		Call ReportMessage(micPass, "A 'User Name' Lable is displaying on Log-in page.", "Test Pass: ""","","")
	      ValidateUserName_Lable = "Pass" 
	 else
		  ValidateUserName_Lable = "Fail"
		Call ReportMessage(micFail, "A 'User Name' Lable is not displaying on Log-in page. 'User Name' lable should display left to the User-Name Field on Log-In page.","Test Failed:","","")
	End If
End Function

Function ValidatePassword_Lable()
     If Browser("Browser").Page("Page").WebElement("Password_Lable_WebElement").Exist(50) then 
		Call ReportMessage(micPass, "A 'Password' Lable is displaying on Log-in page.", "Test Pass: ""","","")
	      ValidatePassword_Lable = "Pass" 
	 else
		  ValidatePassword_Lable = "Fail"
		Call ReportMessage(micFail, "A 'Password' Lable is not displaying on Log-in page. 'Password' lable should display left to the Password Field on Log-In page.","Test Failed:","","")
	End If
End Function

Function ValidateInvalidLogin_Label()
	 If Browser("Browser").Page("Page").WebElement("Invalid_Login_Label").Exist(5) Then
		Call ReportMessage(micPass, "A 'Invalid Login.' Label is displaying on Log-in page.", "Test Pass: ""","","")
	      ValidateInvalidLogin_Label = "Pass"
	  else
		  ValidateInvalidLogin_Label = "Fail"
		'Call ReportMessage(micFail, "A 'Invalid Login.' Label is not displaying on Log-in page' Clicking on login Button Upon providing invalid password and username, a Label 'Invalid Login.' should display below Password Field on Log-In page.", "Test failed: ""","","")
				  'fnSendEmailFromOutlook
	End If
End Function


'*************************************************'**************************************__END__***************************************************************************************


'****************************************************************Disclaimer_Page********************************************************************************************************
Function Verify_PrivacyStatement_Disclaimer_Page()
	 If Browser("Browser").Page("Page").WebElement("Disclamir_PrivacyandSecurity_Agreement_WebElement").Exist(25) Then
		Call ReportMessage(micPass, "A 'Invalid Login.' Disclamir Privacy and Security Agreement is displying Corecctly on the page.", "Test Pass: ""","","")
	      Verify_PrivacyStatement_Disclaimer_Page = "Pass"
	  else
		  Verify_PrivacyStatement_Disclaimer_Page = "Fail"
		Call ReportMessage(micFail, "A 'Invalid Login.' Disclamir Privacy and Security AgreementObject 'webelement' is NOT displying on the page.", "Test Fail: ""","","")
	End If
End Function

Function Verify_LogOut_Link_Disclaimer_Page()
	 If Browser("Browser").Page("Page").Link("Logout_Link").Exist(25) Then
		Call ReportMessage(micPass, "A 'Invalid Login.'  is displying Corecctly on disclaimer page.","Test Pass: ""","","")
	      Verify_LogOut_Link_Disclaimer_Page = "Pass"
	  else
		  Verify_LogOut_Link_Disclaimer_Page = "Fail"
		Call ReportMessage(micFail, "A 'Invalid Login.' Log-Out link is displying on the page.see Screenshot","Test Fail: ""","","")
	End If
	
End Function
'*************************************************'**************************************__END__***************************************************************************************







'*************************************************Close Borwser************************************************************************************************************************
Function CloseBrowser()
	   Browser("Browser").Close
	   'Close any existing IE browser windows
	   systemutil.CloseProcessByName "iexplore.exe"
	If Browser("Browser").Page("Page").Exist(1) then
	    CloseBrowser = "Fail"
	 else
		CloseBrowser = "Pass"
	End If
End Function
'*************************************************'**************************************__END__***************************************************************************************

