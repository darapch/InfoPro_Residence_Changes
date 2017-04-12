Dim int_rowCount, int_currentRow
Dim str_currentFunc, str_appPath
Dim str_query, str_dataSheet
Dim str_MenuSelection, str_SelectRegion, str_DivisionRegion
Dim  int_lenAcctNum
Dim str_selectRegionQuery, str_divisionDataTable
Dim str_PrimarySelection, str_SecondarySelection
Dim str_sendKey
Dim str_path, int_pathCount, arr_path
Dim obj_fso
Dim obj_service, str_process
Dim arr_BIDDS035Fields

'Driver
Environment.Value("ErrorScreenshot") = ""
Environment.Value("UName") = ""
Environment.Value("Password") = ""
Environment.Value("ACCOUNTNUMBER") = ""
Environment.Value("AccountNumber") = ""
Environment.Value("TrimAccountNumber") = ""
Environment.Value("OrderNum") = ""
Environment.Value("DivisionNumber") = ""
Environment.Value("ProjectId") = ""
Environment.Value("DivisionCode") = ""
Environment.Value("PrimarySelection") = ""
Environment.Value("SecondarySelection") = ""
Environment.Value("StreetName") = ""
Environment.Value("Path") = ""
Environment.Value("End2EndFlow") = ""
Environment.Value("SecondFlow") = ""

Environment.Value("PassFlag") = ""
Environment.Value("Warnings") = ""
Environment.Value("WarningFlag") = ""
Environment.Value("FailureReason") = ""


Environment.Value("QuoteNum") =   Environment("QuoteNumber")
Environment.Value("LOB") =  Environment("LOB")

Environment.Value("UName") =  Environment("UsName") '"darapch"
Environment.Value("Password") =  Environment("PWORD")



Environment.Value("End2EndFlow") = "YES"
Environment.Value("WarningFlag") = 1

arr_path = Split(Environment.Value("TestDir"), "\")

For int_pathCount = 0 To Ubound(arr_path) - 2
	If (int_pathCount = 0) Then
		str_path = 	arr_path(int_pathCount) & "\"
	Else
		str_path = 	str_path & arr_path(int_pathCount) & "\"
	End If 'If (int_pathCount = 0) Then
	
	'str_excelFilePath = str_path & "DataSheet\ResidenceSimpleFlow_B89005.xls"
Next 'For int_pathCount = 0 To Ubound(arr_path) - 2

Environment.Value("Path") = str_path

'Dynamically Associating the Object Repository & Function Libraries
RepositoriesCollection.Add Environment.Value("Path") & "ObjectRepository\InforProOR.tsr"
LoadFunctionLibrary Environment.Value("Path") & "FunctionLibrary\GenericFunction.vbs"
LoadFunctionLibrary Environment.Value("Path") & "FunctionLibrary\GenericFunction.qfl"
LoadFunctionLibrary Environment.Value("Path") & "FunctionLibrary\ReportingFunction.qfl"


Reporter.ReportEvent micDone, "START RUN", Now()

Call func_closeApplication()

str_appPath = "C:\Program Files (x86)\IBM\Client Access\Emulator\pcsws.exe" '"C:\Program Files\IBM\Client Access\Emulator\pcsws.exe" '
Call func_invokeapplication(str_appPath)
Wait(5)

If Dialog("Configure PC5250").Exist Then
	Call func_handleLoginPopup("Configure PC5250")
End If 'If Dialog("Configure PC5250").Exist Then

If Dialog("System i signon").Exist Then
	Call func_handleLoginPopup("System i signon")
End If 'If Dialog("System i signon").Exist Then

If TEWindow("InfoProWindow").TEScreen("Login").TEField("SignOn").Exist(10) Then
	Call func_Login()
Else
	Call func_reportFailureScreenshot()
	Call func_reportStatus("FAIL", "Login Screen", "Login Screen does not exist")
End If 'If TEWindow("InfoProWindow").TEScreen("Login").TEField("SignOn").Exist(10) Then

str_MenuSelection = "InfoPro"
TEWindow("InfoProWindow").TEScreen("Menu").TEField(str_MenuSelection).Set 1
Call func_SendKey("ENTER")


'Added By Krishna
'********************************

Environment.Value("ProjectName")=Mid(Environment.Value("QuoteNum"),2)

arrQuoteNum = Split(Environment.Value("QuoteNum"),"-")
If UBound(arrQuoteNum)>0 Then
	Environment.Value("QuoteNum")=Trim(arrQuoteNum(0))
End If
'********************************

'Commented and Changed By Krishna
'********************************
str_query = "SELECT * FROM cufile.CUPAAPRO WHERE pvarchar2 = '"&Environment.Value("QuoteNum")&"'"
'str_query = "SELECT * FROM cufile.CUPAAPRO WHERE PROJNAME='" & Environment.Value("ProjectName") & "'"
'********************************

str_dataSheet = "AccountInfo"
Call func_retrieveData(str_query, str_dataSheet)
Environment.Value("QuoteNum") = Ucase(Trim(DataTable.Value("PVARCHAR2", "ACCOUNTINFO")))
Environment.Value("OrderNum") = Mid(Environment.Value("QuoteNum"),2)
Environment.Value("DivisionNumber") = Trim(DataTable.Value("DIVISION", "ACCOUNTINFO"))
Environment.Value("ProjectId") = Ucase(Trim(DataTable.Value("PROJECTID", "ACCOUNTINFO")))
Environment.Value("ProjectName") = "B" & Ucase(Trim(DataTable.Value("PROJNAME", "ACCOUNTINFO")))


If ((Trim(DataTable.Value("AAACCT", "ACCOUNTINFO"))) <> "") Then
	Environment.Value("AccountNumber") = Ucase(Trim(DataTable.Value("AAACCT", "ACCOUNTINFO")))
	intSpaces = 7-Len(Environment.Value("AccountNumber"))
	Environment.Value("AccountNumber") = Space(intSpaces) & Environment.Value("AccountNumber")
	
	'Commented By Krishna
'	int_lenAcctNum = Len(Trim(Environment.Value("AccountNumber")))
'	If (int_lenAcctNum = 4) Then
'		Environment.Value("TrimAccountNumber") = "   "&Trim(Environment.Value("AccountNumber"))
'	ElseIf (int_lenAcctNum = 5) Then
'		Environment.Value("TrimAccountNumber") = "  "&Trim(Environment.Value("AccountNumber"))
'	ElseIf (int_lenAcctNum = 6) Then
'		Environment.Value("TrimAccountNumber") = " "&Trim(Environment.Value("AccountNumber"))
'	Else
'		'Msgbox("Account Number less than 4 digits. Please Check")
'		Call func_reportStatus("FAIL", "Account Number less than 4 digits. Please Check.", "Account Number :"&Environment.Value("AccountNumber"))
'	End If 'If (int_lenAcctNum = 4) Then
End If 'If ((Trim(DataTable.Value("AAACCT", "ACCOUNTINFO"))) <> "") Then

Reporter.ReportEvent micDone, "QUOTE NUMBER : ", Environment.Value("QuoteNum")
Environment.Value("ErrorScreenshot") = str_path & Environment.Value("QuoteNum") & "_error.png"

Set obj_fso = CreateObject("Scripting.FileSystemObject")
If (obj_fso.FileExists(Environment.Value("ErrorScreenshot"))) Then
	obj_fso.DeleteFile(Environment.Value("ErrorScreenshot"))
End If 'If (obj_fso.FileExists(Environment.Value("ErrorScreenshot"))) Then
Set obj_fso = Nothing

str_selectRegionQuery = "Select * from cufile.BIPIC where ICCOMP = '  "&Environment.Value("DivisionNumber")&"'"
str_divisionDataTable = "DIVISION"
Call func_retrieveData(str_selectRegionQuery, str_divisionDataTable)
Environment.Value("DivisionCode") = Ucase(Trim(DataTable.Value("ICREG", "DIVISION")))

Call func_RegionSelection()

Call func_DivisionSelection(Environment.Value("DivisionNumber"))

Environment.Value("PrimarySelection") = "CustomerMaintenance"
Call func_PrimarySelection(Environment.Value("PrimarySelection"))

If (UCASE(Trim(DataTable.Value("CONS_AAE", "ACCOUNTINFO"))) = "Y") Then
	Call func_SecondrySelection("ConsolidatedAutoAccountEntry")
Else
	Call func_SecondrySelection("AutoAccountEntryMaintainAccts")
End If 'If (UCASE(Trim(DataTable.Value("CONS_AAE", "ACCOUNTINFO"))) = "Y") Then

If (Environment.Value("LOB") = "Residence_Monthly") Then
	LoadAndRunAction str_path & "TestScript\ResidenceMonthly_Controller", "Action1", oneIteration
ElseIf (Environment.Value("LOB") = "Residence_Monthly_Container") Then
	LoadAndRunAction str_path & "TestScript\Residence_Monthly_Container", "Action1", oneIteration

ElseIf (Environment.Value("LOB") = "Residence_24Months") Then
	LoadAndRunAction str_path & "TestScript\Residence24Months_Controller", "Action1", oneIteration
ElseIf (Environment.Value("LOB") = "Residence_24Months_Container") Then
	LoadAndRunAction str_path & "TestScript\Residence_24Months_Container", "Action1", oneIteration

ElseIf (Environment.Value("LOB") = "Residence_Container") Then
	LoadAndRunAction str_path & "TestScript\ResidenceContainer_Controller", "Action1", oneIteration

ElseIf (Environment.Value("LOB") = "Business") Then
	LoadAndRunAction str_path & "TestScript\Business_Controller", "Action1", oneIteration
Else
	Call func_reportStatus("FAIL", "No Valid LOB present", "LOB : "&Environment.Value("LOB"))
End If 'If (Environment.Value("LOB") = "Residence Container") Then







