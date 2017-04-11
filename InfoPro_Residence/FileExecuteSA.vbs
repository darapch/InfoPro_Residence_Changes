Dim arr_Arg, str_quoteNumber, str_lob
Dim obj_QTP, obj_outputFile, obj_shell, obj_service, obj_readFile
Dim int_accountNumber, str_passFlag, int_warningFlag, str_warnings, str_failurereason
Dim str_process
Dim str_fileName, int_readCount, str_passFlagLine, str_status, str_accountNumberLine
Dim obj_restReq
Dim str_url, str_testResult
Set arr_Arg = WScript.Arguments

str_quoteNumber = arr_Arg(0)
str_lob = arr_Arg(1)
str_driverPath = arr_Arg(2)
'strTestId = arr_Arg(3)
str_testResult = ""

Set obj_QTP = Createobject("quicktest.application")
obj_QTP.launch
obj_QTP.visible=False
obj_QTP.Options.Run.RunMode = "Fast"
'obj_QTP.open "C:\Users\sreerga\Desktop\InfoPro_Residence\TestScript\E2EDriver", True
obj_QTP.open str_driverPath&"\TestScript\E2EDriver", True

obj_QTP.Test.Environment("QuoteNumber") = str_quoteNumber
obj_QTP.Test.Environment("LOB") = str_lob

obj_QTP.Test.run
obj_QTP.Test.Close
obj_QTP.quit

Set obj_shell = CreateObject("WScript.Shell")
Set obj_service = GetObject ("winmgmts:")
For Each str_process in obj_service.InstancesOf ("Win32_Process")
	If Ucase(Trim(str_process.Name)) = "PCSWS.EXE" Then
		obj_shell.Run "taskkill /f /im pcsws.exe", , True
		obj_shell.Run "taskkill /f /im pcsws.exe", , True
		'SystemUtil.CloseProcessByName("pcsws.exe")
		'SystemUtil.CloseProcessByName("pcscm.exe")
	End If 'If str_process.Name = "pcsws.exe *32" Then
Next 'For Each str_process in obj_service.InstancesOf ("Win32_Process")

Set obj_outputFile = CreateObject("Scripting.FileSystemObject")
str_fileName = str_driverPath&"\Result\"&str_quoteNumber&".txt"
If (obj_outputFile.FileExists(str_fileName)) Then
	Set obj_readFile = obj_outputFile.openTextFile(str_fileName,1)
	
	For int_readCount = 1 to 4
		If (int_readCount = 1) Then
			str_passFlagLine = obj_readFile.ReadLine
			str_status = Trim(Mid(str_passFlagLine, 10, 6))
		ElseIf (int_readCount = 2) Then
			If (str_status = "PASSED") Then
				str_accountNumberLine = obj_readFile.ReadLine
				int_accountNumber = Trim(Mid(str_accountNumberLine, 19, 6))
				Exit For
			Else
				obj_readFile.ReadLine
			End If 'If (str_status = "PASSED") Then
		ElseIf (int_readCount = 4) Then
			If (str_status = "FAILED") Then
				str_failureReason = obj_readFile.ReadLine
			End If 'If (str_status = "PASSED") Then
		Else
			obj_readFile.ReadLine
		End If 'If (int_readCount = 1) Then
	Next
Else
	Msgbox("Output file "&str_quoteNumber&".txt does not exists")
End If 'If (obj_outputFile.FileExists(str_path)) Then

'Set obj_restReq = CreateObject("Microsoft.XMLHTTP")
	
'If (str_status = "PASSED") Then
'	str_testResult = "{""AccountNumber"":"""&int_accountNumber &""",""Result"":true,""Output"":""Infopro test output set to true by vbscript""}"
'Else
'	str_testResult = "{""OrderNumber"":"""&str_quoteNumber &""",""Result"":false,""Output"":"""&str_failureReason &"""}"
'End If 'If (str_status = "PASSED") Then
	
'str_url = "http://localhost:8080/api/endtoend/infoproresult?ordernumber=76300"
'str_url = "http://127.0.0.1:8080/api/endtoend/infoproresult?testid="&str_testId
'str_url = "http://localhost:8080/api/endtoend/infoproresult?testid=" & strTestId

'obj_restReq.open "PUT", str_url, false
'obj_restReq.setRequestHeader "Content-Type", "application/json"
'obj_restReq.send str_testResult

'WScript.echo obj_restReq.responseText

'Set obj_restReq = NOTHING
Set obj_readFile = NOTHING
Set obj_outputFile = NOTHING
Set obj_QTP = NOTHING
Set obj_service = NOTHING