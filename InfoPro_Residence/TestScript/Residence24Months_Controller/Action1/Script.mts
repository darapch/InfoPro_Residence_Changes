RepositoriesCollection.Add Environment.Value("Path") & "ObjectRepository\InforProOR.tsr"

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIDAA000", "Action1", oneIteration
'RunAction "Action1 [BIDAA000]", oneIteration
Call BIDAA000()

Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA001", "Action1", oneIteration
'RunAction "Action1 [BIGAA001]", oneIteration
Call BIGAA001()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01", "Action1", oneIteration
'RunAction "Action1 [CUGAACST01]", oneIteration
Call CUGAACST01()
Wait(1)

Environment.Value("BIGAA014Fields") = "RevenueDist:30;TaxCode:0000"
'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014", "Action1", oneIteration
'RunAction "Action1 [BIGAA014]", oneIteration
Call BIGAA014()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01_2", "Action1", oneIteration
'RunAction "Action1 [CUGAACST01_2]", oneIteration
Call CUGAACST01_2()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014_Override", "Action1", oneIteration
'RunAction "Action1 [BIGAA014_Override]", oneIteration
Call BIGAA014_Override()
Wait(1)

Call func_SendKey("F10")
'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014R", "Action1", oneIteration
'RunAction "Action1 [BIGAA014R]", oneIteration
Call BIGAA014R()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014R", "Action1", oneIteration
'RunAction "Action1 [BIGAA014R]", oneIteration
Call BIGAA014R()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\RateValidation", "Action1", oneIteration
'RunAction "Action1 [RateValidation]", oneIteration
Call RateValidation()
Wait(1)

Call func_SendKey("F3")
Wait(1)

Call func_setScreenProperty("BIGAA014")
If (TEWindow("InfoProWindow").TEScreen("BIGAA014").TEField("ResidentialServiceInfo").Exist(5)) Then
	Environment.Value("BIGAA014Fields") = "RevenueDist:30;TaxCode:0000"
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014", "Action1", oneIteration
	'RunAction "Action1 [BIGAA014]", oneIteration
	Call BIGAA014()
	Wait(1)
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01_2", "Action1", oneIteration
	'RunAction "Action1 [CUGAACST01_2]", oneIteration
	Call CUGAACST01_2()
	Wait(1)
	
	Call func_setScreenProperty("BIGAA014")	
	Call func_SendKey("ENTER")
	
	If (TEWindow("InfoProWindow").TEScreen("BIGAA014").TEField("DeliveryContConfOvride").Exist(5)) Then
		Call func_SendKey("F10")
	End If 'If (TEWindow("InfoProWindow").TEScreen("BIGAA014").TEField("DeliveryContConfOvride").Exist(5)) Then
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014R", "Action1", oneIteration
	'RunAction "Action1 [BIGAA014R]", oneIteration
	Call BIGAA014R()
	Wait(1)
	
	Call func_SendKey("F3")
	wait(2)
	strStatusMsg = TEWindow("InfoProWindow").TEScreen("BIGAA014R").TeField("field id:=1842").GetROProperty("text")
	If Trim(strStatusMsg)="One rate for the container group MUST BE ADDED." Then
		TEWindow("InfoProWindow").TEScreen("BIGAA014R").SendKey TE_ENTER
		wait(2)
		If TEWindow("InfoProWindow").TEScreen("BIGAA014R").TEField("SalesTransactionDetails").Exist(3) Then
			TEWindow("InfoProWindow").TEScreen("BIGAA014R").SendKey TE_ENTER
			wait(2)
			TEWindow("InfoProWindow").TEScreen("BIGAA014R").SendKey TE_PF3
		End If
	End If
	Wait(1)
End If

''LoadAndRunAction Environment.Value("Path") & "TestScript\CREATEACCOUNT", "Action1", oneIteration
Call CREATEACCOUNT()
