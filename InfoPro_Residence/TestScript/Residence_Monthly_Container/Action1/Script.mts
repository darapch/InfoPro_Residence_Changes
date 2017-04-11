RepositoriesCollection.Add Environment.Value("Path") & "ObjectRepository\InforProOR.tsr"


'LoadAndRunAction Environment.Value("Path") & "TestScript\BIDAA000", "Action1", oneIteration
Call BIDAA000()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA001", "Action1", oneIteration
Call BIGAA001()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01", "Action1", oneIteration
Call CUGAACST01()
Wait(1)

Environment.Value("BIGAA014Fields") = "RevenueDist:30;TaxCode:0000"
'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014", "Action1", oneIteration
Call BIGAA014()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01_2", "Action1", oneIteration
Call CUGAACST01_2()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014_Override", "Action1", oneIteration
Call BIGAA014_Override
Wait(1)

Call func_SendKey("F10")

'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014R", "Action1", oneIteration
Call BIGAA014R()
Wait(1)

'LoadAndRunAction Environment.Value("Path") & "TestScript\RateValidation", "Action1", oneIteration
Call RateValidation()
Wait(1)

Call func_SendKey("F3")
Wait(1)

Call func_setScreenProperty("BIGAA002")
Call func_setScreenProperty("BIGAA014")
If (TEWindow("InfoProWindow").TEScreen("BIGAA002").TEField("SiteInformation").Exist(5)) Then
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002", "Action1", oneIteration
	Call BIGAA002()
	Wait(1)
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01_2", "Action1", oneIteration
	Call CUGAACST01_2()
	Wait(1)
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONTAINER", "Action1", oneIteration
	Call BIGAA002_CONTAINER()
	Wait(1)
	
	Call func_SendKey("F10")
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONT_RATE_INFO", "Action1", oneIteration
	Call BIGAA002_CONT_RATE_INFO()
	Wait(1)
	
	Call func_SendKey("F3")
	Wait(1)
	
ElseIf (TEWindow("InfoProWindow").TEScreen("BIGAA014").TEField("ResidentialServiceInfo").Exist(5)) Then
	Environment.Value("BIGAA014Fields") = "RevenueDist:30;TaxCode:0000"
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014", "Action1", oneIteration
	Call BIGAA014()
	Wait(1)
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01_2", "Action1", oneIteration
	Call CUGAACST01_2()
	Wait(1)
	
	Call func_setScreenProperty("BIGAA014")	
	Call func_SendKey("ENTER")
	
	If (TEWindow("InfoProWindow").TEScreen("BIGAA014").TEField("DeliveryContConfOvride").Exist(5)) Then
		Call func_SendKey("F10")
	End If 'If (TEWindow("InfoProWindow").TEScreen("BIGAA014").TEField("DeliveryContConfOvride").Exist(5)) Then
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA014R", "Action1", oneIteration
	Call BIGAA014R()
	Wait(1)
	
	Call func_SendKey("F3")
	Wait(1)
	
	Call func_setScreenProperty("BIGAA002")
	If (TEWindow("InfoProWindow").TEScreen("BIGAA002").TEField("SiteInformation").Exist(5)) Then
		'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002", "Action1", oneIteration
		Call BIGAA002()
		Wait(1)
	
		'LoadAndRunAction Environment.Value("Path") & "TestScript\CUGAACST01_2", "Action1", oneIteration
		Call CUGAACST01_2()
		Wait(1)
	
		'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONTAINER", "Action1", oneIteration
		Call BIGAA002_CONTAINER()
		Wait(1)
	
		Call func_SendKey("F10")
		'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONT_RATE_INFO", "Action1", oneIteration
		Call BIGAA002_CONT_RATE_INFO()
		Wait(1)
		
		'Call func_SendKey("F3")
		Wait(1)
	End If 'If (TEWindow("InfoProWindow").TEScreen("BIGAA002").TEField("SiteInformation").Exist(5)) Then
End If 'If (TEWindow("InfoProWindow").TEScreen("BIGAA002").TEField("SiteInformation").Exist(5)) Then

'LoadAndRunAction Environment.Value("Path") & "TestScript\CREATEACCOUNT", "Action1", oneIteration
Call CREATEACCOUNT()
