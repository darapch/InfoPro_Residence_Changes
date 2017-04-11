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

Call func_setScreenProperty("BIGAA002_CONTAINER")
If (TEWindow("InfoProWindow").TEScreen("BIGAA002_CONTAINER").TEField("ContainerInformation").Exist(5)) Then
	Environment.Value("SecondFlow") = "YES"
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONTAINER", "Action1", oneIteration
	Call BIGAA002_CONTAINER()
	Wait(1)
	
	'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONT_RATE_INFO", "Action1", oneIteration
	Call BIGAA002_CONT_RATE_INFO()
	Wait(1)
	
	Call func_setScreenProperty("BIGAA002_CONTAINER")
	If (TEWindow("InfoProWindow").TEScreen("BIGAA002_CONTAINER").TEField("ContainerInformation").Exist(5)) Then
		'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONTAINER", "Action1", oneIteration
		Call BIGAA002_CONTAINER()
		Wait(1)
		
		'LoadAndRunAction Environment.Value("Path") & "TestScript\BIGAA002_CONT_RATE_INFO", "Action1", oneIteration
		Call BIGAA002_CONT_RATE_INFO()
		Wait(1)
	End If 'If (TEWindow("InfoProWindow").TEScreen("BIGAA002_CONTAINER").TEField("ContainerInformation").Exist(5)) Then
	
End If 'If (TEWindow("InfoProWindow").TEScreen("BIGAA002_CONTAINER").TEField("ContainerInformation").Exist(5)) Then

'LoadAndRunAction Environment.Value("Path") & "TestScript\CREATEACCOUNT", "Action1", oneIteration
Call CREATEACCOUNT()
