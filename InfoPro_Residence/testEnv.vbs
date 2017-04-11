Dim arr_Arg, str_quoteNumber, str_lob
Dim obj_QTP

Set arr_Arg = WScript.Arguments
str_quoteNumber = arr_Arg(0)
str_lob = arr_Arg(1)

Set obj_QTP = Createobject("quicktest.application")
obj_QTP.launch
obj_QTP.visible=False
obj_QTP.Options.Run.RunMode = "Fast"
obj_QTP.open "C:\Users\sreerga\Desktop\InfoPro_Residence\TestScript\testEnv", True

obj_QTP.Test.Environment("QuoteNumber") = str_quoteNumber
obj_QTP.Test.Environment("LOB") = str_lob

obj_QTP.Test.Environment.Value("OutputQuote") = ""
obj_QTP.Test.Environment.Value("OutputLOB") = ""

obj_QTP.Test.run
Msgbox("Before")
Msgbox(obj_QTP.Test.Environment.Value("OutputQuote"))
Msgbox(obj_QTP.Test.Environment.Value("OutputLOB"))
Msgbox("After")
obj_QTP.Test.Close
obj_QTP.quit

Msgbox("DONE")

Set obj_QTP = NOTHING
Set Arg = NOTHING