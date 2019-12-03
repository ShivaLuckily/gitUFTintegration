 Dim objQTP
 Set objQTP = CreateObject("QuickTest.Application")
 objQTP.Visible = True

 objQTP.Launch

 objQTP.Open "C:\Users\Shipra Mandal\Documents\Unified Functional Testing\Smoke\cal"
 objQTP.Test.Run
 

 
 objQTP.Open "C:\Users\Shipra Mandal\Documents\Unified Functional Testing\Smoke\GUITest4"
 objQTP.Test.Run
 objQTP.Test.Close

 objQTP.Quit
 Set objQTP = Nothing
