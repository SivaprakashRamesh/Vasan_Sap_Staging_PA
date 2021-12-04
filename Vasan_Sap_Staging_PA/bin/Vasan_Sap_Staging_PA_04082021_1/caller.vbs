Set WshShell = WScript.CreateObject("WScript.Shell")
param1= WScript.Arguments.Named.Item("param1")
param2= WScript.Arguments.Named.Item("param2")
URL = "C:\Users\praveen.e\Desktop\Vasan_Final_Source_22022012\ProActive\Vasan_Sap_Staging\Vasan_Sap_Staging\bin\Debug\Vasan_Sap_Staging_PA " &param1 &" " &param2
WshShell.Run(URL)