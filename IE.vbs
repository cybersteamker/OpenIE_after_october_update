Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_")
objie.navigate "home"
objIE.Visible = 1
objShell.AppActivate objIE
