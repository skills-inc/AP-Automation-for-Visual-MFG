'AP Automation for Invoices
'Update the URL on line 29 to the Webserver where your going to publish the Click Once Application
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim AppPath
Dim statusCode
Dim appDataLoc
Dim Company
Dim PO
Dim DocType
Dim DocNo


appDataLoc = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
Company = VENDOR_ID
PO = ""
DocType = "Invoice"
DocNo = INVOICE_ID

APSoftware = """" + appDataLoc + "\Desktop\AP Automation.appref-ms"" "
APSoftware1 = appDataLoc + "\Desktop\AP Automation.appref-ms"

If Not (oFSO.FileExists(APSoftware1)) Then
    Dim objExplorer
    Set objExplorer = CreateObject ("InternetExplorer.Application")
    objExplorer.Visible = 1
    objExplorer.Navigate "http://YourWebserverToHostClickOnceApps/AP%20Automation/AP%20Automation.application#AP Automation.application"
    Set objExplorer = Nothing
    MsgBox "AP Automation Software Installed"
else
    AppPath = """" + appDataLoc + "\Desktop\AP Automation.appref-ms"" " + Company + "," + PO + "," + DocType + "," + DocNo
    statusCode = WshShell.Run (AppPath, 1, true)
end if 
