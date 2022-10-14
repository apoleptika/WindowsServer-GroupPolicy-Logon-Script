
REM http://support.microsoft.com/kb/244677/tr
REM http://msdn.microsoft.com/en-us/library/8kst88h6(v=vs.84).aspx
REM Abdurrahim Hacioglu (Antalya)
REM Windows 2008 Active Directory and Domain Controller tests are ok (ALIVELI.local)
REM Company and Organization Unit is Bogacay
REM Windows 7 x64 test is ok
REM Windows XP x86 x64 SP2,SP3 tests are ok.


Const GRP_IT = "cn=grp-it,ou=ou-bogacay,ou=ou-sirket,dc=aliveli,dc=local"
Const GRP_FINANCE = "cn=grp-finance,ou=ou-bogacay,ou=ou-sirket,dc=aliveli,dc=local"
Const GRP_HR = "cn=grp-hr,ou=ou-bogacay,ou=ou-sirket,dc=aliveli,dc=local"
Const GRP_PRODUCTION = "cn=grp-production,ou=ou-bogacay,ou=ou-sirket,dc=aliveli,dc=local"
Const GRP_ACCOUNTING = "cn=grp-accounting,ou=ou-bogacay,ou=ou-sirket,dc=aliveli,dc=local"
Const GRP_MARKETING = "cn=grp-marketing,ou=ou-bogacay,ou=ou-sirket,dc=aliveli,dc=local"


REM map shared network path in local computer to local drive letter F: G: H: ...
Set wshNetwork = CreateObject("WScript.Network")
Set WshShell = CreateObject("WScript.Shell")
Set ADSysInfo = CreateObject("ADSystemInfo")
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName)
strGroups = LCase(Join(CurrentUser.MemberOf))
REM WScript.Echo "This is test message, Hi"
REM For test show group names in Active Directory  WScript.Echo strGroups

If  InStr(strGroups, GRP_IT) Then
REM map shared folder on network to local drive Z:
wshNetwork.MapNetworkDrive "Z:",         "\\FILESERVER01\it"
wshNetwork.AddWindowsPrinterConnection   "\\PRINTSERVER\accounting"
wshNetwork.AddWindowsPrinterConnection   "\\PRINTSERVER\marketing"
wshNetwork.AddWindowsPrinterConnection   "\\PRINTSERVER\production"
wshNetwork.AddWindowsPrinterConnection   "\\PRINTSERVER\photocopy"
wshNetWork.SetDefaultPrinter             "\\PRINTSERVER\photocopy"
REM create shortcut on desktop for shared folder on network Z: 
REM Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Shared Documents.LNK")
oUrlLink.TargetPath = "\\FILESERVER01\it"
oUrlLink.Save
oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Web Mail.URL")
oUrlLink.TargetPath="http://webmail.aliveli.com"
oUrlLink.Save

ElseIf  InStr(strGroups, GRP_ACCOUNTING) Then
wshNetwork.MapNetworkDrive "Z:",         "\\FILESERVER01\accounting"
wshNetwork.AddWindowsPrinterConnection   "\\PRINTSERVER\accounting"
wshNetwork.AddWindowsPrinterConnection   "\\PRINTSERVER\photocopy"
wshNetWork.SetDefaultPrinter             "\\PRINTSERVER\accounting"
REM Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Shared Documents.LNK")
oUrlLink.TargetPath = "\\FILESERVER01\accounting"
oUrlLink.Save
oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Web Mail.URL")
oUrlLink.TargetPath ="http://webmail.aliveli.com"
oUrlLink.Save

ElseIf  InStr(strGroups, GRP_PRODUCTION) Then
wshNetwork.MapNetworkDrive "Z:",        "\\FILESERVER01\production"
wshNetwork.AddWindowsPrinterConnection  "\\PRINTSERVER\production"
wshNetwork.AddWindowsPrinterConnection  "\\PRINTSERVER\photocopy"
wshNetWork.SetDefaultPrinter            "\\PRINTSERVER\production"
REM Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Shared Documents.LNK")
oUrlLink.TargetPath = "\\FILESERVER01\production"
oUrlLink.Save
oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Web Mail.URL")
oUrlLink.TargetPath = "http://webmail.aliveli.com"
oUrlLink.Save
REM production department using too much Autocad, create Autocad shortcut on desktop
oMyShortCut= WshShell.CreateShortcut(strDesktop+"\AutoCad.lnk")
oMyShortCut.WindowStyle = 7  &&Minimized 0=Maximized  4=Normal
oMyShortCut.TargetPath = "C:\Program Files\Autodesk\AutoCAD\acad.exe"
oMyShortCut.Save

ElseIf  InStr(strGroups, GRP_MARKETING) Then
wshNetwork.MapNetworkDrive "Z:",        "\\FILESERVER01\marketing"
wshNetwork.AddWindowsPrinterConnection  "\\PRINTSERVER\marketing"
wshNetwork.AddWindowsPrinterConnection  "\\PRINTSERVER\photocopy"
wshNetWork.SetDefaultPrinter            "\\PRINTSERVER\marketing"
REM Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Shared Documents.LNK")
oUrlLink.TargetPath = "\\FILESERVER01\marketing"
oUrlLink.Save
oUrlLink = WshShell.CreateShortcut(strDesktop+"\Company Web Mail.URL")
oUrlLink.TargetPath = "http://webmail.aliveli.com"
oUrlLink.Save
End If

