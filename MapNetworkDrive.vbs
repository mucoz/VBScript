Option Explicit
Dim objNetwork, strDriveLetter, strRemotePath, WshShell, username, password
 
' Set credentials & network share to variables.
strDriveLetter = "F:"
strRemotePath = "\\corp.hds.com\dfs01\krk"
 
username = Inputbox("Enter your username", "Username")
if typename(username) = "Empty" then WScript.Quit
password = Inputbox("Enter your password", "Password")
if typename(password) = "Empty" then WScript.Quit

' Create a network object (objNetwork) do apply MapNetworkDrive Z:
Set objNetwork = WScript.CreateObject("WScript.Network")
objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, True, username, password
 
' Open message box, enable remove the apostrophe at the beginning.
' WScript.Echo "Map Network Drive " & strDriveLetter
MsgBox " Explorer launch Network Drive " & strDriveLetter, vbInformation, "Network Drive Mapping"
' Explorer will open the mapped network drive.
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "explorer.exe /e," & strDriveLetter, 1, false
WScript.Quit
