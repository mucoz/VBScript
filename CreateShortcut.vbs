'******************************************************
'Script Name : CreateShortcut.vbs
'Author : Mustafa Can Ozturk
'Created : 25.06.2021
'Description : This script creates a description for another script
'******************************************************

dim objWshShell, DesktopPath, objNewShortcut

set objWshShell = WScript.CreateObject("WScript.Shell")

DesktopPath = objWshShell.SpecialFolders("Desktop")

set objNewShortcut = objWshShell.CreateShortcut(DesktopPath & _
                    "\\Yeni Kisayol.lnk")

objNewShortcut.TargetPath = "C:\Users\moeztuerk\Desktop\VBS\CreateDescription.vbs"
'objNewShortcut.Description = "This script creates a description for another script"
'objNewShortcut.HotKey = "Ctrl+Alt+G"

objNewShortcut.Save()
