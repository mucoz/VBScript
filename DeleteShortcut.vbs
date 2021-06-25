'******************************************************
'Script Name : DeleteShortcut.vbs
'Author : Mustafa Can Ozturk
'Created : 25.06.2021
'Description : This script deletes a shortcut object from a folder
'******************************************************

dim objWshShell, DesktopPath, FSO, ShortCut 

set objWshShell = WScript.CreateObject("WScript.Shell")

DesktopPath = objWshShell.SpecialFolders("Desktop")

set FSO = CreateObject("Scripting.FileSystemObject")
set ShortCut = FSO.GetFile(DesktopPath & _
                    "\\Yeni Kisayol.lnk")

ShortCut.Delete

