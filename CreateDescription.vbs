'******************************************************
'Script Name : CreateShortcut.vbs
'Author : Mustafa Can Ozturk
'Created : 25.06.2021
'Description : This script creates a description for another script
'******************************************************

Option Explicit

Dim objWshArgs, Name, Auth, Desc

set objWshArgs = Wscript.Arguments

If objWshArgs.Count <> 3 then
    Wscript.Echo "To be able to use this script, You need to specify Script Name, Author, and Description"
    Wscript.Quit
End If

Name = objWshArgs.Item(0)
Auth = objWshArgs.Item(1)
Desc = objWshArgs.Item(2)

Wscript.Echo "'" & "******************************************************" & vbCrLf & _
             "'Script Name : " & Name & vbCrLf & _
             "'Author : " & Auth & vbCrLf & _
             "'Created : " & FormatDateTime(Now, vbShortDate) & vbCrLf & _
             "'Description : " & Desc & vbCrLf & _
             "'" & "******************************************************"


set objWshArgs = Nothing

Wscript.Quit
