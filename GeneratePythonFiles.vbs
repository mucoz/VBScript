Dim objFSO, objFile, strScriptPath, strReqFilePath
Dim strRunRobotFilePath, strInstallerFilePath, strPythonMainFilePath
Dim strRunRobotContent, strInstallerContent, strPythonMainContent
Dim arrLibraries, strLibrary, strVersion
Dim allLibrariesFound

Call GenerateFiles

Sub GenerateFiles()
    If Not initializationSuccessful() Then 
        WScript.Echo "Initialization failed."
        WScript.Quit
    End If

    If Not requirementsFound() Then 
        WScript.Echo "Requirements file not found."
        WScript.Quit
    End If

    If Not requirementsRead() Then 
        WScript.Echo "Requirements could not be read."
        WScript.Quit
    End If

    If Not requirementsParsed() Then 
        WScript.Echo "Requirements could not be parsed."
        WScript.Quit
    End If

    If Not filesGenerated() Then 
        WScript.Echo "Files could not be generated."
        WScript.Quit
    End If

    WScript.Echo "Process completed successfully."
End Sub

Private Function initializationSuccessful()
    On Error Resume Next
    result = False

    ' Create a FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Get the path of the script
    strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

    ' Assign the file paths
    strReqFilePath = strScriptPath & "\requirements.txt"
    strRunRobotFilePath = strScriptPath & "\Launcher.bat"
    strInstallerFilePath = strScriptPath & "\Installer.bat"
    strPythonMainFilePath = strScriptPath & "\main_test.py"

    ' Generate content for RunRobot.bat file
    strRunRobotContent = "CALL C:\Programs\Miniconda3_x64\Scripts\activate.bat" & vbCrLf & _
                         "CALL conda activate ""%~dp0app_env""" & vbCrLf & _
                         "cmd /k CALL python main_test.py"

    ' Generate content for Installer.bat file
    strInstallerContent = "CALL C:\Programs\Miniconda3_x64\Scripts\activate.bat" & vbCrLf & _
                          "set ENV_NAME=app_env""" & vbCrLf & _
                          "set FOLDER_PATH=""%~dp0" & vbCrLf & _
                          "CALL conda create -y -p %FOLDER_PATH%%ENV_NAME% python=3.7" & vbCrLf & _
                          "CALL conda activate %FOLDER_PATH%%ENV_NAME%" & vbCrLf & _
                          "pip install -r requirements.txt" & vbCrLf & _
                          "DEL *.vbs" & vbCrLf & _
                          "cmd /k echo Environment created and libraries installed successfully."

    ' Generate content of the Python test file (main_test.py)
    strPythonMainContent = "print('TEST')"

    ' Debugging the content after assignment
    ' WScript.Echo "RunRobot Content: " & strRunRobotContent
    ' WScript.Echo "Installer Content: " & strInstallerContent
    ' WScript.Echo "Python Main Content: " & strPythonMainContent

    ' Check if everything is initialized properly
    If Err.Number <> 0 Then
        WScript.Echo "Failed to initialize the script! Error: " & Err.Description & " (Error Code: " & Err.Number & ")"
    Else
        WScript.Echo "Initialized the script successfully."
        result = True
    End If

    On Error GoTo 0
    initializationSuccessful = result
End Function

Private Function requirementsFound()
    ' Check if requirements.txt exists
    On Error Resume Next
    result = False
    If Not objFSO.FileExists(strReqFilePath) Then
        WScript.Echo "requirements.txt file not found!"
    Else
        WScript.Echo "Requirements file found."
        result = True
    End If
    On Error GoTo 0
    requirementsFound = result
End Function

Private Function requirementsRead()
    ' Read requirements.txt file
    On Error Resume Next
    result = False
    Set objFile = objFSO.OpenTextFile(strReqFilePath)
    strContents = objFile.ReadAll
    objFile.Close
    If Err.Number = 0 Or Err.Number = 62 Then
        WScript.Echo "Collected requirements successfully."
        result = True
    Else
        WScript.Echo "Error occurred while reading requirements."
    End If
    On Error GoTo 0
    requirementsRead = result
End Function

Private Function requirementsParsed()
    On Error Resume Next
    result = False
    ' If requirements.txt is empty, ask the user whether to continue
    If Len(strContents) = 0 Then
        answer = MsgBox("requirements.txt file is empty. Do you want to continue?", vbYesNo, "Empty File")
        If answer = vbYes Then
            result = True
        End If
    Else
        ' Split the contents by line
        arrLibraries = Split(strContents, vbCrLf)
        ' Flag to track if all required libraries are found
        allLibrariesHaveVersions = True
        ' Loop through each line in requirements.txt
        For Each line In arrLibraries
            If Trim(line) <> "" Then
                If InStr(line, "==") = 0 Then
                    allLibrariesHaveVersions = False
                Else
                    ' Split the line by '=='
                    arrLine = Split(line, "==")
                    strLibrary = arrLine(0)
                    strVersion = arrLine(1)
                    ' Check if the library with correct version
                    If Trim(strLibrary) = "" Or Trim(strVersion) = "" Then
                        allLibrariesHaveVersions = False
                    End If
                End If
            End If
        Next
        If Not allLibrariesHaveVersions Then
            WScript.Echo "Warning: All libraries in requirements must have versions!"
        Else
            result = True
        End If
    End If
    If Err.Number <> 0 Then
        WScript.Echo "Error occurred while parsing requirements file!"
        result = False
    Else
        WScript.Echo "File parsed successfully."
    End If
    On Error GoTo 0
    requirementsParsed = result
End Function

Private Function filesGenerated()
    On Error Resume Next
    result = False

    ' Debug content before writing
    ' WScript.Echo "RunRobot Content: " & vbCrLf & strRunRobotContent
    ' WScript.Echo "Installer Content: " & vbCrLf & strInstallerContent
    ' WScript.Echo "Python Main Content: " & vbCrLf & strPythonMainContent

    ' Create RunRobot.bat file
    If Len(strRunRobotFilePath) > 0 Then
        WScript.Echo "Generating " & strRunRobotFilePath
        Set objFile = objFSO.CreateTextFile(strRunRobotFilePath)
        If Err.Number <> 0 Then
            WScript.Echo "Error creating " & strRunRobotFilePath & ": " & Err.Description & " (Error Code: " & Err.Number & ")"
            Exit Function
        End If
        objFile.Write strRunRobotContent
        objFile.Close
    Else
        WScript.Echo "Error: strRunRobotFilePath is empty!"
        Exit Function
    End If

    ' Create Installer.bat file
    If Len(strInstallerFilePath) > 0 Then
        WScript.Echo "Generating " & strInstallerFilePath
        Set objFile = objFSO.CreateTextFile(strInstallerFilePath)
        If Err.Number <> 0 Then
            WScript.Echo "Error creating " & strInstallerFilePath & ": " & Err.Description & " (Error Code: " & Err.Number & ")"
            Exit Function
        End If
        objFile.Write strInstallerContent
        objFile.Close
    Else
        WScript.Echo "Error: strInstallerFilePath is empty!"
        Exit Function
    End If

    ' Create main.py file
    If Len(strPythonMainFilePath) > 0 Then
        WScript.Echo "Generating " & strPythonMainFilePath
        Set objFile = objFSO.CreateTextFile(strPythonMainFilePath)
        If Err.Number <> 0 Then
            WScript.Echo "Error creating " & strPythonMainFilePath & ": " & Err.Description & " (Error Code: " & Err.Number & ")"
            Exit Function
        End If
        objFile.Write strPythonMainContent
        objFile.Close
    Else
        WScript.Echo "Error: strPythonMainFilePath is empty!"
        Exit Function
    End If

    If Err.Number <> 0 Then
        WScript.Echo "Error occurred while generating files: " & Err.Description & " (Error Code: " & Err.Number & ")"
    Else
        WScript.Echo "Generated files successfully."
        result = True
    End If
    On Error GoTo 0
    filesGenerated = result
End Function
