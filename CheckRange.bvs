Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Author  : Mustafa Can Ozturk                                                                   '
'      Purpose : This script checks values in General Data file for the range from column R to X      '
'              : If there is any value inside the range, it returns True                              '
'      Input   : General data filepath                                                                '
'      Output  : True, False, Error                                                                   '
'      Date    : 21.09.2021                                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Dim generalPath, XLApp, wb, ws, lr, rng, cell, isOccupied

On Error Resume Next

    generalPath = Wscript.Arguments.Item(0)

    Const xlUp = -4162
    Set XLApp = CreateObject("Excel.Application")
    XLApp.DisplayAlerts = False

    Set wb = XLApp.Workbooks.Open(generalPath)

    Set ws = wb.Sheets("General Data")

    lr = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    Set rng = ws.Range("R2:X" & lr)

    isOccupied = False

    For Each cell in rng

        If Trim(cell.Value) <> "" Then
            isOccupied = True
            Exit For
        End If

    Next

    If Err.Number = 0 Then
        Wscript.StdOut.Write(isOccupied)
    Else
        Wscript.StdOut.Write("ERROR : " & Err.Number & "; " & Err.Description)
    End If

    wb.Close False
    XLApp.Quit

    Set cell = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set XLApp = Nothing

On Error GoTo 0
