Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Author       : Mustafa Can Ozturk                                                                                      '
'      Purpose      : This script opens excel file, deletes rows that contain a keyword in Column E and F                     '
'      Input        : 3 inputs such as path of excel file, name of excel file, and keyword to be deleted                      '
'      Output       : "SUCCESS" or "FAIL" message                                                                             '
'      Demonstration: Open terminal and type => cscript Delete.vbs "C:\Users\moeztuerk\Desktop\" "Deneme.xlsx" "Accrual"      '
'      Date         : 05.08.2021                                                                                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim excelFilePath, excelFileName, XLApp, wb, ws, lr, i, keyword

On Error Resume Next

    excelFilePath = Wscript.Arguments.Item(0)
    excelFileName = Wscript.Arguments.Item(1)
    keyword = Wscript.Arguments.Item(2)

    Set XLApp = CreateObject("Excel.Application")

    XLApp.DisplayAlerts = True

    const xlUp = -4162
    set wb = XLApp.Workbooks.Open(excelFilePath + excelFileName)
    if wb.Application.ProtectedViewWindows.Count > 0 then
        wb.Application.ActiveProtectedViewWindow.Edit
    end if

    set ws = wb.Sheets("Sheet1")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = lr to 1 Step -1

        if trim(ws.range("E" & i)) = keyword then
            ws.range("E" & i).EntireRow.Delete
        end if

        if instr(ws.range("F" & i), keyword) > 0 then
            ws.range("F" & i).EntireRow.Delete
        end if

    Next

    wb.SaveAs excelFilePath + Left(excelFileName, len(excelFileName) - 5) + " edited.xlsx"  

if Err.Number = 0 then
    Wscript.StdOut.Write("SUCCESS")
else
    Wscript.StdOut.Write("FAIL")
end if

    wb.Close False
    XLApp.Quit
    set XLApp = nothing

On Error GoTo 0
