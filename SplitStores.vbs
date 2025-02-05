Sub testfirst()
    Dim objExcel, objWorkbook, objSheet, objFSO, outputFolder, storeList, objLogFile
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False 
    Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Leonid.ksenchuk\Desktop\TaskArtur2\testData (4).xlsx")
    Set objSheet = objWorkbook.Sheets(1)
    outputFolder = "C:\Users\Leonid.ksenchuk\Desktop\TaskArtur2\output"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(outputFolder) Then objFSO.CreateFolder outputFolder
    Set storeList = CreateObject("System.Collections.ArrayList")
    Set objLogFile = objFSO.CreateTextFile(outputFolder & "\log.txt", True)

    Dim row, store
    For row = 2 To GetLastRow(objSheet)
        store = Trim(objSheet.Cells(row, 4).Value)
        If Not storeList.Contains(store) And store <> "" Then
            storeList.Add store
        End If
    Next
    LogMessage objLogFile, "Total unique stores: " & storeList.Count

    Dim objNewWorkbook, objNewSheet, col, newRow
    For Each store In storeList
        Set objNewWorkbook = objExcel.Workbooks.Add
        Set objNewSheet = objNewWorkbook.Sheets(1)

        For col = 1 To objSheet.UsedRange.Columns.Count
            objNewSheet.Cells(1, col).Value = objSheet.Cells(1, col).Value
        Next

        newRow = 2
        For row = 2 To GetLastRow(objSheet)
            If Trim(objSheet.Cells(row, 4).Value) = store Then
                For col = 1 To objSheet.UsedRange.Columns.Count
                    objNewSheet.Cells(newRow, col).Value = objSheet.Cells(row, col).Value
                    If col = 6 Then objNewSheet.Cells(newRow, col).Value = Round(objSheet.Cells(row, col).Value, 3)
                Next
                newRow = newRow + 1
            End If
        Next

        RemoveEmptyColumns objNewSheet

        If newRow > 2 Then
            objNewWorkbook.SaveAs outputFolder & "\" & store & ".xlsx"
            objNewWorkbook.Close False
            LogMessage objLogFile, "File created: " & store & ".xlsx"
        Else
            LogMessage objLogFile, "File for " & store & " is empty, not created."
        End If
    Next

    objWorkbook.Close False
    objExcel.Quit
    objLogFile.Close
End Sub

Function GetLastRow(sheet)
    GetLastRow = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
End Function

Sub RemoveEmptyColumns(sheet)
    Dim col As Integer, lastRow As Integer, isEmpty As Boolean
    lastRow = GetLastRow(sheet)
    
    For col = sheet.UsedRange.Columns.Count To 1 Step -1
        isEmpty = True
        For row = 2 To lastRow
            If sheet.Cells(row, col).Value <> "" Then
                isEmpty = False
                Exit For
            End If
        Next
        If isEmpty Then sheet.Columns(col).Delete
    Next
End Sub

Sub LogMessage(logFile, message)
    logFile.WriteLine Now & " - " & message
End Sub
