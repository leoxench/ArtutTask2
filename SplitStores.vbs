Sub testfirst()
    Dim objExcel, objWorkbook, objSheet, objFSO, outputFolder, storeDict, objLogFile
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False 
    Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Leonid.ksenchuk\Desktop\TaskArtur2\testData (4).xlsx")
    Set objSheet = objWorkbook.Sheets(1)
    outputFolder = "C:\Users\Leonid.ksenchuk\Desktop\TaskArtur2\output\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(outputFolder) Then objFSO.CreateFolder outputFolder
    Set storeDict = CreateObject("Scripting.Dictionary")
    Set objLogFile = objFSO.CreateTextFile(outputFolder & "log.txt", True)

    ' Збір унікальних магазинів
    Dim row, store
    For row = 2 To GetLastRow(objSheet)
        store = Trim(objSheet.Cells(row, 4).Value)
        If Not storeDict.Exists(store) And store <> "" Then
            storeDict.Add store, True
        End If
    Next
    LogMessage objLogFile, "Total unique stores: " & storeDict.Count

    ' Експорт даних для кожного магазину
    Dim objNewWorkbook, objNewSheet, col, newRow, objTxtFile
    For Each store In storeDict.Keys
        Set objNewWorkbook = objExcel.Workbooks.Add
        Set objNewSheet = objNewWorkbook.Sheets(1)

        ' Запис заголовків
        For col = 1 To GetLastCol(objSheet)
            objNewSheet.Cells(1, col).Value = objSheet.Cells(1, col).Value
        Next

        newRow = 2
        For row = 2 To GetLastRow(objSheet)
            If Trim(objSheet.Cells(row, 4).Value) = store Then
                For col = 1 To GetLastCol(objSheet)
                    objNewSheet.Cells(newRow, col).Value = objSheet.Cells(row, col).Value
                    If col = 6 Then objNewSheet.Cells(newRow, col).Value = Round(objSheet.Cells(row, col).Value, 3)
                Next
                newRow = newRow + 1
            End If
        Next

        If newRow > 2 Then
            objNewWorkbook.SaveAs outputFolder & store & ".xlsx"
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

' Функція отримання останнього рядка
Function GetLastRow(sheet)
    GetLastRow = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
End Function

' Функція отримання останньої колонки
Function GetLastCol(sheet)
    GetLastCol = sheet.Cells(1, sheet.Columns.Count).End(-4159).Column
End Function

' Функція логування
Sub LogMessage(logFile, message)
    logFile.WriteLine Now & " - " & message
End Sub
