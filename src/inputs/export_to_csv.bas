Attribute VB_Name = "Module1"

Sub ExportToCSV()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim filePath As String

    ' Define the folder path for CSV export
    folderPath = ThisWorkbook.Path & "\..\outputs\"
    filePath = folderPath & "Data.csv"

    ' Ask for confirmation to proceed
    confirmation = MsgBox("Your file will be saved and data will be exported to " & filePath & ". Do you want to continue?", vbYesNo + vbQuestion, "Confirmation")
    
    ' Check if the user confirmed
    If confirmation = vbNo Then
        Exit Sub
    End If

    ' Suppress the warning message about personal information
    ThisWorkbook.RemovePersonalInformation = True
    ThisWorkbook.Save

    ' Custom functions : clean the last worksheet and copy the relevant data from previous pages
    ClearLastWorksheetCells
    CopyRangesToSheet

    ' Create the output folder if it doesn't exist
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Get the last worksheet
    Set ws = Worksheets(Worksheets.Count)
    
    ' Create a new workbook
    Set newWb = Workbooks.Add
    Set newWs = newWb.Worksheets(1)
    
    ' Create a copy of the worksheet with values only
    ws.Cells.Copy
    newWs.Cells.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Clear clipboard
    
    ' Disable display alerts to suppress confirmation dialogs
    Application.DisplayAlerts = False
    
    ' Export data to CSV without opening the file
    newWs.SaveAs Filename:=filePath, FileFormat:=xlCSV
    
    ' Re-enable display alerts
    Application.DisplayAlerts = True
    
    ' Close the Data.csv
    newWb.Close SaveChanges:=False

    ' Clear the contents of all cells
    ClearLastWorksheetCells
    
    ' Notify user
    MsgBox "Data exported to " & filePath, vbInformation
    
End Sub

Sub ClearLastWorksheetCells()
    Dim ws As Worksheet
    
    ' Check if there are any worksheets in the workbook
    If Worksheets.Count = 0 Then
        MsgBox "There are no worksheets in this workbook.", vbExclamation
        Exit Sub
    End If
    
    ' Get the last worksheet
    Set ws = Worksheets(Worksheets.Count)
    
    ' Clear the contents of all cells
    ws.Cells.Clear
End Sub

Sub CopyRangesToSheet()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim targetWorksheet As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    ' Set the source and target workbooks and worksheets
    Set sourceWorkbook = ActiveWorkbook
    Set targetWorkbook = ActiveWorkbook
    Set sourceWorksheet = sourceWorkbook.Worksheets(1) ' Change to the first worksheet
    Set targetWorksheet = targetWorkbook.Worksheets(targetWorkbook.Sheets.Count) ' Change to the fifth worksheet
    ' Set targetWorksheet = targetWorkbook.Worksheets(5) ' Change to the fifth worksheet
    
    ' Loop through the first 4 worksheets
    For i = 1 To targetWorkbook.Sheets.Count - 1
        Set sourceWorksheet = sourceWorkbook.Worksheets(i)

        ' Loop the rows on "name" column
        For j = 4 To sourceWorksheet.Cells(Rows.Count, "E").End(xlUp).Row

            ' Write the class
            targetWorksheet.Cells(targetWorksheet.Rows.Count, 1).End(xlUp).Offset(1, 0) = sourceWorksheet.Cells(2, 2).Value
            
            ' Loop the columns E:Q
            For k = 5 To 17

                Set sourceRange = sourceWorksheet.Cells(j, k)
                Set targetRange = targetWorksheet.Cells(targetWorksheet.Rows.Count, k - 3).End(xlUp).Offset(1, 0)
                
                If IsEmpty(sourceWorksheet.Cells(j, k)) Then
                    targetRange.Value = "-"
                Else
                    ' Copy the value from the source range to the target range
                    targetRange.Value = sourceRange.Value
                End If

            Next k
        Next j
    Next i
End Sub

Sub ResetResults()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    ' Display a confirmation box with "Yes" and "No" options
    confirmation = MsgBox("Are you sure you want to reset all notes ?", vbYesNo + vbQuestion, "Confirmation")
    
    If confirmation = vbYes Then
        ' Set the source and target workbooks and worksheets
        Set wb = ActiveWorkbook
        
        ' Loop through the first 4 worksheets
        For i = 1 To wb.Sheets.Count - 1
            Set ws = wb.Worksheets(i)
    
    
            ' Loop the rows on "name" column
            For j = 2 To ws.Cells(Rows.Count, "E").End(xlUp).Row
                ' Loop the columns E:Q
                For k = 8 To 17
    
                    If j < 5 Then
                        ws.Cells(j, k) = ""
                    ElseIf j > 5 Then
                        ws.Cells(j, k) = "-"
                    End If
                            
                Next k
            Next j
        Next i
        
        MsgBox "Notes successfully reset on all pages !"
    End If
End Sub

