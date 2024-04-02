Attribute VB_Name = "NewMacros"


Private Sub Document_Close()
'
' Save the Document in PDF in folder .\PDF, except if it's the template
'
    DocName = ActiveDocument.Name
    DocNameSplitted = Split(DocName, ".") ' Remove extension
    DocNameSplitted = Split(DocNameSplitted(0), "_") ' Remove extension
    
    ' Check if the document name contains "[template]"
    If InStr(DocName, "[template]") > 0 Or InStr(DocName, "_") < 3 Then
        ' Exit the function if the document name contains "[template]"
        Exit Sub
    End If

    ' Define the folder path for PDF export
    pdfFolderPath = ThisDocument.Path & "\PDF\"
    pdfPath = pdfFolderPath & DocNameSplitted(0) & " " & DocNameSplitted(1) & ".pdf"
    
    ' Create the PDF folder if it doesn't exist
    If Dir(pdfFolderPath, vbDirectory) = "" Then
        MkDir pdfFolderPath
    End If
    
    ' Save as PDF in a specific folder
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        pdfPath _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    


End Sub

Private Sub Document_Open()
'
' Retrieve_name_and_update_properties based on the folliwing syntax:
' DocName = NOM_PRENOM_CLASSE__T1L1Status_T1L2Status_...__T2L1Status_T1L2Status.docm
'
    
    DocName = ActiveDocument.Name
    
    ' Check if the document name contains "[template]"
    If InStr(DocName, "[template]") > 0 Or InStr(DocName, "_") < 3 Then
        ' Exit the function if the document name contains "[template]"
        Exit Sub
    End If
    
    DocNameSplitted = Split(DocName, ".") ' Remove extension
    ' DocNameSplitted = Split(DocNameSplitted(0), "_")
    DocNameSplitted = Split(DocNameSplitted(0), "__")
    DocNameSplittedTable1 = Split(DocNameSplitted(1), "_")
    DocNameSplittedTable2 = Split(DocNameSplitted(2), "_")
    
    
    Dim NbLinesinTable1
    Dim NbLinesinTable2
    
    '-------------------------
    NbLinesinTable1 = ActiveDocument.Tables(1).Rows.Count - 1
    NbLinesinTable2 = ActiveDocument.Tables(2).Rows.Count - 1
    '-------------------------
    
    
    ' Result (OK, NOK) for each line.
    ' e.g. Result(0,0) = "X" if line 1 is OK
    ' e.g. Result(0,1) = "" if line 1 is OK
    ' e.g. Result(3,1) = "X" if line 4 is NOK
    Dim Result1()
    ReDim Result1(NbLinesinTable1 - 1, 2)
    
    
    ' Write a cross in the right column. 0=Success / 1=Failed / 2=Absent
    Dim i As Integer
    For i = 0 To NbLinesinTable1 - 1
        If DocNameSplittedTable1(i) = 0 Then
            Result1(i, 0) = ""
            Result1(i, 1) = "X"
        ElseIf DocNameSplittedTable1(i) = 1 Then
            Result1(i, 0) = "X"
            Result1(i, 1) = ""
        ElseIf DocNameSplittedTable1(i) = 2 Then
            Result1(i, 0) = ""
            Result1(i, 1) = "ABS"
        End If
        
        Debug.Print "Result1(" & i & ", 0) : " & Result1(i, 0)
        Debug.Print "Result1(" & i & ", 1) : " & Result1(i, 1)
        Debug.Print "-------"
        
    Next i
    
    
    Dim Result2()
    ReDim Result2(NbLinesinTable2 - 1, 2)
    For i = 0 To NbLinesinTable2 - 1
            If DocNameSplittedTable2(i) = 0 Then
            Result2(i, 0) = ""
            Result2(i, 1) = "X"
        ElseIf DocNameSplittedTable2(i) = 1 Then
            Result2(i, 0) = "X"
            Result2(i, 1) = ""
        ElseIf DocNameSplittedTable2(i) = 2 Then
            Result2(i, 0) = ""
            Result2(i, 1) = "ABS"
        End If
        
        Debug.Print "Result2(" & i & ", 0) : " & Result2(i, 0)
        Debug.Print "Result2(" & i & ", 1) : " & Result2(i, 1)
        Debug.Print "-------"
        
    Next i
    
    
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("NOM").Value = DocNameSplitted(0)
    ActiveDocument.CustomDocumentProperties("PRENOM").Value = DocNameSplitted(1)
    ActiveDocument.CustomDocumentProperties("CLASSE").Value = DocNameSplitted(2)
    ActiveDocument.CustomDocumentProperties("T1_L1_OK").Value = Result1(0, 0)
    ActiveDocument.CustomDocumentProperties("T1_L1_NOK").Value = Result1(0, 1)
    ActiveDocument.CustomDocumentProperties("T1_L1_OK").Value = Result1(0, 0)
    ActiveDocument.CustomDocumentProperties("T1_L1_NOK").Value = Result1(0, 1)
    ActiveDocument.CustomDocumentProperties("T1_L2_OK").Value = Result1(1, 0)
    ActiveDocument.CustomDocumentProperties("T1_L2_NOK").Value = Result1(1, 1)
    ActiveDocument.CustomDocumentProperties("T1_L3_OK").Value = Result1(2, 0)
    ActiveDocument.CustomDocumentProperties("T1_L3_NOK").Value = Result1(2, 1)
    ActiveDocument.CustomDocumentProperties("T1_L4_OK").Value = Result1(3, 0)
    ActiveDocument.CustomDocumentProperties("T1_L4_NOK").Value = Result1(3, 1)
    ActiveDocument.CustomDocumentProperties("T1_L5_OK").Value = Result1(4, 0)
    ActiveDocument.CustomDocumentProperties("T1_L5_NOK").Value = Result1(4, 1)
    
    ActiveDocument.CustomDocumentProperties("T2_L1_OK").Value = Result2(0, 0)
    ActiveDocument.CustomDocumentProperties("T2_L1_NOK").Value = Result2(0, 1)
    ActiveDocument.CustomDocumentProperties("T2_L2_OK").Value = Result2(1, 0)
    ActiveDocument.CustomDocumentProperties("T2_L2_NOK").Value = Result2(1, 1)
    ActiveDocument.CustomDocumentProperties("T2_L3_OK").Value = Result2(2, 0)
    ActiveDocument.CustomDocumentProperties("T2_L3_NOK").Value = Result2(2, 1)
    ActiveDocument.CustomDocumentProperties("T2_L4_OK").Value = Result2(3, 0)
    ActiveDocument.CustomDocumentProperties("T2_L4_NOK").Value = Result2(3, 1)
    ActiveDocument.CustomDocumentProperties("T2_L5_OK").Value = Result2(4, 0)
    ActiveDocument.CustomDocumentProperties("T2_L5_NOK").Value = Result2(4, 1)

End Sub











