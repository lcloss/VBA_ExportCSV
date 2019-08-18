Attribute VB_Name = "modExport"
Sub GravaCSV()
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim strPrf As String
    Dim intCount As Integer
    
    Set wkb = ActiveWorkbook
    strPrf = FileBaseName(wkb.Name)
    
    For Each wks In wkb.Worksheets
        With wks
            .SaveAs Filename:=strPrf & "_" & .Name & ".csv", FileFormat:=xlCSV
        End With
        intCount = intCount + 1
    Next wks
    
    MsgBox "Finalizado com " & intCount & " ficheiros exportados."
End Sub

Private Function FileBaseName(strFileName As String) As String
    Dim dotFound As Boolean
    Dim intInd As Integer
    
    dotFound = False
    intInd = Len(strFileName)
    While (dotFound = False And intInd > 1)
        If Mid(strFileName, intInd, 1) = "." Then
            dotFound = True
        End If
        intInd = intInd - 1
    Wend
    
    If dotFound Then
        FileBaseName = Left(strFileName, intInd)
    Else
        FileBaseName = strFileName
    End If
End Function
