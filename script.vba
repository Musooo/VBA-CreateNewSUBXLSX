Sub MacroName()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim uniqueNames As Variant
    Dim newBook As Workbook
    Dim newWS As Worksheet
    Dim name As Variant
    Dim outputFolder As String
    Dim currentName As String
    Dim lastRow As Long
    
    outputFolder = ThisWorkbook.Path & "\output_files\"
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder
    End If
    
    Set ws = ThisWorkbook.Sheets(1)
    ws.Columns.AutoFit
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A2:C" & lastRow).SpecialCells(xlCellTypeVisible)
    
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In rng.Columns(1).Cells
        currentName = cell.Value
        If Not dict.exists(currentName) And currentName <> "" Then
            dict.Add currentName, Nothing
        End If
    Next cell
    
    uniqueNames = dict.keys
    
    For Each name In uniqueNames
        Set newBook = Workbooks.Add
        Set newWS = newBook.Sheets(1)
        
        ws.Rows(1).Copy Destination:=newWS.Rows(1)
        
        For Each cell In rng.Columns(1).Cells
            If cell.Value = name Then
                cell.EntireRow.Copy Destination:=newWS.Cells(newWS.Cells(newWS.Rows.Count, "A").End(xlUp).Row + 1, 1)
            End If
        Next cell
        newWS.Columns.AutoFit
        
        newBook.SaveAs Filename:=outputFolder & name & ".xlsx"
        newBook.Close SaveChanges:=False
    Next name
    
    MsgBox "File separati creati nella cartella: " & outputFolder

End Sub

