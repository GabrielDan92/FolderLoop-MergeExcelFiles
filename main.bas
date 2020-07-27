Sub merge()
Dim path As String, lastRow As Long
    
    'clear the previous values from the destination sheet
    Sheet1.Range("A1").CurrentRegion.ClearFormats
    Sheet1.Range("A1").CurrentRegion.ClearContents
    
    'this is the path to the folder with the targeted .xlsx files
    path = "C:\Users\Gabriel\Desktop"
    
    'call the private Sub containing the loop and merge logic, with two arguments (path and the extension needed)
    Call folderFilesLoop(path, "*.xlsx*")
    
    MsgBox "The files have been imported!"

End Sub

'===================================================================================================================

Private Sub folderFilesLoop(ByVal strDir As String, strType As String)
Dim file As Variant, wb As Workbook, counter As Long, lastRow As Long, rng As Range, arr As Variant
    
    If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
    file = Dir(strDir & strType)
    counter = 1
    
On Error GoTo errHandler

    While (file <> "")
'        If Left(file, 4) = "CFR_" Then        'in addition to the loop, you can target specific files using their filename as criteria (with left/rigth/mid/instr/etc)
            Debug.Print "Opening template file: " & file
            Set wb = Workbooks.Open(strDir & file, 2)
            With wb.Sheets(1)
                If .AutoFilterMode Then .AutoFilterMode = False
                If counter = 1 Then
                    Set rng = .Range("A1").CurrentRegion
                    arr = rng
                    Sheet1.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
                Else
                    Set rng = .Range("A1").CurrentRegion.Offset(1)
                    arr = rng
                    Sheet1.Range("A" & lastRow + 1).Resize(UBound(arr, 1), UBound(arr, 2)) = arr
                End If
            End With
            Erase arr
            wb.Close (False)
            Set wb = Nothing
nextWhile:
            lastRow = Sheet1.Range("A1").CurrentRegion.Rows.Count
            If lastRow <> 1 Then counter = counter + 1
'        End If
        file = Dir
    Wend
    
Exit Sub
errHandler:
    MsgBox "Error found for file: " & wb.Name
    If IsEmpty(arr) = False Then Erase arr
    If wb Is Nothing = False Then
        wb.Close (False)
        Set wb = Nothing
    End If
    Resume nextWhile

End Sub
