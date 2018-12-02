Option Explicit

Public Function ProcessMergedSheet()
    Dim MergedSheet As Worksheet
    Dim i As Long
    Dim CellData As String
    
    Set MergedSheet = ActiveWorkbook.Sheets("MergedSheet")
    MergedSheet.Activate
    
    For i = 1 To 100 'MergedSheet.UsedRange.Rows.Count
        CellData = CStr(MergedSheet.Cells(i, 2)) 'go down each row, getting the data in the second column (b)
    
        If StringIsFound("Store", CellData) Then
            Debug.Print "ayy on " + CStr(i)
            If RowIsBlank(i) Then
            'If IsEmpty(Range("A" + CStr(i))) Then
                Debug.Print "zz"
            End If
        End If
        
        Next i
End Function

 
Sub ProcessSalesTransactions()
    Debug.Print ("Starting...")
    
    Call MySetup
    'Debugging = False
    
    If Not WorksheetExists("MergedSheet") Then
        Call MergeSheets(ReturnSheetNames("Sheet"), "MergedSheet")
    End If
    
    If Not WorksheetExists("MergedSheet") Then
        Call Err.Raise(1, "SalesTransactions", "Error creating merged sheet")
    End If
    
    Call ProcessMergedSheet
    
End Sub


Sub test()
End Sub

Function test2()
    
End Function
