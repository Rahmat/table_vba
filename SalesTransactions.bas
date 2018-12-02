Option Explicit

Public Function MergeTransactionSheets()
    If Not WorksheetExists("MergedSheet") Then
        Call MergeSheets(ReturnSheetNames("Sheet"), "MergedSheet")
    End If
    
    If Not WorksheetExists("MergedSheet") Then
        Call Err.Raise(1, "SalesTransactions", "Error creating merged sheet")
    End If
End Function

Public Function ProcessMergedSheet()
    Dim MergedSheet As Worksheet
    Dim i As Long
    Dim CellData As String
    Dim t As Range
    t.FormulaR1C1
    Set MergedSheet = ActiveWorkbook.Sheets("MergedSheet")
    MergedSheet.Activate
    
    For i = 1 To 100 'MergedSheet.UsedRange.Rows.Count
        CellData = CStr(MergedSheet.Cells(i, 2)) 'go down each row, getting the data in the second column (b)
    
        If StringIsFound("Store", CellData) Then
            Debug.Print "'Store' found at line " + CStr(i)
            If RowIsBlank(i) Then
            'If IsEmpty(Range("A" + CStr(i))) Then
                Debug.Print "Also, this row blank!"
            End If
        End If
        
        Next i
End Function

Sub ProcessSalesTransactions()
    Debug.Print ("Starting! " + CStr(Now))
    
    Call MySetup    'Debugging = False
    
    Call MergeTransactionSheets
    Call ProcessMergedSheet
End Sub


Sub test()
    Debug.Print ("Is blank?: " + CStr(RowIsBlank(13)))
End Sub

Function test2()
    
End Function
