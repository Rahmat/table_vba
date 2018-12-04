Option Explicit

Public Function MergeTransactionSheets()
    Call MergeSheets("MergedSheet") ', ReturnSheetNames("Sheet"))
    
    If Not WorksheetExists("MergedSheet") Then
        Call Err.Raise(1, "SalesTransactions", "Error creating merged sheet")
    End If
    
    ActiveWorkbook.Sheets("MergedSheet").Activate
End Function

Public Function ProcessMergedSheet()
    Dim MergedSheet As Worksheet
    Dim i As Long
    Dim CellData As String

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

Todo: make this? later...if needed.
'Public Function AddDataToSheet(MyData As Collection, MySheet As Worksheet)
'    Dim t As Long
'
'    GetLastRow (MySheet.Name)
'
'    MySheet.
'End Function

Public Function ProcessMergedSheetMichael()
    Dim MergedSheet As Worksheet
    Dim i As Long, j As Long
    Dim CellData As String, StoreCoupon As String
    Dim StoreCoupons As New Collection
    'Dim StoreCouponFound As Boolean

    Set MergedSheet = ActiveWorkbook.Sheets("MergedSheet")
    MergedSheet.Activate
    
    For i = 1 To MergedSheet.UsedRange.Rows.Count
        CellData = CStr(MergedSheet.Cells(i, 1)) 'go down each row, getting the data in the second column (b)
    
        If StringIsFound("store coupon", CellData) Then
            If RowIsBlank(i) Then
                Debug.Print "Also, this row blank! (dafuq?)"
            End If
            
            For j = 1 To NumberOfColumns(i)
                    StoreCoupon = MergedSheet.Cells(i, j)
                    If IsNumeric(StoreCoupon) Then
                        Debug.Print ("Store Coupon found!! row #" + CStr(i) + " column #" + CStr(j))
                        StoreCoupons.Add (CStr(StoreCoupon))
                        Exit For
                    End If
            Next j
        End If
    Next i
    
    Debug.Print ("STORE COUPONS:")
    For i = 1 To StoreCoupons.Count
        Debug.Print (StoreCoupons.Item(i))
    Next i
End Function

Sub ProcessSalesTransactions()
    Debug.Print ("")
    Debug.Print ("Starting! " + CStr(Now))
    
    Call MySetup    'Debugging = False
    
    'Call MergeTransactionSheets
    'Call ProcessMergedSheet
    Call ProcessMergedSheetMichael
End Sub


Function test()
    Dim i As Long
    Dim tast As Boolean
    
    tast = True
    
    For i = 1 To 10
        Do
            'Do everything in here and
        Debug.Print "asdf"
            If tast Then
                Debug.Print "!!!"
                Exit Do
            End If
            
            Debug.Print "gadgasdg"
    
            'Of course, if I do want to finish it,
            'I put more stuff here, and then...
    
        Loop While False 'quit after one loop
    Next i
End Function

Function test2()
    Debug.Print (IsNumeric("5.1d"))
End Function
