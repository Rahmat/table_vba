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
            'Debug.Print "'Store' found at line " + CStr(i)
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

Public Function PrintArray(TheArray As Variant)
    Dim i As Long
    
    For i = LBound(TheArray) To UBound(TheArray)
        Debug.Print (TheArray(i))
    Next i
End Function

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
    
        'If StringIsFound("store coupon", CellData) Then
        If StringIsFound("Snap", CellData) Or StringIsFound("store coupon", CellData) Then
            'If RowIsBlank(i) Then
            '    Debug.Print "Also, this row blank! (dafuq?)"
            'End If
            
            For j = 1 To NumberOfColumns(i)
                    StoreCoupon = MergedSheet.Cells(i, j)
                    If IsNumeric(StoreCoupon) Then
                        'Debug.Print ("Store Coupon found!! row #" + CStr(i) + " column #" + CStr(j))
                        StoreCoupons.Add (CStr(StoreCoupon))
                        Exit For
                    End If
            Next j
        End If
    Next i
    
    Debug.Print ("STORE COUPONS:")
    
    'Dim arr As Variant
    'arr = CollectionToArray(StoreCoupons)
    'PrintArray (arr)
    
    'ClearSheet ("MergedSheet")
    
    Call DeleteSheet(MergedSheet)
    
    CreateWorksheet ("Results")
    Dim ResultsSheet As Worksheet
    Set ResultsSheet = ActiveWorkbook.Sheets("Results")
    
    Call WriteCollectionToSheet(StoreCoupons, ResultsSheet)
End Function

Public Function DeleteSheet(MySheet As Worksheet)
    Application.DisplayAlerts = False
    MySheet.Delete
    Application.DisplayAlerts = True
End Function

Public Function WriteCollectionToSheet(MyColl As Collection, MySheet As Worksheet)
    Dim i As Long
    
    For i = 1 To MyColl.Count
        MySheet.Cells(i, 1).Value = MyColl.Item(i)
    Next i
End Function

Public Sub writeArrToWS(arr() As Variant, startCell As Range)
    Dim targetRange As Range
    Set targetRange = startCell ' assumes startCell is a single cell. Could do error checking here!
    Set targetRange = targetRange.Resize(UBound(arr, 1), UBound(arr, 2))
    targetRange.ClearContents ' don't even think this is necessary.
    targetRange = arr
End Sub

Function GetCoupons()
    Debug.Print ("")
    Debug.Print ("Starting! " + CStr(Now))
    
    Call MySetup    'Debugging = False
    
    Call MergeTransactionSheets
    'Call ProcessMergedSheet
    'Call CalculateSnapDiscounts
End Function

'Go down column B, looking for "Total Net Sales"
'Between where you started from and where you found "Total Net Sales" we'll be assuming is where this transaction is
'
'For the rows of this transaction, go down column F, looking for "Produce (PD)"
'Rows where you found Produce (PD), the last number on the row is where we assume the Net Sales for this item was
'Sum these numbers and divide by 2, for each transaction
'
'Output Results
Public Function ProcessSnapDiscounts()
    Dim MergedSheet As Worksheet
    Dim i As Long, j As Long, NumOfRows As Long, TransactionStartingRow As Long, TransactionEndingRow As Long
    Dim CellData As String, StoreCoupon As String
    Dim StoreCoupons As New Collection
    Dim TmpStoreCoupons As New Collection
    'Dim StoreCouponFound As Boolean

    Set MergedSheet = ActiveWorkbook.Sheets("MergedSheet")
    MergedSheet.Activate
    
    TransactionStartingRow = 1
    TransactionEndingRow = 1
    NumOfRows = MergedSheet.UsedRange.Rows.Count
    For i = 1 To NumOfRows
        CellData = CStr(MergedSheet.Cells(i, 2))
        
        'If StringIsFound("Produce (PD)", CellData) Then
        If StringIsFound("Total Net Sales", CellData) Or (i = NumOfRows) Then
            TransactionEndingRow = i
            
            For j = TransactionStartingRow To TransactionEndingRow
                CellData = CStr(MergedSheet.Cells(j, 6))
                If StringIsFound("Produce (PD)", CellData) Then
                    CellData = CStr(MergedSheet.Cells(j, 13))
                    Debug.Print (CellData)
                    TmpStoreCoupons.Add (CellData)
                End If
            Next j
            
            StoreCoupons.Add (SumColl(TmpStoreCoupons))
            Set TmpStoreCoupons = Nothing
            Debug.Print ("EmptiedTmpColl")
            
            TransactionStartingRow = i + 1
        End If
        
    Next i
    
    Call DeleteSheet(MergedSheet)
    
    CreateWorksheet ("Results")
    Dim ResultsSheet As Worksheet
    Set ResultsSheet = ActiveWorkbook.Sheets("Results")
    
    Call WriteCollectionToSheet(StoreCoupons, ResultsSheet)
End Function

Sub CalculateSnapDiscounts()
    Debug.Print ("")
    Debug.Print ("Starting! " + CStr(Now))
    
    Call MySetup    'Debugging = False
    
    If WorksheetExists("MergedSheet") Then
        Call DeleteSheet(ActiveWorkbook.Sheets("MergedSheet"))
    End If
    
    If WorksheetExists("Results") Then
        Call DeleteSheet(ActiveWorkbook.Sheets("Results"))
    End If
    
    Call MergeTransactionSheets
    'Call ProcessMergedSheet
    Call ProcessSnapDiscounts
End Sub

Public Function CollectionToArray(myCol As Collection) As Variant
    Dim result  As Variant
    Dim cnt     As Long

    ReDim result(myCol.Count)
    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt
    CollectionToArray = result
End Function


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
    Dim t As New Collection
    
    t.Add (1)
    t.Add (3)
    t.Add ("3")
    
    Debug.Print (SumCol(t) / 2)
End Function

Public Function SumColl(Coll As Collection) As Double
    Dim i As Double
    SumColl = 0
    
    For i = 1 To Coll.Count
        If IsNumeric(Coll.Item(i)) Then
            SumColl = SumColl + Coll.Item(i)
        Else
            SumColl = -1
        End If
    Next i
End Function
