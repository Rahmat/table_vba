Option Explicit

Public Function ArrayLen(arr As Variant) As Integer 'credit: https://stackoverflow.com/a/48627091
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function inc(ByRef data As Integer) 'credit: https://stackoverflow.com/a/46728639
    data = data + 1
    inc = data
End Function

'Example:
'    Dim SheetsToMerge As Variant
'    SheetsToMerge = Array("Sheet1", "Sheet2")
'    Call MergeSheets(SheetsToMerge, "Output")
Public Function MergeSheets(SheetsToMerge As Variant, OutputSheetName As String)
    Application.CutCopyMode = True
    
    Dim Sheet As Variant
    
    If WorksheetExists(OutputSheetName) Then
        ClearSheet (OutputSheetName)
    Else
        CreateWorksheet (OutputSheetName)
    End If
    
    For Each Sheet In SheetsToMerge
        If Debugging Then
            Debug.Print ("Sheet name is: " + Sheet)
            Debug.Print ("Last row in OutputSheet currently is: " + CStr(GetLastRow(OutputSheetName)))
        
            'Debug.Print (Sheets(Sheet).UsedRange.Rows.Count)
            Debug.Print ("Last col in OutputSheet currently is: " + CStr(Sheets(Sheet).UsedRange.Columns.Count))
        End If
        
        'so that we can access the data
        Sheets(Sheet).Select
        
        Dim RowCount As Long
        Dim ColCount As Long
        RowCount = Sheets(Sheet).UsedRange.Rows.Count
        ColCount = Sheets(Sheet).UsedRange.Columns.Count
        
        'test.Range(.cells(1, 1), .cells(RowCount, ColCount).copy
        Dim tempWorksheet As Worksheet
        'Dim TempRange As Range
        Set tempWorksheet = Sheets(Sheet)
        
        'tempWorksheet.Range
        With tempWorksheet
            'Set TempRange = Range(.Cells(1, 1), .Cells(RowCount, ColCount))
            'TempRange.Select
            Range(.Cells(1, 1), .Cells(RowCount, ColCount)).Select
        End With
        
        'tempWorksheet.Range(Cells(1, 1), Cells(RowCount, ColCount)).Select
        Selection.Copy
        'test.Range()
        Sheets(OutputSheetName).Range("A" + CStr(GetLastRow(OutputSheetName) + 1)).PasteSpecial xlPasteValues
        
        Next Sheet
    
    Application.CutCopyMode = False
End Function

Public Function GetLastRow(SheetName) As String
    Dim MySheet As Worksheet
    Set MySheet = ActiveWorkbook.Sheets(SheetName)
    
    GetLastRow = MySheet.UsedRange.Rows(MySheet.UsedRange.Rows.Count).row 'Credit: https://www.thespreadsheetguru.com/blog/2014/7/7/5-different-ways-to-find-the-last-row-or-last-column-using-vba
    'does Sheets(SheetName).UsedRange.Rows.Count not work?
End Function

'Testing:
    'StringIsFound("t", "tt")   True
    'StringIsFound("t", "TTT")  False
    'StringIsFound("t", "zzzz") False
Public Function StringIsFound(Needle As String, HayStack As String) As Boolean
    If InStr(HayStack, Needle) > 0 Then
        StringIsFound = True
    End If
End Function

'Testing
    'Lets say there's 3 sheets in the workbook, they're named: "Sheet1", "Sheet2", & "Other"
    'ReturnSheetNames()         Collection("Sheet1", "Sheet2", "Other")
    'ReturnSheetNames("Sheet")  Collection("Sheet1", "Sheet2")
    'ReturnSheetNames("Oth")    Collection("Other")
Public Function ReturnSheetNames(Optional WithString As String = "NOSTRINGSUPPLIEDBYUSER") As Collection
    Dim Sheet As Worksheet
    Dim Result As New Collection
    Dim CheckForString As Boolean
    
    If WithString <> "NOSTRINGSUPPLIEDBYUSER" Then
        CheckForString = True
    End If
    
    For Each Sheet In ActiveWorkbook.Sheets
        If CheckForString Then
            If Not StringIsFound(WithString, Sheet.Name) Then
                'pass
            Else
                Result.Add Sheet.Name
            End If
        Else
            Result.Add Sheet.Name
        End If
        Next Sheet
    
    Set ReturnSheetNames = Result
End Function

 
Sub ProcessSalesTransactions()
    Debug.Print ("Starting...")
    
    Call MySetup
    'Debugging = False
    
    Call MergeSheets(ReturnSheetNames("Sheet"), "Output")
End Sub

Function test()
    
End Function

Function test2()
    
End Function
