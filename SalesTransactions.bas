Option Explicit

Public Function ArrayLen(arr As Variant) As Integer 'credit: https://stackoverflow.com/a/48627091
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function inc(ByRef data As Long) 'credit: https://stackoverflow.com/a/46728639
    data = data + 1
    inc = data
End Function

'Example:
'    Dim SheetsToMerge As Variant
'    SheetsToMerge = Array("Sheet1", "Sheet2")
'    Call MergeSheets(SheetsToMerge, "Output")
Public Function MergeSheets(SheetsToMerge As Variant, OutputSheetName As String)
    Application.CutCopyMode = True
    
    Dim sheet As Variant
    
    If WorksheetExists(OutputSheetName) Then
        ClearSheet (OutputSheetName)
    Else
        CreateWorksheet (OutputSheetName)
    End If
    
    For Each sheet In SheetsToMerge
        If Debugging Then
            Debug.Print ("Sheet name is: " + sheet)
            Debug.Print ("Last row in OutputSheet currently is: " + CStr(GetLastRow(OutputSheetName)))
        
            'Debug.Print (Sheets(Sheet).UsedRange.Rows.Count)
            Debug.Print ("Last col in OutputSheet currently is: " + CStr(Sheets(sheet).UsedRange.Columns.Count))
        End If
        
        'so that we can access the data
        Sheets(sheet).Select
        
        Dim RowCount As Long
        Dim ColCount As Long
        RowCount = Sheets(sheet).UsedRange.Rows.Count
        ColCount = Sheets(sheet).UsedRange.Columns.Count
        
        'test.Range(.cells(1, 1), .cells(RowCount, ColCount).copy
        Dim tempWorksheet As Worksheet
        'Dim TempRange As Range
        Set tempWorksheet = Sheets(sheet)
        
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
        
        Next sheet
    
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
    Dim sheet As Worksheet
    Dim Result As New Collection
    Dim CheckForString As Boolean
    
    If WithString <> "NOSTRINGSUPPLIEDBYUSER" Then
        CheckForString = True
    End If
    
    For Each sheet In ActiveWorkbook.Sheets
        If CheckForString Then
            If Not StringIsFound(WithString, sheet.Name) Then
                'pass
            Else
                Result.Add sheet.Name
            End If
        Else
            Result.Add sheet.Name
        End If
        Next sheet
    
    Set ReturnSheetNames = Result
End Function


Public Function RowIsBlank(RowNumber As Long, Optional sheet As String = "NOSTRINGSUPPLIEDBYUSER", Optional Debugging As Boolean = False) As Boolean
    Dim sh As Worksheet
    Dim UsedCols As Long, BlankCols As Long
    Dim cell As Variant
    
    RowIsBlank = True
    
    If sheet = "NOSTRINGSUPPLIEDBYUSER" Then
        Set sh = ActiveWorkbook.ActiveSheet
    Else
        Set sh = ActiveWorkbook.Sheets(sheet)
    End If
    
    Debug.Print (1)
    Debug.Print ("asdf: " + CStr(Application.WorksheetFunction.CountBlank(Range("A" + CStr(13)))))
    Debug.Print (2)
    UsedCols = sh.UsedRange.Rows(RowNumber).Columns.Count
    'Debug.Print (Application.WorksheetFunction.CountBlank(Range("A" + CStr(RowNumber))))
    Debug.Print (CStr(RowNumber) + " " + CStr(sh.UsedRange.Rows(RowNumber).Columns.Count))
    Debug.Print (CStr(RowNumber) + " " + CStr(WorksheetFunction.CountBlank(sh.UsedRange.Rows(RowNumber))))
    'debug.Print(
    BlankCols = 1 ' WorksheetFunction.CountBlank(sh.UsedRange.Rows(RowNumber))
    If UsedCols = BlankCols Then
        'Exit Function
    End If
    
    Debug.Print (3)
    
    For Each cell In sh.UsedRange.Rows(RowNumber).Cells
        If cell.Value <> vbNullString Then
            If Debugging Then
                Debug.Print "Row #" & CStr(RowNumber) & " is not blank! Found value '" & cell.Value & "' in column " & CStr(cell.Column)
            End If
            
            RowIsBlank = False
            Exit Function
        End If
    Next cell
    
    Call Err.Raise(1, "My Application", "If the program ever hits this line, then there's a problem with how we're checking for blank rows!")
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
            Debug.Print "ayy on " + CStr(i)
            If isB(i) Then
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

Public Function isB(RowNumber As Long) As Boolean
    Dim MySheet As Worksheet
    Dim ColCount As Long, i As Long
    
    isB = True 'We assume it's blank... until we can find a reason that it's not
    Set MySheet = ActiveWorkbook.ActiveSheet
    
    Debug.Print (CStr(MySheet.UsedRange.Rows(RowNumber).Cells.SpecialCells(xlCellTypeBlanks).Count))
    Debug.Print (MySheet.UsedRange.Rows(RowNumber).Cells.Count)
    
    ColCount = NumberOfColumns(RowNumber)
    
    For i = 1 To ColCount 'MySheet.UsedRange.Rows(RowNumber).End(xlToLeft).Column '.Columns.Count
        Debug.Print (CStr(i) + ": " + CStr(MySheet.Cells(RowNumber, i)))
    Next i
    
End Function

Public Function NumberOfColumns(RowNumber As Long, Optional SheetName As String) As Long
    Dim MySheet As Worksheet

    If SheetName = "" Then
        Set MySheet = ActiveWorkbook.ActiveSheet
    Else
        Set MySheet = ActiveWorkbook.Sheets(SheetName)
    End If
    
    Debug.Print (MySheet.Name)
    
    With MySheet
        'credit: https://stackoverflow.com/a/35945397
        NumberOfColumns = .UsedRange.Column + .UsedRange.Columns.Count - 1
    End With
    
    Debug.Print (NumberOfColumns)
End Function

Sub test()
    'NumberOfColumns (22)
    Debug.Print (isB(22))
End Sub

Function test2()
    
End Function
