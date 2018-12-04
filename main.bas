Option Explicit

#Const developMode = True

Public Debugging As Boolean

Dim ScreenUpdateState As Boolean
Dim StatusBarState  As Boolean
Dim CalcState As Variant
Dim EventsState  As Boolean

Dim InputSheetName As String
Dim OutputSheetName As String

Dim OutputWS As Worksheet
Dim InputWS As Worksheet

Dim CurrentRow As Integer

Function SetGlobalsToDefault()
    Dim EmptyWS As Worksheet

    InputSheetName = ""
    OutputSheetName = ""
    
    'If Not TypeName(OutputWS) = "Nothing" Then
    '    OutputWS.Copy After:=EmptyWS
    '    OutputWS = Nothing
    'End If
    '
    'If Not TypeName(InputWS) = "Nothing" Then
    '    InputWS = EmptyWS
    'End If
    
    CurrentRow = 0
End Function

Function LastRow(sh As Worksheet) 'Credit: https://www.rondebruin.nl/win/s3/win002.htm
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row
    On Error GoTo 0
End Function


Function LastCol(sh As Worksheet) 'Credit: https://www.rondebruin.nl/win/s3/win002.htm
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Public Function MakeRowBold(RowNumber As Long)
    Range("A" + CStr(RowNumber)).EntireRow.Font.Bold = True
End Function

Public Function setrowfont(RowNumber As Long, rowfont As String)
    Range("A" + CStr(RowNumber)).EntireRow.Font.Name = rowfont
End Function

Public Function MySetup()
    SetGlobalsToDefault
    Debugging = True

    'Save parameters
    ScreenUpdateState = Application.ScreenUpdating
    StatusBarState = Application.DisplayStatusBar
    CalcState = Application.Calculation
    EventsState = Application.EnableEvents

    'Turn them off
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Function

Public Function MyOnTerminate()
    'Turn them back to normal
    Application.ScreenUpdating = ScreenUpdateState
    Application.DisplayStatusBar = StatusBarState
    Application.Calculation = CalcState
    Application.EnableEvents = EventsState
End Function

'Function SetColumnsWidth(SheetName As String, WidthSize As Integer)
    'ActiveWorkbook.Sheets(SheetName).UsedRange.ColumnWidth = WidthSize
'End Function

Function SetColumnWidth(ColumnName As String, WidthSize As Integer)
    Columns(ColumnName).ColumnWidth = WidthSize
End Function

Function ClearSheet(SheetName As String)
    ActiveWorkbook.Sheets(SheetName).UsedRange.Clear
End Function

Function CreateWorksheet(NewSheetName As String)
    Dim NewSheet As Object
    
    Set NewSheet = ActiveWorkbook.Sheets.Add(Type:=xlWorksheet)
    NewSheet.Name = NewSheetName
End Function

Function WorksheetExists(sName As String) As Boolean
    If (sName = vbNullString) Then
        WorksheetExists = False
        Exit Function
    End If
    
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function
    
Function ExportSheet(FromSheet As String, ToSheet As String, Optional ImportRange As String = vbNullString, Optional ExportRange As String = vbNullString, Optional KeepFormatting As Boolean = True)
    Application.CutCopyMode = True
    
    If ImportRange = vbNullString Then
        Sheets(FromSheet).Cells.Copy
    Else
        Sheets(FromSheet).Range(ImportRange).Copy
    End If
    
    If ExportRange = vbNullString Then
        ExportRange = "A1"
    End If
     
    Sheets(ToSheet).Range(ExportRange).PasteSpecial Paste:=xlPasteValues
    
    If KeepFormatting Then
        Sheets(ToSheet).Range(ExportRange).PasteSpecial Paste:=xlPasteFormats
    End If
    
    Application.CutCopyMode = False
End Function

Function SetupSheets()
    If Not ActiveWorkbook.ActiveSheet.Name = "Output" Then
        InputSheetName = ActiveWorkbook.ActiveSheet.Name
    ElseIf ActiveWorkbook.Worksheets.Count = 1 Then 'If just 1 sheet, then that's the input sheet
        InputSheetName = ActiveWorkbook.Sheets(1).Name
    ElseIf ActiveWorkbook.Worksheets.Count = 2 Then 'If 2 sheets, then if one is named Output, we can assume the other is Input
        If ActiveWorkbook.Sheets(1).Name = "Output" Then 'check first sheet for output
            InputSheetName = ActiveWorkbook.Sheets(2).Name
        ElseIf ActiveWorkbook.Sheets(2).Name = "Output" Then 'check second
            InputSheetName = ActiveWorkbook.Sheets(1).Name
        End If
    End If
        
    If InputSheetName = "" Then 'If more than 2 sheets or failed to get a sheetname
        InputSheetName = InputBox("Multiple Sheets detected. What is the Input sheet's name?", "Enter Input sheet's name", InputSheetName)
    End If
    
    If WorksheetExists(InputSheetName) Then
        Set InputWS = ActiveWorkbook.Sheets(InputSheetName)
    Else
        MsgBox ("'" + InputSheetName + "' InputSheet was not found in this workbook (" + ActiveWorkbook.Name + "). Terminating.")
        End
    End If
    
    OutputSheetName = "Output"
    'OutputSheetName = InputBox("What sheet name do you want to output to?", "Enter Output sheet's name", OutputSheetName)
        
    If OutputSheetName = vbNullString Then
        MsgBox ("No Output Sheet Name detected, using 'Output' as the Name")
        OutputSheetName = "Output"
    End If
    If Not WorksheetExists(OutputSheetName) Then
        If Debugging Then
            Debug.Print ("'" + OutputSheetName + "' OutputSheet was not found in this workbook (" + ActiveWorkbook.Name + "). Creating that now.")
        End If
        CreateWorksheet (OutputSheetName)
    End If
    
    If WorksheetExists(OutputSheetName) Then
        Set OutputWS = ActiveWorkbook.Sheets(OutputSheetName)
        ClearSheet (OutputSheetName)
    Else
        Call Err.Raise(0, "My Application", "Error finding " + OutputSheetName + " this code should be fixed.")
    End If
    
    OutputWS.Activate
End Function

Public Function ExportRows(BeginningRow As Integer, EndingRow As Integer, Optional KeepFormatting As Boolean = True)
    Dim MyRange As String
    MyRange = "A" + CStr(BeginningRow) + ":DZ" + CStr(EndingRow) 'e.g. "A1:DZ4"
    Call ExportSheet(InputSheetName, OutputSheetName, MyRange, KeepFormatting:=KeepFormatting)
End Function

Public Function ExportCell(InputCell As String, OutputCell As String, Optional KeepFormatting As Boolean = True)
    Call ExportSheet(InputSheetName, OutputSheetName, InputCell, OutputCell, KeepFormatting)
End Function

Function InsertColumnTitles()
    OutputWS.Range("A5") = "Code"
    OutputWS.Range("B5") = "Description"
    OutputWS.Range("C5") = "Dept Name"
    OutputWS.Range("D5") = "Dept code"
    OutputWS.Range("E5") = "Qty/Weight"
    OutputWS.Range("F5") = "Amount"
    
    CurrentRow = 5
End Function

Function InsertNextItemRow(Code As String, Description As String, DeptName As String, DeptCode As String, QtyOrWeight As String, Amount As String)
    CurrentRow = CurrentRow + 1
    
    OutputWS.Cells(CurrentRow, 1) = Code
    OutputWS.Cells(CurrentRow, 2) = Description
    OutputWS.Cells(CurrentRow, 3) = DeptName
    OutputWS.Cells(CurrentRow, 4) = DeptCode
    OutputWS.Cells(CurrentRow, 5) = QtyOrWeight
    OutputWS.Cells(CurrentRow, 6) = Amount
End Function

Function IsDeptCode(Code) As Boolean
    If Code < 1101 Or Code = 9999 Then
        IsDeptCode = True
    End If
End Function

Function IsItemCode(Code) As Boolean
    If Code > 1101 And Not Code = 9999 Then
        IsItemCode = True
    End If
End Function

Function InsertItemMultiTotalsBySubDepartment()
    Dim i As Long
    
    Dim CurrentDeptName As String
    Dim CurrentDeptCode As String
    Dim OldDeptName As String
    Dim OldDeptCode As String
    Dim Code As String
    Dim Description As String
    Dim QtyOrWeight As String
    Dim Amount As String
    Dim Value As Variant
    
    For i = 6 To InputWS.UsedRange.Rows.Count
        Value = InputWS.Cells(i, 1)
        
        If Value = vbNullString Or Not IsNumeric(Value) Then 'Nothing there, or it's not a number. skip.
            'Do Nothing (There is no continue statement in VBA)
        ElseIf IsDeptCode(Value) Then
            OldDeptCode = CurrentDeptCode
            OldDeptName = CurrentDeptName
            CurrentDeptCode = InputWS.Cells(i, 1)
            CurrentDeptName = InputWS.Cells(i, 1 + 1)
            
            If CurrentDeptName = vbNullString Then
                CurrentDeptName = OldDeptName 'https://i.imgur.com/kLFsrI4.png line 1303s code error would set CurrentDeptName to nothing
            End If
            'Debug.Print (CurrentDeptCode + CurrentDeptName)
        ElseIf IsItemCode(Value) Then
            Code = InputWS.Cells(i, 1)
            Description = InputWS.Cells(i, 1 + 2)
            QtyOrWeight = InputWS.Cells(i + 1, 1 + 7)
            Amount = InputWS.Cells(i + 1, 1 + 8)
            'Debug.Print (Code + Description + QtyOrWeight + Amount)
            'Debug.Print (Len(Code))
            Call InsertNextItemRow(Code, Description, CurrentDeptName, CurrentDeptCode, QtyOrWeight, Amount)
        End If
    Next i
    
    'formatting
    Call SetColumnWidth("A", 13)
    Call SetColumnWidth("B", 26)
    Call SetColumnWidth("C", 16)
    Call SetColumnWidth("D", 11)
    Call SetColumnWidth("E", 11)
    Call SetColumnWidth("F", 8)
    MakeRowBold (5)
    Columns("A").NumberFormat = 0
End Function

'Version 0.2
Function ItemMultiTotals() 'ItemMultiTotalsBySubDepartment()
    Debug.Print ("hi")
    Debug.Print (TypeName(OutputWS))
    
    'Debug.Print (OutputWS.Name)
    Call MySetup
    
    Call SetupSheets
    Call ExportRows(1, 4)
    Call InsertColumnTitles
    Call InsertItemMultiTotalsBySubDepartment
    
    OutputWS.Activate
    
    Call MyOnTerminate
End Function

'Version 0.0
Function CustomerMultiTotals()
    Call MySetup
    
    Call SetupSheets
    Call ExportRows(1, 5)
    'Call InsertColumnTitles
    'Call InsertItemMultiTotalsBySubDepartment
    
    Call MyOnTerminate
End Function

'Version 0.1
Function OptimizedItemNetSales()
    Call MySetup
    
    Dim i As Long
    
'    For i = 1 To 10
'        Debug.Print ("Row #" & CStr(i) & " has a Height of " & Rows(i).RowHeight)
'    Next i
'
'    For i = 1 To 10
'        Debug.Print ("Column #" & CStr(i) & " has a width of " & Columns(i).ColumnWidth)
'    Next i
    
    Call SetupSheets
    
    Call ExportCell("C4", "C1")
    Call ExportCell("C13", "C3")
    Call ExportCell("B13", "B3")
    Call ExportCell("B17", "B4")
    Call ExportCell("G19", "E5")
    Call ExportCell("D19", "D5")
    Call ExportCell("C19", "C5")
    Call ExportCell("B19", "B5")
    
    Rows(1).RowHeight = 18.75
    Rows(2).RowHeight = 0.75
    Rows(3).RowHeight = 24
    Rows(4).RowHeight = 15
    Rows(5).RowHeight = 13.5
    Columns(1).ColumnWidth = 1.14
    Columns(2).ColumnWidth = 14
    Columns(3).ColumnWidth = 31.29
    Columns(4).ColumnWidth = 10.29
    Columns(5).ColumnWidth = 10.14
    
    Call MakeRowBold(5)
    
    Range("B5:E5").AutoFilter 'adds filter thing there
    
    Rows(1).HorizontalAlignment = xlLeft
    Rows(5).HorizontalAlignment = xlLeft
    
    ActiveWindow.DisplayGridlines = False
    
    Dim StartingRow As Long
    Dim Value As Variant
    For i = 1 To InputWS.UsedRange.Rows.Count
        Value = InputWS.Cells(i, "B")
        If IsNumeric(Value) And Not Value = vbNullString Then
            StartingRow = i
            Debug.Print ("ayylmao: " & InputWS.Cells(i, "B"))
            Exit For
        End If
    Next i
    Debug.Print ("StartingRow is: " & CStr(StartingRow))
    Debug.Print ("Nani: " & CStr(InputWS.Cells(StartingRow, "B")))
    
    Dim ValueC As Variant
    Dim NextValueC As Variant
    Dim OutputRow As Long

    OutputRow = 6
    For i = StartingRow To InputWS.UsedRange.Rows.Count
        ValueC = InputWS.Cells(i, "C")
        NextValueC = InputWS.Cells(i + 1, "C")
        
        If Not ValueC = vbNullString Then
            OutputWS.Cells(OutputRow, 2) = InputWS.Cells(i, 2) 'itemid
            OutputWS.Cells(OutputRow, 3) = InputWS.Cells(i, 3) 'receipt alias
            OutputWS.Cells(OutputRow, 4) = InputWS.Cells(i, 4) 'net qty sold
            OutputWS.Cells(OutputRow, 5) = InputWS.Cells(i, 7) 'net sales
            OutputRow = OutputRow + 1
        ElseIf NextValueC = vbNullString Then
            Debug.Print ("we did ittt! at line #" & CStr(i))
            Exit For
        End If
    Next i
    
    OutputWS.Columns("E").NumberFormat = "[$$ -en-US]#,##0.00_);([$$ -en-US]#,##0.00)"

    Call MyOnTerminate
End Function

Public Function NumberOfColumns(RowNumber As Long, Optional SheetName As String) As Long
    Dim MySheet As Worksheet

    If SheetName = "" Then
        Set MySheet = ActiveWorkbook.ActiveSheet
    Else
        Set MySheet = ActiveWorkbook.Sheets(SheetName)
    End If
    
    With MySheet
        'credit: https://stackoverflow.com/a/35945397
        NumberOfColumns = .UsedRange.Column + .UsedRange.Columns.Count - 1
    End With
    
End Function

Public Function RowIsBlank(RowNumber As Long) As Boolean
    Dim MySheet As Worksheet
    Dim ColCount As Long, i As Long
    
    RowIsBlank = True 'We assume it's blank... until we can find a reason that it's not
    Set MySheet = ActiveWorkbook.ActiveSheet
        
    ColCount = NumberOfColumns(RowNumber)
    
    For i = 1 To ColCount 'MySheet.UsedRange.Rows(RowNumber).End(xlToLeft).Column '.Columns.Count
        'Debug.Print (CStr(i) + ": " + CStr(MySheet.Cells(RowNumber, i).Formula))
        If Not MySheet.Cells(RowNumber, i).Formula = "" Then
            RowIsBlank = False
            Exit Function
        End If
    Next i
End Function

'Testing
    'Lets say there's 3 sheets in the workbook, they're named: "Sheet1", "Sheet2", & "Other"
    'ReturnSheetNames()         Returns: Collection("Sheet1", "Sheet2", "Other")
    'ReturnSheetNames("Sheet")  Returns: Collection("Sheet1", "Sheet2")
    'ReturnSheetNames("Oth")    Returns: Collection("Other")
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

'Testing:
    'StringIsFound("t", "tt")   True
    'StringIsFound("t", "TTT")  False
    'StringIsFound("t", "zzzz") False
Public Function StringIsFound(Needle As String, HayStack As String) As Boolean
    If InStr(HayStack, Needle) > 0 Then
        StringIsFound = True
    End If
End Function

Public Function GetLastRow(SheetName) As String
    Dim MySheet As Worksheet
    Set MySheet = ActiveWorkbook.Sheets(SheetName)
    
    GetLastRow = MySheet.UsedRange.Rows(MySheet.UsedRange.Rows.Count).row 'Credit: https://www.thespreadsheetguru.com/blog/2014/7/7/5-different-ways-to-find-the-last-row-or-last-column-using-vba
    'does Sheets(SheetName).UsedRange.Rows.Count not work?
End Function

Public Function ArrayLen(Arr As Variant) As Integer 'credit: https://stackoverflow.com/a/48627091
    ArrayLen = UBound(Arr) - LBound(Arr) + 1
End Function

Public Function inc(ByRef data As Long) 'credit: https://stackoverflow.com/a/46728639
    data = data + 1
    inc = data
End Function

'credit: https://stackoverflow.com/a/30025752 should be called IsArrayAllocated
Function IsVarAllocated(Arr As Variant) As Boolean
        On Error Resume Next
        IsVarAllocated = IsArray(Arr) And _
                           Not IsError(LBound(Arr, 1)) And _
                           LBound(Arr, 1) <= UBound(Arr, 1)
End Function

'Example:
    'Call MergeSheets("MergedSheet", ReturnSheetNames("Sheet"))
    'Call MergeSheets("MergedSheet")
    'Call MergeSheets()
'Docs:
    'SheetsToMerge: Collection expection
Public Function MergeSheets(Optional OutputSheetName As String = "DefaultValue", Optional SheetsToMerge As Variant)
    Dim sheet As Variant
    
    If OutputSheetName = "DefaultValue" Then
        OutputSheetName = "MergedSheet"
    End If
    
    If Not IsVarAllocated(SheetsToMerge) Then
        Set SheetsToMerge = ReturnSheetNames()
    End If
   
    If WorksheetExists(OutputSheetName) Then
        ClearSheet (OutputSheetName)
    Else
        CreateWorksheet (OutputSheetName)
    End If
    
    Application.CutCopyMode = True
    
    For Each sheet In SheetsToMerge
        If Debugging Then
            'Debug.Print ("Sheet name is: " + sheet)
            'Debug.Print ("Last row in OutputSheet currently is: " + CStr(GetLastRow(OutputSheetName)))
        
            ''''Debug.Print (Sheets(Sheet).UsedRange.Rows.Count)
            'Debug.Print ("Last col in OutputSheet currently is: " + CStr(Sheets(sheet).UsedRange.Columns.Count))
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

Function t()
    Columns("A").NumberFormat = 0
    'MakeRowBold (5)
    
    'Call SetColumnWidth("A", 13)
    
    'Debug.Print (TypeName(Application.Calculation))
    
    'Debug.Print (ActiveWorkbook.Sheets(1).Name)
    'Debug.Print (ActiveWorkbook.Worksheets.Count)
End Function





