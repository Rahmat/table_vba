Option Explicit

Dim Debugging As Boolean

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

Public Function MakeRowBold(rownumber As Long)
    Range("A" + CStr(rownumber)).EntireRow.Font.Bold = True
End Function

Public Function setrowfont(rownumber As Long, rowfont As String)
    Range("A" + CStr(rownumber)).EntireRow.Font.Name = rowfont
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
Sub ItemMultiTotals() 'ItemMultiTotalsBySubDepartment()
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
End Sub

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
Sub OptimizedItemNetSales()
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
End Sub


Function t()
    Columns("A").NumberFormat = 0
    'MakeRowBold (5)
    
    'Call SetColumnWidth("A", 13)
    
    'Debug.Print (TypeName(Application.Calculation))
    
    'Debug.Print (ActiveWorkbook.Sheets(1).Name)
    'Debug.Print (ActiveWorkbook.Worksheets.Count)
End Function



