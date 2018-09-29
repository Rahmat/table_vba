Dim Debugging As Boolean

Dim InputSheetName As String
Dim OutputSheetName As String

Dim OutputWS As Worksheet
Dim InputWS As Worksheet

Dim CurrentRow As Integer

'Function SetColumnsWidth(SheetName As String, WidthSize As Integer)
    'ActiveWorkbook.Sheets(SheetName).UsedRange.ColumnWidth = WidthSize
'End Function

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
    
Function ExportSheet(FromSheet As String, ToSheet As String, Optional ImportRange As String = vbNullString, Optional ExportRange As String = vbNullString)
    Application.CutCopyMode = True
    
    If ImportRange = vbNullString Then
        Sheets(FromSheet).Cells.Copy
    Else
        Sheets(FromSheet).Range(ImportRange).Copy
    End If
    
    If ExportRange = vbNullString Then
        ExportRange = "A1"
    End If
     
    Sheets(ToSheet).Range("A1").PasteSpecial Paste:=xlPasteValues
    Sheets(ToSheet).Range("A1").PasteSpecial Paste:=xlPasteFormats
    
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
        ElseIf ActiveWorkboo.Sheets(2).Name = "Output" Then 'check second
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
End Function

Function ExportFirst4Rows()
    Call ExportSheet(InputSheetName, OutputSheetName, "A1:Z4")
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

Function InsertItemMultiTotalsBySubDepartment()
    Dim i As Long
    
    Dim CurrentDeptName As String
    Dim CurrentDeptCode As String
    Dim Code As String
    Dim Description As String
    Dim QtyOrWeight As String
    Dim Amount As String
    
    For i = 6 To InputWS.UsedRange.Rows.Count
        Value = InputWS.Cells(i, 1)
        
        If Value = vbNullString Or Not IsNumeric(Value) Then 'No value or not a number, skip
            'Do Nothing (There is no continue statement in VBA)
        ElseIf Value < 10000 Then 'Dept Code found
            CurrentDeptCode = InputWS.Cells(i, 1)
            CurrentDeptName = InputWS.Cells(i, 1 + 1)
            'Debug.Print (CurrentDeptCode + CurrentDeptName)
        ElseIf Value > 10000 Then 'Item Code found
            Code = InputWS.Cells(i, 1)
            Description = InputWS.Cells(i, 1 + 2)
            QtyOrWeight = InputWS.Cells(i + 1, 1 + 7)
            Amount = InputWS.Cells(i + 1, 1 + 8)
            'Debug.Print (Code + Description + QtyOrWeight + Amount)
            Call InsertNextItemRow(Code, Description, CurrentDeptName, CurrentDeptCode, QtyOrWeight, Amount)
        End If
    Next i
    
End Function

Sub Main()
    Debugging = True
    
    Application.ScreenUpdating = False
    
    Call SetupSheets
    Call ExportFirst4Rows
    Call InsertColumnTitles
    Call InsertItemMultiTotalsBySubDepartment
    
    Application.ScreenUpdating = True
End Sub

Sub t()
    Call SetColumnWidth("Output", 20)
    'Debug.Print (ActiveWorkbook.Sheets(1).Name)
    'Debug.Print (ActiveWorkbook.Worksheets.Count)
End Sub
