Dim Debugging As Boolean

Dim InputSheetName As String
Dim OutputSheetName As String

Dim OutputWS As Worksheet
Dim InputWS As Worksheet

Dim CurrentRow As Integer

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

Function GetSheetNames()
    Debug.Print (ActiveWorkbook.Worksheets.Count)
    If ActiveWorkbook.Worksheets.Count = 1 Then
        InputSheetName = ActiveWorkbook.Sheets(1).Name
    ElseIf ActiveWorkbook.Worksheets.Count = 2 Then
        Dim i As Integer
        Dim WorksheetNumber As Integer
        
        For i = 1 To 2
            Debug.Print (ActiveWorkbook.Sheets(i).Name)
            Debug.Print (ActiveWorkbook.Sheets(i).Name = "Output")
            If ActiveWorkbook.Sheets(i).Name = "Output" Then
                If i = 1 Then
                    WorksheetNumber = 2
                Else
                    WorksheetNumber = 1
                End If
                InputSheetName = ActiveWorkbook.Sheets(WorksheetNumber).Name
                Debug.Print ("t: " + InputSheetName)
            End If
        Next i
        If InputSheetName = "" Then
            Debug.Print (1)
            InputSheetName = InputBox("Multiple Sheets detected. What is the Input sheet's name?", "Enter Input sheet's name", InputSheetName)
        End If
    Else
        Debug.Print (2)
        InputSheetName = InputBox("Multiple Sheets detected. What is the Input sheet's name?", "Enter Input sheet's name", InputSheetName)
    End If
    
    If WorksheetExists(InputSheetName) Then
        Set InputWS = ActiveWorkbook.Sheets(InputSheetName)
    Else
        MsgBox ("'" + InputSheetName + "' InputSheet was not found in this workbook (" + ActiveWorkbook.Name + "). Terminating.")
        End
    End If
    
    If Not Debugging Then
        'OutputSheetName = InputBox("What sheet name do you want to output to?", "Enter Output sheet's name", OutputSheetName)
        OutputSheetName = "Output"
    End If
    If OutputSheetName = vbNullString Then
        MsgBox ("No Output Sheet Name detected, using 'Output' as the Name")
        OutputSheetName = "Output"
    End If
    If Not WorksheetExists(OutputSheetName) Then
        If Debugging Then
            MsgBox ("'" + OutputSheetName + "' OutputSheet was not found in this workbook (" + ActiveWorkbook.Name + "). Creating that now.")
        End If
        CreateWorksheet (OutputSheetName)
    End If
    
    If WorksheetExists(OutputSheetName) Then
        Set OutputWS = ActiveWorkbook.Sheets(OutputSheetName)
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
    OutputWS.Range("D5") = "Dept Name"
    OutputWS.Range("E5") = "Dept code"
    OutputWS.Range("F5") = "Qty/Weight"
    OutputWS.Range("G5") = "Amount"
    
    CurrentRow = 5
End Function

Function InsertNextItemRow(Code As String, Description As String, DeptName As String, DeptCode As String, QtyOrWeight As String, Amount As String)
    CurrentRow = CurrentRow + 1
    
    OutputWS.Cells(CurrentRow, 1) = Code
    OutputWS.Cells(CurrentRow, 2) = Description
    OutputWS.Cells(CurrentRow, 4) = DeptName
    OutputWS.Cells(CurrentRow, 5) = DeptCode
    OutputWS.Cells(CurrentRow, 6) = QtyOrWeight
    OutputWS.Cells(CurrentRow, 7) = Amount
End Function

Function InsertItemMultiTotalsBySubDepartment()
    Dim i As Long
    
    Dim CurrentDeptName As String
    Dim CurrentDeptCode As String
    Dim Code As String
    Dim Description As String
    Dim QtyOrWeight As String
    Dim Amount As String
    
    Debug.Print (InputWS.UsedRange.Rows.Count)
    For i = 6 To 39784
        i = i
    Next i
    For i = 6 To InputWS.UsedRange.Rows.Count
        Value = InputWS.Cells(i, 1)
        
        If Value = vbNullString Or Not IsNumeric(Value) Then 'No value or not a number, skip
            'Do Nothing (There is no continue statement in VBA)
        ElseIf Value < 10000 Then 'Dept Code found
            CurrentDeptCode = InputWS.Cells(i, 1)
            CurrentDeptName = InputWS.Cells(i, 1 + 1)
            Debug.Print (CurrentDeptCode + CurrentDeptName)
        ElseIf Value > 10000 Then 'Item Code found
            Code = InputWS.Cells(i, 1)
            Description = InputWS.Cells(i, 1 + 2)
            QtyOrWeight = InputWS.Cells(i + 1, 1 + 7)
            Amount = InputWS.Cells(i + 1, 1 + 8)
            Debug.Print (Code + Description + QtyOrWeight + Amount)
            Call InsertNextItemRow(Code, Description, CurrentDeptName, CurrentDeptCode, QtyOrWeight, Amount)
        End If
    Next i
    
End Function

Sub Main()
    Debugging = True
    
    Application.ScreenUpdating = False
    
    Call GetSheetNames
    Call ExportFirst4Rows
    Call InsertColumnTitles
    Call InsertItemMultiTotalsBySubDepartment
    
    Application.ScreenUpdating = True
End Sub

Function t()
    'Debug.Print (ActiveWorkbook.Sheets(1).Name)
    'Debug.Print (ActiveWorkbook.Worksheets.Count)
End Function
