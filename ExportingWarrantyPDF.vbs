Sub ExportingWarrantyPDF()

'Defining worksheets
Dim trackerSheet As Worksheet
Dim formSheet As Worksheet

Set trackerSheet = ActiveWorkbook.Sheets("Tracker")
Set formSheet = ActiveWorkbook.Sheets("Form")

Application.ScreenUpdating = True

'Looping the through each row
'For i = 2 To 20

'Prompt for row number
Dim inputNum As String
Dim lastRow As Long
Dim trackerRange As Range
Dim requiredColumns As Variant
requiredColumns = Array("B", "C", "E", "F", "H", "M")
Dim column As Variant

lastRow = trackerSheet.Cells(trackerSheet.Rows.Count, "A").End(xlUp).Row
Set trackerRange = trackerSheet.Range("A2:A" & lastRow)
inputNum = InputBox("Please type in the case number (in the Number column):")
    If Not IsNumeric(inputNum) Then
        MsgBox ("You must enter a number to continue.")
        Exit Sub
        ElseIf inputNum = "" Then
            MsgBox ("You must enter a valid number to continue.")
            Exit Sub
            ElseIf WorksheetFunction.CountIf(trackerRange, inputNum) = 0 Then
                MsgBox ("You must enter a valid number to continue.")
                Exit Sub
    End If

'Assigning values
rowNum = CInt(inputNum) + 1
Set selectedRow = trackerSheet.Rows(rowNum)
For Each column In requiredColumns
    If Trim(selectedRow.Cells(1, column).Value) = "" Then
        MsgBox "Column " & column & " is empty in the selected row.", vbExclamation
        Exit Sub
    End If
Next column

'MsgBox rowNum
SType = trackerSheet.Cells(rowNum, 2)
'MsgBox SType
SCustomer = trackerSheet.Cells(rowNum, 5)
SSN = trackerSheet.Cells(rowNum, 3)
SDate = trackerSheet.Cells(rowNum, 6)
STech = trackerSheet.Cells(rowNum, 8)
SIssue = trackerSheet.Cells(rowNum, 13)
Dim outputPath As String
    outputPath = Environ$("USERPROFILE") & "\Downloads" & Application.PathSeparator & SSN & "_" & Format(SDate, "MM_DD_YYYY") & ".pdf"
    MsgBox outputPath

'Generating the output
If SType = "via Returns" Then
    Sheets("Form").Range("A2").Interior.Color = RGB(255, 255, 0)
    ElseIf SType = "via Manufacturer" Then
        Sheets("Form").Range("A2").Interior.Color = RGB(255, 0, 0)
End If

formSheet.Cells(2, 1).Value = "Process " & SType
formSheet.Cells(3, 2).Value = SCustomer
formSheet.Cells(4, 2).Value = SSN
formSheet.Cells(5, 2).Value = SDate
formSheet.Cells(6, 2).Value = STech
formSheet.Cells(8, 1).Value = SIssue


'Save the PDF file
formSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    outputPath, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=True

'Next i

'Application.ScreenUpdating = True

End Sub



