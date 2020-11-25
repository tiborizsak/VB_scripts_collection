Attribute VB_Name = "Module1"
Sub MI_standard_v1()
Attribute MI_standard_v1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MI_standard_v1 Makró
'

'
Dim client As String
Dim datum As String
    'Original data copy
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Munka2").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Sheets("Munka3").Select
    ActiveWindow.SelectedSheets.Delete
    'Datum
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Today"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=NOW()"
    Range("A1").Select
    Selection.NumberFormat = "yyyy.mm.dd.hh.mm"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    datum = Replace(Range("B1").Value, ":", ".")
    datum = Replace(datum, " ", "")
    client = Range("B3")
    'Loading/unloading date format
    Columns("N:O").Select
    Selection.NumberFormat = "m/d/yyyy"
    'Data to unloading searchstring
    Range("F3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("S3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Unloading countrycode
    Range("T3").Select
    ActiveCell.FormulaR1C1 = "HU"
    Range("T3").Select
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Selection.AutoFill Destination:=Range("T3:T" & lastrow)
Call CreateCSV("\\bszburo\departments\DOMESTIC TRANSPORT\NEW CHW\Mass import\CSV\" & client & "_" & datum & ".csv")
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close False
    Application.DisplayAlerts = True
End Sub
Sub CreateCSV(Path As String)

    Dim rCell As Range
    Dim rRow As Range
    Dim sOutput As String
    Dim sFname As String, lFnum As Long

    'Open a text file to write
    sFname = Path
    lFnum = FreeFile

     Open sFname For Output As lFnum
    'Loop through the rows'
        For Each rRow In ActiveSheet.UsedRange.Rows
        'Loop through the cells in the rows'
        For Each rCell In rRow.Cells
            sOutput = sOutput & rCell.Value & ";"
        
        
        
        
        Next rCell
         'remove the last comma'
        sOutput = Left(sOutput, Len(sOutput) - 1)

        'write to the file and reinitialize the variables'
        Print #lFnum, sOutput
        sOutput = ""
     Next rRow

    'Close the file'
    Close lFnum

End Sub

