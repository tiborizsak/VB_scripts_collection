Attribute VB_Name = "Module4"
Sub UJRA()
Attribute UJRA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UJRA Makró
'

'
    
    Range("K:K").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Ujra").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "_UJRA"
    Range("B3").Select
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Selection.AutoFill Destination:=Range("B3:B" & lastrow), Type:=xlFillDefault
    Range("B3:B" & lastrow).Select
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C" & lastrow), Type:=xlFillDefault
    Range("C3:C" & lastrow).Select
    Selection.Copy
    Sheets("Rendeles").Select
    Range("K3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Ujra sheet clear
    Sheets("Ujra").Select
    Columns("A:K").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Rendeles").Select
    Range("A2").Select
End Sub
