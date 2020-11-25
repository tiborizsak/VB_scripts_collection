Attribute VB_Name = "Module2"
Sub Makró3()
Attribute Makró3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makró3 Makró
'

'

'Régi adatok törlése
    'Rows("2:5001").Select
    'Selection.Delete Shift:=xlUp
    'Range("B1").Select

'Alapadatok átmásolása
    Rows("2:5000").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Sheets("alap").Select
    Range("A1:T5000").Select
    Selection.Copy
    Sheets("számol").Select
    Range("AA3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Túraszám betöltése
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[24],8)=""Túraszám"",RIGHT(RC[24],8),R[-1]C)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D5000"), Type:=xlFillDefault
    Range("D2:D5000").Select
    
'Sofõr nevének betöltése
    Range("F2").Select
        ActiveCell.FormulaR1C1 = "=IF(R[2]C[25]=""Sofõr neve"",R[2]C[26],R[-1]C)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F5000"), Type:=xlFillDefault
    Range("F2:F5000").Select


'Rendszám betölt
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[1]C[24]=""Rendszám"",R[1]C[25],R[-1]C)"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G5000"), Type:=xlFillDefault
    Range("G2:G5000").Select


'Megrendelõ és szállítási cím
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[20]"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=RC[22]"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("M3").Select
    Range("H2:M2").Select
    Selection.AutoFill Destination:=Range("H2:M5000"), Type:=xlFillDefault
    Range("H2:M5000").Select


'Dobozszám, raklap, súly
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:S2"), Type:=xlFillDefault
    Range("N2:S2").Select
    Selection.AutoFill Destination:=Range("N2:S5000"), Type:=xlFillDefault
    Range("N2:S5000").Select
    
'Értékként beilleszt
    Columns("C:V").Select
    Selection.Copy
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

'Felesleges sorok törlése
    ActiveSheet.Range("$C$1:$V$5001").AutoFilter Field:=11, Criteria1:="Cím"
    Rows("2:5001").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$C$1:$V$5001").AutoFilter Field:=11
    Range("C1").Select

    ActiveSheet.Range("$C$1:$V$5001").AutoFilter Field:=9, Criteria1:="<1", _
        Operator:=xlAnd
    Rows("2:5001").Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    ActiveSheet.Range("$C$1:$V$5001").AutoFilter Field:=9
    Range("B1").Select
    
'Megjegyzés fkeres (T)
    Range("W1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "Megjegyzés"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-14],alap!C[-18]:C[-3],16,0)"
    Range("W2").Select
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Selection.AutoFill Destination:=Range("W2:W" & lastrow)
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Reference fkeres (T)
    Range("Y1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "Reference"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-16],alap!C[-20]:C[-5],3,0)"
    Range("Y2").Select
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Selection.AutoFill Destination:=Range("Y2:Y" & lastrow)
    Columns("Y:Y").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'nulla torol
    ActiveSheet.Range("$C$1:$W$100000").AutoFilter Field:=21, Criteria1:="0"
    Range("W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("W1").Select
    ActiveSheet.Range("$C$1:$W$100000").AutoFilter Field:=21
    
'oszlop rendezes + ido
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    ActiveWindow.SmallScroll ToRight:=10
    Range("X2").Select
    ActiveSheet.Paste
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("F2").Select
    ActiveSheet.Paste
    Range("X2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    ActiveWindow.SmallScroll ToRight:=-5
    Range("G2").Select
    ActiveSheet.Paste
    Range("P2").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("O2:S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("P2").Select
    ActiveSheet.Paste
    Range("W2:W100000").Select
    Selection.Cut
    Range("O2").Select
    ActiveSheet.Paste
    Columns("O:O").Select
    Selection.NumberFormat = "h:mm"
    Range("W1").Select
    Selection.ClearContents
    Range("B1").Select
     
'Alap és felesleges oszlopok törlése
    'Sheets("alap").Select
    'Cells.Select
    'Columns("A:AV").Select
    'Selection.Delete Shift:=xlToLeft
    'Range("A1").Select
    'Sheets("számol").Select
    'Columns("AA:AV").Select
    'Selection.Delete Shift:=xlToLeft
    'Range("B1").Select
    'Columns("Z:Z").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    'Selection.Delete Shift:=xlToLeft
    'Range("B1").Select
End Sub


