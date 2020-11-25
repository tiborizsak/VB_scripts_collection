Attribute VB_Name = "Module1"
Sub Makró1()
Attribute Makró1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makró1 Makró
'

'
    Rows("2:2791").Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    Sheets("alap").Select
    Range("A1:V2500").Select
    Selection.Copy
    Sheets("számol").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.SmallScroll ToRight:=6
    Range("AA4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("D3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[24],8)=""Túraszám"",RIGHT(RC[24],8),R[-1]C)"
    Range("D3").Select

    Selection.AutoFill Destination:=Range("D3:D2504"), Type:=xlFillDefault
    Range("D3:D2504").Select

    Range("D2").Select

    Range("F3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(R[2]C[25],10)=""Sofõr neve"",R[2]C[26],R[-1]C)"
    Range("F3").Select

    Selection.AutoFill Destination:=Range("F3:F2504"), Type:=xlFillDefault
    Range("F3:F2504").Select


    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(R[1]C[24],8)=""Rendszám"",R[1]C[25],R[-1]C)"
    Range("G3").Select
    Selection.AutoFill Destination:=Range("G3:G2504"), Type:=xlFillDefault
    Range("G3:G2504").Select

    Range("H3").Select

    ActiveCell.FormulaR1C1 = "=RC[20]"
    Range("H3").Select

    Selection.AutoFill Destination:=Range("H3:H2504"), Type:=xlFillDefault
    Range("H3:H2504").Select

    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=RC[22]"
    Range("I3").Select

    Selection.AutoFill Destination:=Range("I3:I2504"), Type:=xlFillDefault
    Range("I3:I2504").Select

    Range("I3").Select
    
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("J4").Select

    Range("J3").Select
    Selection.AutoFill Destination:=Range("J3:r3"), Type:=xlFillDefault
    Range("J3:r3").Select
    Selection.AutoFill Destination:=Range("J3:r2504"), Type:=xlFillDefault
    Range("J3:r2504").Select
    Range("I3").Select

    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("J4").Select

    Range("J3").Select
    Selection.AutoFill Destination:=Range("J3:r3"), Type:=xlFillDefault
    Range("J3:r3").Select
    Selection.AutoFill Destination:=Range("J3:r17"), Type:=xlFillDefault
    Range("J3:r17").Select
    Selection.AutoFill Destination:=Range("J3:r2504"), Type:=xlFillDefault
    Range("J3:r2504").Select

    Range("s3").Select
    ActiveCell.FormulaR1C1 = "=RC[24]"
    Range("s3").Select
    Selection.AutoFill Destination:=Range("s3:s2503"), Type:=xlFillDefault
    Range("s3:s2503").Select
    ActiveWindow.SmallScroll ToRight:=-2
    Selection.AutoFill Destination:=Range("s3:s2504"), Type:=xlFillDefault
    Range("R3:R2504").Select


    Columns("D:s").Select
    Selection.Copy
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("w3").Select

    Rows("3:3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6, Criteria1:="=0", _
        Operator:=xlAnd
    Rows("2:2506").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6, Criteria1:="=EUR pal" _
        , Operator:=xlAnd
    Rows("2:2506").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6, Criteria1:= _
        "=Egyutas pal", Operator:=xlAnd
    ActiveWindow.SmallScroll Down:=-12
    Rows("2:2502").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6, Criteria1:="=Ügyfél" _
        , Operator:=xlAnd
    Rows("2:2575").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6

    Range("I2").Select
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=7, Criteria1:= _
        "=Indulás:", Operator:=xlAnd
    Rows("2:2505").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=7
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6, Criteria1:= _
        "=Összesen", Operator:=xlAnd
    Rows("2:2505").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$C$1:$v$2504").AutoFilter Field:=6
    Range("I4").Select
    ActiveWindow.SmallScroll Down:=-30

    'Sheets("alap").Select
    'Cells.Select
    'Selection.ClearContents
    'Range("A1").Select
    Sheets("számol").Select
    Range("B1").Select
End Sub

