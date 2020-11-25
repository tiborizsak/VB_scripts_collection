Attribute VB_Name = "Module2"
Sub Schname_gen()
Attribute Schname_gen.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Schname_gen Makró
'

'
    Sheets("Rendeles").Select
    Columns("G:J").Select
    Selection.Copy
    Sheets("Keresonev").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-5],4)"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=RC[-5]"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-5],3)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-4],4)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-5],4)"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]&RC[-3]&RC[-2]&RC[-1]"
    Range("F3:J3").Select
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Selection.AutoFill Destination:=Range("F3:J" & lastrow), Type:=xlFillDefault
    Range("F3:J" & lastrow).Select
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=UPPER(RC[-1])"
    Range("K3").Select
    Selection.AutoFill Destination:=Range("K3:K" & lastrow), Type:=xlFillDefault
    Range("K3:K" & lastrow).Select
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'letter change
    Selection.Replace What:="Á", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="É", Replacement:="E", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Í", Replacement:="I", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Ö", Replacement:="O", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Õ", Replacement:="O", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Ó", Replacement:="O", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Ú", Replacement:="U", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Ü", Replacement:="U", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Û", Replacement:="U", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'character change
     Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     Selection.Replace What:="/", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     Selection.Replace What:="&", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Rendeles").Select
    Range("F3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'SCHgen sheet clear
    Sheets("Keresonev").Select
    Columns("A:K").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Rendeles").Select
    Range("A2").Select
End Sub
