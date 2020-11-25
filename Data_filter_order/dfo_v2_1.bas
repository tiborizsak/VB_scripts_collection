Attribute VB_Name = "Module3"
Sub Wpool_color()
'
' Wpool_color Makró
'

'
    Range("C1:W1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=6, Criteria1:= _
        "*WHIRLPOOL*"
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*EXTREME*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*MEDIA*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*AUCHAN*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:V" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*INTERNET*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:V" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*czovek*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*ELECTRO*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:V" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*DANTE*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*SELEX*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("C1:V1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*SVEA*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*VOROSKO*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*PREMIUM*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*MS E*"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("B1").Select
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8, Criteria1:= _
        "*BIGI *"
    lastrow = Cells(Rows.Count, 4).End(xlUp).Row
    Range("C1:W" & lastrow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Range("$C$1:$W$10000").AutoFilter Field:=8
    Range("C1:W1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("C1:W1").Select
    Selection.AutoFilter
    Range("C1:W1").Select
    Selection.AutoFilter
    Range("B1").Select
End Sub



