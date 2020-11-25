Attribute VB_Name = "Module1"
Sub Start()
Attribute Start.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Start Makró
'

'
   

Dim client As String
Dim period As String
'Dim datum As String
'Dim ua As String
Dim hely As String
Dim helyeleje As String
Dim ev As String

    ActiveSheet.PivotTables("Kimutatás1").PivotCache.Refresh
    ActiveSheet.PivotTables("Kimutatás1").PivotFields("customer").ClearAllFilters
    
    With ActiveSheet.PivotTables("Kimutatás1").PivotFields("customer")

    End With
    Range("A7").Select
    ActiveSheet.PivotTables("Kimutatás1").PivotCache.Refresh

'  Range("g2").Select
'  hely = Selection.Value
'    Range("f2") = Range("g2").Value
  
'mentés helyének meghatározása
  Range("g2").Select
  hely = Selection.Value
'periódus/hónapmeghatározása
  Range("b4").Select
  period = Selection.Value
  Range("b5").Select
  ev = Selection.Value
  
  
'mentés számlakiállítás keltének meghatározása
'  Range("b1").Select
'  datum = Selection.Value
'mentés üa % meghatározása
'  Range("S2").Select
'  ua = Selection.Value
  
'  Range("f1") = Range("g1").Value
   helyeleje = Range("f1").Value
 '(üres)

    Range("a8").Select
Application.DisplayAlerts = False
        Do Until Selection.Value = "X" Or Selection.Value = "Végösszeg" Or Selection.Value = "(üres)"

            client = Selection.Value

Sheets("Inv_data").Select

'   ActiveSheet.Range("$A$1:$t$65000").AutoFilter Field:=1, Criteria1:=period
    ActiveSheet.Range("$A$1:$t$65000").AutoFilter Field:=4, Criteria1:=client

                   
'másolandó terület kijelõlése

     Range("A1:T1").Select
     Range(Selection, Selection.End(xlDown)).Select
     Selection.Copy


 '   Range("A4:U504").Copy
'    Selection.SpecialCells(xlCellTypeVisible).Select
'    Application.CutCopyMode = False
'    Selection.Copy

'template elérési út elejének meghatározása

    'Workbooks.Open Filename:=helyeleje & "Templates\InvTemplate" & client & ".xlsx"
'    SendKeys "{ENTER}", True
    '\\bszburo\departments\DOMEST~1\Kriszti_számlázás\Ügyfél számlázás\Templates\InvTemplateAGRI.xlsx
'    Workbooks.Open Filename:="\\bszburo\departments\DOMEST~1\Kriszti_számlázás\Ügyfél számlázás\Templates\InvTemplateAGRI.xlsx"
    
'template megnyitása (T)

    Workbooks.Open Filename:="C:\Users\tizsak\Desktop\FM data_2\DATA\Inv_makro\Templates\InvTemplateSTANDARD.xlsx"
    Sheets("Data").Select
    
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
 'táblázat formátum (T)
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            Range("A1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:T" & lastrow), , xlYes).Name = _
        "Táblázat3"
    Range("Táblázat3[#All]").Select
    ActiveSheet.ListObjects("Táblázat3").TableStyle = "TableStyleMedium16"
    Columns("A:A").Select
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True
    Range("E10").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("A2").Select
    Columns("A:A").ColumnWidth = 11.14
    Columns("B:B").ColumnWidth = 9.14
    Columns("C:C").ColumnWidth = 9.14
    Columns("F:F").ColumnWidth = 12.43
    Columns("G:G").ColumnWidth = 2.29
    Columns("H:H").ColumnWidth = 3.86
    Columns("I:I").ColumnWidth = 16.29
    Columns("J:J").ColumnWidth = 22.71
    Columns("H:H").ColumnWidth = 3.29
    Columns("K:K").ColumnWidth = 4.14
    Columns("L:L").ColumnWidth = 16.29
    Columns("N:N").ColumnWidth = 2.29
    Columns("O:O").ColumnWidth = 4.57
    Columns("P:P").ColumnWidth = 2.29
    Columns("Q:Q").ColumnWidth = 7.57
    Columns("R:R").ColumnWidth = 3.86
    Columns("S:S").ColumnWidth = 12.29
    Columns("T:T").ColumnWidth = 9.43
    Columns("T:T").Select
    Selection.Font.Bold = True
    ActiveWindow.Zoom = 90
    Range("Táblázat3[[#Headers],[Unloading date]]").Select
    
    'nem szükséges tariff type törlés (T)
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveSheet.Range("A1:I" & lastrow).AutoFilter Field:=7, Criteria1:=Array("FM"), Operator:=xlFilterValues
    'ActiveSheet.AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
    Range("A2").Select
    'Range(Selection, Selection.End(xlDown)).Select
    Range("A2:T1000").Select
    Selection.EntireRow.Delete
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error Resume Next
    
    'extra charge törlés (T)
    'lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'ActiveSheet.Range("A1:I" & lastrow).AutoFilter Field:=7, Criteria1:=Array(""), Operator:=xlFilterValues
    'ActiveSheet.AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
    'Range("A2").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range("A2:T1000").Select
    'Selection.EntireRow.Delete
    'On Error Resume Next
    'ActiveSheet.ShowAllData
    'On Error Resume Next
    
        Sheets("Total").Select
        Range("c3").Select
        Range("c3").Value = period
            
    Range("B1").Select
         ActiveWorkbook.SaveAs Filename:= _
                helyeleje & ev & "\" & period & "\" & client & "_" & period & ".xls", _
        FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            ActiveWindow.Close
            
            Sheets("Pivot").Select
            ActiveCell.Offset(1, 0).Select
            
            
            
'    Rows("2:2").Select
'    Selection.Delete Shift:=xlUp

'            Range("k2").Select
'            Range("k2").Value = client
            'Range("s4").Select
            'Range("s4").Value = ua * 1
'            Range("k1").Select
'            Range("k1").Value = datum
'                SendKeys "{F2}", True
'                SendKeys "{ENTER}", True
            

 ''''''        Selection.ClearContents
        
         
         
         
         
         
'         Range("A8:k8").Select
'         Selection.Copy
'         Range("A9:k253").Select
'         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'     Application.CutCopyMode = False
'     Range("A8").Select
'felesleges sorok törlése
' i = 253 'Sor
'    Cells(i, 21).Select
'    Do Until Selection.Value <> ""
'        i = i - 1
'        Cells(i, 11).Select

'    Loop
'        Cells(i + 1, 11).Select
'Dim row As Range
'For Each row In Range("u8:u200")
'If UCase(row.Value) = "" Then
'row.EntireRow.Delete Shift:=xlUp
'End If
'Next

'Range("u" & CStr(i + 1) & ":k253").EntireRow.Delete
'        Rows(i).(254).Select
'    Selection.Delete Shift:=xlUp

                
        '            ActiveWorkbook.SaveAs Filename:= _
        '                hely & period & client & ".xlsm", _
        '        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        
                'hely & Left(Range("T2"), 4) & Mid(Range("T2"), 5, 3) & "_" & client & ".xlsm", _

                    'ActiveWorkbook.SaveAs Filename:="E:\Docs\Dropbox\RRS\Tesco\Invoice\aaa.xls", _
                    '        FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
                    '        ReadOnlyRecommended:=False, CreateBackup:=False
   

        Loop
        
    Sheets("inv_data").Select
'    ActiveSheet.Range("$A$1:$t$65000").AutoFilter Field:=1
    ActiveSheet.Range("$A$1:$t$65000").AutoFilter Field:=4
    Sheets("Pivot").Select

Application.DisplayAlerts = True
    




End Sub
Sub Pivot_refresh_all()
Attribute Pivot_refresh_all.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Pivot_refresh_all Makró
'

'
    ActiveWorkbook.RefreshAll
End Sub
