VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PDA_form 
   Caption         =   "PDA kiadási ürlap"
   ClientHeight    =   2580
   ClientLeft      =   -135
   ClientTop       =   -390
   ClientWidth     =   3825
   OleObjectBlob   =   "PDA_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PDA_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub btn_rendben_Click()

Dim munkafuzet As Worksheet

Set munkafuzet = ThisWorkbook.Sheets("Fõoldal")

ujsor = munkafuzet.Cells(Rows.Count, 1).End(xlUp).Row + 1

munkafuzet.Cells(ujsor, 1) = Me.ki_be_lista
munkafuzet.Cells(ujsor, 3) = Me.WMS_kod
munkafuzet.Cells(ujsor, 4) = Me.PDA_kod

End Sub

Private Sub ki_be_lista_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub kibe_Click()

End Sub

Private Sub PDA_kod_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

'ki_be_list töltése
For Each elem In [ki_be]
    Me.ki_be_lista.AddItem elem
Next elem

End Sub
