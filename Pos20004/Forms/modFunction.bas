Attribute VB_Name = "modFunction"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public calther As String
Public Qty As Long
Sub Main()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\POS.mdb"
Splash.Show
End Sub
Public Function getquantity() As Single
    getquantity = Qty
End Function
Public Function TxtBoxIsEmpty(t As Object, tcounter As Integer) As Boolean
Dim i As Integer
For i = 0 To tcounter - 1
    If t(i).Text = "" Then
        TxtBoxIsEmpty = True
    End If
Next i
End Function
Public Sub Cleartext(t As Object, tcounter As Integer)
Dim i As Integer
For i = 0 To tcounter - 1
    t(i).Text = ""
Next i

End Sub
Public Sub AllowNumbersOnly(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Str("8")
Case Else
    KeyAscii = 0
End Select
End Sub
Public Sub AllowNumbersOnlyDot(KeyAscii As Integer)
Select Case KeyAscii

   Case Asc("0") To Asc("9")
   Case Asc(".")
   Case Str("8")
   Case Else
       KeyAscii = 0
End Select
End Sub
