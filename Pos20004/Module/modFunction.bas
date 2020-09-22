Attribute VB_Name = "modFunction"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public calther As String
Public Qty As Long
Sub Main()
'con.Open "DSN=Steve"
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\POS.mdb"
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
