Attribute VB_Name = "conecta"
Public DBConnection  As ADODB.Connection
Private DatabasePath As String

Public Function ConnectLocalDatabase()
    ' On Error Resume Next
    Screen.MousePointer = vbHourglass

    DatabasePath = Trim("C:\VB6_Demo\Db.mdb")

    Set DBConnection = New ADODB.Connection
    DBConnection.CursorLocation = adUseClient
    DBConnection.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & DatabasePath & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"

    Screen.MousePointer = 0
End Function

Public Sub Main()
   ConnectLocalDatabase
   Inicial.Show
End Sub


