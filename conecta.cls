VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "conecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

