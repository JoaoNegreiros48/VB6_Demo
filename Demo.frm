VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Demo 
   BorderStyle     =   0  'None
   Caption         =   "Demo"
   ClientHeight    =   8535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   16695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   15055
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CLIENTES"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cod_cadastro"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome"
         Caption         =   "Nome completo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "email"
         Caption         =   "Email"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "telefone"
         Caption         =   "Telefone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "idade"
         Caption         =   "Idade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Size            =   474
         BeginProperty Column00 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4919,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4694,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2624,882
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim rs As New ADODB.Recordset

    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2

    If DBConnection.State = 0 Then DBConnection.Open ' se a conexão não estiver aberta, abre
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Tabela1", DBConnection, adOpenStatic, adLockBatchOptimistic, adCmdText

    Set rs.ActiveConnection = Nothing ' Manter os dados do Rs mesmo quando a conexão é fechada
    Set DataGrid1.DataSource = rs ' Adiciona os dados no datagrid
End Sub

