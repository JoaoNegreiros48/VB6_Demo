VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Demo 
   BorderStyle     =   0  'None
   Caption         =   "Demo"
   ClientHeight    =   12510
   ClientLeft      =   975
   ClientTop       =   1755
   ClientWidth     =   22380
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
   ScaleHeight     =   12510
   ScaleWidth      =   22380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameAdicionar 
      Caption         =   "Adicionar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   1440
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   16695
      Begin VB.TextBox txtTelefone 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10440
         TabIndex        =   9
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10440
         TabIndex        =   8
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txtIdade 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         TabIndex        =   7
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         TabIndex        =   6
         Top             =   2040
         Width           =   3015
      End
      Begin lvButton.lvButtons_H cmdFecharFrameAdicionar 
         Height          =   600
         Left            =   9120
         TabIndex        =   15
         Top             =   4920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         Caption         =   "Fechar Adicionar"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdAdicionarCliente 
         Height          =   600
         Left            =   6960
         TabIndex        =   16
         Top             =   4920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         Caption         =   "Adicionar"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblIdade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Idade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4680
         TabIndex        =   5
         Top             =   3360
         Width           =   1650
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10440
         TabIndex        =   4
         Top             =   3360
         Width           =   1650
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10440
         TabIndex        =   3
         Top             =   1680
         Width           =   1650
      End
      Begin VB.Label lblNomeCompleto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome completo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4680
         TabIndex        =   2
         Top             =   1680
         Width           =   1650
      End
   End
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
   Begin lvButton.lvButtons_H cmdImprimir 
      Height          =   600
      Left            =   14760
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1058
      Caption         =   "Imprimir"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdicionar 
      Height          =   600
      Left            =   14760
      TabIndex        =   11
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1058
      Caption         =   "Adicionar"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H LvBEditar 
      Height          =   600
      Left            =   14760
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1058
      Caption         =   "Editar"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExcluir 
      Height          =   600
      Left            =   14760
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1058
      Caption         =   "Excluir"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFechar 
      Height          =   600
      Left            =   14760
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1058
      Caption         =   "Fechar"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_cadastrar As New ADODB.Recordset

Private Sub cmdAdicionarCliente_Click()
    If DBConnection.State = 0 Then DBConnection.Open ' se a conexão não estiver aberta, abre
    Set rs = New ADODB.Recordset
    rs.Open "Tabela1", DBConnection, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs("Nome").Value = txtNome
    rs("Email").Value = txtEmail
    rs("Telefone").Value = txtTelefone
    rs("Idade").Value = txtIdade
    rs.Update
    rs.Close
    
    MsgBox "Cliente cadastrado com sucesso"
    cmdFecharFrameAdicionar_Click
    atualizarDataGrid
End Sub

Private Sub cmdExcluir_Click()
    Dim sql As String
    
    sql = "SELECT * FROM Tabela1 WHERE cod_cadastro = " + DataGrid1.Columns(0).Text
    If DBConnection.State = 0 Then DBConnection.Open ' se a conexão não estiver aberta, abre
    rs.Close
    rs.Open sql, DBConnection, adOpenStatic, adLockOptimistic
    rs.Delete
    rs.Update
    rs.Close
    MsgBox "Cliente excluido com sucesso"
    atualizarDataGrid
End Sub

Private Sub cmdFechar_Click()
    Unload Me
    Unload Inicial
End Sub

Private Sub cmdFecharFrameAdicionar_Click()
    frameAdicionar.Visible = False
End Sub

Private Sub Form_Load()
    Me.Width = 16695
    Me.Height = 8535
    
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    
    Set rs = New ADODB.Recordset
    atualizarDataGrid
End Sub

Private Sub cmdImprimir_Click()
    Call imprime_dados
End Sub

Private Sub cmdAdicionar_Click()
    frameAdicionar.Visible = True
    frameAdicionar.Left = 0
    frameAdicionar.Top = 0
End Sub

Private Sub imprime_dados()
    If DBConnection.State = 0 Then DBConnection.Open ' se a conexão não estiver aberta, abre
    rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Tabela1", DBConnection, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    Set rs.ActiveConnection = Nothing ' Manter os dados do Rs mesmo quando a conexão é fechada
    
    Dim tamanhofolha As Integer
    Dim i            As Integer
    'define a fonte e o tamanhao da fonte
    
    'Printer.DeviceName = Printer.PDF
    Printer.FontName = "Arial"
    Printer.FontSize = "10"
    tamanhofolha = Printer.ScaleHeight - 1440 'define o tamanho da folha
    rs.MoveFirst 'movimenta o ponteiro para o primeiro registro

    contapagina = 0 'inicia o variável
    
    'Linhas para fazer cabeçalho
    Printer.Print Tab(100); "";
    Printer.Print Tab(10); "Nome";
    Printer.Print Tab(40); "Email";
    Printer.Print Tab(85); "Telefone";
    Printer.Print Tab(115); "Idade";
    Printer.Print Tab(100); "";

    Do While Not rs.EOF '
  
        If Printer.CurrentY >= tamanhofolha Then 'verifica se se folha já 'encheu'
            Printer.NewPage
            
        End If
  
        '---------------imprime os dados da tabela----------------------------
        Printer.Print Tab(10); rs("Nome");
        Printer.Print Tab(40); rs("Email");
        Printer.Print Tab(85); rs("Telefone");
        Printer.Print Tab(115); rs("Idade");
  
        '--------------------------------------------
  
        rs.MoveNext 'vai para o proximo registro

    Loop

    Printer.EndDoc 'envia os dados para a impressora

    MsgBox "Os dados foram enviados para a impressora ... ! "

End Sub

Function atualizarDataGrid()
    If DBConnection.State = 0 Then DBConnection.Open ' se a conexão não estiver aberta, abre
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Tabela1", DBConnection, adOpenStatic, adLockBatchOptimistic, adCmdText

    Set rs.ActiveConnection = Nothing ' Manter os dados do Rs mesmo quando a conexão é fechada
    Set DataGrid1.DataSource = rs ' Adiciona os dados no datagrid
    
    DBConnection.Close
End Function
