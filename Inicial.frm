VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Inicial 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Demonstração"
   ClientHeight    =   15735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   28680
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
   ScaleHeight     =   15735
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   15735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   28695
      Begin lvButton.lvButtons_H cmdIniciarDemo 
         Height          =   600
         Left            =   2400
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         Caption         =   "Iniciar Demo"
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
      Begin lvButton.lvButtons_H cmdFecharDemo 
         Height          =   600
         Left            =   4440
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         Caption         =   "Fechar demo"
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
      Begin VB.Label lblTituloInicial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Olá, que bom que decidiu executar essa demonstração. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7140
      End
   End
End
Attribute VB_Name = "Inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFecharDemo_Click()
    Unload Me
End Sub

Private Sub cmdIniciarDemo_Click()
    lblTituloInicial.Visible = False
    cmdIniciarDemo.Visible = False
    cmdFecharDemo.Visible = False
    Inicial.Enabled = False
    Demo.Show
End Sub

Private Sub Form_Load()
    Frame1.Left = 0
    Frame1.Top = 0
    Frame1.Width = Me.Width
    lblTituloInicial.Left = (Screen.Width - lblTituloInicial.Width) \ 2
    lblTituloInicial.Top = (Screen.Height - lblTituloInicial.Height) \ 2
    cmdIniciarDemo.Left = ((Screen.Width - cmdIniciarDemo.Width) \ 2) - 1000
    cmdIniciarDemo.Top = ((Screen.Height - lblTituloInicial.Height) \ 2) + 500
    cmdFecharDemo.Top = cmdIniciarDemo.Top
    cmdFecharDemo.Left = cmdIniciarDemo.Left + 1920
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Demo
End Sub
