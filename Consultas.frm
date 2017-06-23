VERSION 5.00
Begin VB.Form Consultas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consultas"
   ClientHeight    =   4755
   ClientLeft      =   6450
   ClientTop       =   3930
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Consultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Consultas.frx":0BC2
   MousePointer    =   99  'Custom
   Picture         =   "Consultas.frx":1784
   ScaleHeight     =   4755
   ScaleWidth      =   5985
   Begin VB.CommandButton Command5 
      Caption         =   "Produtos/Clientes"
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Menu"
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clientes"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Produtos"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Encomendas"
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   2640
      Picture         =   "Consultas.frx":1AC6
      Top             =   0
      Width           =   3510
   End
End
Attribute VB_Name = "Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Consultas_clientes.Show
End Sub

Private Sub Command2_Click()
Consultas_produtos.Show
End Sub

Private Sub Command3_Click()
Consultas_encomendas.Show
End Sub

Private Sub Command4_Click()
Index.Show
Consultas.Hide
End Sub

Private Sub Command5_Click()
Consultas_Produtos_Clientes.Show
End Sub

