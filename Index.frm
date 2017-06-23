VERSION 5.00
Begin VB.Form Index 
   BackColor       =   &H80000008&
   Caption         =   "Estima Diária"
   ClientHeight    =   7950
   ClientLeft      =   3390
   ClientTop       =   2190
   ClientWidth     =   12000
   Icon            =   "Index.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Index.frx":0BC2
   MousePointer    =   99  'Custom
   Picture         =   "Index.frx":1784
   ScaleHeight     =   7950
   ScaleWidth      =   12000
   Begin VB.CommandButton Command8 
      Caption         =   "Serviços"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   9
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Orçamentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Galeria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3600
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Encomendas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   0
      Picture         =   "Index.frx":194AB
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Produtos.Show
End Sub

Private Sub Command2_Click()
Clientes.Show
End Sub

Private Sub Command3_Click()
Encomenda.Show
End Sub

Private Sub Command4_Click()
Consultas.Show
End Sub

Private Sub Command5_Click()
Dim resp As Integer, mens As String
mens = "Deseja sair?"
resp = MsgBox(mens, vbYesNo, "Sair")
If resp = vbNo Then
    MsgBox "Encerramento não confirmado", 64
Else
End
End If
End Sub

Private Sub Command6_Click()
Galeria.Show
End Sub

Private Sub Command7_Click()
Orçamentos.Show
End Sub

Private Sub Command8_Click()
Serviços.Show
End Sub

Private Sub Timer1_Timer()
Label1 = Time()
Label2 = Date
End Sub

