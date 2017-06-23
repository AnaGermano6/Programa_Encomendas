VERSION 5.00
Begin VB.Form User 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Encomendas"
   ClientHeight    =   3630
   ClientLeft      =   5880
   ClientTop       =   3930
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "User.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "User.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   3630
   ScaleWidth      =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3360
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   3360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   3360
      Picture         =   "User.frx":1784
      Top             =   120
      Width           =   3510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizador"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "estimadiaria" And Text2.Text = "1060423" Then
Index.Show
User.Hide
Else
MsgBox "Utilizador errado", 16, "Entrar"
End If
End Sub

