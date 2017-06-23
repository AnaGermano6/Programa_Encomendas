VERSION 5.00
Begin VB.Form Galeria 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Galeria"
   ClientHeight    =   8190
   ClientLeft      =   3960
   ClientTop       =   1995
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Galeria.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Galeria.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   8190
   ScaleWidth      =   12120
   Begin VB.CommandButton Command1 
      Caption         =   "Menu"
      Height          =   615
      Left            =   11040
      TabIndex        =   4
      Top             =   7440
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   480
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1755
      Left            =   8640
      Picture         =   "Galeria.frx":1784
      Top             =   0
      Width           =   3510
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   240
      Y1              =   1680
      Y2              =   6360
   End
   Begin VB.Line Line3 
      X1              =   3240
      X2              =   240
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   3240
      Y1              =   1680
      Y2              =   6360
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Galeria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   6975
   End
End
Attribute VB_Name = "Galeria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Index.Show
Galeria.Hide
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    SelectedFile = File1.Path & "\" & File1.FileName
    Image1.Picture = LoadPicture(SelectedFile)
End Sub

