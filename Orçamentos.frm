VERSION 5.00
Begin VB.Form Orçamentos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Orçamentos"
   ClientHeight    =   6570
   ClientLeft      =   3765
   ClientTop       =   2385
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Orçamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Orçamentos.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   6570
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdCalc 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   5880
      TabIndex        =   38
      Top             =   4440
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   5880
      TabIndex        =   37
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5880
      TabIndex        =   36
      Top             =   3960
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6480
      TabIndex        =   33
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   32
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6480
      TabIndex        =   31
      Top             =   3960
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   6480
      TabIndex        =   30
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   6480
      TabIndex        =   29
      Top             =   4440
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   28
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   8280
      TabIndex        =   27
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   7680
      TabIndex        =   26
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   7080
      TabIndex        =   25
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   8280
      TabIndex        =   24
      Top             =   4440
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   7680
      TabIndex        =   23
      Top             =   4440
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   7080
      TabIndex        =   22
      Top             =   4440
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8280
      TabIndex        =   21
      Top             =   3960
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7680
      TabIndex        =   20
      Top             =   3960
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   7080
      TabIndex        =   19
      Top             =   3960
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8280
      TabIndex        =   18
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   17
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   15
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   14
      Top             =   3000
      Width           =   795
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8040
      TabIndex        =   13
      Top             =   3000
      Width           =   795
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Calcular"
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Menu"
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   390
      Left            =   1800
      TabIndex        =   9
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   390
      Left            =   2760
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   390
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   35
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   34
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblDisplay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   2520
      Width           =   2955
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de vezes"
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Funcionárias"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Custo em horas:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de horas:"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   6000
      Picture         =   "Orçamentos.frx":1784
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Orçamentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Orçamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdblResult           As Double
Private mdblSavedNumber      As Double
Private mstrDot              As String
Private mstrOp               As String
Private mstrDisplay          As String
Private mblnDecEntered       As Boolean
Private mblnOpPending        As Boolean
Private mblnNewEquals        As Boolean
Private mblnEqualsPressed    As Boolean
Private mintCurrKeyIndex    As Integer

Private Sub Command16_Click()
Index.Show
Orçamentos.Hide
End Sub

Private Sub Command17_Click()
Text6.Text = Text1.Text * Text2.Text * Text4.Text * Text5.Text

End Sub

Private Sub Form_Load()

    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intIndex    As Integer
    
    Select Case KeyCode
        Case vbKeyBack:             intIndex = 0
        Case vbKeyDelete:           intIndex = 1
        Case vbKeyEscape:           intIndex = 2
        Case vbKey0, vbKeyNumpad0:  intIndex = 18
        Case vbKey1, vbKeyNumpad1:  intIndex = 13
        Case vbKey2, vbKeyNumpad2:  intIndex = 14
        Case vbKey3, vbKeyNumpad3:  intIndex = 15
        Case vbKey4, vbKeyNumpad4:  intIndex = 8
        Case vbKey5, vbKeyNumpad5:  intIndex = 9
        Case vbKey6, vbKeyNumpad6:  intIndex = 10
        Case vbKey7, vbKeyNumpad7:  intIndex = 3
        Case vbKey8, vbKeyNumpad8:  intIndex = 4
        Case vbKey9, vbKeyNumpad9:  intIndex = 5
        Case vbKeyDecimal:          intIndex = 20
        Case vbKeyAdd:              intIndex = 21
        Case vbKeySubtract:         intIndex = 16
        Case vbKeyMultiply:         intIndex = 11
        Case vbKeyDivide:           intIndex = 6
        Case Else:                  Exit Sub
    End Select
    
    cmdCalc(intIndex).SetFocus
    cmdCalc_Click intIndex
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim intIndex    As Integer
    
    Select Case Chr$(KeyAscii)
        Case "S", "s":  intIndex = 7
        Case "P", "p":  intIndex = 12
        Case "R", "r":  intIndex = 17
        Case "X", "x":  intIndex = 11
        Case "=":       intIndex = 22
        Case Else:      Exit Sub
    End Select
    
    cmdCalc(intIndex).SetFocus
    cmdCalc_Click intIndex

End Sub

Private Sub cmdCalc_Click(Index As Integer)

    Dim strPressedKey   As String
    
    mintCurrKeyIndex = Index
    
    If mstrDisplay = "ERROR" Then
        mstrDisplay = ""
    End If
    
    strPressedKey = cmdCalc(Index).Caption
    
    Select Case strPressedKey
        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9"
            If mblnOpPending Then
                mstrDisplay = ""
                mblnOpPending = False
            End If
            If mblnEqualsPressed Then
                mstrDisplay = ""
                mblnEqualsPressed = False
            End If
            mstrDisplay = mstrDisplay & strPressedKey
        Case "."
            If mblnOpPending Then
                mstrDisplay = ""
                mblnOpPending = False
            End If
            If mblnEqualsPressed Then
                mstrDisplay = ""
                mblnEqualsPressed = False
            End If
            If InStr(mstrDisplay, ".") > 0 Then
                Beep
            Else
                mstrDisplay = mstrDisplay & strPressedKey
            End If
        Case "+", "-", "X", "/"
            mdblResult = Val(mstrDisplay)
            mstrOp = strPressedKey
            mblnOpPending = True
            mblnDecEntered = False
            mblnNewEquals = True
        Case "%"
            mdblSavedNumber = (Val(mstrDisplay) / 100) * mdblResult
            mstrDisplay = Format$(mdblSavedNumber)
        Case "="
            If mblnNewEquals Then
                mdblSavedNumber = Val(mstrDisplay)
                mblnNewEquals = False
            End If
            Select Case mstrOp
                Case "+"
                    mdblResult = mdblResult + mdblSavedNumber
                Case "-"
                    mdblResult = mdblResult - mdblSavedNumber
                Case "X"
                    mdblResult = mdblResult * mdblSavedNumber
                Case "/"
                    If mdblSavedNumber = 0 Then
                        mstrDisplay = "ERROR"
                    Else
                        mdblResult = mdblResult / mdblSavedNumber
                    End If
                Case Else
                    mdblResult = Val(mstrDisplay)
            End Select
            If mstrDisplay <> "ERROR" Then
                mstrDisplay = Format$(mdblResult)
            End If
            mblnEqualsPressed = True
        Case "+/-"
            If mstrDisplay <> "" Then
                If Left$(mstrDisplay, 1) = "-" Then
                    mstrDisplay = Right$(mstrDisplay, 2)
                Else
                    mstrDisplay = "-" & mstrDisplay
                End If
            End If
        Case "Backspace"
            If Val(mstrDisplay) <> 0 Then
                mstrDisplay = Left$(mstrDisplay, Len(mstrDisplay) - 1)
                mdblResult = Val(mstrDisplay)
            End If
        Case "CE"
            mstrDisplay = ""
        Case "C"
            mstrDisplay = ""
            mdblResult = 0
            mdblSavedNumber = 0
        Case "1/x"
            If Val(mstrDisplay) = 0 Then
                mstrDisplay = "ERROR"
            Else
                mdblResult = Val(mstrDisplay)
                mdblResult = 1 / mdblResult
                mstrDisplay = Format$(mdblResult)
            End If
        Case "sqrt"
            If Val(mstrDisplay) < 0 Then
                mstrDisplay = "ERROR"
            Else
                mdblResult = Val(mstrDisplay)
                mdblResult = Sqr(mdblResult)
                mstrDisplay = Format$(mdblResult)
            End If
    End Select
        
    If mstrDisplay = "" Then
        lblDisplay = "0."
    Else
        mstrDot = IIf(InStr(mstrDisplay, ".") > 0, "", ".")
        lblDisplay = mstrDisplay & mstrDot
        If Left$(lblDisplay, 1) = "0" Then
            lblDisplay = Mid$(lblDisplay, 2)
        End If
    End If
    
    If lblDisplay = "." Then lblDisplay = "0."
    
End Sub

