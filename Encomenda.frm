VERSION 5.00
Begin VB.Form Encomenda 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Encomendas"
   ClientHeight    =   6900
   ClientLeft      =   4350
   ClientTop       =   2775
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Encomenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Encomenda.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   6900
   ScaleWidth      =   8820
   Begin VB.TextBox Text4 
      DataField       =   "Cod_cliente"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2280
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   6480
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Menu"
      Height          =   615
      Left            =   7680
      TabIndex        =   13
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "N_encomenda"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "Data_encomenda"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Data_entrega"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      DataField       =   "Referência"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "     Encomendas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sofia\Documents\Programas\Programa Encomendas\ENCOMENDAS.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Encomendas"
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton Cmdadicionar 
      Caption         =   "Adicionar"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Cmdguardar 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Cmdeliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   5040
      Picture         =   "Encomenda.frx":1784
      Top             =   1920
      Width           =   3195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Encomendas"
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
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   5400
      Picture         =   "Encomenda.frx":4AA4
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de encomenda"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data da encomenda"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de entrega"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Código do cliente"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "Encomenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdadicionar_Click()
On Error GoTo trataerro
If Cmdadicionar.Caption = "Adicionar" Then
    Data1.Recordset.AddNew
    Text1.SetFocus
    Cmdeliminar.Enabled = False
    Cmdguardar.Enabled = True
    Cmdadicionar.Caption = "Cancelar"
Else
    Data1.Recordset.CancelUpdate
    Cmdeliminar.Enabled = True
    Cmdguardar.Enabled = False
    Cmdadicionar.Caption = "Adicionar"
End If
Exit Sub
trataerro:
MsgBox Err.Description
End Sub
Private Sub Cmdeliminar_Click()
Dim resp As Integer, mens As String
On Error GoTo trataerro
mens = "Deseja eliminar este registo?"
resp = MsgBox(mens, vbYesNo, "Eliminar")
If resp = vbNo Then
    MsgBox "Registo não eliminado", 64, "Eliminar"
Else
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then
        Data1.Recordset.MovePrevious
        If Data1.Recordset.EOF Then
            MsgBox "Não há registos"
            Cmdeliminar.Enabled = False
        End If
    End If
    MsgBox ("Registo eliminado")
End If
Exit Sub
trataerro: MsgBox Err.Description
End Sub
Private Sub Cmdguardar_Click()
Dim resp As Integer, mens As String
On Error GoTo trataerro
mens = "Deseja guardar os novos dados?"
resp = MsgBox(mens, vbYesNo, "Guardar")
If resp = vbYes Then
 Data1.Recordset.Update
 Cmdeliminar.Enabled = True
 Cmdguardar.Enabled = False
 Cmdadicionar.Caption = "Adicionar"
End If
Exit Sub
trataerro: MsgBox Err.Description
End Sub

Private Sub Command1_Click()
Index.Show
Encomenda.Hide
End Sub

Private Sub Command2_Click()
Relatório_Encomendas.Show

End Sub

Private Sub Form_Load()
Cmdguardar.Enabled = False
End Sub



