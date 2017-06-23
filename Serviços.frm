VERSION 5.00
Begin VB.Form Serviços 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Serviços"
   ClientHeight    =   6780
   ClientLeft      =   4920
   ClientTop       =   2970
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Serviços.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Serviços.frx":0BC2
   ScaleHeight     =   6780
   ScaleWidth      =   9825
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   7320
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu"
      Height          =   615
      Left            =   8640
      TabIndex        =   11
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "Tipo"
      DataSource      =   "Data1"
      Height          =   390
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton Cmdeliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   7320
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Cmdguardar 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   7320
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Cmdadicionar 
      Caption         =   "Adicionar"
      Height          =   615
      Left            =   7320
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "        Serviços"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sofia\Documents\Programas\Programa Encomendas\ENCOMENDAS.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Serviços"
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "Valor"
      DataSource      =   "Data1"
      Height          =   390
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nome_serviço"
      DataSource      =   "Data1"
      Height          =   390
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Serviço"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   5880
      Picture         =   "Serviços.frx":1784
      Top             =   1800
      Width           =   3195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Serviços"
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   6000
      Picture         =   "Serviços.frx":4AA4
      Top             =   0
      Width           =   3510
   End
End
Attribute VB_Name = "Serviços"
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
Relatório_Serviços.Show
End Sub

Private Sub Command3_Click()
Index.Show
Serviços.Hide
End Sub

Private Sub Form_Load()
Cmdguardar.Enabled = False
End Sub




