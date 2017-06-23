VERSION 5.00
Begin VB.Form Clientes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Clientes"
   ClientHeight    =   7200
   ClientLeft      =   4530
   ClientTop       =   3165
   ClientWidth     =   9480
   Icon            =   "Clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Clientes.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7200
   ScaleWidth      =   9480
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      Left            =   7200
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Menu"
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
      Left            =   8400
      TabIndex        =   16
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Cod_cliente"
      DataSource      =   "Data1"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nome"
      DataSource      =   "Data1"
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
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      DataField       =   "Morada"
      DataSource      =   "Data1"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      DataField       =   "Localidade"
      DataSource      =   "Data1"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "N_contribuinte"
      DataSource      =   "Data1"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Contacto"
      DataSource      =   "Data1"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Cmdadicionar 
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
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Cmdguardar 
      Caption         =   "Guardar"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Cmdeliminar 
      Caption         =   "Eliminar"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "         Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sofia\Documents\Programas\Programa Encomendas\ENCOMENDAS.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cliente"
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   5760
      Picture         =   "Clientes.frx":1784
      Top             =   2280
      Width           =   3195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes"
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
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   5880
      Picture         =   "Clientes.frx":4AA4
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo do cliente"
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
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Morada"
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
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Localidade"
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
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de contribuite"
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
      TabIndex        =   10
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto"
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
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
End
Attribute VB_Name = "Clientes"
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
Clientes.Hide
End Sub

Private Sub Command2_Click()
Relatório_Clientes.Show
End Sub

Private Sub Form_Load()
Cmdguardar.Enabled = False
End Sub




