VERSION 5.00
Begin VB.Form Produtos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Produtos"
   ClientHeight    =   7275
   ClientLeft      =   4725
   ClientTop       =   2385
   ClientWidth     =   9885
   Icon            =   "Produtos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Produtos.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7275
   ScaleWidth      =   9885
   Begin VB.CommandButton Command4 
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
      Left            =   7800
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   8880
      TabIndex        =   18
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular Preço com IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   17
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Referência"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nome_produto"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Descrição"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      DataField       =   "Preço"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "IVA"
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
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "Preço_IVA"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "           Produtos"
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Produto"
      Top             =   6240
      Width           =   3495
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
      Left            =   7800
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
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
      Left            =   7800
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
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
      Left            =   7800
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
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
      TabIndex        =   15
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   6480
      Picture         =   "Produtos.frx":1784
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   2055
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
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço"
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
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço com IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   6240
      Picture         =   "Produtos.frx":32D8
      Top             =   2160
      Width           =   3195
   End
End
Attribute VB_Name = "Produtos"
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
Text5.Text = Text4.Text * 0.21
Text5.Text = Format(Text5, "0.00")
End Sub

Private Sub Command2_Click()
Text6.Text = Text4.Text * 1.21
Text6.Text = Format(Text6, "0.00")
End Sub

Private Sub Command3_Click()
Index.Show
Produtos.Hide
End Sub

Private Sub Command4_Click()
Relatório_Produtos.Show
End Sub

Private Sub Form_Load()
Cmdguardar.Enabled = False
End Sub



