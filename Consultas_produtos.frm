VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Consultas_produtos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consulta de Produtos"
   ClientHeight    =   7875
   ClientLeft      =   5115
   ClientTop       =   1995
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Consultas_produtos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Consultas_produtos.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7875
   ScaleWidth      =   9795
   Begin VB.CommandButton Command4 
      Caption         =   "Voltar"
      Height          =   615
      Left            =   8640
      TabIndex        =   5
      Top             =   7080
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   960
      Top             =   4560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Consultas_produtos.frx":1784
      OLEDBString     =   $"Consultas_produtos.frx":180C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select Nome_produto, Descrição, Preço_IVA from produto order by Nome_produto;"
      Caption         =   "           Produtos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Consultas_produtos.frx":1894
      Height          =   2415
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar Tudo"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Número de produtos"
      Height          =   975
      Left            =   6360
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Os preços mais baixos"
      Height          =   975
      Left            =   3480
      TabIndex        =   0
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   5520
      Picture         =   "Consultas_produtos.frx":18A9
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultas de Produtos"
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
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Consultas_produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "select* from Produto"
Adodc1.Refresh
DataGrid1.SetFocus
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select count(*)as Produtos from Produto"
Adodc1.Refresh
DataGrid1.SetFocus
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "select Nome_produto, Preço_IVA from Produto order by Preço_IVA"
Adodc1.Refresh
DataGrid1.SetFocus
End Sub

Private Sub Command4_Click()
Consultas.Show
Consultas_produtos.Hide
End Sub
