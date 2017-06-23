VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Consultas_encomendas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consultas de Encomendas"
   ClientHeight    =   6540
   ClientLeft      =   4350
   ClientTop       =   2580
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Consultas_encomendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Consultas_encomendas.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   6540
   ScaleWidth      =   9960
   Begin VB.CommandButton Command3 
      Caption         =   "Voltar"
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   720
      Top             =   4440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   794
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
      Connect         =   $"Consultas_encomendas.frx":1784
      OLEDBString     =   $"Consultas_encomendas.frx":180C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select N_encomenda, Data_entrega, Nome from Cliente, Encomendas where Encomendas.Cod_cliente=Cliente.Cod_cliente;"
      Caption         =   "      Encomendas"
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
      Bindings        =   "Consultas_encomendas.frx":1894
      Height          =   2295
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   2
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
      Caption         =   "Mostrar tudo"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Número de Encomendas"
      Height          =   735
      Left            =   3720
      TabIndex        =   0
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultas de Encomendas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   6000
      Picture         =   "Consultas_encomendas.frx":18A9
      Top             =   0
      Width           =   3510
   End
End
Attribute VB_Name = "Consultas_encomendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "select* from Encomendas"
Adodc1.Refresh
DataGrid1.SetFocus
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select count(*)as Encomendas from Encomendas"
Adodc1.Refresh
DataGrid1.SetFocus
End Sub

Private Sub Command3_Click()
Consultas.Show
Consultas_encomendas.Hide

End Sub
