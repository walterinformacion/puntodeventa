VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form items 
   BackColor       =   &H00FFFFFF&
   Caption         =   "productos"
   ClientHeight    =   10950
   ClientLeft      =   345
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "Form3"
   ScaleHeight     =   10950
   ScaleWidth      =   15645
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Bindings        =   "Form3.frx":0000
      Height          =   2175
      Left            =   0
      TabIndex        =   15
      Top             =   8775
      Width           =   15645
      _ExtentX        =   27596
      _ExtentY        =   3836
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
            LCID            =   9226
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
            LCID            =   9226
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
      Caption         =   "Command1"
      Height          =   375
      Left            =   12360
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame items 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Items"
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   9015
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6360
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Text            =   "Text4"
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6360
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   7800
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   7800
         Width           =   2415
      End
      Begin VB.TextBox Text12 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6360
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   7800
         Width           =   2415
      End
      Begin VB.Image foto 
         Height          =   2175
         Left            =   360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image foto1 
         Height          =   1935
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image foto2 
         Height          =   1935
         Left            =   6360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image foto3 
         Height          =   1935
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Image foto4 
         Height          =   1935
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Image foto5 
         Height          =   1935
         Left            =   6360
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Image foto6 
         Height          =   1935
         Left            =   360
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Image foto7 
         Height          =   1935
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Image foto8 
         Height          =   1935
         Left            =   6360
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10320
      Top             =   7440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\para motos\punto de venta.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\para motos\punto de venta.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "existencia"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      DataField       =   "precio"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "desciocion"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "desciocion"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "imagen"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   6120
      Width           =   2775
   End
End
Attribute VB_Name = "items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo salida

If Adodc1.Recordset.BOF Then
End If
Dim busqueda As String
busqueda = InputBox("ingrese el nombre del cliente", "SISTEMA DE REGISTRO")
Adodc1.Recordset.Find "codifo='" & Trim(busqueda) & "'"
If Adodc1.Recordset.EOF Then


Exit Sub
End If
Text1.Text = Adodc1.Recordset.Fields(1).Value
Text2.Text = Adodc1.Recordset.Fields(2).Value
Text3.Text = Adodc1.Recordset.Fields(3).Value

Exit Sub
salida:
MsgBox "POR FAVOR INSERTE ALGUN NOMBRE PARA BUSCAR", vbInformation, " SISTEMA DE REGISTRO"
End Sub

Private Sub Form_Load()



ima = App.Path
foto.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text4 = Label1
Adodc1.Recordset.MoveNext
foto1.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text5 = Label1.Caption
Adodc1.Recordset.MoveNext
foto2.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text6 = Label1.Caption
Adodc1.Recordset.MoveNext
foto3.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text7 = Label1.Caption
Adodc1.Recordset.MoveNext
foto4.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text8 = Label1.Caption
Adodc1.Recordset.MoveNext
foto5.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text9 = Label1.Caption
Adodc1.Recordset.MoveNext
foto6.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text10 = Label1.Caption
Adodc1.Recordset.MoveNext
foto7.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text11 = Label1.Caption
Adodc1.Recordset.MoveNext
foto8.Picture = LoadPicture(ima & "\imagenes\" & Label1.Caption & ".jpg")
Text12 = Label1.Caption
Adodc1.Recordset.MoveFirst






End Sub



Private Sub foto_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text4 & ".jpg")


End Sub


Private Sub foto1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text5 & ".jpg")
End Sub

Private Sub foto2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text6 & ".jpg")
End Sub

Private Sub foto3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text7 & ".jpg")

End Sub

Private Sub foto4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text8 & ".jpg")
End Sub

Private Sub foto5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text9 & ".jpg")
End Sub

Private Sub foto6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text10 & ".jpg")
End Sub

Private Sub foto7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text11 & ".jpg")
End Sub

Private Sub foto8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ima = App.Path
Image1.Picture = LoadPicture(ima & "\imagenes\" & Text12 & ".jpg")
End Sub

