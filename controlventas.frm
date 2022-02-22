VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form controlventas 
   Caption         =   "THE EXPERTS/Control De Venta"
   ClientHeight    =   13080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   26760
   Icon            =   "controlventas.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   13080
   ScaleWidth      =   26760
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   12495
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   22040
      _Version        =   393216
      Cols            =   4
      BackColorFixed  =   12632256
      BackColorBkg    =   16777215
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   14640
      TabIndex        =   9
      Top             =   5400
      Width           =   6135
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   3600
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "numerofactura"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1080
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         DataField       =   "cliente"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   600
         TabIndex        =   19
         Text            =   "Text14"
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox Text13 
         DataField       =   "total"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Text            =   "Text13"
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         DataField       =   "fecha"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Text            =   "Text12"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         DataField       =   "numerofactura"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Text            =   "Text10"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         DataField       =   "descuento"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Text            =   "Text11"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         DataField       =   "cantidad"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         DataField       =   "articulo"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         DataField       =   "precio"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         DataField       =   "precioxcantidad"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Text            =   "Text6"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         DataField       =   "numerofactura"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Text            =   "Text7"
         Top             =   3720
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   240
         Top             =   3120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         RecordSource    =   "controlventa"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   3360
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         RecordSource    =   "ventas"
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9855
      Left            =   9480
      TabIndex        =   3
      Top             =   840
      Width           =   4095
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Text            =   "$0"
         Top             =   9600
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Text            =   "$0"
         Top             =   9360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Text            =   "$0"
         Top             =   9120
         Width           =   1695
      End
      Begin VB.ListBox lissubtotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         ItemData        =   "controlventas.frx":E95A
         Left            =   2880
         List            =   "controlventas.frx":E95C
         TabIndex        =   5
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ListBox Listarticulos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   4110
         Left            =   120
         Picture         =   "controlventas.frx":E95E
         Top             =   9000
         Width           =   4035
      End
      Begin VB.Image Image1 
         Height          =   2910
         Left            =   120
         Picture         =   "controlventas.frx":16CBF
         Top             =   120
         Width           =   3885
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   14880
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   14280
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "vender"
      Height          =   375
      Left            =   14280
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "controlventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ventas1.Show

End Sub

Private Sub Command2_Click()
A = MsgBox("Esta Seguro De Eliminar El Registro", vbOKCancel, "Eliminar")
If Val(A) = vbOK Then
Adodc1.Recordset.Delete
End If
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command3_Click()
DataEnvironment1.Command2_Grouping Trim(Text1.Text)
DataReport2.Show 1

End Sub

Private Sub Command4_Click()
Dim i As Integer
lissubtotal.Clear
Listarticulos.Clear
numero = Adodc2.Recordset.RecordCount
For i = 1 To numero
If Text1 = Text7 Then
Listarticulos.AddItem Text4
Listarticulos.AddItem Text3 & " " & Text5 & "     " & Text11
lissubtotal.AddItem ""
W = Text6
lissubtotal.AddItem W
End If
Adodc2.Recordset.MoveNext
Next i
Adodc2.Refresh
End Sub

Private Sub Form_Load()
Adodc1.Refresh
h = Adodc1.Recordset.RecordCount
 
For i = 1 To h
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Text = "FACTURA"
  
    MSFlexGrid1.Col = 1
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Text = "FECHA"
  
    MSFlexGrid1.Col = 2
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Text = "CLIENTE"
  
    MSFlexGrid1.Col = 3
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Text = "VALOR"
  
   MSFlexGrid1.Rows = i + 1
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Row = i
  MSFlexGrid1.Text = Text10.Text
  MSFlexGrid1.ColWidth(0) = 1500
 
  
  MSFlexGrid1.Col = 1
  MSFlexGrid1.Row = i
  MSFlexGrid1.Text = Text12.Text
    MSFlexGrid1.ColWidth(1) = 1500
    
   MSFlexGrid1.Col = 2
  MSFlexGrid1.Row = i
  MSFlexGrid1.Text = Text14.Text
    MSFlexGrid1.ColWidth(2) = 3500
    
  
   MSFlexGrid1.Col = 3
  MSFlexGrid1.Row = i
  MSFlexGrid1.Text = Text13.Text
    MSFlexGrid1.ColWidth(3) = 2500
    
  
  Adodc1.Recordset.MoveNext
Next i
End Sub


Private Sub MSFlexGrid1_Click()
  MSFlexGrid1.Col = 0
  W = MSFlexGrid1.Text
  Text1 = W
   Command4_Click
End Sub
