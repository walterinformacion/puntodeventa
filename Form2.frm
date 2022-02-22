VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form menu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "inicio / THE EXPERTS"
   ClientHeight    =   12570
   ClientLeft      =   -105
   ClientTop       =   750
   ClientWidth     =   22200
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   Picture         =   "Form2.frx":E95A
   ScaleHeight     =   12570
   ScaleWidth      =   22200
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   19680
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         DataField       =   "tiqueck"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   48
         Text            =   "tikeck"
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         DataField       =   "usuarios"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   47
         Text            =   "usuarios"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         DataField       =   "configuracion"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   46
         Text            =   "configuracion"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         DataField       =   "reportes"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   45
         Text            =   "reportes"
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         DataField       =   "controldecompras"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   44
         Text            =   "controlcpmpras"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         DataField       =   "clientes"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   43
         Text            =   "clientes"
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         DataField       =   "proveedores"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   42
         Text            =   "proveedores"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         DataField       =   "cuentasxcobrar"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   41
         Text            =   "cuentasxcobrar"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         DataField       =   "inventario"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   40
         Text            =   "inventario"
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         DataField       =   "combenir"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   39
         Text            =   "combenir"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         DataField       =   "controlventas"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   38
         Text            =   "controlventas"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "cuentasxpagar"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   37
         Text            =   "cuentasxpagar"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "compras"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   36
         Text            =   "compras"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "cotizaciom"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   35
         Text            =   "cotizacion"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "ventas"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   0
         TabIndex        =   34
         Text            =   "ventas"
         Top             =   840
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   240
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
         RecordSource    =   "load"
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
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   720
      X2              =   13920
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Image Image33 
      Height          =   960
      Left            =   15600
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":15862
      Top             =   4080
      Width           =   960
   End
   Begin VB.Image Image32 
      Height          =   1515
      Left            =   16200
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":16426
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "reportes"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   49
      Top             =   8040
      Width           =   945
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "waije software"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   510
      Left            =   16320
      TabIndex        =   32
      Top             =   360
      Width           =   3585
   End
   Begin VB.Image Image31 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   15240
      Picture         =   "Form2.frx":1A347
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "reportes"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   31
      Top             =   8040
      Width           =   945
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
      Height          =   195
      Left            =   12360
      TabIndex        =   30
      Top             =   6720
      Width           =   210
   End
   Begin VB.Image Image30 
      Height          =   1515
      Left            =   12360
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":23201
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "configuracion"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   29
      Top             =   8040
      Width           =   1515
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(C)"
      Height          =   195
      Left            =   9120
      TabIndex        =   28
      Top             =   6720
      Width           =   195
   End
   Begin VB.Image Image29 
      Height          =   1515
      Left            =   9360
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":27122
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "usuarios"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   27
      Top             =   8040
      Width           =   945
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(U)"
      Height          =   195
      Left            =   6120
      TabIndex        =   26
      Top             =   6720
      Width           =   210
   End
   Begin VB.Image Image28 
      Height          =   1515
      Left            =   6240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":2B62B
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ticket"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F12)"
      Height          =   195
      Left            =   3360
      TabIndex        =   24
      Top             =   6720
      Width           =   360
   End
   Begin VB.Image Image27 
      Height          =   1515
      Left            =   3360
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":2FE70
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "inventario"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F11)"
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   6720
      Width           =   360
   End
   Begin VB.Image Image26 
      Height          =   1515
      Left            =   840
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":33EEE
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cuentas X cobrar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   21
      Top             =   5160
      Width           =   1875
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F10)"
      Height          =   195
      Left            =   12000
      TabIndex        =   20
      Top             =   3840
      Width           =   360
   End
   Begin VB.Image Image25 
      Height          =   1515
      Left            =   12240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":36F03
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "proveedores"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   17
      Top             =   5160
      Width           =   1395
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F9)"
      Height          =   195
      Left            =   8880
      TabIndex        =   16
      Top             =   3840
      Width           =   270
   End
   Begin VB.Image Image24 
      Height          =   1440
      Left            =   9000
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":3B71F
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "control compra"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   5160
      Width           =   1680
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F8)"
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   3840
      Width           =   270
   End
   Begin VB.Image Image23 
      Height          =   1515
      Left            =   6240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":3F9E4
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image Image8 
      Height          =   960
      Left            =   6360
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":42945
      Top             =   4080
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F6)"
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   270
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "control ventas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   5160
      Width           =   1560
   End
   Begin VB.Image Image22 
      Height          =   1515
      Left            =   840
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":432A3
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cuentas X pagar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   19
      Top             =   2640
      Width           =   1770
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F5)"
      Height          =   195
      Left            =   11880
      TabIndex        =   18
      Top             =   1320
      Width           =   270
   End
   Begin VB.Image Image21 
      Height          =   1515
      Left            =   12000
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":46204
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F4)"
      Height          =   195
      Left            =   9000
      TabIndex        =   14
      Top             =   1320
      Width           =   270
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "clientes"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image Image20 
      Height          =   1440
      Left            =   9000
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":49507
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "compras"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F3)"
      Height          =   195
      Left            =   5880
      TabIndex        =   6
      Top             =   1320
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cotizacion "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F2)"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   270
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   600
      X2              =   13920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   3240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":4CA80
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F7)"
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      Top             =   3840
      Width           =   270
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "combenir"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Image Image19 
      Height          =   1395
      Left            =   3600
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":50E44
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image Image18 
      Height          =   960
      Left            =   12720
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":55A7A
      Top             =   7080
      Width           =   960
   End
   Begin VB.Image Image17 
      Height          =   960
      Left            =   9600
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5663E
      Top             =   6960
      Width           =   960
   End
   Begin VB.Image Image16 
      Height          =   960
      Left            =   6480
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5721D
      Top             =   6960
      Width           =   960
   End
   Begin VB.Image Image15 
      Height          =   960
      Left            =   3720
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":57DDD
      Top             =   6960
      Width           =   960
   End
   Begin VB.Image Image14 
      Height          =   960
      Left            =   1080
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":58A13
      Top             =   6960
      Width           =   960
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   720
      X2              =   13920
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Image Image13 
      Height          =   960
      Left            =   12480
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":594F0
      Top             =   4080
      Width           =   960
   End
   Begin VB.Image Image12 
      Height          =   960
      Left            =   12240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5A106
      Top             =   1560
      Width           =   960
   End
   Begin VB.Image Image11 
      Height          =   960
      Left            =   9240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5ACD1
      Top             =   4080
      Width           =   960
   End
   Begin VB.Image Image10 
      Height          =   960
      Left            =   9120
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5B81E
      Top             =   1560
      Width           =   960
   End
   Begin VB.Image Image9 
      Height          =   960
      Left            =   3720
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5C4CD
      Top             =   4080
      Width           =   960
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   480
      X2              =   14280
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Image Image7 
      Height          =   960
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5D1A8
      Top             =   4080
      Width           =   960
   End
   Begin VB.Image Image5 
      Height          =   1575
      Left            =   6000
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":5DB06
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   960
      Left            =   3480
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":61C14
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F1)"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ventas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   1515
      Left            =   720
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":62707
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":661D6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image Image6 
      Height          =   960
      Left            =   6240
      MousePointer    =   1  'Arrow
      Picture         =   "Form2.frx":66C2D
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   14655
   End
   Begin VB.Menu INICIO 
      Caption         =   "INICIO"
      WindowList      =   -1  'True
   End
   Begin VB.Menu VENTAS 
      Caption         =   "VENTAS"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Visible = True
End Sub
Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image20.Visible = True
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image24.Visible = True
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image21.Visible = True
End Sub
Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image25.Visible = True
End Sub
Private Sub Image14_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image26.Visible = True
End Sub
Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image27.Visible = True
End Sub
Private Sub Image16_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image28.Visible = True
End Sub
Private Sub Image17_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image29.Visible = True
End Sub
Private Sub Image18_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image30.Visible = True
End Sub

Private Sub Image19_Click()
If Text6 = "si" Then
MsgBox "falta editar"
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image2_Click()
If Text1 = "si" Then
ventas1.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If

End Sub

Private Sub Image20_Click()
If Text8 = "si" Then
clientes.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image21_Click()
If Text4 = "si" Then
MsgBox "falta editar"
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image22_Click()
If Text6 = "si" Then
controlventas.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If

End Sub

Private Sub Image23_Click()
If Text7 = "si" Then
MsgBox "falta editar"
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image24_Click()
If Text9 = "si" Then
proveedores1.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image25_Click()
If Text10 = "si" Then
MsgBox "falta editar"
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image26_Click()
If Text11 = "si" Then
productos.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image27_Click()
If Text12 = "si" Then
MsgBox "falta editar"
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image28_Click()
If Text13 = "si" Then
nivelacceso.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image29_Click()
If Text14 = "si" Then
configuracion.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If

End Sub

Private Sub Image3_Click()
If Text2 = "si" Then
Rem COTIZACION.Show
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image30_Click()
If Text115 = "si" Then
MsgBox "falta editar"
Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image3.Visible = True
End Sub

Private Sub Image5_Click()
If Text3 = "si" Then
COMPRAS.Show

Else
MsgBox "Nivel de acceso no concedido", vbCritical, "Security System"
End If
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Visible = True
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image22.Visible = True
End Sub
Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image23.Visible = True
End Sub
Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image19.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image30.Visible = False
Image29.Visible = False
Image28.Visible = False
Image27.Visible = False
Image26.Visible = False
Image25.Visible = False
Image24.Visible = False
Image23.Visible = False
Image22.Visible = False
Image2.Visible = False
Image19.Visible = False
Image3.Visible = False
Image5.Visible = False
Image20.Visible = False
Image21.Visible = False
End Sub

Private Sub Label35_Click()

End Sub
