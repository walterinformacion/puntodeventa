VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ventas1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "FACTURACION"
   ClientHeight    =   10905
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   20250
   Icon            =   "ventas.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   10905
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Commandpro 
      Caption         =   "productos"
      Height          =   375
      Left            =   2880
      TabIndex        =   103
      Top             =   9480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Textca 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   95
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Commandn 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4440
      TabIndex        =   91
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CONTROLES ADOC"
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   3120
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   15255
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "codifo"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   101
         Text            =   "codifo"
         Top             =   2760
         Width           =   1755
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   100
         Text            =   "descripcion"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9720
         TabIndex        =   99
         Text            =   "precio"
         Top             =   2520
         Width           =   1755
      End
      Begin VB.TextBox IVA 
         Appearance      =   0  'Flat
         DataField       =   "iva"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   9720
         TabIndex        =   98
         Text            =   "IVA"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox text100 
         Appearance      =   0  'Flat
         DataField       =   "cantidades"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   97
         Text            =   "cantidad"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text39 
         Appearance      =   0  'Flat
         DataField       =   "fecha"
         DataSource      =   "Adodc4"
         Height          =   375
         Left            =   6720
         TabIndex        =   94
         Text            =   "fecha"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         DataField       =   "hora"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   93
         Text            =   "hora"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox Textdd 
         Appearance      =   0  'Flat
         DataField       =   "descuento"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   92
         Text            =   "descuento"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "imprimir tiket"
         Height          =   375
         Left            =   3240
         TabIndex        =   90
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text46 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """*""###""*"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   89
         Text            =   "Text4"
         Top             =   3720
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc10 
         Height          =   375
         Left            =   12480
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   "PROVEEDORES"
         Caption         =   "cliente"
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
      Begin VB.TextBox Textelcli 
         DataField       =   "Telefono"
         DataSource      =   "Adodc10"
         Height          =   285
         Left            =   12480
         TabIndex        =   88
         Text            =   "telefo"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Texdircli 
         DataField       =   "Campo6"
         DataSource      =   "Adodc10"
         Height          =   495
         Left            =   12480
         TabIndex        =   87
         Text            =   "dir"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Texnitcli 
         DataField       =   "NIT"
         DataSource      =   "Adodc10"
         Height          =   285
         Left            =   12480
         TabIndex        =   86
         Text            =   "nit"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Texnomcli 
         DataField       =   "nombre"
         DataSource      =   "Adodc10"
         Height          =   285
         Left            =   12480
         TabIndex        =   85
         Text            =   "nombre"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TextNOLOAD 
         DataField       =   "nombre"
         DataSource      =   "Adodc6"
         Height          =   285
         Left            =   9600
         TabIndex        =   84
         Text            =   "NOMBRE"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   375
         Left            =   1800
         TabIndex        =   83
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text38 
         Appearance      =   0  'Flat
         DataField       =   "cont"
         DataSource      =   "Adodc5"
         Height          =   285
         Left            =   960
         TabIndex        =   82
         Text            =   "cont"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   600
         TabIndex        =   81
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text35 
         Appearance      =   0  'Flat
         DataField       =   "descuento"
         DataSource      =   "Adodc5"
         Height          =   375
         Left            =   960
         TabIndex        =   80
         Text            =   "descuento"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         DataField       =   "iva"
         DataSource      =   "Adodc5"
         Height          =   375
         Left            =   960
         TabIndex        =   79
         Text            =   "iva"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         DataField       =   "subtotal"
         DataSource      =   "Adodc5"
         Height          =   375
         Left            =   960
         TabIndex        =   78
         Text            =   "subtotal"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         DataField       =   "cantidad"
         DataSource      =   "Adodc5"
         Height          =   285
         Left            =   960
         TabIndex        =   77
         Text            =   "cantidad"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text36 
         Appearance      =   0  'Flat
         DataField       =   "precio"
         DataSource      =   "Adodc5"
         Height          =   285
         Left            =   960
         TabIndex        =   76
         Text            =   "precio"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text37 
         Appearance      =   0  'Flat
         DataField       =   "descripcion"
         DataSource      =   "Adodc5"
         Height          =   285
         Left            =   960
         TabIndex        =   75
         Text            =   "descripcion"
         Top             =   1200
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   600
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         RecordSource    =   "tikeck"
         Caption         =   "TIKECT"
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
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         DataField       =   "fecha"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   74
         Text            =   "fecha"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         DataField       =   "numerofactura"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   73
         Text            =   "numero factura"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         DataField       =   "subtotal"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   72
         Text            =   "subtotal"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         DataField       =   "iva"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   71
         Text            =   "iva"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         DataField       =   "total"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   70
         Text            =   "total"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         DataField       =   "efectivo"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   69
         Text            =   "efectivo"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         DataField       =   "cambio"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   68
         Text            =   "cambio"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         DataField       =   "cliente"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   67
         Text            =   "cliente"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         DataField       =   "vendedor"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3720
         TabIndex        =   66
         Text            =   "vendedor"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox cambio 
         Height          =   495
         Left            =   360
         TabIndex        =   65
         Top             =   6960
         Width           =   2295
      End
      Begin VB.TextBox EFECTIVO 
         Height          =   495
         Left            =   360
         TabIndex        =   64
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         DataField       =   "iva"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   63
         Text            =   "VALOR IVA"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         DataField       =   "descuento"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   62
         Text            =   "descuento"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         DataField       =   "precioxcantidad"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   61
         Text            =   "PREXCANT"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         DataField       =   "precio"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   60
         Text            =   "PRECIO"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         DataField       =   "articulo"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   59
         Text            =   "DESCRIPCION"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         DataField       =   "cantidad"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   58
         Text            =   "CANTIDAD"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         DataField       =   "numerofactura"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   6720
         TabIndex        =   57
         Text            =   "numero factura"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         DataField       =   "cliente"
         DataSource      =   "Adodc4"
         Height          =   375
         Left            =   6720
         TabIndex        =   56
         Text            =   "cliente"
         Top             =   2640
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   6360
         Top             =   600
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   582
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
         RecordSource    =   "COMPRAS"
         Caption         =   "VENTAS"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   3360
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         RecordSource    =   "CONTROLCOMPRAS"
         Caption         =   "CONTROLDE VENTA"
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
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   375
         Left            =   9600
         Top             =   600
         Visible         =   0   'False
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
         Caption         =   "LOAD"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   9360
         Top             =   1920
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
   End
   Begin VB.CommandButton Command4 
      Caption         =   "otro"
      Height          =   495
      Left            =   11280
      TabIndex        =   54
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox desv 
      Height          =   285
      Left            =   2040
      TabIndex        =   53
      Top             =   10080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox desp 
      Height          =   495
      Left            =   2040
      TabIndex        =   52
      Text            =   "no"
      Top             =   9360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox totaldescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   570
      Left            =   13920
      TabIndex        =   47
      Top             =   8280
      Width           =   3135
   End
   Begin VB.CommandButton des20 
      Caption         =   "20%"
      Height          =   495
      Left            =   10560
      TabIndex        =   46
      Top             =   8280
      Width           =   495
   End
   Begin VB.CommandButton des15 
      Caption         =   "15%"
      Height          =   495
      Left            =   9840
      TabIndex        =   45
      Top             =   8280
      Width           =   495
   End
   Begin VB.CommandButton des10 
      Caption         =   "10%"
      Height          =   495
      Left            =   9120
      TabIndex        =   44
      Top             =   8280
      Width           =   495
   End
   Begin VB.CommandButton des5 
      Caption         =   "5%"
      Height          =   495
      Left            =   8400
      TabIndex        =   43
      Top             =   8280
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cobrar"
      Height          =   375
      Left            =   10080
      TabIndex        =   42
      Top             =   9240
      Width           =   2055
   End
   Begin VB.TextBox buscadorclinetes 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   41
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox codigo 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   40
      Top             =   8160
      Width           =   3855
   End
   Begin VB.TextBox filaa2 
      Height          =   285
      Left            =   720
      TabIndex        =   37
      Text            =   "0"
      Top             =   10080
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   16711680
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox subtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   525
      Left            =   13920
      TabIndex        =   36
      Top             =   7680
      Width           =   3135
   End
   Begin VB.TextBox totaliva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   525
      Left            =   13920
      TabIndex        =   35
      Top             =   8880
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   17880
      Top             =   1080
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox direcion 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12480
      TabIndex        =   24
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox telefono 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   23
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox nitcliente 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   22
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox nombrecliente 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   21
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "cobrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   10080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command17 
      Caption         =   "cancelar venta"
      Height          =   375
      Left            =   9960
      TabIndex        =   19
      Top             =   9720
      Width           =   2175
   End
   Begin VB.TextBox TOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   570
      Left            =   13920
      TabIndex        =   18
      Top             =   9480
      Width           =   3255
   End
   Begin VB.CommandButton Command15 
      Caption         =   "cancelar"
      Height          =   1455
      Left            =   10800
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      Caption         =   "seleccionar"
      Height          =   1455
      Left            =   10800
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   10080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Text            =   "0"
      Top             =   9600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Text            =   "0"
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox filaa 
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Text            =   "0"
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1560
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label LabelFACTURACION 
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   102
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Labeltd 
      BackStyle       =   0  'Transparent
      Caption         =   "stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   96
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image Imagebus 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   12000
      Picture         =   "ventas.frx":E95A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12840
      TabIndex        =   51
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCUENTO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12120
      TabIndex        =   50
      Top             =   8400
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA TOTAL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12240
      TabIndex        =   49
      Top             =   9000
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12360
      TabIndex        =   48
      Top             =   7800
      Width           =   1395
   End
   Begin VB.Image Imagensalidas 
      Height          =   1335
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Command13 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   7200
      Picture         =   "ventas.frx":2A705
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   735
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   6360
      Picture         =   "ventas.frx":43242
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image Command11 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5520
      Picture         =   "ventas.frx":522E8
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image Command10 
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   4680
      Picture         =   "ventas.frx":6E093
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   705
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1320
      TabIndex        =   39
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "venta N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   330
      Left            =   240
      TabIndex        =   38
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   10800
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10920
      TabIndex        =   34
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11880
      TabIndex        =   33
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11040
      TabIndex        =   32
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label nobrelab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11280
      TabIndex        =   31
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   330
      Left            =   3000
      TabIndex        =   30
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label411 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   330
      Left            =   6240
      TabIndex        =   29
      Top             =   240
      Width           =   600
   End
   Begin VB.Label fecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3840
      TabIndex        =   28
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label hora 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6840
      TabIndex        =   27
      Top             =   240
      Width           =   1635
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   120
      Top             =   960
      Width           =   10335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "comentario"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   25
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   7800
      Picture         =   "ventas.frx":8A7A9
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Iva"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "preXcant"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Descipción"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "cantidad"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Menu INICO 
      Caption         =   "INICIO"
   End
   Begin VB.Menu VENTAS 
      Caption         =   "VENTAS"
   End
End
Attribute VB_Name = "ventas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fila As Integer
Dim tot As Double
Dim x As Double
Dim TM As Double
Dim descuentosi As Double






Private Sub botonbus_Click()

End Sub



Private Sub buscadorclinetes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Imagebus_Click
Adodc10.Refresh

End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
Command10_Click
End If
If KeyAscii = 43 Then
Command16_Click
End If

End Sub

Private Sub com_Click()

End Sub

Private Sub Command1_Click()
Me.Hide
Dim ventas2 As New ventas1
ventas2.Show


End Sub

Private Sub Command13_Click()

MSFlexGrid1.Col = 0
MSFlexGrid1.Row = filaa
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = filaa
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = filaa
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = filaa
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = filaa
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = filaa
If totaliva = Empty Then
Else
totaliva = totaliva - MSFlexGrid1.Text
End If
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 6
MSFlexGrid1.Row = filaa
If MSFlexGrid1.Text = "$0" Then
Else
totaldescuento = "$" & totaldescuento - MSFlexGrid1.Text
End If
l = MSFlexGrid1.Text
MSFlexGrid1.Text = ""
MSFlexGrid1.Col = 7
MSFlexGrid1.Row = filaa
If subtotal = "" Then


Else
f = Replace(MSFlexGrid1.Text, "$", "")
ñ = Replace(subtotal, "$", "")
subtotal = ñ - f
End If
MSFlexGrid1.Text = ""


If subtotal = Text Then

End If

Text9 = Replace(Text9, "$", "")
tot = tot - Text9
TOTAL = Format(tot, "$ ##,#")


MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)

 



End Sub

Private Sub Command16_Click()
Dim n As Integer
Dim j As Integer
Dim t As Integer

Rem *********************** buscar ultima factura**************************************
Adodc3.Recordset.MovePrevious
Adodc3.Recordset.MoveLast
t = Text20 + 1
Adodc3.Refresh
Rem *********************** funcion de guardado el control de venta*****************************
Adodc3.Recordset.AddNew
Text20 = t
Text21 = subtotal
Textdd = totaldescuento
Text22 = totaliva
Text23 = TOTAL
Text24 = EFECTIVO
Text25 = cambio
If nombrecliente = Empty Then
Text26 = "Venta Mostrador"
Else
Text26 = nombrecliente
End If
Text27 = hora
Text28 = fecha
Text29 = TextNOLOAD
Adodc3.Recordset.Update
Adodc3.Refresh
Rem *********************** fin de guardado el control de venta*****************************

Rem *********************** inicio de guardado  de venta*****************************

w = MSFlexGrid1.Rows
For j = 1 To w - 1

    With MSFlexGrid1
        Adodc4.Recordset.AddNew
        Text19 = t
        Text18 = (.TextMatrix(j, 0))
        Text17 = (.TextMatrix(j, 1))
        Text16 = (.TextMatrix(j, 2))
        Text15 = (.TextMatrix(j, 3))
        Text14 = (.TextMatrix(j, 6))
        Text13 = (.TextMatrix(j, 5))
               If nombrecliente = Empty Then
                 Text30 = "Venta Mostrador"
                 Else
                 Text30 = nombrecliente
               End If
         Text39 = fecha
        
        Adodc4.Recordset.MoveLast
        Adodc4.Refresh
      End With
Next j
  Adodc4.Refresh
Rem *********************** fin de guardado  de venta*****************************
  Rem ********************* aumenta articulo por compra **************
  Dim Y As String
  Dim canti As Integer
  Dim canti1 As Integer
  
  Y = Adodc1.Recordset.RecordCount
  j = 0
  For j = 1 To w - 1
  With MSFlexGrid1
  canti = (.TextMatrix(j, 0))
  busqueda = (.TextMatrix(j, 1))
  Adodc1.Refresh
  Adodc1.Recordset.Find "desciocion='" & Trim(busqueda) & "'"
  canti1 = text100
  text100 = canti1 - canti
  Adodc1.Recordset.Update
  Adodc1.Recordset.MoveFirst
   End With
  Next j
    Rem *********************fim  aumenta articulo por compra **************
  Command5_Click
Exit Sub

End Sub

Private Sub Command12_Click()
Frame2.Visible = False
Frame1.Visible = True
End Sub





Private Sub Command7_Click()
Dim x1app As New Excel.Application
Dim n As Integer
Dim m As Integer

Rem ********************************** codigo exel*********************************

Workbooks.Open "C:\para motos\tiketes", , False
x1app.Visible = True
x1app.WindowState = xlNormal

Rem ****************************** datos a incertar *******************
 x1app.Cells(13, 4) = Format(Date, "DD/MM/YY")
x1app.Cells(12, 4) = Text20 + 1
x1app.Cells(9, 3) = TextNOLOAD.Text
If nombrecliente = Empty Then
x1app.Cells(11, 2) = "VENTA MOSTRADOR"
Else
x1app.Cells(11, 2) = nombrecliente
x1app.Cells(12, 2) = telefono
x1app.Cells(13, 2) = "NIT:" & nitcliente
x1app.Cells(14, 1) = direcion
End If


m = 16
For n = 1 To filaa
 With MSFlexGrid1
.Row = n
x1app.Cells(m, 2) = (.TextMatrix(.Row, 1))
x1app.Cells(m, 1) = (.TextMatrix(.Row, 0))
h = Format((.TextMatrix(.Row, 3)), "$##,#0.")
x1app.Cells(m, 3) = h
x1app.Cells(m, 3).horizontalAlignment = xlRight
p = Format((.TextMatrix(.Row, 7)), "$##,#0.")
x1app.Cells(m, 4) = p
x1app.Cells(m, 4).horizontalAlignment = xlRight


m = m + 1

 End With
  Next n
  G = 15 + n
  
 x1app.Cells(G, 1) = "-----------------------------------"
 G = G + 1
 x1app.Cells(G, 2) = "SUBTOTAL:"
 x1app.Cells(G, 4) = subtotal
 G = G + 1
  x1app.Cells(G, 2) = "DESCUENTO:"
 x1app.Cells(G, 4) = totaldescuento
 G = G + 1
  x1app.Cells(G, 2) = "IVA:"
 x1app.Cells(G, 4) = totaliva
  G = G + 1
  x1app.Cells(G, 2) = "TOTAL:"
 x1app.Cells(G, 4) = TOTAL
 
 
  

 Rem x1app.Worksheets("HOJA1").PrintOut
 
 Rem x1app.ThisWorkbook.Save
 Rem x1app.ThisWorkbook.Close
  x1app.Application.Quit

 

  
  
  
End Sub

Private Sub Command8_Click()
MSFlexGrid1.AddItem "1" & vbTab & var_awb, MSFlexGrid1.RowSel
End Sub

Private Sub Command9_Click()
If Combo1 <> "" Then
Adodc1.Recordset.MoveFirst
While Not (Adodc1.Recordset.EOF = True)
  If Combo1 = Adodc1.Recordset(1) Then
  Exit Sub
  End If
Adodc1.Recordset.MoveNext
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Visible = False
Combo1.Visible = False
Text2.Visible = True
Wend
End If
End Sub

Private Sub Command14_Click()
If List2 <> "" Then
Adodc1.Recordset.MoveFirst
While Not (Adodc1.Recordset.EOF = True)
  If List2 = Adodc1.Recordset(1) Then
  If Val(Text5) > (0) Then
Command10.Visible = True
Command11.Visible = True


Command14.Visible = False
Command15.Visible = False
Command16.Visible = True
Command17.Visible = True
Else
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True

Command14.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
End If
MSFlexGrid1.Visible = True
List2.Visible = False
    B = InputBox((Text1) + ("    ") + (Text2) + ("    Precio: $") + (Text3) + ("           ¿Cantidad A Vender?"), "Ingresar Cantidad")
    If Val(B) > (0) Then
    Text4 = Val(B) * Val(Text3)
     c = Val(Text4) + Val(Label4)
     Label4 = Val(c)
  
  
     Text5 = Text1
     Text6 = Text2
     Text7 = Val(B)
     Text8 = Text3
     Text9 = Text4
       Text10 = IVA
     fila = fila + 1
     
     MSFlexGrid1.Col = 0
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text7.Text
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text6.Text
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text8.Text
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text9.Text
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text10.Text
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = fila
x = Val(Text7) * Val(Text8)
MSFlexGrid1.Text = x
If Text10 = "0" Then
Else
Text11 = Text8 / ((Text10 / 100) + 1)
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text11.Text, "$ ##,#")

Text12 = Text11 * Text7
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text12.Text, "$ ##,#")
End If


tot = tot + x
TOTAL = Format(tot, "$ ##,#")
filaa = filaa + 1

  If Val(Text5) > (0) Then
Command10.Visible = True
Command11.Visible = True


Command14.Visible = False
Command15.Visible = False
Command16.Visible = True
Command17.Visible = True
Else
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Command14.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
End If
  End If
  Exit Sub
  End If
Adodc1.Recordset.MoveNext

Wend
End If
End Sub

Private Sub Command15_Click()

If Val(Text5) > (0) Then
List2.Clear
Command10.Visible = True
Command11.Visible = True

MSFlexGrid1.Visible = True
Command14.Visible = False
Command15.Visible = False
Command16.Visible = True
Command17.Visible = True

List2.Visible = False
Else
List2.Clear
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Command14.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
DataGrid1.Visible = True
List2.Visible = False
End If
End Sub



Private Sub Command3_Click()
ventanadecambio.Show
ventanadecambio.Text1 = TOTAL

End Sub

Private Sub Command4_Click()

Dim i As Integer
Dim m As Integer
Dim ñ As Integer

If desp = "si" Then
A = InputBox("ingrese el valor a aplicar ", "SISTEMA DE DESCUENTO")
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
net = MSFlexGrid1.Text
Q = MSFlexGrid1.Text
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 6

net = net * (A / 100)
MSFlexGrid1.Text = "$" & net
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
MSFlexGrid1.Text = Format(Q - net, "$##,##")
m = MSFlexGrid1.Rows - 1
For i = 1 To m
If MSFlexGrid1.Text = Empty Then
Else
MSFlexGrid1.Row = i
MSFlexGrid1.Col = 6
ñ = MSFlexGrid1.Text + ñ
totaldescuento = ñ
End If
Next

subtotal = Format(tot, "$##,#0")
TOTAL = Format(tot - ñ, "$##,#0")
tot = TOTAL

Else

 MsgBox "Seleccione un articulo para aplicar el descuento", vbRetryCancel, "ERROR DESCUENTO"
 End If

End Sub

Private Sub Command5_Click()
Dim h As Integer
Dim e As Integer
Dim ñ As Integer

h = Adodc5.Recordset.RecordCount

For ñ = 1 To h
Adodc5.Refresh
Adodc5.Recordset.Delete
Adodc5.Recordset.MoveNext



Next ñ
p = MSFlexGrid1.Rows - 1

For e = 1 To p
 With MSFlexGrid1

Adodc5.Recordset.AddNew
Text37 = (.TextMatrix(e, 1))
Text36 = (.TextMatrix(e, 3))
Text31 = (.TextMatrix(e, 0))
Text32 = (.TextMatrix(e, 7))
Text33 = (.TextMatrix(e, 5))
Text35 = (.TextMatrix(e, 6))
Text38 = e
Adodc5.Recordset.MovePrevious

Adodc5.Refresh
 End With
  Next e
  Text46 = "*" & Label14 & "*"
Adodc5.Refresh
Command6_Click
End Sub

Private Sub Command6_Click()

DataReport1.Sections("ReportFooter").Controls("Etiqueta3").Caption = subtotal
DataReport1.Sections("ReportFooter").Controls("Etiqueta5").Caption = totaldescuento
DataReport1.Sections("ReportFooter").Controls("Etiqueta7").Caption = TOTAL
DataReport1.Sections("ReportHeader").Controls("Etiqueta38").Caption = TextNOLOAD
DataReport1.Sections("ReportHeader").Controls("Etiqueta22").Caption = Label14
DataReport1.Sections("ReportFooter").Controls("Etiqueta43").Caption = Format(Text46, "*###*")
If nombrecliente = Text Then
Else

DataReport1.Sections("ReportHeader").Controls("Etiqueta14").Caption = nombrecliente
DataReport1.Sections("ReportHeader").Controls("Etiqueta16").Caption = nitcliente
DataReport1.Sections("ReportHeader").Controls("Etiqueta18").Caption = telefono
DataReport1.Sections("ReportHeader").Controls("Etiqueta20").Caption = direcion

End If

Rem DataReport1.PrintReport
 DataReport1.Show

Command17_Click

End Sub

Private Sub Command10_Click()
Dim subt
If TOTAL = Empty Then
Else
tot = TOTAL
End If

If totaliva.Text = Empty Then
Else
TM = totaliva
End If


If codigo.Text = Text5 Then
If Text7 >= text100 Then
MsgBox "no hay inventario suficiente"
Exit Sub
Else
End If

On Error GoTo salida

Text7 = Text7 + 1
Text9 = Text7 * Text8
 c = Val(Text8) + Val(Label4)
     Label4 = Val(c)
     
tot = tot + Text8
TOTAL = Format(tot, "$ ##,#")

 
MSFlexGrid1.Col = 0
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text7.Text
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text6.Text
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text8.Text, "$ ##,#")
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text9.Text, "$ ##,#")
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text10.Text, "%##")
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = "$0"
MSFlexGrid1.Col = 6
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = "$0"
MSFlexGrid1.Col = 7
MSFlexGrid1.Row = fila
x = Val(Text7) * Val(Text8)
MSFlexGrid1.Text = Format(x, "$ ##,#")

codigo = ""


If Text10 = "0" Then
Else
Text11 = Text8 / ((Text10 / 100) + 1)
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text11.Text, "$ ##,#")

Text12 = Text11 * Text7
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text12.Text, "$ ##,#")

MSFlexGrid1.Col = 6
MSFlexGrid1.Row = fila
Z = x - Text12
MSFlexGrid1.Text = Format(Z, "$ #,##")
TM = Z
tot = tot + Text8
TOTAL = Format(tot, "$ ##,#")
totaliva = Format(TM, "$ ##,#")
subtotal = TOTAL - totaliva
subtotal = Format(subtotal, "$ ##,#")

codigo = ""

End If



Else

A = codigo.Text
If Val(A) > (0) Then
Adodc1.Recordset.MoveFirst
While Not (Adodc1.Recordset.EOF = True)
  If Val(A) = Adodc1.Recordset(0) Then
    If Val(A) > (0) Then
    B = 1
    If Val(B) > (0) Then
 
 If Text7 >= text100 Then
MsgBox "no hay inventario suficiente"
Exit Sub
Else
End If
 
    Text4 = Val(B) * Val(Text3)
     c = Val(Text4) + Val(Label4)
     Label4 = Val(c)
      
   codigo = ""
   codigo.BackColor = &H80000005
      Text5 = Text1
     Text6 = Text2
     Text7 = Val(B)
     Text8 = Text3
     Text9 = Text4
     Text10 = IVA
     filaa = filaa + 1
     fila = fila + 1
     Textca = text100
     
     
     If MSFlexGrid1.Rows = 1 Then
     MSFlexGrid1.Rows = fila + 1
     Else
   MSFlexGrid1.RowSel = 1
   MSFlexGrid1.AddItem "" & vbTab & var_awb, MSFlexGrid1.RowSel
   fila = fila - 1
     End If
     
     
MSFlexGrid1.Col = 0
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text7.Text
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Text6.Text
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text8.Text, "$ ##,#")
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text9.Text, "$ ##,#")
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text10.Text, "%##")
MSFlexGrid1.Col = 6
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = "$0"
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = "$0"
MSFlexGrid1.Col = 7
MSFlexGrid1.Row = fila
x = Val(Text7) * Val(Text8)
MSFlexGrid1.Text = Format(x, "$ ##,#")
tot = tot + x
TOTAL = Format(tot, "$ ##,#")
If Text10 = "0" Then
totaliva = Format(totaliva, "$ ##,#")
Else
Text11 = Text8 / ((Text10 / 100) + 1)
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text11.Text, "$ ##,#")

Text12 = Text11 * Text7
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = fila
MSFlexGrid1.Text = Format(Text12.Text, "$ ##,#")

MSFlexGrid1.Col = 5
MSFlexGrid1.Row = fila
m = x - Text12
MSFlexGrid1.Text = Format(m, "$ ##,#")
TM = TM + m
totaliva = Format(TM, "$ ##,#")

subtotal = TOTAL - totaliva

subtotal = Format(subtotal, "$ ##,#")


 ima = App.Path
Rem Imagensalidas.Picture = LoadPicture(ima & "\imagenes\" & A & ".jpg")
End If
List2.Visible = False
     
     End If
      Exit Sub
       End If
    End If
      Adodc1.Recordset.MoveNext

Wend
 
 codigo.BackColor = &HC0C0FF
Adodc1.Recordset.MoveFirst

End If
End If
codigo.SetFocus
salida:
End Sub


Private Sub Command11_Click()
 buscador1.Show
 
End Sub

Private Sub Command17_Click()
Me.Hide
Dim ventas2 As New ventas1
ventas2.Show

End Sub



Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Public Sub Commandn_Click()
Command16_Click
End Sub

Public Sub Commandpro_Click()
If TOTAL = Empty Then
Else
tot = TOTAL
End If
Command10_Click
End Sub

Private Sub des10_Click()
Dim i As Integer
Dim m As Integer
Dim ñ As Integer

If desp = "si" Then
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
net = MSFlexGrid1.Text
Q = MSFlexGrid1.Text
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 6

net = net * 0.1
MSFlexGrid1.Text = "$" & net
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
MSFlexGrid1.Text = Format(Q - net, "$##,##")
m = MSFlexGrid1.Rows - 1
For i = 1 To m
If MSFlexGrid1.Text = Empty Then
Else
MSFlexGrid1.Row = i
MSFlexGrid1.Col = 6
ñ = MSFlexGrid1.Text + ñ
totaldescuento = ñ
End If
Next

subtotal = Format(tot, "$##,#0")
TOTAL = Format(tot - ñ, "$##,#0")
tot = TOTAL

Else

 MsgBox "Seleccione un articulo para aplicar el descuento", vbRetryCancel, "ERROR DESCUENTO"
 End If


End Sub

Private Sub des15_Click()
Dim i As Integer
Dim m As Integer
Dim ñ As Integer

If desp = "si" Then
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
net = MSFlexGrid1.Text
Q = MSFlexGrid1.Text
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 6

net = net * 0.15
MSFlexGrid1.Text = "$" & net
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
MSFlexGrid1.Text = Format(Q - net, "$##,##")
m = MSFlexGrid1.Rows - 1
For i = 1 To m
If MSFlexGrid1.Text = Empty Then
Else
MSFlexGrid1.Row = i
MSFlexGrid1.Col = 6
ñ = MSFlexGrid1.Text + ñ
totaldescuento = ñ
End If
Next

subtotal = Format(tot, "$##,#0")
TOTAL = Format(tot - ñ, "$##,#0")
tot = TOTAL

Else

 MsgBox "Seleccione un articulo para aplicar el descuento", vbRetryCancel, "ERROR DESCUENTO"
 End If


End Sub

Private Sub des20_Click()
Dim i As Integer
Dim m As Integer
Dim ñ As Integer

If desp = "si" Then
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
net = MSFlexGrid1.Text
Q = MSFlexGrid1.Text
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 6

net = net * 0.2
MSFlexGrid1.Text = "$" & net
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
MSFlexGrid1.Text = Format(Q - net, "$##,##")
m = MSFlexGrid1.Rows - 1
For i = 1 To m
If MSFlexGrid1.Text = Empty Then
Else
MSFlexGrid1.Row = i
MSFlexGrid1.Col = 6
ñ = MSFlexGrid1.Text + ñ
totaldescuento = ñ
End If
Next

subtotal = Format(tot, "$##,#0")
TOTAL = Format(tot - ñ, "$##,#0")
tot = TOTAL

Else

 MsgBox "Seleccione un articulo para aplicar el descuento", vbRetryCancel, "ERROR DESCUENTO"
 End If


End Sub

Private Sub des5_Click()
Dim i As Integer
Dim m As Integer
Dim ñ As Integer
Dim o As Integer

If desp = "si" Then
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
net = MSFlexGrid1.Text
Q = MSFlexGrid1.Text
MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 6
net = Replace(net, "$", "")
net = net * 0.05
MSFlexGrid1.Text = "$" & net

MSFlexGrid1.Row = filaa
MSFlexGrid1.Col = 7
Q = Replace(Q, "$", "")
MSFlexGrid1.Text = Format(Q - net, "$##,##")
totaldescuento = net
m = MSFlexGrid1.Rows - 1
For i = 1 To m
If MSFlexGrid1.Text = Empty Then

Else
MSFlexGrid1.Row = i
MSFlexGrid1.Col = 6
h = Replace(MSFlexGrid1.Text, "$", "")
ñ = h + ñ
totaldescuento = ñ
End If
Next

subtotal = Format(tot, "$##,#0")
TOTAL = Format(tot - ñ, "$##,#0")
k = Replace(TOTAL, "$", "")
tot = k
Else

 MsgBox "Seleccione un articulo para aplicar el descuento", vbRetryCancel, "ERROR DESCUENTO"
 End If


End Sub


Private Sub EFECTIVO_Change()
If EFECTIVO = Text Then
Command16_Click
End If

End Sub

Private Sub Form_Activate()
Adodc3.Recordset.MovePrevious
Adodc3.Recordset.MoveLast
Adodc4.Recordset.MovePrevious
Adodc4.Recordset.MoveLast

 Label14 = Text20 + 1
tot = 0
Adodc10.Refresh
On Error GoTo salida:
 codigo.SetFocus
salida:
End Sub

Private Sub Form_LinkClose()
Rem ventas1.Close
End Sub

Private Sub Form_Load()
 fila = 0
MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.Rows = 1
MSFlexGrid1.Col = 0
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "CANT"
MSFlexGrid1.ColWidth(1) = 6000
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "DESCRIPCION"
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "PRECIO"
MSFlexGrid1.ColWidth(3) = 2500
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "CANTxPRE"
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "IVA"
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "VALOR"
MSFlexGrid1.Col = 6
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "DESCUENTO"
MSFlexGrid1.Col = 7
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "TOTAL"
nombrecliente = ""
nitcliente = ""
telefono = ""
direcion = ""

End Sub



Private Sub Image3_Click()

End Sub

Private Sub Image1_Click()
Me.Refresh

End Sub

Private Sub ImageBUSK_Click()

End Sub

Private Sub Imagen1_Click()

End Sub

Private Sub Imagebus_Click()
Q = Adodc10.Recordset.RecordCount

For m = 1 To Q
   If Texnitcli = buscadorclinetes Then
    nombrecliente = Texnomcli
    nitcliente = Texnitcli
    telefono = Textelcli
    direcion = Texdircli
    Exit Sub
 Else
 On Error GoTo salida
Adodc10.Recordset.MoveNext
End If
Next m
Dim PREGUNTA As Integer
PREGUNTA = MsgBox("EL CLIENTE NO EXISTE DESEA CREARLO", vbYesNo, "BUSCADOR DE CLIENTES")
If PREGUNTA = vbYes Then
clientes.Show
Else
Exit Sub
End If

salida:
End Sub

Private Sub INICO_Click()
Form2.Show
End Sub

Private Sub List2_Click()
order1 = xlAscending

End Sub

Private Sub MSFlexGrid1_Click()

    With MSFlexGrid1
    If .TextMatrix(.Row, 0) = Text Then
    desp = "no"
    
     MsgBox " este espacio no tiene articulos"
        Command13.Visible = False
        Else
        
        Text6 = .TextMatrix(.Row, 1)
        h = Adodc1.Recordset.RecordCount
        Adodc1.Refresh
        For p = 1 To h
          If Text6 = Text2 Then
             Textca = text100
             Exit For
             End If
             Adodc1.Recordset.MoveNext
        Next p
        
        Text7 = .TextMatrix(.Row, 0)
        Text8 = .TextMatrix(.Row, 2)
        Text9 = .TextMatrix(.Row, 3)
        Text10 = .TextMatrix(.Row, 4)
        desv = .TextMatrix(.Row, 6)
        filaa = .Row
        Command13.Visible = True
        desp = "si"
        
        
        
        
        End If
        End With
        
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Text7 > text100 Then
MsgBox "no hay inventario suficiente"
Exit Sub

Else

If Text7 = "1" Then
Exit Sub
Else
Text9 = Text7 * Replace(Text8, "$", "")
 c = Val(Text8) + Val(Label4)
     Label4 = Val(c)
   Text8 = Replace(Text8, "$", "")
tot = tot + Text8
TOTAL = Format(tot, "$ ##,#")
w = MSFlexGrid1.RowSel

 
MSFlexGrid1.Col = 0
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Text7.Text
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Text6.Text
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Format(Text8.Text, "$ ##,#")
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Format(Text9.Text, "$ ##,#")
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Format(Text10.Text, "%##")
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = w
MSFlexGrid1.Text = "$0"
MSFlexGrid1.Col = 6
MSFlexGrid1.Row = w
MSFlexGrid1.Text = "$0"
MSFlexGrid1.Col = 7
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Format(Text9.Text, "$ ##,#")

codigo = ""


If Text10 = "0" Then
Else
On Error GoTo salida
Text11 = Text8 / ((Text10 / 100) + 1)
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Format(Text11.Text, "$ ##,#")

Text12 = Text11 * Text7
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = w
MSFlexGrid1.Text = Format(Text12.Text, "$ ##,#")

MSFlexGrid1.Col = 6
MSFlexGrid1.Row = w
Z = x - Text12
MSFlexGrid1.Text = Format(Z, "$ #,##")
TM = Z
tot = tot + Text8
TOTAL = Format(tot, "$ ##,#")
totaliva = Format(TM, "$ ##,#")
subtotal = TOTAL - totaliva
subtotal = Format(subtotal, "$ ##,#")
salida:

codigo = ""
End If
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
hora.Caption = Time
fecha.Caption = Date

End Sub

Private Sub totaldescuento_Change()
totaldescuento = Format(totaldescuento, "$##,#0")
End Sub

Private Sub VENTAS_Click()
ventas1.Show
End Sub
