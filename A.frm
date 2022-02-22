VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THE EXPERTS/punto de venta"
   ClientHeight    =   3480
   ClientLeft      =   4620
   ClientTop       =   4035
   ClientWidth     =   12465
   Icon            =   "A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "A.frx":E95A
   ScaleHeight     =   174
   ScaleMode       =   0  'User
   ScaleWidth      =   623.25
   Begin VB.TextBox Text38 
      DataField       =   "turno"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   13800
      TabIndex        =   48
      Text            =   "tunio avierto"
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   47
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text37 
      Appearance      =   0  'Flat
      DataField       =   "tiqueck"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   46
      Text            =   "tikeck"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text36 
      Appearance      =   0  'Flat
      DataField       =   "usuarios"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   45
      Text            =   "usuarios"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text35 
      Appearance      =   0  'Flat
      DataField       =   "configuracion"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   44
      Text            =   "configuracion"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Text34 
      Appearance      =   0  'Flat
      DataField       =   "reportes"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   43
      Text            =   "reportes"
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox Text33 
      Appearance      =   0  'Flat
      DataField       =   "controldecompras"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   42
      Text            =   "controlcpmpras"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text32 
      Appearance      =   0  'Flat
      DataField       =   "clientes"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   41
      Text            =   "clientes"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text31 
      Appearance      =   0  'Flat
      DataField       =   "proveedores"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   40
      Text            =   "proveedores"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text30 
      Appearance      =   0  'Flat
      DataField       =   "cuentasxcobrar"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   39
      Text            =   "cuentasxcobrar"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text29 
      Appearance      =   0  'Flat
      DataField       =   "inventario"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   38
      Text            =   "inventario"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      DataField       =   "tikeck"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   37
      Text            =   "tikeck"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text27 
      Appearance      =   0  'Flat
      DataField       =   "usuarios"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   36
      Text            =   "usuarios"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text26 
      Appearance      =   0  'Flat
      DataField       =   "configuracion"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   35
      Text            =   "configuracion"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Text25 
      Appearance      =   0  'Flat
      DataField       =   "reportes"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   34
      Text            =   "reportes"
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox Text24 
      Appearance      =   0  'Flat
      DataField       =   "controldecompras"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   33
      Text            =   "controlcpmpras"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text23 
      Appearance      =   0  'Flat
      DataField       =   "clientes"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   32
      Text            =   "clientes"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text22 
      Appearance      =   0  'Flat
      DataField       =   "proveedores"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   31
      Text            =   "proveedores"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text21 
      Appearance      =   0  'Flat
      DataField       =   "cuentasxcobrar"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   30
      Text            =   "cuentasxcobrar"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      DataField       =   "inventario"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   29
      Text            =   "inventario"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      DataField       =   "contraseña"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14640
      TabIndex        =   27
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      DataField       =   "nombre"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   14640
      TabIndex        =   26
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   25
      Text            =   "Text4"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      DataField       =   "ventas"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   24
      Text            =   "ventas"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      DataField       =   "cotizaciom"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   23
      Text            =   "cotizacion"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      DataField       =   "compras"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   22
      Text            =   "compras"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      DataField       =   "cuentasxpagar"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   21
      Text            =   "cuentasx pagar"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      DataField       =   "controlventas"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   20
      Text            =   "controlventas"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      DataField       =   "combenir"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   14640
      TabIndex        =   19
      Text            =   "combenir"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      DataField       =   "cobenir"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   18
      Text            =   "combenir"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      DataField       =   "controldeventa"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   17
      Text            =   "controlventas"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "cuentasxpagar"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   16
      Text            =   "cuentasxpagar"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      DataField       =   "compras"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   15
      Text            =   "compras"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      DataField       =   "cotizacion"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   14
      Text            =   "cotizacion"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "ventas"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12600
      TabIndex        =   13
      Text            =   "ventas"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   12480
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   240
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "A.frx":12905
      Height          =   975
      Left            =   600
      TabIndex        =   11
      Top             =   6240
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1720
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "A.frx":1291A
      Height          =   975
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1720
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4680
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Adodc2"
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
      Appearance      =   0  'Flat
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12480
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5640
      TabIndex        =   5
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "145236987"
      Top             =   1080
      Width           =   4485
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   6000
      TabIndex        =   1
      Top             =   1920
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00808080&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   8160
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   600
      Top             =   2640
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "contraseña"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12480
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7320
      Top             =   3720
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   873
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
      RecordSource    =   "usuaios"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " waije software"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   10200
      TabIndex        =   28
      Top             =   3120
      Width           =   2040
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   10680
      Picture         =   "A.frx":1292F
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "contraseña"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "usuario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   1440
      Picture         =   "A.frx":1B7E9
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu ejecutar 
      Caption         =   "menu ventas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ventas 
         Caption         =   "vender"
      End
      Begin VB.Menu historialv 
         Caption         =   "historial"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo salida
    'comprobar si la contraseña es correcta
    Adodc1.Recordset.Find "Nombre='" & Trim(Text1) & "'"
  
    If txtPassword = Text2 And Text1 = Text3 Then
          LoginSucceeded = True
          ProgressBar1.Visible = True
          Text19 = ""
          Text19 = Text2
          Text18 = ""
          Text18 = Text3
          Text17 = ""
          Text17 = Text4
          Text16 = ""
          Text16 = Text5
          Text15 = ""
          Text15 = Text6
          Text14 = ""
          Text14 = Text7
          Text13 = ""
          Text13 = Text8
          Text12 = ""
          Text12 = Text9
          Text11 = ""
          Text11 = Text10
          Text33 = ""
          Text33 = Text24
          Text32 = ""
          Text32 = Text23
          Text31 = ""
          Text31 = Text22
          Text30 = ""
          Text30 = Text21
          Text29 = ""
          Text29 = Text20
          Text37 = ""
          Text37 = Text28
          Text36 = ""
          Text36 = Text27
          Text35 = ""
          Text35 = Text26
          Text34 = ""
          Text34 = Text25
          
                Adodc2.Recordset.MovePrevious
        
        
      Else

        MsgBox "La contraseña o el ususario no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        Adodc1.Refresh
        Text1 = Text3
        txtPassword = ""
       
     
   End If
salida:
    Adodc1.Refresh
End Sub



Private Sub Command1_Click()
ventatipo2.Show


End Sub

Private Sub Command2_Click()


End Sub

Private Sub Form_Load()
If Text38 = "si" Then
Me.Hide
menu.Show

End If

End Sub

Private Sub Timer1_Timer()
If LoginSucceeded = True Then


ProgressBar1.Min = 0
ProgressBar1.Max = 20
ProgressBar1.Value = ProgressBar1 + 1
Label1.Caption = ProgressBar1 & "%"
If ProgressBar1 = 20 Then
Timer1.Enabled = False

      menu.Show
        Me.Hide
        End If
        End If
        
End Sub

