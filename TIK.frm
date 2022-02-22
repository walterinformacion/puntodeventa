VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TIK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THE EXPERTS/codigo de barras"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15075
   DrawMode        =   6  'Mask Pen Not
   FillStyle       =   2  'Horizontal Line
   Icon            =   "TIK.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text16 
      DataField       =   "logo"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   6960
      TabIndex        =   33
      Text            =   "Text16"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   6480
      TabIndex        =   21
      Top             =   6840
      Width           =   7815
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "nombre"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   35
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   6000
         TabIndex        =   32
         Top             =   1560
         Width           =   1695
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   495
         Left            =   7200
         TabIndex        =   31
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text15 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   1095
         Left            =   720
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   6480
      TabIndex        =   20
      Top             =   3960
      Width           =   7815
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   38
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "codifo1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Code 128"
            Size            =   36
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2640
         TabIndex        =   37
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2640
         TabIndex        =   36
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "nombre"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Top             =   1560
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   495
         Left            =   7320
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text14 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   28
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   6000
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   720
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   5175
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "prederminar la impresora "
      Height          =   495
      Left            =   10440
      TabIndex        =   19
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "configuracion de impresoras"
      Height          =   495
      Left            =   12360
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   6480
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   4560
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   495
         Left            =   1800
         Max             =   20
         TabIndex        =   25
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox Text13 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   24
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "nombre"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "codifo1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Code 128"
            Size            =   26.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5520
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "nombre"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "codifo1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Code 128"
            Size            =   26.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2880
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "nombre"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "codifo1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Code 128"
            Size            =   26.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   3  'Align Left
      Bindings        =   "TIK.frx":E95A
      Height          =   9825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   17330
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6120
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7320
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "configuracion"
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
      BackColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "TIK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
IMPRESORAS.Show
End Sub

Private Sub Command2_Click()
cd1.ShowPrinter

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(Text16)
Image2.Picture = LoadPicture(Text16)
End Sub

Private Sub Text13_Change()
Text13 = Format(Text13, "numero")
End Sub

Private Sub Text21_Change()
Text21 = Format(Text21, "$##,#")
End Sub

Private Sub VScroll1_Change()
Text13 = VScroll1.Value
End Sub
