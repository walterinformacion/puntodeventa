VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tiquetes1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THE EXPERTS/codigo de barras"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15945
   DrawMode        =   6  'Mask Pen Not
   FillStyle       =   2  'Horizontal Line
   Icon            =   "tiquetes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   6480
      TabIndex        =   1
      Top             =   240
      Width           =   9135
      Begin VB.TextBox Text4 
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "codifo"
         DataSource      =   "Adodc1"
         Height          =   525
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Height          =   1455
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Height          =   1455
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Height          =   1455
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   3  'Align Left
      Height          =   8760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   15452
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\walter\Desktop\para motos\punto de venta.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\walter\Desktop\para motos\punto de venta.mdb;Persist Security Info=False"
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
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "tiquetes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label3_Click()

End Sub
