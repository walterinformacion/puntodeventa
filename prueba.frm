VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form buscador1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "buscador productos"
   ClientHeight    =   10995
   ClientLeft      =   -465
   ClientTop       =   1185
   ClientWidth     =   15675
   LinkTopic       =   "Form2"
   ScaleHeight     =   10995
   ScaleWidth      =   15675
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "vender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   28
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "REFRESCAR"
      Height          =   495
      Left            =   480
      TabIndex        =   25
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "BUSCAR"
      Height          =   495
      Left            =   2280
      TabIndex        =   24
      Top             =   1920
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   360
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   14850
      _ExtentX        =   26194
      _ExtentY        =   5318
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   16320
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   4200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox imag 
         DataField       =   "imagen"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "codifo"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   120
         Top             =   240
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
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   14895
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   4080
      Top             =   240
      Width           =   11175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13680
      TabIndex        =   27
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      DataField       =   "cantidades"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13560
      TabIndex        =   26
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   2205
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3120
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      DataField       =   "codifo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   20
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DESCIPCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "desciocion"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "costo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BUCADOR "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "cantidades"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "precio"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "codifo"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$"" #.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   2400
      Width           =   3495
   End
End
Attribute VB_Name = "buscador1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fila As String
Dim contcaracter As Integer
Dim extraer As String

Dim i As Integer
Dim x As String

Private Sub Command1_Click()
On Error GoTo salida
List2.Clear
For i = 1 To fila - 1
w = UCase(Text2)
h = UCase(Text3)
x = w Like "*" & h & "*"

If x = "Verdadero" Then
  List2.AddItem w
  
  End If
Adodc1.Recordset.MoveNext
Next i
salida:
Adodc1.Refresh

End Sub

Private Sub Command2_Click()
Command1_Click




End Sub

Private Sub Command3_Click()
ventas1.codigo = Label1
ventas1.Commandpro_Click
End Sub

Private Sub Form_Activate()
Text3.SetFocus

End Sub

Private Sub Form_Load()

List1.Clear
Adodc1.Recordset.MoveFirst


   While Not Adodc1.Recordset.EOF
    List1.AddItem UCase(Text2)
    List2.AddItem UCase(Text2)

   Adodc1.Recordset.MoveNext
   If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
Exit Sub
     End If
    Wend

End Sub

Private Sub Label2_Click()
Label2 = Format(Label2, "currency")
End Sub

Private Sub List2_Click()
Adodc1.Refresh
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
End If
Dim busqueda As String
busqueda = List2
Adodc1.Recordset.Find "desciocion='" & Trim(busqueda) & "'"
If Adodc1.Recordset.EOF Then


Exit Sub
End If
Text1.Text = Adodc1.Recordset.Fields(0).Value
Text2.Text = Adodc1.Recordset.Fields(1).Value
On Error GoTo salida
 mopri = App.Path
     Image1.Picture = LoadPicture(mopri & "\imagenes\" & imag & ".jpg")

salida:

Exit Sub





End Sub

Private Sub List2_DblClick()
ventas1.codigo = Label1
ventas1.Commandpro_Click
Me.Hide

End Sub

Private Sub Text3_Change()

 



Rem ******************************** contador de caracter**********************
contcaracter = Len(Text3.Text)
Text4 = contcaracter
Rem ****************************** extraer cararter**************
extraer = Left(Text2, contcaracter)
Text5 = UCase(extraer)
Text7 = UCase(extraer)

Rem **********************************cantidad fe fila *********************
fila = List1.ListCount
Text6 = fila



End Sub

Private Sub Text8_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
