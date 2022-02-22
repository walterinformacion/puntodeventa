VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form clientes 
   Caption         =   "THE EXPERTS/clientes"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   Icon            =   "clientes.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6270
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "MODIFICAR"
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   4920
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "clieentes"
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
   Begin VB.CommandButton Command6 
      Caption         =   "BUSCAR"
      Height          =   615
      Left            =   12480
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   10680
      TabIndex        =   14
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ANTERIOR"
      Height          =   495
      Left            =   8640
      TabIndex        =   13
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SIGUIENTE "
      Height          =   495
      Left            =   6840
      TabIndex        =   12
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GUARDAR REGISTRO"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUEVO REGISTRO"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "e-mail"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      DataField       =   "tel"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "direccion"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      DataField       =   "nit"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1440
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "clientes.frx":E95A
      Height          =   3495
      Left            =   6360
      TabIndex        =   16
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6165
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
   Begin VB.Label Label6 
      Caption         =   "MODIFICAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "CLIENTE NUEVO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "CORREO"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "TELEFONO"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "DIRECCION "
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NIT"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label nombre 
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   705
   End
End
Attribute VB_Name = "clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox " PUEDE EMPEZAR A DILIJENCIAR EL FORMATO", vbInformation, " SISTEMA DE REGISTRO"

Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Label5.Visible = True
Text1.SetFocus

Exit Sub
salida:
MsgBox "vas dando clic dos veces en nuevo registro anterior", vbInformation, "SISTEMA DE REGISTROS"

End Sub

Private Sub Command2_Click()
On Error GoTo salida
Adodc1.Recordset.Update
MsgBox "se guardaron los datos correctamente hemos corrido al dato anterior", vbInformation, " SISTEMA DE REGISTRO"
Adodc1.Recordset.MovePrevious
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Label5.Visible = False
Label6.Visible = False
If Adodc1.Recordset.BOF Then
End If
Exit Sub
salida:
MsgBox "los campos estan vacios no se puede guardar hasta llenarlos ", vbInformation, " SISTEMA DE REGISTRO"


End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.BOF Then
ado.Recordset.MovePrevious
End If




End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
ado.Recordset.MoveNext
End If
End Sub

Private Sub Command5_Click()
On Error GoTo salida
Adodc1.Recordset.Delete
MsgBox "se elimino el clinete correctamente", vbInformation, " SISTEMA DE REGISTRO"

Exit Sub
salida:
MsgBox "los campos estan vacios busuqe los datos a eliminar ", vbInformation, " SISTEMA DE REGISTRO"


End Sub

Private Sub Command6_Click()
On Error GoTo salida
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
End If
Dim busqueda As String
busqueda = InputBox("ingrese el nombre del cliente", "SISTEMA DE REGISTRO")
Adodc1.Recordset.Find "NOMBRE='" & Trim(busqueda) & "'"
If Adodc1.Recordset.EOF Then


Exit Sub
End If
Text1.Text = Adodc1.Recordset.Fields(1).Value
Text2.Text = Adodc1.Recordset.Fields(2).Value
Text3.Text = Adodc1.Recordset.Fields(3).Value
Text5.Text = Adodc1.Recordset.Fields(4).Value
Exit Sub
salida:
MsgBox "POR FAVOR INSERTE ALGUN NOMBRE PARA BUSCAR", vbInformation, " SISTEMA DE REGISTRO"

End Sub

Private Sub Command7_Click()
Label6.Visible = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True






End Sub

