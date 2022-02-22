VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form proveedores1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16935
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   16935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   9360
      TabIndex        =   23
      Top             =   2520
      Width           =   4695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "BUSCAR RAPIDO"
      Height          =   735
      Left            =   6720
      TabIndex        =   22
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3885
      Left            =   14640
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   240
      Width           =   5655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11880
      Top             =   3240
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "PROVEEDORES"
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
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   6720
      TabIndex        =   20
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      DataField       =   "NIT"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      DataField       =   "Campo6"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      DataField       =   "Telefono"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      DataField       =   "Correo"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUEVO REGISTRO"
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GUARDAR REGISTRO"
      Height          =   735
      Left            =   9240
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SIGUIENTE "
      Height          =   735
      Left            =   6720
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ANTERIOR"
      Height          =   735
      Left            =   11880
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ELIMINAR"
      Height          =   735
      Left            =   9240
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "BUSCAR NIT"
      Height          =   615
      Left            =   9480
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "MODIFICAR"
      Height          =   735
      Left            =   11880
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Bindings        =   "proveedores.frx":0000
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   4365
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   4895
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
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   1080
      Y2              =   4080
   End
   Begin VB.Label nombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      Height          =   195
      Left            =   840
      TabIndex        =   19
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NIT"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION "
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CORREO"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR NUEVO"
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
      Left            =   2040
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      Left            =   2520
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "proveedores1"
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


Adodc1.Recordset.Find "Nombre='" & Trim(Text6) & "'"
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

Private Sub Command8_Click()
List1.Clear
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.MoveNext
   While Not Adodc1.Recordset.EOF
   List1.AddItem (Text1)
   Adodc1.Recordset.MoveNext
   If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst




Exit Sub
     End If
    Wend
End Sub

