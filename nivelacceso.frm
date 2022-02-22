VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form nivelacceso 
   BackColor       =   &H00FFFFFF&
   Caption         =   "NIVEL DE ACCESO"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8175
   Icon            =   "nivelacceso.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   3840
      TabIndex        =   32
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6000
      TabIndex        =   31
      Text            =   "Text9"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CREAR USUARIOS"
      Height          =   495
      Left            =   1920
      TabIndex        =   29
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "CANCELAR"
      Height          =   495
      Left            =   2160
      TabIndex        =   28
      Top             =   5760
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   480
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BUSCAR USUARIO"
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   3840
      TabIndex        =   24
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CheckBox Check7 
      Caption         =   "opcion si"
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      DataField       =   "reportes"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "opcion si"
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "usuarios"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "opcion si"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "proveedores"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "opcion si"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "clientes"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "opcion si"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "inventario"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "ventas"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "opcion si"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox text2 
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5760
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   720
      Y2              =   5520
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUEVO USUARIO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   30
      Top             =   120
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER"
      Height          =   195
      Left            =   1320
      TabIndex        =   23
      Top             =   5040
      Width           =   675
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LADING"
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDORES"
      Height          =   195
      Left            =   840
      TabIndex        =   17
      Top             =   4080
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTES"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTARIO"
      Height          =   195
      Left            =   1080
      TabIndex        =   11
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURA"
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ver"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE DE USUARIO"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1755
   End
End
Attribute VB_Name = "nivelacceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
If Check1 = Text Then
Text2.PasswordChar = "*"
Else
Text2.PasswordChar = ""
End If
End Sub

Private Sub Check2_Click()
If Check2 = Text Then
Text3.Text = "no"
Else
Text3.Text = "si"
End If
End Sub

Private Sub Check3_Click()
If Check3 = Text Then
Text4.Text = "no"
Else
Text4.Text = "si"
End If
End Sub

Private Sub Check4_Click()
If Check4 = Text Then
Text5.Text = "no"
Else
Text5.Text = "si"
End If
End Sub

Private Sub Check5_Click()
If Check5 = Text Then
Text6.Text = "no"
Else
Text6.Text = "si"
End If
End Sub

Private Sub Check6_Click()
If Check6 = Text Then
Text7.Text = "no"
Else
Text7.Text = "si"
End If
End Sub

Private Sub Check7_Click()
If Check7 = Text Then
Text8.Text = "no"
Else
Text8.Text = "si"
End If
End Sub

Private Sub Command1_Click()
On Error GoTo salida
Adodc1.Recordset.Update
Adodc1.Recordset.MovePrevious
MsgBox "LOS DATOS GUARDARON EXITOSAMENTE", vbInformation, "SISTEMA DE USUARIOS"
Adodc1.Recordset.MoveNext

salida:
If Label10.Visible = True Then
If Text1 = "" Then
MsgBox ("LOS CAMPOS NOMBRE Y CONTRASEÑA DEBE ESTAR DILIGENCIADO"), vbCritical, ("ERROR AL GUARDAR")
Else
If Text2 = "" Then
Text2 = "1234"
MsgBox "EL CAMPO NO FUE DILIGENCIADO CONTRASEÑA TEMPORAL 1234", vbCritical, "CONTASEÑA TEMPORAL"
End If
List2.AddItem Text1
m = List2.ListCount
Text9 = m + 1
Adodc1.Recordset.Update
Adodc1.Recordset.MovePrevious
MsgBox "LOS DATOS GUARDARON EXITOSAMENTE", vbInformation, "SISTEMA DE USUARIOS"
Adodc1.Recordset.MoveNext
Label10.Visible = False
Text1.Enabled = False
Text2.Enabled = False

End If
End If
End Sub

Private Sub Command2_Click()
List1.Visible = True
List1.Clear
Adodc1.Refresh
Dim i As Integer
Dim x As String
x = List2.ListCount
For i = 1 To x
List1.AddItem Text1
On Error GoTo salida
Adodc1.Recordset.MoveNext

Next i
salida:
    Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text3 = "no"
Text4 = "no"
Text5 = "no"
Text6 = "no"
Text7 = "no"
Text8 = "no"
Check1 = "1"
Label10.Visible = True


End Sub

Private Sub Command5_Click()
On Error GoTo salida
Adodc1.Recordset.Delete
MsgBox "se elimino el clinete correctamente", vbInformation, " SISTEMA DE REGISTRO"
Adodc1.Refresh
Exit Sub
salida:
MsgBox "los campos estan vacios busuqe los datos a eliminar ", vbInformation, " SISTEMA DE REGISTRO"

End Sub

Private Sub Form_Load()
List2.Clear
Adodc1.Recordset.MoveFirst


   While Not Adodc1.Recordset.EOF
    
    List2.AddItem (Text1)

   Adodc1.Recordset.MoveNext
   If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
Exit Sub
     End If
    Wend
End Sub

Private Sub List1_Click()
Adodc1.Refresh

Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
End If
Dim busqueda As String
busqueda = List1
Adodc1.Recordset.Find "nombre='" & Trim(busqueda) & "'"
If Adodc1.Recordset.EOF Then


Exit Sub
End If
Text1.Text = Adodc1.Recordset.Fields(1).Value
Text2.Text = Adodc1.Recordset.Fields(2).Value

List1.Visible = False
Exit Sub

End Sub

Private Sub Text3_Change()
If Text3 = "si" Then
Check2 = "1"
Else
Check2 = "0"

End If
End Sub

Private Sub Text4_Change()
If Text4 = "si" Then
Check3 = "1"
Else
Check3 = "0"

End If
End Sub

Private Sub Text5_Change()
If Text5 = "si" Then
Check4 = "1"
Else
Check4 = "0"

End If
End Sub

Private Sub Text6_Change()
If Text6 = "si" Then
Check5 = "1"
Else
Check5 = "0"

End If
End Sub

Private Sub Text7_Change()
If Text7 = "si" Then
Check6 = "1"
Else
Check6 = "0"

End If
End Sub

Private Sub Text8_Change()
If Text8 = "si" Then
Check7 = "1"
Else
Check7 = "0"

End If
End Sub
