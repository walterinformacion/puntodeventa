VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form buscador 
   Caption         =   "Buscador De Articulos"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13860
   LinkTopic       =   "Form2"
   ScaleHeight     =   8640
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox imag 
      DataField       =   "imagen"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   11640
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   11760
      TabIndex        =   9
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   12000
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   12120
      TabIndex        =   7
      Top             =   5040
      Width           =   975
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
      Height          =   7980
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   12120
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "codifo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   11880
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "desciocion"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   11880
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11880
      Top             =   1680
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
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   10680
      TabIndex        =   0
      Top             =   8040
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   3735
   End
End
Attribute VB_Name = "buscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fila As String
Dim contcaracter As Integer
Dim extraer As String

Dim i As Integer
Private Sub Command1_Click()
On Error GoTo salida
List2.Clear
For i = 1 To fila
extraer = Left(Text2, contcaracter)
Text7 = UCase(extraer)
If Text7 = UCase(Text3) Then
List2.AddItem UCase(Text2)

End If

Adodc1.Recordset.MoveNext
Next i
salida:
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
ventas1.codigo = Text1
Me.Hide

ventas1.Command10




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

Private Sub List2_Click()
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
On Error GoTo salir
 ima = App.Path
     Image1.Picture = LoadPicture(ima & "\imagenes\" & imag & ".jpg")
salir:
     
Exit Sub



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

Command1_Click


End Sub

Private Sub Text8_Change()

End Sub
