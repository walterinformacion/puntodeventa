VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form productos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "THE EXPERTS/inventario"
   ClientHeight    =   9000
   ClientLeft      =   1035
   ClientTop       =   1545
   ClientWidth     =   17565
   Icon            =   "proo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   17565
   Begin VB.TextBox Text81 
      DataField       =   "cantidades"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   14280
      TabIndex        =   35
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "inventario en 0"
      Height          =   495
      Left            =   13800
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   195
      Left            =   2640
      TabIndex        =   33
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   18240
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   720
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text4A 
         Height          =   495
         Left            =   0
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text5A 
         Height          =   495
         Left            =   0
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text6A 
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text7A 
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.ListBox List2 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      ItemData        =   "proo.frx":E95A
      Left            =   240
      List            =   "proo.frx":E95C
      TabIndex        =   26
      Top             =   5160
      Width           =   12735
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   15480
      TabIndex        =   25
      Top             =   1110
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   13320
      TabIndex        =   24
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "buscar"
      Height          =   495
      Left            =   13560
      TabIndex        =   23
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      DataField       =   "codifo1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox Text6 
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
      Height          =   450
      Left            =   6600
      TabIndex        =   21
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "codifo"
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
      Height          =   435
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      DataField       =   "desciocion"
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
      Height          =   465
      Left            =   4560
      TabIndex        =   10
      Top             =   1560
      Width           =   8175
   End
   Begin VB.TextBox Text3 
      DataField       =   "precio"
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
      Height          =   465
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "modificar"
      Height          =   495
      Left            =   13680
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "guardar"
      Height          =   495
      Left            =   13680
      TabIndex        =   7
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cancelar"
      Height          =   495
      Left            =   13680
      TabIndex        =   6
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "eliminar"
      Height          =   615
      Left            =   13680
      TabIndex        =   5
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox TextBUS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   13200
      TabIndex        =   4
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "crear un articulo"
      Height          =   495
      Left            =   13560
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "iva2"
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
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "costo"
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
      Height          =   450
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   16320
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "proo.frx":E95E
      Height          =   2415
      Left            =   4440
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
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
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   3360
      Picture         =   "proo.frx":E973
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   6345
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   19095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1200
      TabIndex        =   17
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6120
      TabIndex        =   16
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "precio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   15
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "costo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      TabIndex        =   14
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "iva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      DataField       =   "imagen"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Image ImagenSalidas 
      Height          =   2175
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "utilidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7200
      TabIndex        =   2
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      DrawMode        =   6  'Mask Pen Not
      Height          =   3735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   12855
   End
End
Attribute VB_Name = "productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filaSI As String
Dim contcaracter As Integer
Dim extraer As String

Dim i As Integer
Dim x1 As String



Private Sub Command16_Click()
A = InputBox(("Cantidad A Cobrar $") + (Label4) + (",                           Ingrese Pago Recibido:"), "Mensaje De Caja")
If Val(A) > (0) And Val(A) >= Val(Label4) Then
B = MsgBox("Total: $ " & Val(Label4) & "                    Efectivo: $ " & Val(A) & "                    Cambio: $ " & Val(A) - Val(Label4) & "                                   ¿Desea Imprimir El Ticket?", vbYesNo, "Mensaje De Caja")
If Val(B) = vbYes Then
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.MoveNext
Printer.Print "              ********Gracias Pro Su Preferencia********          "
Printer.Print "                                          "
Printer.Print "                                          "
Printer.Print "=========================================="
 While Not Adodc2.Recordset.EOF
 Printer.Print (Text5) + (" ") + (Text6)
Printer.Print (" $ ") + (Text8) + ("     X     ") + (Text7) + (" Pz.     =     $ ") + (Text9)
   Adodc2.Recordset.MoveNext
   If Adodc2.Recordset.EOF Then
Printer.Print "=========================================="
Printer.Print ("Total: $ ") + (Label4)
Printer.Print ("Efectivo: $ ") + (A)
Printer.Print ("Cambio: $ "); A - Label4
Printer.Print "=========================================="
Printer.Print "                                                                                                "
Printer.Print "                          * * Vuelva Pronto * *                                   "
Printer.EndDoc
Adodc2.Recordset.MoveFirst
List1.Clear
Label4 = ""
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Command14.Visible = False
Command15.Visible = False
Command17.Visible = False
Command16.Visible = False
   Adodc2.Recordset.MoveFirst
   Adodc2.Recordset.MoveNext
  If Text5 > (0) Then
    While Not Adodc2.Recordset.EOF
    Adodc2.Recordset.Delete
    Adodc2.Recordset.MoveNext
    If Adodc2.Recordset.EOF Then
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Update
    Exit Sub
    End If
    Wend
    End If
Exit Sub
End If
Wend
Else
List1.Clear
Label4 = ""
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Command14.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
   Adodc2.Recordset.MoveFirst
   Adodc2.Recordset.MoveNext
  If Text5 > (0) Then
    While Not Adodc2.Recordset.EOF
    Adodc2.Recordset.Delete
    Adodc2.Recordset.MoveNext
    If Adodc2.Recordset.EOF Then
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Update
    Exit Sub
    End If
    Wend
    End If
End If

End If
End Sub

Private Sub Command12_Click()
Frame2.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command13_Click()
End
End Sub



Private Sub Command7_Click()
    calcular

End Sub

Private Sub Command8_Click()

Adodc1.Refresh
h = Adodc1.Recordset.RecordCount
For i = 1 To h
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
Next i
Adodc1.Recordset.MoveLast
Adodc1.Recordset.MovePrevious

End Sub


Private Sub Command14_Click()
If List2 <> "" Then
Adodc1.Recordset.MoveFirst
While Not (Adodc1.Recordset.EOF = True)
  If List2 = Adodc1.Recordset(1) Then
  If Val(Text5) > (0) Then


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
List1.Visible = True
List2.Visible = False
    B = InputBox((Text1) + ("    ") + (Text2) + ("    Precio: $") + (Text3) + ("           ¿Cantidad A Vender?"), "Ingresar Cantidad")
    If Val(B) > (0) Then
    Text4 = Val(B) * Val(Text3)
     c = Val(Text4) + Val(Label4)
     Label4 = Val(c)
     List1.AddItem (Text1) + ("    ") + (Text2)
    List1.AddItem ("Precio: $") + (Text3) + (" Cantidad: ") + (B) + (" Subtotal: $") + (Text4)
     Adodc2.Recordset.AddNew
     Text5 = Text1
     Text6 = Text2
     Text7 = Val(B)
     Text8 = Text3
     Text9 = Text4
Adodc2.Recordset.Update
  If Val(Text5) > (0) Then
Command10.Visible = True
Command11.Visible = True
Command12.Visible = False
Command13.Visible = False
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
Command12.Visible = False
Command13.Visible = False
Command14.Visible = False
Command15.Visible = False
Command16.Visible = True
Command17.Visible = True
List1.Visible = True
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
List1.Visible = True
List2.Visible = False
End If
End Sub

Private Sub Command1_Click()
Adodc1.Refresh
If Option1 = True Then

On Error GoTo salida
List2.Clear
For i = 1 To filaSI - 1
w = UCase(Text2)
h = UCase(TextBUS)
x1 = w Like "*" & h & "*"

If x1 = "Verdadero" Then
  List2.AddItem w
  
  End If
Adodc1.Recordset.MoveNext
Next i
salida:
Adodc1.Refresh
Else
Adodc1.Recordset.Find "codifo='" & Trim(TextBUS) & "'"
If Adodc1.Recordset.EOF Then


Exit Sub
End If

End If


End Sub

Private Sub Command2_Click()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False


Rem Command8.Enabled = False
End Sub

Private Sub Command3_Click()
If Text2 > "" And Text3 > "" Then
Adodc1.Recordset.Update
Adodc1.Recordset.MoveFirst
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False



Text2.Enabled = False
Text3.Enabled = False
Else
c = MsgBox("Información Incompleta, Inténtelo Nuevamente ", vbCritical, "Error De Registro")
End If
End Sub

Private Sub Command4_Click()
A = MsgBox("¿Desea Cancelar Los Cambios Efectuados?", vbYesNo, "Cancelando")
If Val(A) = vbYes Then
Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.MoveFirst
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True



Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
End If
End Sub

Private Sub Command5_Click()
A = MsgBox("Esta Seguro De Eliminar El Registro", vbOKCancel, "Eliminar")
If Val(A) = vbOK Then
Adodc1.Recordset.Delete
End If
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
 

Adodc1.Recordset.AddNew
Text1 = InputBox("Ingrese El codigo Del Producto", "Captura")
Text2 = InputBox("Ingrese La Descripción Del Producto", "Captura")
Text3 = InputBox("Ingrese El Precio Del Producto", "Captura")
Text4 = InputBox("Ingrese El costo Del Producto", "Captura")
Text5 = InputBox("Ingrese El iva Del Producto de lo contrario poga (0)", "Captura")


End Sub

Private Sub Command10_Click()
A = InputBox("Ingrese El Código Del Producto", "Ingresar Producto")
If Val(A) > (0) Then
Adodc1.Recordset.MoveFirst
While Not (Adodc1.Recordset.EOF = True)
  If Val(A) = Adodc1.Recordset(0) Then
    If Val(A) > (0) Then
    B = InputBox((Text1) + ("    ") + (Text2) + ("    Precio: $") + (Text3) + ("           ¿Cantidad A  Vender?"), "Ingresar Cantidad")
    If Val(B) > (0) Then
    Text4 = Val(B) * Val(Text3)
     c = Val(Text4) + Val(Label4)
     Label4 = Val(c)
     List1.AddItem (Text1) + ("    ") + (Text2)
    List1.AddItem ("Precio: $") + (Text3)
     List1.AddItem (" Cantidad: ") + (B)
      List1.AddItem (Text4)
     Adodc2.Recordset.AddNew
     Text5 = Text1
     Text6 = Text2
     Text7 = Val(B)
     Text8 = Text3
     Text9 = Text4
Command10.Visible = True
Command11.Visible = True
Command12.Visible = False
Command13.Visible = False
Command14.Visible = False
Command15.Visible = False
Command16.Visible = True
Command17.Visible = True
List1.Visible = True
List2.Visible = False
     Adodc2.Recordset.Update
     End If
      Exit Sub
       End If
    End If
      Adodc1.Recordset.MoveNext

Wend
 D = MsgBox("Codigo De Producto Incorrecto", vbCritical, "Error De Captura")
Adodc1.Recordset.MoveFirst
End If
End Sub


Private Sub Command11_Click()
List2.Clear
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.MoveNext
   While Not Adodc1.Recordset.EOF
   List2.AddItem (Text2)
   Adodc1.Recordset.MoveNext
   If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
Command10.Visible = False
Command11.Visible = False
Command12.Visible = False
Command13.Visible = False
Command14.Visible = True
Command15.Visible = True
Command16.Visible = False
Command17.Visible = False
List1.Visible = False
List2.Visible = True
Exit Sub
     End If
    Wend
End Sub

Private Sub Command17_Click()
A = MsgBox("¿Esta Seguro De Cancelar Esta Venta?", vbYesNo, "Cancelando")
If Val(A) = vbYes Then
List1.Clear
Label4 = ""
Command10.Visible = True
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Command14.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
If Text5 > (0) Then
  Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.MoveNext
    While Not Adodc2.Recordset.EOF
    Adodc2.Recordset.Delete
    Adodc2.Recordset.MoveNext
    If Adodc2.Recordset.EOF Then
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Update
    Exit Sub
    End If
    Wend
    End If
End If
End Sub


Private Sub Image2_Click()

End Sub

Private Sub CommandBU_Click()

End Sub

Private Sub DataGrid1_Click()
On Error GoTo salida
ima = App.Path
Imagensalidas.Picture = LoadPicture(ima & "\imagenes\" & Label4.Caption & ".jpg")
salida:
End Sub



Private Sub Form_Activate()

Option1 = True
    TextBUS.SetFocus

Rem Imagensalidas.Picture = LoadPicture("C:\para motos\imagenes\" & Label4.Caption & ".jpg")

 List1.Clear
Rem Adodc1.Recordset.MoveFirst


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

Sub calcular()

Dim v1 As Double
Dim v2 As Double
Dim vd As Double
Dim vp As Integer
On Error GoTo salida:
v1 = Text3
v2 = Text4


vp = v2 / v1 * 100


Text6 = vp & "%"
salida:
End Sub

Private Sub Form_Load()
Command7_Click
End Sub

Private Sub List2_Click()
order1 = xlAscending
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
End If
Dim busqueda As String
Adodc1.Refresh
busqueda = List2
Adodc1.Recordset.Find "desciocion='" & Trim(busqueda) & "'"
If Adodc1.Recordset.EOF Then


Exit Sub
End If



Exit Sub
End Sub

Private Sub Text4_Change()
If Text4 = "0" Then
Else
Command7_Click
End If

End Sub

Private Sub TextBUS_Change()

Rem ******************************** contador de caracter**********************
contcaracter = Len(Text3.Text)
Text4A = contcaracter
Rem ****************************** extraer cararter**************
extraer = Left(Text2, contcaracter)
Text5A = UCase(extraer)
Text7A = UCase(extraer)

Rem **********************************cantidad fe fila *********************
filaSI = List1.ListCount
Text6A = filaSI

End Sub

Private Sub Text1_Change()
If Val(Text1) = (0) Then
Command2.Enabled = False
Command5.Enabled = False
Else
Command2.Enabled = True
Command5.Enabled = True
End If
End Sub
Private Sub ColocarImagen(Label4 As Long)
        Dim Variable As String
        
        Variable = Trim(ThisWorkbook.Path & "\Images\" & Label4 & ".jpg")
        If Dir(Variable) <> "" Then
            Set Imagensalidas.Picture = LoadPicture(Variable)
        Else
            Set Imagensalidas.Picture = Nothing
        End If
        FrameImagen.Repaint

End Sub

Private Sub TextBUS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
