VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "punto de venta"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "caja rejistradora"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7935
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   12255
      Begin VB.CommandButton Command17 
         Caption         =   "cancelar venta"
         Height          =   255
         Left            =   8280
         TabIndex        =   34
         Top             =   4920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command16 
         Caption         =   "cobrar"
         Height          =   375
         Left            =   8280
         TabIndex        =   33
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command15 
         Caption         =   "cancelar"
         Height          =   375
         Left            =   8280
         TabIndex        =   32
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command14 
         Caption         =   "seleccionar"
         Height          =   375
         Left            =   8280
         TabIndex        =   31
         Top             =   3120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command13 
         Caption         =   "salir"
         Height          =   255
         Left            =   8280
         TabIndex        =   30
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command12 
         Caption         =   "almacen"
         Height          =   375
         Left            =   8280
         TabIndex        =   29
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "buscar por lista"
         Height          =   255
         Left            =   8280
         TabIndex        =   28
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "buscar por codigo"
         Height          =   375
         Left            =   8280
         TabIndex        =   27
         Top             =   840
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   3120
         Top             =   3960
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         Connect         =   $"puntode venta.frx":0000
         OLEDBString     =   $"puntode venta.frx":008C
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tikeck"
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
      Begin VB.TextBox Text10 
         DataField       =   "subtotal"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Text            =   "Text10"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         DataField       =   "cantidad"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         Text            =   "Text9"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         DataField       =   "precio"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Text            =   "Text8"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         DataField       =   "descipcion"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   2640
         TabIndex        =   22
         Text            =   "Text7"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         DataField       =   "codigo"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Text            =   "Text6"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   5580
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   6855
      End
      Begin VB.TextBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   480
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   840
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000008&
         Caption         =   "Label4"
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   8040
         TabIndex        =   26
         Top             =   5520
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "almacen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   12135
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   855
         Left            =   1680
         Top             =   5760
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1508
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
         Connect         =   $"puntode venta.frx":0118
         OLEDBString     =   $"puntode venta.frx":01A4
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
      Begin VB.CommandButton Command9 
         Caption         =   "buscar por codigo"
         Height          =   255
         Left            =   9600
         TabIndex        =   16
         Top             =   5760
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "seleccionar"
         Height          =   375
         Left            =   9960
         TabIndex        =   15
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "buscar lista"
         Height          =   255
         Left            =   9840
         TabIndex        =   14
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "caja rejistradora"
         Height          =   255
         Left            =   9960
         TabIndex        =   13
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "eliminar"
         Height          =   255
         Left            =   9840
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "cancelar"
         Height          =   375
         Left            =   9960
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "guardar"
         Height          =   495
         Left            =   9960
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "modificar"
         Height          =   375
         Left            =   9840
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "buscar"
         Height          =   495
         Left            =   9720
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   480
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         DataField       =   "precio"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   4560
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         DataField       =   "desciocion"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         DataField       =   "codifo"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2520
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "descipcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   2
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   2040
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
A = InputBox("Ingrese El Código Del Producto", "Buscar En La Base De Datos")
If Val(A) > (0) Then
Adodc1.Recordset.MoveFirst
While Not (Adodc1.Recordset.EOF = True)
  If Val(A) = Adodc1.Recordset(0) Then
  Exit Sub
  End If
      Adodc1.Recordset.MoveNext
Wend
  B = MsgBox("El Producto Que Busca No Esta En La Base De Datos, ¿Desea Agregarlo?", vbYesNo, "Error De Captura")
If Val(B) = vbYes Then
Adodc1.Recordset.AddNew
Text1 = Val(A)
Text2 = InputBox("Ingrese La Descripción Del Producto", "Captura")
Text3 = InputBox("Ingrese El Precio Del Producto", "Captura")
If Text2 > "" And Text3 > "" Then
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Else
C = MsgBox("Información Incompleta, Inténtelo Nuevamente", vbCritical, "Error De Registro")
Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Adodc1.Recordset.MoveFirst
End If
End If

End If

End Sub

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
A = MsgBox("¿Esta Seguro Que Desea Salir?", vbYesNo, "Saliendo")
 If Val(A) = vbYes Then
 End
 End If
End Sub

Private Sub Command8_Click()
Combo1.Clear
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.MoveNext
   While Not Adodc1.Recordset.EOF
   Combo1.AddItem (Text2)
   Adodc1.Recordset.MoveNext
   If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Visible = True
Combo1.Visible = True
Text2.Visible = False

 Exit Sub
     End If
    Wend
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
List1.Visible = True
List2.Visible = False
    B = InputBox((Text1) + ("    ") + (Text2) + ("    Precio: $") + (Text3) + ("           ¿Cantidad A Vender?"), "Ingresar Cantidad")
    If Val(B) > (0) Then
    Text4 = Val(B) * Val(Text3)
     C = Val(Text4) + Val(Label4)
     Label4 = Val(C)
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

Private Sub Command2_Click()
Text2.Enabled = True
Text3.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
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
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Else
C = MsgBox("Información Incompleta, Inténtelo Nuevamente ", vbCritical, "Error De Registro")
End If
End Sub

Private Sub Command4_Click()
A = MsgBox("¿Desea Cancelar Los Cambios Efectuados?", vbYesNo, "Cancelando")
If Val(A) = vbYes Then
Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.MoveFirst
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
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
Frame1.Visible = False
Frame2.Visible = True
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
     C = Val(Text4) + Val(Label4)
     Label4 = Val(C)
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


Private Sub Text1_Change()
If Val(Text1) = (0) Then
Command2.Enabled = False
Command5.Enabled = False
Else
Command2.Enabled = True
Command5.Enabled = True
End If
End Sub



