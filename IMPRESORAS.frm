VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form IMPRESORAS 
   Caption         =   "IMPRESORAS"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   Icon            =   "IMPRESORAS.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3930
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6720
      Top             =   3480
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
   Begin VB.CommandButton Command2 
      Caption         =   "guardar los cambios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   14
      Top             =   3240
      Width           =   3735
   End
   Begin VB.OptionButton Option4 
      Height          =   375
      Left            =   8880
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Text4 
      DataField       =   "impresora4"
      DataSource      =   "Adodc1"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   2640
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      DataField       =   "impresora3"
      DataSource      =   "Adodc1"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      DataField       =   "impresora2"
      DataSource      =   "Adodc1"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   5295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   9720
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Impresora Predeterminada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      DataField       =   "impresora1"
      DataSource      =   "Adodc1"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora  Reportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora  Etiquetar "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora  Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora  Tickets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "IMPRESORAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim r As Long
Dim Buffer As String
Dim DeviceName As String
Dim DriverName As String
Dim PrinterPort As String
Dim PrinterName As String
If List1.ListIndex > -1 Then
Buffer = Space(1024)
PrinterName = List1.Text
r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))
GetDriverAndPort Buffer, DriverName, PrinterPort
If DriverName <> "" And PrinterPort <> "" Then
SetDefaultPrinter List1.Text, DriverName, PrinterPort
End If
End If
End Sub



Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
MsgBox "LOS DATOS SE GUARDARON ", vbInformation, "GUARDAR"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Dim r As Long
Dim Buffer As String
Buffer = Space(8192)
r = GetProfileString("PrinterPorts", vbNullString, "", Buffer, Len(Buffer))
ParseList List1, Buffer
End Sub
Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
Dim DeviceLine As String
Dim r As Long
Dim l As Long
DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
r = WriteProfileString("windows", "Device", DeviceLine)
l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub
Sub ParseList(lstCtl As Control, ByVal Buffer As String)
Dim i As Integer
Do
i = InStr(Buffer, Chr(0))
If i > 0 Then
lstCtl.AddItem Left(Buffer, i - 1)
Buffer = Mid(Buffer, i + 1)
Else
lstCtl.AddItem Buffer
Buffer = ""
End If
Loop While i > 0
End Sub
Sub GetDriverAndPort(ByVal Buffer As String, DriverName As String, PrinterPort As String)
Dim r As Integer
Dim iDriver As Integer
Dim iPort As Integer
DriverName = ""
PrinterPort = ""
iDriver = InStr(Buffer, ",")
If iDriver > 0 Then
DriverName = Left(Buffer, iDriver - 1)
iPort = InStr(iDriver + 1, Buffer, ",")
If iPort > 0 Then
PrinterPort = Mid(Buffer, iDriver + 1, iPort - iDriver - 1)
End If
End If
End Sub

Private Sub Label5_Click()

End Sub

Private Sub List1_Click()
  For Contador = 0 To List1.ListCount - 1
        If List1.Selected(Contador) Then
           If Option1 = True Then
              Text1 = List1
           End If
           If Option2 = True Then
              Text2 = List1
           End If
           If Option3 = True Then
              Text3 = List1
           End If
           If Option4 = True Then
              Text4 = List1
           End If
        
    End If
    Next
End Sub

Private Sub Text1_click()
Option1 = True

End Sub

Private Sub Text2_click()
Option2 = True
End Sub

Private Sub Text3_click()
Option3 = True
End Sub

Private Sub Text4_click()
Option4 = True
End Sub
