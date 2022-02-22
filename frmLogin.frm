VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   7605
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4493.285
   ScaleMode       =   0  'User
   ScaleWidth      =   13196.88
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   840
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   720
      Picture         =   "frmLogin.frx":3406
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   840
      Picture         =   "frmLogin.frx":69B4
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   5535
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   720
      Picture         =   "frmLogin.frx":9D20
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   8040
      Picture         =   "frmLogin.frx":D217
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   8160
      Picture         =   "frmLogin.frx":10967
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5535
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Image1_Click()
controlventas.Show

End Sub

Private Sub Image2_Click()
productos.Show

End Sub

Private Sub Image3_Click()
clientes.Show

End Sub

Private Sub Image4_Click()
proveedores1.Show
End Sub

Private Sub ventas_Click()

ventas1.Show

End Sub

