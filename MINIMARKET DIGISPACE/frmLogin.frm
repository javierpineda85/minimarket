VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de usuarios"
   ClientHeight    =   1770
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   1045.775
   ScaleMode       =   0  'User
   ScaleWidth      =   4070.33
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Me.Hide
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'comprobar si la contraseña es correcta
    If txtPassword = "admin" Or txtPassword = "minimarket" Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        LoginSucceeded = True
        Me.Hide
        frmSplash.Show
        'MDIForm1.Show
        
    Else
        MsgBox "La contraseña no es válida o está activado la tecla BLOQ MAYUS. Vuelva a intentarlo o desactive la tecla de las mayúsculas", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub



