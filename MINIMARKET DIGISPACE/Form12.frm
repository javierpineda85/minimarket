VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Cliente Nuevo"
   ClientHeight    =   1890
   ClientLeft      =   4260
   ClientTop       =   2835
   ClientWidth     =   6270
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   1890
   ScaleWidth      =   6270
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   4440
      Picture         =   "Form12.frx":1D3D0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Tarjeta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre y Apellido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Nuevo"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

a = "select * from monevi_nombre where tarjeta = " & Val(Text2) & ""
    conexion_tablas.Open a, conexion_basedatos
    If conexion_tablas.EOF = True Then
        f = "insert into monevi_nombre values ( " & Val(Text2) & ",'" & Text1 & "')"
        conexion_basedatos.Execute f
        MsgBox " El cliente ha sido guardado", vbOKOnly, "DIGISPACE"
        Text1 = ""
        Text2 = ""
        conexion_tablas.Close
    Else
        MsgBox "El Nº de tarjeta esta asignado a otro cliente", vbOKOnly, "DIGISPACE"
    conexion_tablas.Close
    End If
End Sub
