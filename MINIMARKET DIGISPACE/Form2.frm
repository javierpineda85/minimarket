VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Listado de Reposicion"
   ClientHeight    =   7245
   ClientLeft      =   900
   ClientTop       =   960
   ClientWidth     =   9975
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   9975
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6000
      Picture         =   "Form2.frx":2C6D1
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   8760
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   7440
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   6120
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   8160
      Picture         =   "Form2.frx":33C43
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   8160
      Picture         =   "Form2.frx":36D91
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE REPOSICION"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MINIMO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub image2_Click()
List1.Clear: List2.Clear: List3.Clear: List4.Clear: List5.Clear
datos = "select * from articulos order by codigo"
conexion_tablas.Open datos, conexion_basedatos
Do While Not conexion_tablas.EOF
If conexion_tablas!cantidad < conexion_tablas!minimo Then
    List1.AddItem conexion_tablas!codigo
    List2.AddItem conexion_tablas!detalle
    List3.AddItem conexion_tablas!precio
    List4.AddItem conexion_tablas!cantidad
    List5.AddItem conexion_tablas!minimo
End If
conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
If MsgBox("Esta la impresora OK?", vbInformation + vbYesNo, "listado") = vbYes Then
    datos = "select * from articulos"
    conexion_tablas.Open datos, conexion_basedatos
    Printer.CurrentX = 1
    Printer.CurrentY = 5
      
    Do While Not conexion_tablas.EOF
        If conexion_tablas!cantidad < conexion_tablas!minimo Then
            Printer.FontSize = 10
            Printer.Print
            Printer.Print conexion_tablas!codigo & "" & conexion_tablas!detalle & "" & conexion_tablas!precio&; "" & conexion_tablas!cantidad & "" & conexion_tablas!minimo
        End If
    conexion_tablas.MoveNext
    Loop
    conexion_tablas.Close
    Printer.EndDoc
End If

End Sub
