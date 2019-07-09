VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Compras"
   ClientHeight    =   6270
   ClientLeft      =   555
   ClientTop       =   795
   ClientWidth     =   9240
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9240
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   960
      Picture         =   "Form5.frx":0000
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   4680
      Picture         =   "Form5.frx":2F0C
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   5880
      Picture         =   "Form5.frx":617F
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   4680
      Picture         =   "Form5.frx":92CE
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Line Line8 
      X1              =   8640
      X2              =   8640
      Y1              =   2640
      Y2              =   4920
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   8640
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   8640
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   120
      Y1              =   2640
      Y2              =   4920
   End
   Begin VB.Line Line4 
      X1              =   8640
      X2              =   8640
      Y1              =   840
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8640
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   2400
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FACTURA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPRAS"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
If Val(Text3) < 1 Then
    MsgBox "No se puede buscar sino carga primero el codigo", vbOKOnly, "ERROR!"
    Text3.SetFocus
Else
    X = "select * from articulos where codigo=" & Text3 & ""
    conexion_tablas.Open X, conexion_basedatos
    If conexion_tablas.EOF = True Then
        If MsgBox("Codigo inexistente, desea crearlo?", vbInformation + vbYesNo, "DIGISPACE") = vbYes Then
            Form1.Show
            conexion_tablas.Close
        Else
            conexion_tablas.Close
            Text3 = ""
            Text3.SetFocus
        End If
    Else
        Label9 = conexion_tablas!detalle
        conexion_tablas.Close
    End If
End If

End Sub

Private Sub image3_Click()
A = "insert into compra values (" & Val(Text1) & ",'" & Text2 & "'," & Val(Text3) & "," & Val(Text5) & "," & Val(Label10) & ",'" & Label11 & "')"
conexion_basedatos.Execute A
r = "select * from articulos where codigo= " & Val(Text3) & ""
conexion_tablas.Open r, conexion_basedatos
cant = conexion_tablas!cantidad
suma = cant + Val(Text5)
X = "update articulos set cantidad=" & suma & ", precio = " & Text4 & " where codigo= " & Val(Text3) & ""
conexion_basedatos.Execute X
MsgBox "Los datos han sido guardados", vbOKOnly, "DIGISPACE"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Label9 = ""
Label10 = ""
conexion_tablas.Close

End Sub

Private Sub Image4_Click()
Unload Me
End Sub

Private Sub image2_Click()
Label10 = Val(Text4) * Val(Text5)
End Sub

Private Sub Form_Load()
Label11 = Format(Date, "short date")

End Sub

