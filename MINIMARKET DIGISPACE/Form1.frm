VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Gestión de Artículos"
   ClientHeight    =   4440
   ClientLeft      =   1230
   ClientTop       =   1470
   ClientWidth     =   7110
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   7110
   Begin VB.PictureBox Picture_1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   4920
      Picture         =   "Form1.frx":1D3D0
      ScaleHeight     =   450
      ScaleWidth      =   1575
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      ToolTipText     =   "Para numeros decimales colocar PUNTO en lugar de coma"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   3495
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   360
      Picture         =   "Form1.frx":205A1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   2640
      Picture         =   "Form1.frx":23814
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   4920
      Picture         =   "Form1.frx":26A2E
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   4920
      Picture         =   "Form1.frx":29A26
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   4920
      Picture         =   "Form1.frx":2CA4D
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de Articulos"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image5.Visible = True
Picture_1.Visible = False
End Sub

Private Sub image5_Click()
X = "insert into articulos values(" & Val(Text1) & ",'" & Text2 & "'," _
& Val(Text4) & "," & Val(Text3) & "," & Val(Text5) & ")"
conexion_basedatos.Execute X
MsgBox "El articulo ha sido guardado", vbOKOnly, "DIGISPACE"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text1.SetFocus


End Sub
Private Sub Image4_Click()
res = MsgBox("Esta seguro de eliminar el registro?", vbYesNo, "CUIDADO!")
If res = vbYes Then
    X = "delete * from articulos where codigo=" & Val(Text1)
    conexion_basedatos.Execute X
    MsgBox "El articulo ha sido eliminado", vbOKOnly, "DIGISPACE"
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text1.SetFocus
    Image5.Visible = True
End If
End Sub

Private Sub Picture_1_click()

X = " update articulos set detalle='" & Text2 & "' ,precio=" & Val(Text3) & ",cantidad= " _
& Val(Text4) & ",minimo=" & Val(Text5) & " where codigo=" & Val(Text1) & ""
conexion_basedatos.Execute X
MsgBox "El artículo ha sido modificado!", vbOKOnly, "DIGISPACE"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text1.SetFocus
Image5.Visible = True
End Sub

Private Sub Image1_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text1.SetFocus
Image5.Visible = True
Picture_1.Visible = False

End Sub

Private Sub image2_Click()
If Val(Text1) < 1 Then
    MsgBox "Error, no se puede buscar sino carga primero el codigo", vbOKOnly, "DIGISPACE"
    Text1.SetFocus
Else
    X = "select * from articulos where codigo=" & Val(Text1) & ""
    conexion_tablas.Open X, conexion_basedatos
    If conexion_tablas.EOF = True Then
        If MsgBox("Codigo inexistente, desea crearlo?", vbInformation + vbYesNo, "DIGISPACE") = vbYes Then
            Text2.SetFocus
            conexion_tablas.Close
        Else
            conexion_tablas.Close
            Text1 = ""
            Text1.SetFocus
        End If
    Else
        Text2 = conexion_tablas!detalle
        Text3 = conexion_tablas!precio
        Text4 = conexion_tablas!cantidad
        Text5 = conexion_tablas!minimo
        conexion_tablas.Close
        Image5.Visible = False
        Picture_1.Visible = True
    End If
End If
        
End Sub

Private Sub image3_Click()
Unload Me

End Sub


