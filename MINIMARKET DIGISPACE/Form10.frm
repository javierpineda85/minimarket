VERSION 5.00
Object = "{4561DB7F-ADD4-11D4-A550-400005860166}#3.0#0"; "Bouton3D.ocx"
Begin VB.Form Form10 
   Caption         =   "Cierre de caja"
   ClientHeight    =   5760
   ClientLeft      =   735
   ClientTop       =   960
   ClientWidth     =   7635
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   7635
   Begin bouton3D.Command Command3 
      Height          =   375
      Left            =   2640
      TabIndex        =   33
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ColorBack       =   16777088
      ColorLight      =   8421376
      ColorShade      =   8421376
      ColorText       =   0
      BorderSize      =   5
      Caption         =   "Guardar"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin bouton3D.Command Command2 
      Height          =   375
      Left            =   4680
      TabIndex        =   31
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ColorBack       =   16777152
      ColorLight      =   8421376
      ColorShade      =   8421376
      ColorText       =   0
      BorderSize      =   5
      Caption         =   "Calcular"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1320
      TabIndex        =   29
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1920
      TabIndex        =   27
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   6600
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Para números con decimales colocar PUNTO en lugar de coma"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin bouton3D.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   32
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ColorBack       =   16777152
      ColorLight      =   8421376
      ColorShade      =   8421376
      ColorText       =   0
      BorderSize      =   5
      Caption         =   "Guardar"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierre de caja"
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
      Left            =   360
      TabIndex        =   30
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto: $"
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
      Left            =   360
      TabIndex        =   28
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social:"
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
      Left            =   360
      TabIndex        =   26
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagos diarios de Proveedores"
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
      Left            =   480
      TabIndex        =   25
      Top             =   3360
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      X1              =   7320
      X2              =   240
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Salidas de caja: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Monedas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "$2 x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "$5 x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "$10 x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "$ 20 x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "$ 50 x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "$ 100 x "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total ingresado en caja: $"
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
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
c = MsgBox("Una vez realizada esta operacion no puede eliminarla. Desea continuar?", vbOKCancel, "ATENCION!!!")
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    a = MsgBox("Dede completar todos los campos", vbOKOnly, "Faltan datos")
    
Else
If c = vbOK Then
    Label10 = 100 * Val(Text1)
    Label11 = 50 * Val(Text2)
    Label12 = 20 * Val(Text3)
    Label13 = 10 * Val(Text4)
    Label14 = 5 * Val(Text5)
    Label15 = 2 * Val(Text6)
    Label17 = Val(Label10) + Val(Label11) + Val(Label12) + Val(Label13) + Val(Label14) + Val(Label15) + Val(Text7) + Val(Text8)
    g = "insert into cierre values ( '" & Label2 & "','" & Label17 & "')"
    conexion_basedatos.Execute g
    MsgBox " Los datos han sido guardados"
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Label10 = ""
    Label11 = ""
    Label12 = ""
    Label13 = ""
    Label14 = ""
    Label15 = ""
    Text9.SetFocus
End If
End If
End Sub

Private Sub Command2_Click()
Label10 = 100 * Val(Text1)
Label11 = 50 * Val(Text2)
Label12 = 20 * Val(Text3)
Label13 = 10 * Val(Text4)
Label14 = 5 * Val(Text5)
Label15 = 2 * Val(Text6)
Label17 = Val(Label10) + Val(Label11) + Val(Label12) + Val(Label13) + Val(Label14) + Val(Label15) + Val(Text7) + Val(Text8)

End Sub


Private Sub Command3_Click()
c = MsgBox("Una vez realizada esta operacion no puede eliminarla. Desea continuar?", vbOKCancel, "ATENCION!!!")
If c = vbOK Then
    g = "insert into pagos values ( '" & Label2 & "','" & Text9 & "','" & Text10 & "')"
    conexion_basedatos.Execute g
    MsgBox " Los datos han sido guardados"
    Text9 = ""
    Text10 = ""
    Text9.SetFocus
End If
End Sub

Private Sub Form_Load()
Label2 = Date

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label10 = 100 * Val(Text1)
    Text2.SetFocus
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label11 = 50 * Val(Text2)
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label12 = 20 * Val(Text3)
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label13 = 10 * Val(Text4)
    Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label14 = 5 * Val(Text5)
    Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label15 = 2 * Val(Text6)
    Text7.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text8.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(Keyascci As Integer)
If KeyAscii = 13 Then
    Label17 = Val(Label10) + Val(Label11) + Val(Label12) + Val(Label13) + Val(Label14) + Val(Label15) + Val(Text7) + Val(Text8)
 End If
 
End Sub
