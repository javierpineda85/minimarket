VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H0080FF80&
   Caption         =   "Datos de Proveedores"
   ClientHeight    =   4230
   ClientLeft      =   4065
   ClientTop       =   2325
   ClientWidth     =   7770
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   4230
   ScaleWidth      =   7770
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   5880
      Picture         =   "Form6.frx":2B17A
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   360
      Picture         =   "Form6.frx":2E34B
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   5880
      Picture         =   "Form6.frx":31280
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   5880
      Picture         =   "Form6.frx":342A7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   3960
      Picture         =   "Form6.frx":371B3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   2160
      Picture         =   "Form6.frx":3A3CD
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   360
      Picture         =   "Form6.frx":3DA1C
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DE PROVEEDORES"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR DE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Visible = True

End Sub

Private Sub image5_Click()
If Text1 = "" Then
    MsgBox "No se puede buscar sino carga primero el nombre", vbOKOnly, "DIGISPACE"
    Text1.SetFocus
Else
    b = "select*from clientes where cliente like '" & Text1 & "'"
    conexion_tablas.Open b, conexion_basedatos
    If conexion_tablas.EOF = True Then
        MsgBox "No se encontro el cliente la base de datos! Por favor carguelo!", vbOKOnly, "ATENCION!"
        Text2.SetFocus
        conexion_tablas.Close
    Else
        Text2 = conexion_tablas!domicilio
        Text6 = conexion_tablas!telefono
        
        conexion_tablas.Close
    End If
   
End If
Image1.Visible = False
End Sub

Private Sub image3_Click()
res = MsgBox(" Desea eliminar el registro de manera permanente?", vbYesNo, "CUIDADO!!!")
If res = vbYes Then
    X = "delete *from clientes where cliente = '" & (Text1) & "'"
    conexion_basedatos.Execute X
    MsgBox "El registro ha sido eliminado!", vbOKOnly, "DIGISPACE"
    Text1 = ""
    Text6 = ""
End If
conexion_basedatos.Close
End Sub

Private Sub image7_Click()
X = "update clientes set domicilio='" & Text2 & "', telefono ='" & Val(Text6) & "'"
conexion_basedatos.Execute X
MsgBox "Los datos han sido modificados correctamente!", vbOKOnly, "DIGISPACE"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Image1.Visible = True
conexion_basedatos.Close
End Sub

Private Sub Image1_Click()
d = "insert into clientes values('" & Text1 & "','" & Text2 & "'," & Val(Text6) & ")"
conexion_basedatos.Execute d
Text1 = ""
Text2 = ""
Text6 = ""
Text1.SetFocus
r = MsgBox("Los datos se guardaron correctamente! Desea cargar otro proveedor?", vbYesNo, "DIGISPACE")
If r = vbNo Then
    Unload Me
End If

End Sub

Private Sub Image4_Click()
Unload Me
End Sub

Private Sub image6_Click()
c = MsgBox("Si ha realizado algún cambio, debera presionar cancelar y salir del formulario para actualizar la base de datos", vbOKCancel, "IMPORTANTE!")
If c = vbOK Then
    Form7.Show
End If
Image1.Visible = True
End Sub

Private Sub image2_Click()
Text1 = ""
Text2 = ""
Text6 = ""
Image1.Visible = True
End Sub

