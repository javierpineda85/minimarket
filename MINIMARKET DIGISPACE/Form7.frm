VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form7 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Listado de Clientes"
   ClientHeight    =   5835
   ClientLeft      =   270
   ClientTop       =   960
   ClientWidth     =   7320
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   7320
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7223
      _Version        =   393216
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   5280
      Picture         =   "Form7.frx":3A997
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE PROVEEDORES"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call cabecera
datos = "select * from clientes order by cliente"
conexion_tablas.Open datos, conexion_basedatos
lineas = 0
Do While Not conexion_tablas.EOF
        lineas = lineas + 1
        MSFlexGrid1.TextMatrix(lineas, 0) = conexion_tablas!cliente
        MSFlexGrid1.TextMatrix(lineas, 1) = conexion_tablas!domicilio
        MSFlexGrid1.TextMatrix(lineas, 2) = conexion_tablas!telefono
        conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End Sub
Private Sub cabecera()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 3
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 100
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "CLIENTE"
MSFlexGrid1.TextMatrix(0, 1) = "PROVEEDOR DE"
MSFlexGrid1.TextMatrix(0, 2) = "TELEFONO"

MSFlexGrid1.ColWidth(0) = 2500
MSFlexGrid1.ColWidth(1) = 2500
MSFlexGrid1.ColWidth(2) = 1200
End Sub


Private Sub MSFlexGrid1_Click()

End Sub
