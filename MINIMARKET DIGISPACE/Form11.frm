VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form11 
   Caption         =   "Cuentas corrientes y saldos"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2778
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Detalle de pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   7095
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total: $"
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
         Left            =   4800
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   1920
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   5040
      Picture         =   "Form11.frx":1D3D0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Atendio:"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
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
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese Monto:"
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
      Left            =   4560
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuentas Corrientes y Saldos"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call cabecera
Label4 = Date
'Text5.SetFocus
End Sub

Private Sub cabecera()


MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 4
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 200
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "FECHA"
MSFlexGrid1.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 2) = "MONTO"
MSFlexGrid1.TextMatrix(0, 3) = "ATENDIO"

MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1000
MSFlexGrid1.ColWidth(3) = 1500

End Sub



Private Sub Image1_Click()
j = "insert into monevi values (" & Val(Text5) & ",'" & Text2 & "','" & Label4 & "','" & Text4 & "','" & Text3 & "')"
conexion_basedatos.Execute j
MsgBox "Los datos han sido guardado correctamente", vbOKOnly, "DIGISPACE"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Label8 = ""
MSFlexGrid1.Clear
Call cabecera
Text5.SetFocus
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    e = "select * from monevi_nombre where tarjeta=" & Val(Text5)
    conexion_tablas.Open e, conexion_basedatos
    If conexion_tablas.EOF = False Then
        Text1 = conexion_tablas!nombre
        conexion_tablas.Close
        d = "select * from monevi where tarjeta = " & Val(Text5)
        conexion_tablas.Open d, conexion_basedatos

        lineas = 0
        MSFlexGrid1.Clear
        Call cabecera
        Label8 = ""
        total2 = 0
        Do While Not conexion_tablas.EOF
            lineas = lineas + 1
    
            MSFlexGrid1.TextMatrix(lineas, 0) = conexion_tablas!fecha
            MSFlexGrid1.TextMatrix(lineas, 1) = conexion_tablas!detalle
            MSFlexGrid1.TextMatrix(lineas, 2) = conexion_tablas!monto
            MSFlexGrid1.TextMatrix(lineas, 3) = conexion_tablas!atendio
            total2 = total2 + conexion_tablas!monto
            conexion_tablas.MoveNext
    
        Loop
        Label8 = total2
        conexion_tablas.Close
    Else
        MsgBox "El cliente no esta cargado", vbOKOnly, "DIGISPACE"
        Form12.Show
        conexion_tablas.Close
    End If
End If

End Sub
