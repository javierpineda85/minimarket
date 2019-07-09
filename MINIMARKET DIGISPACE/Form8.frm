VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form8 
   BackColor       =   &H00C0C000&
   Caption         =   "Total de ventas"
   ClientHeight    =   7350
   ClientLeft      =   390
   ClientTop       =   795
   ClientWidth     =   10020
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   10020
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   3495
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      CausesValidation=   0   'False
      Height          =   3015
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   6
      Cols            =   6
      BackColor       =   14737632
      BackColorBkg    =   14737632
      GridColor       =   16777215
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CIERRE DE TICKETS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CIERRE DE FACTURAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   6120
      Picture         =   "Form8.frx":2C6D1
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4680
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VENTAS TOTALES"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double

Private Sub Command1_Click()
x = "select fecha from ticket where fecha between  #" & fecha1 & "# and #" & fecha2 & "#"
conexion_tablas.Open x, conexion_basedatos
lineas = 0
Do While Not conexion_tablas.EOF
   If Trim(conexion_tablas!fecha) = Trim(Date) Then
        lineas = lineas + 1
        MSFlexGrid2.TextMatrix(lineas, 0) = conexion_tablas!nticket
        MSFlexGrid2.TextMatrix(lineas, 1) = conexion_tablas!cliente
        MSFlexGrid2.TextMatrix(lineas, 2) = conexion_tablas!fecha
        MSFlexGrid2.TextMatrix(lineas, 3) = conexion_tablas!total
        total2 = total2 + conexion_tablas!total

    End If
    
    conexion_tablas.MoveNext
Loop
Label7 = total + total2
conexion_tablas.Close
End Sub

Private Sub Form_Load()
'Call cabecera
'd = "select * from factura"
'conexion_tablas.Open d, conexion_basedatos
lineas = 0
Label7 = ""
fecha1 = Date
fecha2 = Date
Call cabecera2
't = "select * from ticket"
'conexion_tablas.Open t, conexion_basedatos
'lineas = 0
'Do While Not conexion_tablas.EOF
'    If Trim(conexion_tablas!fecha) = Trim(Date) Then
'        lineas = lineas + 1
'        MSFlexGrid2.TextMatrix(lineas, 0) = conexion_tablas!nticket
'        MSFlexGrid2.TextMatrix(lineas, 1) = conexion_tablas!cliente
'        MSFlexGrid2.TextMatrix(lineas, 2) = conexion_tablas!fecha
'        MSFlexGrid2.TextMatrix(lineas, 3) = conexion_tablas!total
'       total2 = total2 + conexion_tablas!total
'
'    End If
    
'    conexion_tablas.MoveNext
'Loop
'Label7 = total + total2
'conexion_tablas.Close
End Sub

Private Sub cargaflex1()
Do While Not conexion_tablas.EOF
    If Trim(conexion_tablas!fecha) = Trim(Date) Then
        lineas = lineas + 1
        MSFlexGrid1.TextMatrix(lineas, 0) = conexion_tablas!nfactura
        MSFlexGrid1.TextMatrix(lineas, 1) = conexion_tablas!cliente
        MSFlexGrid1.TextMatrix(lineas, 2) = conexion_tablas!fecha
        MSFlexGrid1.TextMatrix(lineas, 3) = conexion_tablas!total
        MSFlexGrid1.TextMatrix(lineas, 4) = conexion_tablas!fpago
        MSFlexGrid1.TextMatrix(lineas, 5) = conexion_tablas!tarjeta
        MSFlexGrid1.TextMatrix(lineas, 6) = conexion_tablas!cuotas
        total = total + conexion_tablas!total

    End If
    
    conexion_tablas.MoveNext
Loop
conexion_tablas.Close


End Sub
Sub cabecera()

MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 7
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 100
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "Nº Factura"
MSFlexGrid1.TextMatrix(0, 1) = "Cliente"
MSFlexGrid1.TextMatrix(0, 2) = "Fecha"
MSFlexGrid1.TextMatrix(0, 3) = "Total"
MSFlexGrid1.TextMatrix(0, 4) = "F de pago"
MSFlexGrid1.TextMatrix(0, 5) = "Tarjeta"
MSFlexGrid1.TextMatrix(0, 6) = "Cuotas"

MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 2500
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 800
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(6) = 1000

End Sub

Sub cabecera2()

MSFlexGrid2.FixedCols = 0
MSFlexGrid2.Cols = 4
MSFlexGrid2.FixedRows = 1
MSFlexGrid2.Rows = 500
MSFlexGrid2.Clear
MSFlexGrid2.TextMatrix(0, 0) = "Nº Ticket"
MSFlexGrid2.TextMatrix(0, 1) = "Cliente"
MSFlexGrid2.TextMatrix(0, 2) = "Fecha"
MSFlexGrid2.TextMatrix(0, 3) = "Total"

MSFlexGrid2.ColWidth(0) = 1000
MSFlexGrid2.ColWidth(1) = 2500
MSFlexGrid2.ColWidth(2) = 1200
MSFlexGrid2.ColWidth(3) = 800


End Sub



Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
End Sub

