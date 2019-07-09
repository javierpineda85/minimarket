VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Stock General"
   ClientHeight    =   7500
   ClientLeft      =   555
   ClientTop       =   615
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   9990
   Begin VB.TextBox Text2 
      Height          =   480
      Left            =   2760
      TabIndex        =   4
      Text            =   "Buscar por codigo"
      Top             =   6120
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   2760
      TabIndex        =   3
      Text            =   "Buscar por nombre"
      Top             =   5520
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4695
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      SelectionMode   =   1
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
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6120
      Picture         =   "Form3.frx":2C6D1
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   8040
      Picture         =   "Form3.frx":33C43
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   8040
      Picture         =   "Form3.frx":36B78
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO GENERAL DE STOCK"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Visible = False
Call cabecera
Call cargar
End Sub

Private Sub cargar()
d = "select * from articulos order by detalle"
conexion_tablas.Open d, conexion_basedatos
lineas = 0
Do While Not conexion_tablas.EOF
        lineas = lineas + 1
        MSFlexGrid1.TextMatrix(lineas, 0) = conexion_tablas!codigo
        MSFlexGrid1.TextMatrix(lineas, 1) = conexion_tablas!detalle
        MSFlexGrid1.TextMatrix(lineas, 2) = conexion_tablas!precio
        MSFlexGrid1.TextMatrix(lineas, 3) = conexion_tablas!cantidad
        conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End Sub


Private Sub Image1_Click()
Call cabecera
Call cargar
Image1.Visible = False
End Sub

Private Sub image2_Click()
Unload Me
End Sub


Private Sub Picture1_Click()
If MsgBox("Esta la impresora OK", vbInformation + vbYesNo, "listado") = vbYes Then
    datos = "select * from articulos"
    conexion_tablas.Open datos, conexion_basedatos
    Printer.CurrentX = 1
    Printer.CurrentY = 10
    Printer.Print "Codigo";
    Printer.CurrentX = 800
    Printer.Print "Detalle";
    Printer.CurrentX = 3600
    Printer.Print "Precio";
    Printer.CurrentX = 4600
    Printer.Print "Cantidad";
    Printer.CurrentX = 6000
    Printer.Print "Minimo"
    
    Do While Not conexion_tablas.EOF
            Printer.FontSize = 10
            Printer.Print
            Printer.Print conexion_tablas!codigo;
            Printer.CurrentX = 800
            Printer.Print conexion_tablas!detalle;
            Printer.CurrentX = 3600
            Printer.Print conexion_tablas!precio;
            Printer.CurrentX = 4600
            Printer.Print conexion_tablas!cantidad;
            Printer.CurrentX = 6000
            Printer.Print conexion_tablas!minimo;
    
    conexion_tablas.MoveNext
    Loop
    conexion_tablas.Close
    Printer.EndDoc
    End If

End Sub


Sub cabecera()

MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 4
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 1000
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "CODIGO"
MSFlexGrid1.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 2) = "PRECIOS"
MSFlexGrid1.TextMatrix(0, 3) = "CANT"

MSFlexGrid1.ColWidth(0) = 2000
MSFlexGrid1.ColWidth(1) = 5000
MSFlexGrid1.ColWidth(2) = 1000
MSFlexGrid1.ColWidth(3) = 1000


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
Image1.Visible = True
datos = "select * from articulos order by detalle"
conexion_tablas.Open datos, conexion_basedatos
lineas = 0
MSFlexGrid1.Clear
Call cabecera
Do While Not conexion_tablas.EOF
If UCase(Left(conexion_tablas!detalle, Len(Text1))) = UCase(Text1) Then
    
    lineas = lineas + 1
    MSFlexGrid1.TextMatrix(lineas, 0) = conexion_tablas!codigo
    MSFlexGrid1.TextMatrix(lineas, 1) = conexion_tablas!detalle
    MSFlexGrid1.TextMatrix(lineas, 2) = conexion_tablas!precio
    MSFlexGrid1.TextMatrix(lineas, 3) = conexion_tablas!cantidad
    
   
End If
conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Image1.Visible = True
datos = "select * from articulos order by detalle"
conexion_tablas.Open datos, conexion_basedatos
lineas = 0
MSFlexGrid1.Clear
Call cabecera
Do While Not conexion_tablas.EOF
If UCase(Left(conexion_tablas!codigo, Len(Text2))) = UCase(Text2) Then
    
    lineas = lineas + 1
    MSFlexGrid1.TextMatrix(lineas, 0) = conexion_tablas!codigo
    MSFlexGrid1.TextMatrix(lineas, 1) = conexion_tablas!detalle
    MSFlexGrid1.TextMatrix(lineas, 2) = conexion_tablas!precio
    MSFlexGrid1.TextMatrix(lineas, 3) = conexion_tablas!cantidad
    
   
End If
conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End If
End Sub
