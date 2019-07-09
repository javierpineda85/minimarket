VERSION 5.00
Begin VB.Form Form9 
   AutoRedraw      =   -1  'True
   Caption         =   "Tickets"
   ClientHeight    =   7140
   ClientLeft      =   390
   ClientTop       =   615
   ClientWidth     =   10890
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   10890
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingresar nombre:"
      Height          =   975
      Left            =   600
      TabIndex        =   28
      Top             =   5880
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Busqueda por nombre"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6135
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   2160
         Picture         =   "Form9.frx":300A0
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   25
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   24
      Text            =   "0"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   7800
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   9120
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.ListBox List5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6240
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL A PAGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   26
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   2160
      Picture         =   "Form9.frx":33098
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   4080
      Picture         =   "Form9.frx":35FA4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   240
      Picture         =   "Form9.frx":39514
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   6000
      Picture         =   "Form9.frx":3CDDB
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   23
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Nº: 1000-0000"
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
      Left            =   5160
      TabIndex        =   22
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tiburcio Benegas y Joaquin V Gonzalez. Godoy Cruz - Mendoza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   20
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "CANT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "$ x UN."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7800
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9120
      TabIndex        =   15
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "CUIT 20-31816334-1 Inicio de Actividades: Enero 2011"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "COD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C000&
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
      Left            =   8040
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      X1              =   240
      X2              =   10200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Codigo:"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "MINIMARKET"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por nombre:"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Frame2.Visible = True

End Sub

Private Sub com_Click()
Y = "insert into ticket values(" & Val(Label19) & ",'" & Text2 & "','" & Label9 & "','" & Text7 & "')"
conexion_basedatos.Execute Y
If existe = True Then
z = " insert into clientes values(' & Consumidor final & ', '& sin nombre &','& sin numero&')"
conexion_basedatos.Execute z
End If
For I = 0 To List1.ListCount - 1
    'x = "insert into ventas values(" & Val(List1.List(I)) & "," & Val(List5.List(I)) & ")"
    'conexion_basedatos.Execute x
    r = "select * from articulos where codigo= " & Val(List5.List(I)) & ""
    conexion_tablas.Open r, conexion_basedatos
    cant = conexion_tablas!cantidad
    conexion_tablas.Close
    resta = cant - Val(List1.List(I))
    p = "update articulos set cantidad=" & resta & " where codigo= " & Val(List5.List(I))
    conexion_basedatos.Execute p
Next

Label19 = ""
Text5 = ""
Text6 = ""
Text7 = ""
List1.Clear: List2.Clear: List3.Clear: List4.Clear: List5.Clear:

res = MsgBox("Desea cargar otro Ticket?", vbYesNo, "DIGISPACE")
If res = vbYes Then
    conexion_tablas.Open datos, conexion_basedatos
    Text8.SetFocus
    If conexion_tablas.EOF = True Then
        Label19 = 1
    Else
        Label19 = conexion_tablas.Fields(0) + 1
    End If
    conexion_tablas.Close
Else
Unload Me
End If
'Else
 '   Frame2.Visible = True
    
'End If
End Sub

Private Sub Form_Load()
Frame2.Visible = False
Label9 = Format(Date, "short date")
datos = "select max(nticket) from ticket"
conexion_tablas.Open datos, conexion_basedatos
If conexion_tablas.EOF = True Then
    Label19 = 1
Else
    Label19 = conexion_tablas.Fields(0) + 1
End If
conexion_tablas.Close
End Sub


Private Sub List6_Click()
 x = "select*from articulos where codigo = " & Val(List6.Text)
    conexion_tablas.Open x, conexion_basedatos
    
        cant = InputBox("Ingresar cantidad. IMPORTANTE: solo debe ingresar numeros!", "IMPORTANTE!")
        cc = conexion_tablas!cantidad
        If cc = 0 Then
            MsgBox "No dispone del stock para realizar la venta", vbOKOnly, "DIGISPACE"
            conexion_tablas.Close
        Else
            List1.AddItem cant
            List5.AddItem conexion_tablas!codigo
            List2.AddItem conexion_tablas!detalle
            List3.AddItem conexion_tablas!precio
            subtot = conexion_tablas!precio * cant
            List4.AddItem subtot
            Text5 = Text5 + subtot
            total = Text5
            Text7 = Text5
            conexion_tablas.Close
            Text1 = ""
        End If
       
Frame1.Visible = False
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Frame1.Visible = True
List6.Clear
datos = "select * from articulos order by detalle"
conexion_tablas.Open datos, conexion_basedatos
Do While Not conexion_tablas.EOF
If UCase(Left(conexion_tablas!detalle, Len(Text1))) = UCase(Text1) Then
    List6.AddItem conexion_tablas!codigo & " " & conexion_tablas!detalle & "  $" & conexion_tablas!precio
    
   
End If
conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    a = "select * from clientes where cliente like '" & Text2 & "'"
    conexion_tablas.Open a, conexion_basedatos
    If conexion_tablas.EOF = True Then
        existe = False
        conexion_tablas.Close
        MsgBox "El cliente no esta en la base de datos! Por favor cargue los datos para continuar con la venta", vbOKOnly, "DIGISPACE"
        Form6.Show
    Else
            existe = True
            'f= "insert into
            conexion_tablas.Close
    End If
End If
End Sub



Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text8) < 1 Then
    MsgBox "Error, cargar codigo primero", vbOKOnly, "DIGISPACE"
    
    Text8.SetFocus
    Text8 = ""
Else
    x = "select*from articulos where codigo = " & Val(Text8)
    conexion_tablas.Open x, conexion_basedatos
    If conexion_tablas.EOF = True Then
        MsgBox " el codigo no existe, intente de nuevo!", vbOKOnly, "DIGISPACE"
        Text8.SetFocus
        conexion_tablas.Close
    Else
        
        cant = InputBox("Ingresar cantidad. IMPORTANTE: solo debe ingresar numeros!", "IMPORTANTE!")
        cc = conexion_tablas!cantidad
        If cc = 1 Then
            MsgBox "No dispone del stock para realizar la venta", vbOKOnly, "DIGISPACE"
            conexion_tablas.Close
        Else
            List1.AddItem cant
            List5.AddItem conexion_tablas!codigo
            List2.AddItem conexion_tablas!detalle
            List3.AddItem conexion_tablas!precio
            subtot = conexion_tablas!precio * cant
            List4.AddItem subtot
            Text5 = Text5 + subtot
            Text7 = Text5
            conexion_tablas.Close
            Text8 = ""
        End If
       
    End If
    
End If
Text8 = ""
Text8.SetFocus
End If

End Sub

    
Sub imprimir()
fecha = Date
datos = "select max(nticket) from ticket"
conexion_tablas.Open datos, conexion_basedatos
If conexion_tablas.EOF = True Then
    Label19 = 1
Else
    Label19 = conexion_tablas.Fields(0) + 1
End If
conexion_tablas.Close
conexion_tablas.Open
Printer.CurrentX = 10
Printer.CurrentY = 50
Printer.FontSize = 20
Printer.Print Tab(30); "X"
Printer.Print Tab(2); "MINIMARKET"
Printer.CurrentX = 100
Printer.FontSize = 10
Printer.Print Tab(2); "T Benegas y J V Gonzalez, Godoy Cruz;"
Printer.Print Tab(80); "TicketNº: 1000-0000 "; Label19;
Printer.Print Tab(2); "Inicio de actividades Enero de 2011";
Printer.Print Tab(80); "Fecha: "; fecha;
Printer.Print Tab(2), "Sede timbrado 1. CUIT 20-31816334-1";
Printer.Print Tab(80); "Cliente / Razon Social: "; "Consumidor final"
Printer.Print Tab(80); "Domicilio: "; "sin datos"
Printer.Print Tab(80); "CUIL/ CUIT: "; "sin datos"
Printer.Print
Printer.Print "Cantidad"; Tab(15); "Codigo"; Tab(30); "Detalle"; Tab(85); "Precio"; Tab(110); "Total"
Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

tot = 0
For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: cant_vendida = List1.List(a)
    List5.ListIndex = a: codigovendido = List5.List(a)
    List2.ListIndex = a: detalle = List2.List(a)
    List3.ListIndex = a: precio = List3.List(a)
    List4.ListIndex = a: total = List4.List(a)

Printer.Print cant_vendida; Tab(15); codigovendido; Tab(30); detalle; Tab(85); precio; Tab(110); total
tot = tot + total
Next
Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Printer.Print Tab(90); "total General $ "; tot
Printer.Print
conexion_tablas.Close

Printer.EndDoc

End Sub

Private Sub Image1_Click()
Frame1.Visible = False
End Sub

Private Sub Image4_Click()
'Call imprimir
'f = MsgBox("El pago es fiado?", vbYesNo, "ATENCION")
'If f = vbNo Then
Y = "insert into ticket values(" & Val(Label19) & ",'" & Text2 & "','" & Label9 & "','" & Text7 & "')"
conexion_basedatos.Execute Y
If existe = True Then
z = " insert into clientes values(' & Consumidor final & ', '& sin nombre &','& sin numero&')"
conexion_basedatos.Execute z
End If
For I = 0 To List1.ListCount - 1
    'x = "insert into ventas values(" & Val(List5.List(I)) & "," & Val(List1.List(I)) & ")"
    'conexion_basedatos.Execute x
    r = "select * from articulos where codigo= " & Val(List5.List(I)) & ""
    conexion_tablas.Open r, conexion_basedatos
    cant = conexion_tablas!cantidad
    conexion_tablas.Close
    resta = cant - Val(List1.List(I))
    p = "update articulos set cantidad=" & resta & " where codigo= " & Val(List5.List(I))
    conexion_basedatos.Execute p
Next

'MsgBox "Gracias por comprar en MINIMARKET"

Label19 = ""
Text5 = ""
Text6 = ""
Text7 = ""
List1.Clear: List2.Clear: List3.Clear: List4.Clear: List5.Clear:

res = MsgBox("Desea cargar otro Ticket?", vbYesNo, "DIGISPACE")
If res = vbYes Then
    conexion_tablas.Open datos, conexion_basedatos
    Text8.SetFocus
    If conexion_tablas.EOF = True Then
        Label19 = 1
    Else
        Label19 = conexion_tablas.Fields(0) + 1
    End If
    conexion_tablas.Close
Else
Unload Me
End If
'Else
 '   Frame2.Visible = True
    
'End If
End Sub

Private Sub image5_Click()
b = InputBox("Ingrese numero de ticket a eliminar!", "DIGISPACE")
t = "delete*from ticket where nticket=" & Val(b) & ""
conexion_basedatos.Execute t
MsgBox "El ticket ha sido eliminado", vbOKOnly, "DIGISPACE"


End Sub

Private Sub image3_Click()
MsgBox "Recuerde que debe tener seleccionado al menos un item para eliminar", vbOKCancel, "ATENCION"
If vbOK Then
    List1.RemoveItem (List1.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    List3.RemoveItem (List3.ListIndex)
    List4.RemoveItem (List4.ListIndex)
    List5.RemoveItem (List5.ListIndex)

    MsgBox "El articulo ha sido eliminado", vbOKOnly, "DIGISPACE"
    Text5 = ""
    Text6 = ""
    Text7 = Text5 + subtot
Else
    Unload Form4
End If
End Sub

Private Sub image2_Click()
MDIForm1.Show
Unload Me
End Sub


