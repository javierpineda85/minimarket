VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Facturación"
   ClientHeight    =   7995
   ClientLeft      =   210
   ClientTop       =   600
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   10920
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
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
      Left            =   3480
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   6735
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
         TabIndex        =   40
         Top             =   360
         Width           =   6495
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   2280
         Picture         =   "Form4.frx":2895A
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   8160
      TabIndex        =   38
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808080&
      Caption         =   "Resp Inscripto"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "Cons Final"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form4.frx":2B952
      Left            =   8760
      List            =   "Form4.frx":2B954
      TabIndex        =   32
      Text            =   "Cuotas"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form4.frx":2B956
      Left            =   6720
      List            =   "Form4.frx":2B958
      TabIndex        =   31
      Text            =   "Tarjeta"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form4.frx":2B95A
      Left            =   4200
      List            =   "Form4.frx":2B95C
      TabIndex        =   30
      Text            =   "Forma de Pago"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text8 
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
      Height          =   360
      Left            =   2760
      TabIndex        =   28
      Top             =   2880
      Width           =   2415
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
      Left            =   1200
      TabIndex        =   27
      Top             =   3720
      Width           =   1695
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
      Left            =   9120
      TabIndex        =   23
      Text            =   "0"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   22
      Text            =   "0"
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   21
      Text            =   "0"
      Top             =   6600
      Width           =   855
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
      Left            =   9360
      TabIndex        =   18
      Top             =   3720
      Width           =   855
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
      Left            =   8040
      TabIndex        =   17
      Top             =   3720
      Width           =   975
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
      Left            =   3000
      TabIndex        =   16
      Top             =   3720
      Width           =   4695
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
      Left            =   480
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   5
      Top             =   810
      Width           =   4095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por nombre:"
      Height          =   375
      Left            =   5400
      TabIndex        =   37
      Top             =   2880
      Width           =   2415
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
      Left            =   720
      TabIndex        =   36
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   6360
      Picture         =   "Form4.frx":2B95E
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   600
      Picture         =   "Form4.frx":2EFAD
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   4440
      Picture         =   "Form4.frx":32874
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   2520
      Picture         =   "Form4.frx":35DE4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Codigo:"
      Height          =   375
      Left            =   720
      TabIndex        =   35
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   10200
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      X1              =   240
      X2              =   10200
      Y1              =   2760
      Y2              =   2760
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
      Left            =   8280
      TabIndex        =   29
      Top             =   240
      Width           =   975
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
      Left            =   1680
      TabIndex        =   26
      Top             =   3360
      Width           =   735
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
      Left            =   480
      TabIndex        =   25
      Top             =   1440
      Width           =   3495
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
      Left            =   8040
      TabIndex        =   24
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "IVA 21%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   6600
      Width           =   735
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
      Left            =   9360
      TabIndex        =   14
      Top             =   3360
      Width           =   855
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
      Left            =   8040
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3360
      Width           =   2655
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
      Left            =   480
      TabIndex        =   11
      Top             =   3360
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
      Left            =   9840
      TabIndex        =   10
      Top             =   240
      Width           =   1095
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
      Left            =   9240
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "CUIL/ CUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente/Razon Social:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   2055
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
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Factura Nº: 1000-0000"
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
      Left            =   5400
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "C"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   -120
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double
Private Sub Combo1_Click()
If Combo1 = "Efectivo" Then
    Combo2.Visible = False
    Combo3.Visible = False
    Combo2.AddItem "Tarjeta"
    Combo3.AddItem "Cuotas"
End If
    If Combo1 = "Credito" Then
        Combo2.Clear
        Combo3.Clear
        Combo2.Visible = True
        Combo3.Visible = True
        Combo2.AddItem "Visa"
        Combo2.AddItem "Mastercard"
        Combo2.AddItem "Nevada"
        
End If

If Combo1 = "Debito" Then
    Combo2.Clear
    Combo3.Clear
    Combo2.Visible = True
    Combo3.Visible = False
    Combo2.AddItem "Visa"
    Combo2.AddItem "Maestro"
End If


End Sub


Private Sub Combo2_click()
If Combo2 = "Nevada" Then
            Combo3.Clear
            Combo3.AddItem "Nevaplan"
            Combo3.AddItem "6"
            Combo3.AddItem "9"
            Combo3.AddItem "12"
            Else
                Combo3.Clear
                Combo3.AddItem "1"
                Combo3.AddItem "2"
                Combo3.AddItem "3"
                Combo3.AddItem "6"
                Combo3.AddItem "9"
                Combo3.AddItem "12"
        End If
End Sub

Private Sub Command1_Click()
Form7.Show

End Sub

Private Sub Image1_Click()
Frame1.Visible = False
End Sub

Private Sub Image4_Click()
Label1.Visible = False
Text8.Visible = False
Label21.Visible = False
Text1.Visible = False
Image4.Visible = False
Image2.Visible = False
Image3.Visible = False
Image5.Visible = False
'Call imprimir
'PrintForm
Y = "insert into factura values(" & Val(Label19) & ",'" & Text2 & "','" & Label9 & "','" & Text7 & "', '" & Combo1 & "','" & Combo2 & "','" & Combo3 & "')"
conexion_basedatos.Execute Y
If existe = True Then
    z = " insert into clientes values('" & Text2 & "','" & Text3 & "','" & Text4 & "')"
    conexion_basedatos.Execute z
End If
For I = 0 To List1.ListCount - 1
    X = "insert into ventas values(" & Val(List1.List(I)) & "," & Val(List5.List(I)) & ")"
    conexion_basedatos.Execute X
    r = "select * from articulos where codigo= " & Val(List5.List(I)) & ""
    conexion_tablas.Open r, conexion_basedatos
    cant = conexion_tablas!cantidad
    conexion_tablas.Close
    resta = cant - Val(List1.List(I))
    p = "update articulos set cantidad=" & resta & " where codigo= " & Val(List5.List(I))
    conexion_basedatos.Execute p
Next

MsgBox " Gracias por su compra"

Label19 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
List1.Clear: List2.Clear: List3.Clear: List4.Clear: List5.Clear:
Label1.Visible = True
Text8.Visible = True
Label21.Visible = True
Text1.Visible = True
Image4.Visible = True
Image2.Visible = True
Image3.Visible = True
Image5.Visible = True

res = MsgBox("Desea cargar otra factura?", vbYesNo, "DIGISPACE")
If res = vbYes Then
    conexion_tablas.Open datos, conexion_basedatos
    Text2.SetFocus
    If conexion_tablas.EOF = True Then
        Label19 = 1
    Else
        Label19 = conexion_tablas.Fields(0) + 1
    End If
    conexion_tablas.Close
Else
Unload Me
End If

End Sub

Private Sub image5_Click()
b = InputBox(" Ingrese numero de factura a eliminar!", "DIGISPACE")
t = "delete*from factura where nfactura=" & Val(b) & ""
conexion_basedatos.Execute t
MsgBox "La factura ha sido eliminada", vbOKOnly, "DIGISPACE"


End Sub

Private Sub image3_Click()
MsgBox "Recuerde que debe tener seleccionado al menos un item para eliminar", vbOKCancel, "DIGISPACE"
If vbOK Then
    List1.RemoveItem (List1.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    List3.RemoveItem (List3.ListIndex)
    List4.RemoveItem (List4.ListIndex)
    List5.RemoveItem (List5.ListIndex)

    MsgBox "El articulo ha sido eliminado", vbOKOnly, "DIGISPACE"
    Text5 = ""
    Text6 = ""
    Text7 = ""
Else
    Unload Form4
End If
End Sub

Private Sub image2_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Command5_Click()
If Val(Text8) < 1 Then
    MsgBox "Error, cargar codigo primero", vbOKOnly, "CUIDADO!"
    Text8.SetFocus
    Text8 = ""
Else
    X = "select*from articulos where codigo = " & Val(Text8)
    conexion_tablas.Open X, conexion_basedatos
    If conexion_tablas.EOF = True Then
        MsgBox "El codigo no existe, intente de nuevo!", vbOKOnly, "ATENCION!"
        Text8.SetFocus
        
    Else
        
        cant = InputBox("Ingresar cantidad", "DIGISPACE")
        cc = conexion_tablas!cantidad
        If cc = 1 Then
            MsgBox "No dispone del stock para realizar la venta", vbOKOnly, "DIGISPACE2"
            conexion_tablas.Close
        Else
            List1.AddItem cant
            List5.AddItem conexion_tablas!codigo
            List2.AddItem conexion_tablas!detalle
            List3.AddItem conexion_tablas!precio
            subtot = conexion_tablas!precio * cant
            List4.AddItem subtot
            Text5 = Val(Text5) + subtot
            Text6 = (Val(Text5) * 21 / 100)
            Text7 = CDbl(Text5) + CDbl(Text6) 'cdbl=convertir a doble= n con coma'
            conexion_tablas.Close
            Text8 = ""
        End If
       
    End If
    
End If
Text8 = ""
Text8.SetFocus

End Sub
Private Sub Form_Load()
Label9 = Format(Date, "short date")
datos = "select max(nfactura) from factura"
conexion_tablas.Open datos, conexion_basedatos
If conexion_tablas.EOF = True Then
    Label19 = 1
Else
    Label19 = conexion_tablas.Fields(0) + 1
End If
Combo1.AddItem "forma de Pagos"
Combo1.AddItem "Efectivo"
Combo1.AddItem "Credito"
Combo1.AddItem "Debito"
conexion_tablas.Close
End Sub

Private Sub List6_Click()
 X = "select*from articulos where codigo = " & Val(List6.Text)
    conexion_tablas.Open X, conexion_basedatos
    
        cant = InputBox("Ingresar cantidad. IMPORTANTE: solo debe ingresar numeros!", "DIGISPACE")
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

Private Sub Option1_Click()
If Option1 = True Then
    Label7.Visible = False
    Text4.Visible = False
    
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
    Label7.Visible = True
    Text4.Visible = True
Else
    Label7.Visible = False
    Text4.Visible = False
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Frame1.Visible = True
List6.Clear
datos = "select * from articulos order by detalle"
conexion_tablas.Open datos, conexion_basedatos
Do While Not conexion_tablas.EOF
If UCase(Left(conexion_tablas!detalle, Len(Text1))) = UCase(Text1) Then
    List6.AddItem conexion_tablas!codigo & " " & conexion_tablas!detalle & "   $" & conexion_tablas!precio
   
End If
conexion_tablas.MoveNext
Loop
conexion_tablas.Close

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    A = "select * from clientes where cliente like '" & Text2 & "'"
    conexion_tablas.Open A, conexion_basedatos
    If conexion_tablas.EOF = True Then
        existe = False
        conexion_tablas.Close
        MsgBox "El cliente no esta en la base de datos! Por favor cargue los datos para continuar con la venta", vbOKOnly, "DIGISPACE"
        Form6.Show
    Else
            existe = True
            Text3 = conexion_tablas!domicilio
            Text4 = conexion_tablas!cuit
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
    X = "select*from articulos where codigo = " & Val(Text8)
    conexion_tablas.Open X, conexion_basedatos
    If conexion_tablas.EOF = True Then
        MsgBox " El codigo no existe, intente de nuevo!", vbOKOnly, "DIGISPACE"
        Text8.SetFocus
        conexion_tablas.Close
    Else
        
        cant = InputBox("Ingresar cantidad. IMPORTANTE: solo debe ingresar numeros!", "DIGISPACE")
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
            'Text6 = (Text5 * 21 / 100)
            'total = Text5 '+ CDbl(Text6)
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
datos = "select max(nfactura) from factura"
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
Printer.Print Tab(30); "C"
Printer.Print Tab(2); "DIGI- SP@CE"
Printer.CurrentX = 100
Printer.FontSize = 10
Printer.Print Tab(2); "Acceso Este y Costanera, Dorrego, Guaymallen";
Printer.Print Tab(80); "Factura Nº: 1000-0000 "; Label19;
Printer.Print Tab(2); "Inicio de actividades Octubre de 2010";
Printer.Print Tab(80); "Fecha: "; fecha;
Printer.Print Tab(2), "Sede timbrado 1. CUIT 33-12345678-0";
Printer.Print Tab(80); "Cliente / Razon Social: "; Text2
Printer.Print Tab(80); "Domicilio: "; Text3
Printer.Print Tab(80); "CUIL/ CUIT: "; Text4
Printer.Print
Printer.Print "Cantidad"; Tab(15); "Codigo"; Tab(30); "Detalle"; Tab(85); "Precio"; Tab(110); "Total"
Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

tot = 0
For A = 0 To List1.ListCount - 1
    List1.ListIndex = A: cant_vendida = List1.List(A)
    List5.ListIndex = A: codigovendido = List5.List(A)
    List2.ListIndex = A: detalle = List2.List(A)
    List3.ListIndex = A: precio = List3.List(A)
    List4.ListIndex = A: total = List4.List(A)

Printer.Print cant_vendida; Tab(15); codigovendido; Tab(30); detalle; Tab(85); precio; Tab(110); total
tot = tot + total
Next
Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Printer.Print Tab(90); "total General $ "; tot
Printer.Print
conexion_tablas.Close

Printer.EndDoc

End Sub


