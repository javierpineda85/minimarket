VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "DIGI-SPACE"
   ClientHeight    =   6645
   ClientLeft      =   360
   ClientTop       =   900
   ClientWidth     =   9900
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu Articulos 
      Caption         =   "Articulos"
   End
   Begin VB.Menu listado 
      Caption         =   "Listado"
      Begin VB.Menu General 
         Caption         =   "General"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Reposicion 
         Caption         =   "Reposicion"
      End
   End
   Begin VB.Menu compras 
      Caption         =   "Compras"
   End
   Begin VB.Menu proveedores 
      Caption         =   "Proveedores"
      Begin VB.Menu nuevo_prov 
         Caption         =   "Nuevo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu lista 
         Caption         =   "Listado"
      End
   End
   Begin VB.Menu ventas 
      Caption         =   "Ventas"
      Begin VB.Menu ticket 
         Caption         =   "Ticket"
         Shortcut        =   {F3}
      End
      Begin VB.Menu factura 
         Caption         =   "Factura"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Clientes 
      Caption         =   "Clientes"
      Begin VB.Menu nuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu ctacte 
         Caption         =   "Ctas corrientes"
      End
   End
   Begin VB.Menu Cierre 
      Caption         =   "Cierre"
      Begin VB.Menu cdecaja 
         Caption         =   "Cierre de caja"
      End
      Begin VB.Menu tdeventas 
         Caption         =   "Total de ventas"
      End
   End
   Begin VB.Menu salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Articulos_Click()
Form1.Show
End Sub


Private Sub cdecaja_Click()
Form10.Show
End Sub

Private Sub compras_Click()
Form5.Show
End Sub

Private Sub ctacte_Click()
Form11.Show
End Sub

Private Sub factura_Click()
Form4.Show
End Sub

Private Sub General_Click()
Form3.Show
End Sub

Private Sub lista_Click()
Form7.Show
End Sub

Private Sub MDIForm_Load()
Call abrir

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call cerrar

End Sub

Private Sub nose_Click()
Form9.Show
End Sub

Private Sub nuevo_Click()
Form12.Show
End Sub

Private Sub nuevo_prov_Click()
Form6.Show
End Sub

Private Sub Reposicion_Click()
Form2.Show
End Sub

Private Sub salir_Click()
a = MsgBox("Esta seguro que desea cerrar el programa?", vbYesNo, "Cerrando aplicacion")
If a = vbYes Then
    End
End If

End Sub

Private Sub tdeventas_Click()
Form8.Show
End Sub

Private Sub ticket_Click()
Form9.Show
End Sub


