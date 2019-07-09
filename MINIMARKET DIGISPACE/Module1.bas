Attribute VB_Name = "Module1"
Public conexion_basedatos As New ADODB.Connection
Public conexion_tablas As New ADODB.Recordset
Global datos As String
Sub abrir()
    conexion_basedatos.ConnectionString = App.Path + "\stock.mdb"
    conexion_basedatos.Provider = "microsoft.jet.oledb.4.0"
    conexion_basedatos.Open
End Sub

Sub cerrar()
    conexion_basedatos.Close
End Sub
