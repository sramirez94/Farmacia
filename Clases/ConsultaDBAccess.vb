Imports System.Data
Imports System.Data.OleDb
Public Class ConsultaDBAccess

    Private consulta As String
    Private STRconex As String
   
   'Esta clase genera consultas a una base de datos de Access
   
    Private Conex As OleDb.OleDbConnection ' Acces 2007
    Private Adapt As OleDb.OleDbDataAdapter ' Acces 2007
    Private Tabla As DataSet
    Private nRegistros As Long
    Private NombT As String
    Public Sub SQL(ByVal Con_SQL As String)
        ' Tipo de Conexion
        STRconex = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + DBruta + "conex.accdb;"  

        'Abrir la Conexion
        Conex = New OleDb.OleDbConnection(STRconex) 

        ' Abrir el Adaptador
        Adapt = New OleDb.OleDbDataAdapter(Con_SQL, Conex) 

        ' Poner los Datos en un DataSet
        Tabla = New DataSet

        ' Leer los Datos (Se cierra automaticamente
        Adapt.Fill(Tabla, NombT)

        nRegistros = Tabla.Tables(0).Rows.Count

    End Sub
    Public ReadOnly Property Tablas() As DataSet
        Get
            Return Tabla
        End Get
    End Property
    Public Property Atabla() As String
        Get
            Return NombT
        End Get
        Set(ByVal value As String)
            NombT = value
        End Set
    End Property

End Class
