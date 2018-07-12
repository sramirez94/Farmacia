Imports MySql.Data.MySqlClient

Public Class ConsultaDBRP
    Private consulta As String
    Private STRconex As String
    Private Conex As MySqlConnection
    Private Adapt As MySqlDataAdapter
    Private Tabla As DataSet1
    Private nRegistros As Long
    Private NombT As String

    Public Sub SQL(ByVal Con_SQL As String)
        ' Tipo de Conexion
        STRconex = "server=" & server & ";user=" & user & ";password=" & password & ";database=" & db & ";port=" & port & ";"

        'Abrir la Conexion
        Conex = New MySqlConnection(STRconex)

        ' Abrir el Adaptador
        Adapt = New MySqlDataAdapter(Con_SQL, Conex)

        ' Poner los Datos en un DataSet
        Tabla = New DataSet1
        ' Leer los Datos (Se cierra automaticamente
        Adapt.Fill(Tabla, NombT)

        nRegistros = Tabla.Tables(0).Rows.Count

    End Sub
    Public ReadOnly Property Tablas() As DataSet1
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
