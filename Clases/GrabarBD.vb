Imports MySql.Data.MySqlClient
Public Class GrabarBD
    Private STRconex As String
    Private Conex As MySqlConnection = New MySqlConnection

    Public Sub SQL(ByVal Con_SQL As String)
        ' Tipo de Conexion
        STRconex = "server=" & server & ";user=" & user & ";password=" & password & ";database=" & db & ";port=" & port & ";"


        'Abrir la Conexion
        Conex = New MySqlConnection(STRconex) ' Acces 2007

        Conex.Open()
        ' Realizamos el Comando
        Dim comando As New MySql.Data.MySqlClient.MySqlCommand(Con_SQL, Conex) ' Access 2007
        comando.ExecuteNonQuery()

        ' Cerramos la conexion
        Conex.Close()
    End Sub
End Class
