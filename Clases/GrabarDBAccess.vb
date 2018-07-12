Imports System.Data
Imports System.Data.OleDb
Public Class GrabarDBAccess

    Private STRconex As String
    Private Conex As OleDbConnection = New OleDbConnection ' Access 2007
    
    Public Sub SQL(ByVal Con_SQL As String)
        ' Tipo de Conexion
        STRconex = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + DBruta + "conex.accdb;"  ' 2007

        'Abrir la Conexion
        Conex.ConnectionString = STRconex

        ' Abrimos la Conexion
        Conex.Open()

        ' Realizamos el Comando
        Dim comando As New OleDb.OleDbCommand(Con_SQL, Conex) ' Access 2007
        comando.ExecuteNonQuery()

        ' Cerramos la conexion
        Conex.Close()
    End Sub
End Class

'En esta clase se genera in INSERT a una base de datos de Access, la cual, tenemos para cuardar las conexiones que se harán a diferentes
'Bases de datos, por ejemplo, si hay dos o tres server a donde nos podemos conectar, el usuario puede decidir en donde hacerlo
'por medio de un formulario 
