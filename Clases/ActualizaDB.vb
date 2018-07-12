Public Class ActualizarDB
    Private Registros As Long = 0
    Private I_Campos As String = ""

    Public Sub Agregar(ByVal Campo As String, ByVal Informacion As String, ByVal Tipo As String)
        Informacion = Replace(Informacion, "'", "''")
        Select Case Tipo
            Case "N"
            Case "C"
                Informacion = "'" + Informacion + "'"
            Case "B"
            Case "F"
                Informacion = "'" + Informacion + "'"
            Case Else
                MsgBox("Los Valores Permitidos son N,C,B,F", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Tipos")
                Exit Sub
        End Select
        If Registros = 0 Then
            I_Campos = I_Campos + Campo + "=" + Informacion
            Registros = Registros + 1
        Else
            I_Campos = I_Campos + "," + Campo + "=" + Informacion
            Registros = Registros + 1
        End If
    End Sub
    Public Sub Limpiar()
        I_Campos = ""
        Registros = 0
    End Sub
    Public Property campos() As String
        Get
            Return I_Campos
        End Get
        Set(ByVal value As String)
        End Set
    End Property
End Class
