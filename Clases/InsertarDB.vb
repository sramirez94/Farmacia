Public Class InsertarDB

    Private Registros As Long = 0
    Private I_Campos As String = ""
    Private I_Datos As String = ""

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
            I_Campos = I_Campos + Campo
            I_Datos = I_Datos + Informacion
            Registros = Registros + 1
        Else
            I_Campos = I_Campos + "," + Campo
            I_Datos = I_Datos + "," + Informacion
            Registros = Registros + 1
        End If
    End Sub

    Public Sub Limpiar()
        I_Campos = ""
        I_Datos = ""
        Registros = 0
    End Sub
    Public Property campos() As String
        Get
            Return I_Campos
        End Get
        Set(ByVal value As String)
        End Set
    End Property
    Public Property datos() As String
        Get
            Return I_Datos
        End Get
        Set(ByVal value As String)
        End Set
    End Property
End Class
