Public Class buscarreceta

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim nSQL, SQL As String
        Dim regis, lista As Long
        Dim datos, data As DataSet
        Dim DB1, DB2 As New ConsultaDB
        Dim frmconsulta As New consultareceta

        nSQL = "SELECT idrec,cantidad,clave,descrip,periodic FROM detallerec WHERE idrec = " + CStr(TextEdit1.Text)
        DB1.Atabla = "detallerec"
        DB1.SQL(nSQL)
        datos = DB1.Tablas
        regis = datos.Tables(0).Rows.Count

        SQL = "SELECT * FROM recetas WHERE id = " + CStr(TextEdit1.Text)
        DB2.Atabla = "recetas"
        DB2.SQL(SQL)
        data = DB2.Tablas
        lista = data.Tables(0).Rows.Count
        If regis > 0 Then
            If lista > 0 Then
                frmconsulta.id = data.Tables(0).Rows(0).Item(0).ToString
                frmconsulta.txtnombre.Text = data.Tables(0).Rows(0).Item(1).ToString
                frmconsulta.txtdirec.Text = data.Tables(0).Rows(0).Item(2).ToString
                frmconsulta.txttelefono.Text = data.Tables(0).Rows(0).Item(3).ToString
                frmconsulta.dtpfnac.Text = data.Tables(0).Rows(0).Item(4).ToString
                frmconsulta.txtpeso.Text = data.Tables(0).Rows(0).Item(5).ToString
                frmconsulta.txtestatura.Text = data.Tables(0).Rows(0).Item(6).ToString
                frmconsulta.txtalergias.Text = data.Tables(0).Rows(0).Item(7).ToString
                frmconsulta.txtsintomas.Text = data.Tables(0).Rows(0).Item(8).ToString
                frmconsulta.txtdiag.Text = data.Tables(0).Rows(0).Item(9).ToString


                frmconsulta.DataGridView1.DataSource = datos.Tables("detallerec")
                frmconsulta.DataGridView1.Columns(0).Width = 40
                frmconsulta.DataGridView1.Columns(1).Width = 50
                frmconsulta.DataGridView1.Columns(2).Width = 70
                frmconsulta.DataGridView1.Columns(3).Width = 180
                frmconsulta.DataGridView1.Columns(4).Width = 200
                frmconsulta.DataGridView1.Columns(0).HeaderText = "ID"
                frmconsulta.DataGridView1.Columns(1).HeaderText = "Cantidad"
                frmconsulta.DataGridView1.Columns(2).HeaderText = "Clave"
                frmconsulta.DataGridView1.Columns(3).HeaderText = "Descripción"
                frmconsulta.DataGridView1.Columns(4).HeaderText = "Periodicidad"
                frmconsulta.Text = "Consulta de receta médica"

                frmconsulta.ShowDialog()

            
            End If
        Else
            MsgBox("No existe el registro de esa receta", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error de busqueda")
            TextEdit1.Text = 0
            TextEdit1.Focus()
        End If
        
    End Sub
End Class
