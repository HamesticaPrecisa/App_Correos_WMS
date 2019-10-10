Public Class Frm_Pedidos


    Private Sub Frm_Pedidos_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Timer1.Stop()
    End Sub

    Private Sub Frm_Pedidos_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.Close()
        End If
    End Sub
 
    Private Sub Frm_Pedidos_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
        carga_pedidos()

    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > -1 AndAlso e.ColumnIndex = 0 Then
            Frm_PedidosDetalle.CODIGO_PEDIDO = Me.DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
            Frm_PedidosDetalle.cliente = Me.DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Remove(0, 14)
            Frm_PedidosDetalle.ShowDialog()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick

    End Sub



    Sub carga_pedidos()

        Dim sql As String = "SELECT  pedido, Orden, cliente+ '  '+ cli_nomb AS cliente, fecha, hora, destino, destino AS sopo FROM pedidos_ficha, clientes " +
                            "WHERE cli_rut=cliente AND terminado ='0' ORDER BY orden DESC"

        Dim Data As DataTable = ListarTablasSQL(sql)

        For i As Integer = 0 To Data.Rows.Count - 1

            Dim sqlTotal As String = "SELECT COUNT(*) FROM pedidos_detalle WHERE pedido='" + Data.Rows(i)(0).ToString() + "'"
            Dim tablatotal As DataTable = ListarTablasSQL(sqlTotal)

            Dim sqlTotal2 As String = "SELECT COUNT(*) FROM pedidos_detalle WHERE pedido='" + Data.Rows(i)(0).ToString() + "' AND est='1'"
            Dim tablatotal2 As DataTable = ListarTablasSQL(sqlTotal2)

            If tablatotal.Rows.Count > 0 AndAlso tablatotal2.Rows.Count > 0 Then
                Dim todos As String = tablatotal.Rows(0)(0).ToString()
                Dim de As String = tablatotal2.Rows(0)(0).ToString()
                Data.Rows(i)(6) = (de + " / " + todos).ToString()
            End If
        Next
        DataGridView1.DataSource = Data
    End Sub
End Class