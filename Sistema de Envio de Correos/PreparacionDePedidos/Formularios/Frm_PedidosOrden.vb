Public Class Frm_PedidosOrden

    Dim fila As Integer = 0
    Dim BindingSource1 As New BindingSource

    Private Sub Frm_PedidosOrden_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.Close()
        End If
    End Sub


    Private Sub Frm_PedidosOrden_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = (ListarTablasSQL("SELECT pedido, cliente+ '  '+ cli_nomb AS cliente, fecha, hora, destino, orden FROM pedidos_ficha, clientes WHERE " +
                                                    "cli_rut=cliente AND terminado='0' ORDER BY orden ASC"))




    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > -1 AndAlso e.ColumnIndex = 0 Then
            Dim frm As New Frm_PedidosDetalle
            frm.CODIGO_PEDIDO = Me.DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
            frm.ShowDialog()
        End If
    End Sub

 

End Class