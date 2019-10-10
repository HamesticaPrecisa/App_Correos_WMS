Public Class Frm_PedidosTerminados

    Private Sub Frm_PedidosTerminados_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Timer1.Stop()
    End Sub

    Private Sub Frm_PedidosTerminados_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.Close()
        End If
    End Sub

    Private Sub Frm_PedidosTerminados_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ListarTablasSQL("SELECT pedido, cliente+ '  '+ cli_nomb AS cliente, fecha, hora, destino FROM pedidos_ficha, clientes WHERE " +
                                                "cli_rut=cliente AND terminado='1'")
        Timer1.Start()
    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > -1 AndAlso e.ColumnIndex = 0 Then
            Frm_PedidosDetalle.CODIGO_PEDIDO = Me.DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString()
            Frm_PedidosDetalle.ShowDialog()
        ElseIf e.RowIndex > -1 AndAlso e.ColumnIndex = 1 Then
            If MsgBox("Desea eliminar este pedido de esta lista", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Aviso") = vbYes Then
                Dim sql As String = "UPDATE pedidos_ficha SET terminado='2' WHERE pedido='" + Me.DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString() + "'"
                If MovimientoSQL(sql) > 0 Then
                    MsgBox("Pedido quitado de la lista", MsgBoxStyle.Information, "Aviso")
                    Frm_PedidosTerminados_Load(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        DataGridView1.DataSource = ListarTablasSQL("SELECT pedido, cliente+ '  '+ cli_nomb AS cliente, fecha, hora, destino FROM pedidos_ficha, clientes WHERE " +
                                                "cli_rut=cliente AND terminado='1'")
    End Sub
End Class