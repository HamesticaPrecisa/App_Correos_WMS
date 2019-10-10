Public Class Frm_PedidosDetalle

    Public CODIGO_PEDIDO As String
    Public cliente As String
    Dim fila = -1

    Private Sub Frm_PedidosDetalle_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.Close()
        End If
    End Sub


    Private Sub Frm_PedidosDetalle_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim sql As String = "SELECT pallet, drec_sopocli, mae_descr, racd_unidades, racd_peso, racd_ca, racd_ba, racd_co, racd_pi, racd_ni " +
                            "FROM Pedidos_detalle, rackdeta, detarece, maeprod " +
                            "WHERE pedido = '" + CODIGO_PEDIDO.ToString() + "' AND pallet = rackdeta.racd_codi " +
                            "AND drec_codi=racd_codi AND mae_codi=racd_codpro"

        Me.DataGridView1.DataSource = ListarTablasSQL(sql)
        Sumar()
    End Sub

    Sub Sumar()
        If DataGridView1.Rows.Count > 0 Then
            Dim suma_cajas As Integer = 0
            Dim suma_kilos As Double = 0.0

            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                suma_cajas = suma_cajas + Convert.ToInt32(DataGridView1.Rows(i).Cells(3).Value.ToString())
                suma_kilos = suma_kilos + Convert.ToDouble(DataGridView1.Rows(i).Cells(4).Value.ToString())
            Next

            txtsopo.Text = DataGridView1.Rows.Count
            txtcajas.Text = suma_cajas.ToString()
            txtkilos.Text = suma_kilos.ToString()

        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex > -1 AndAlso e.ColumnIndex > -1 Then
            If DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString() = "91" Or DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString() = "92" Then
                MsgBox("No puede cambiar el pallet, ya se encuentra en la camara pulmon", MsgBoxStyle.Information, "Aviso")
                fila = -1
            Else
                If MsgBox("Seguro de cambiar el pallet " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString() + "", _
                           MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Aviso") = vbYes Then
                    Dim Frm As New Frm_ProductosCliente
                    Frm.Numero_Pallet = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString()
                    Frm.Rut_Cliente.Text = cliente.ToString()
                    Frm.ShowDialog()

                    Dim sql As String = "SELECT pallet, drec_sopocli, mae_descr, racd_unidades, racd_peso, racd_ca, racd_ba, racd_co, racd_pi, racd_ni " +
                                "FROM Pedidos_detalle, rackdeta, detarece, maeprod " +
                                "WHERE pedido = '" + CODIGO_PEDIDO.ToString() + "' AND pallet = rackdeta.racd_codi " +
                                "AND drec_codi=racd_codi AND mae_codi=racd_codpro"

                    Me.DataGridView1.DataSource = ListarTablasSQL(sql)
                    Sumar()
                End If
            End If
        End If
    End Sub
End Class