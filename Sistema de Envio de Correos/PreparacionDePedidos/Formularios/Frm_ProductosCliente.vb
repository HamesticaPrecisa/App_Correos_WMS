Public Class Frm_ProductosCliente

    Public cliente As String = ""
    Public Numero_Pallet As Integer = 0

    Private Sub Frm_ProductosCliente_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "SELECT racd_codi,  drec_sopocli , mae_descr, racd_unidades, racd_peso, racd_ca, racd_ba, racd_co, racd_pi, racd_ni " +
                            "FROM rackdeta, maeprod, detarece WHERE racd_codi=drec_codi AND mae_codi=racd_codpro AND drec_rutcli='" + Rut_Cliente.Text + "' AND racd_estado<>'2'"
        Data_Productos.DataSource = ListarTablasSQL(sql)

        Dim sql_combo As String = "SELECT mae_codi, mae_descr FROM maeprod WHERE mae_clirut='" + Rut_Cliente.Text + "'"
        Cmbo_Producto.DataSource = ListarTablasSQL(sql_combo)
        Cmbo_Producto.DisplayMember = "mae_descr"
        Cmbo_Producto.ValueMember = "mae_codi"

    End Sub

    Private Sub Cmbo_Producto_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles Cmbo_Producto.SelectedValueChanged
        Dim sql As String = "SELECT racd_codi,  drec_sopocli , mae_descr, racd_unidades, racd_peso, racd_ca, racd_ba, racd_co, racd_pi, racd_ni " +
                           "FROM rackdeta, maeprod, detarece WHERE racd_codi=drec_codi AND mae_codi=racd_codpro " +
                           "AND drec_rutcli='" + Rut_Cliente.Text + "' AND mae_codi='" + Cmbo_Producto.SelectedValue.ToString() + "' AND racd_estado<>'2'"
        Data_Productos.DataSource = ListarTablasSQL(sql)
    End Sub

    Private Sub Txt_Soportante_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Txt_Soportante.KeyPress

        If e.KeyChar = ChrW(13) Then
            Dim sql As String = "SELECT racd_codi,  drec_sopocli , mae_descr, racd_unidades, racd_peso, racd_ca, racd_ba, racd_co, racd_pi, racd_ni " +
                   "FROM rackdeta, maeprod, detarece WHERE racd_codi=drec_codi AND mae_codi=racd_codpro AND racd_estado<>'2' " +
                   "AND drec_rutcli='" + Rut_Cliente.Text + "' " +
                   "AND (racd_codi LIKE'%" + Txt_Soportante.Text + "%' OR drec_sopocli LIKE'%" + Txt_Soportante.Text + "%')"

            Data_Productos.DataSource = ListarTablasSQL(sql)

        End If

    End Sub

    Private Sub Btn_Todos_Click(sender As System.Object, e As System.EventArgs) Handles Btn_Todos.Click
        Dim sql As String = "SELECT racd_codi,  drec_sopocli , mae_descr, racd_unidades, racd_peso, racd_ca, racd_ba, racd_co, racd_pi, racd_ni " +
                           "FROM rackdeta, maeprod, detarece WHERE  racd_estado<>'2' AND racd_codi=drec_codi AND mae_codi=racd_codpro AND drec_rutcli='" + Rut_Cliente.Text + "'"
        Data_Productos.DataSource = ListarTablasSQL(sql)
    End Sub

    Private Sub Data_Productos_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Data_Productos.CellClick
        If e.RowIndex > -1 AndAlso e.ColumnIndex = 0 Then
            If MsgBox("Seguro de cambiar por este pallet", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Aviso") = vbYes Then
                'Cambia el pallet por
                Dim sql As String = "UPDATE rackdeta SET racd_estado='0' WHERE racd_codi='" + CerosAnteriorString(Numero_Pallet.ToString(), 9) + "'"
                If MovimientoSQL(sql) > 0 Then
                    Dim sql2 As String = "UPDATE rackdeta SET racd_estado='2' WHERE racd_codi='" + Me.Data_Productos.Rows(e.RowIndex).Cells(1).Value.ToString() + "'"
                    If MovimientoSQL(sql2) > 0 Then
                        MsgBox("Pallet Cambiado", MsgBoxStyle.Information, "Aviso")

                        Cambia_Pallet_pedido(Me.Data_Productos.Rows(e.RowIndex).Cells(1).Value.ToString(), CerosAnteriorString(Numero_Pallet.ToString(), 9) _
                                             , Me.Data_Productos.Rows(e.RowIndex).Cells(4).Value.ToString(), Me.Data_Productos.Rows(e.RowIndex).Cells(5).Value.ToString() _
                                             , Me.Data_Productos.Rows(e.RowIndex).Cells(6).Value.ToString() + "-" + Me.Data_Productos.Rows(e.RowIndex).Cells(7).Value.ToString() + "-" + _
                                             Me.Data_Productos.Rows(e.RowIndex).Cells(8).Value.ToString() + "-" + Me.Data_Productos.Rows(e.RowIndex).Cells(9).Value.ToString() + "-" + _
                                             Me.Data_Productos.Rows(e.RowIndex).Cells(10).Value.ToString())

                        Me.Close()
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub Cambia_Pallet_pedido(ByVal Pallet_Nuevo As String, ByVal Pallet_antiguo As String, ByVal cajas As String, ByVal kilos As String, ByVal posicion As String)

        Dim sqlselecciona As String = "SELECT linea, pedido FROM pedidos_detalle WHERE pallet='" + Pallet_antiguo + "'"
        Dim tabla As DataTable = ListarTablasSQL(sqlselecciona)

        If tabla.Rows.Count > 0 Then

            Dim sql As String = "DELETE FROM pedidos_detalle WHERE pallet='" + Pallet_antiguo + "'"

            Dim sql2 As String = "INSERT INTO pedidos_detalle(pedido, linea, pallet, cajas, kilos, posicion, est) " +
                "VALUES('" + tabla.Rows(0)(1).ToString() + "','" + tabla.Rows(0)(0).ToString() + "', '" + Pallet_Nuevo.ToString() + "'," +
                "'" + cajas + "','" + kilos.Replace(",", ".") + "','" + posicion + "','0')"
            MovimientoSQL(sql)
            MovimientoSQL(sql2)
        End If


    End Sub

End Class