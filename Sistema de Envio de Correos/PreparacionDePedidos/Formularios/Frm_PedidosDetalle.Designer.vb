<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_PedidosDetalle
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Folio_precisa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Folio_Cliente = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Producto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Unidades = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Kilos = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Camara = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Banda = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Columna = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Piso = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nivel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CambioPallet = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CAMBIARPALLET = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtkilos = New System.Windows.Forms.TextBox()
        Me.txtcajas = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtsopo = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CambioPallet.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Folio_precisa, Me.Folio_Cliente, Me.Producto, Me.Unidades, Me.Kilos, Me.Camara, Me.Banda, Me.Columna, Me.Piso, Me.Nivel})
        Me.DataGridView1.ContextMenuStrip = Me.CambioPallet
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 30)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(822, 364)
        Me.DataGridView1.TabIndex = 0
        '
        'Folio_precisa
        '
        Me.Folio_precisa.DataPropertyName = "pallet"
        Me.Folio_precisa.HeaderText = "Folio Precisa"
        Me.Folio_precisa.Name = "Folio_precisa"
        Me.Folio_precisa.ReadOnly = True
        Me.Folio_precisa.Width = 80
        '
        'Folio_Cliente
        '
        Me.Folio_Cliente.DataPropertyName = "drec_sopocli"
        Me.Folio_Cliente.HeaderText = "Folio Cliente"
        Me.Folio_Cliente.Name = "Folio_Cliente"
        Me.Folio_Cliente.ReadOnly = True
        Me.Folio_Cliente.Width = 80
        '
        'Producto
        '
        Me.Producto.DataPropertyName = "mae_descr"
        Me.Producto.HeaderText = "Producto"
        Me.Producto.Name = "Producto"
        Me.Producto.ReadOnly = True
        Me.Producto.Width = 250
        '
        'Unidades
        '
        Me.Unidades.DataPropertyName = "racd_unidades"
        Me.Unidades.HeaderText = "Unidades"
        Me.Unidades.Name = "Unidades"
        Me.Unidades.ReadOnly = True
        Me.Unidades.Width = 80
        '
        'Kilos
        '
        Me.Kilos.DataPropertyName = "racd_peso"
        Me.Kilos.HeaderText = "Kilos"
        Me.Kilos.Name = "Kilos"
        Me.Kilos.ReadOnly = True
        Me.Kilos.Width = 80
        '
        'Camara
        '
        Me.Camara.DataPropertyName = "racd_ca"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Camara.DefaultCellStyle = DataGridViewCellStyle1
        Me.Camara.HeaderText = "Ca"
        Me.Camara.Name = "Camara"
        Me.Camara.ReadOnly = True
        Me.Camara.Width = 30
        '
        'Banda
        '
        Me.Banda.DataPropertyName = "racd_ba"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Banda.DefaultCellStyle = DataGridViewCellStyle2
        Me.Banda.HeaderText = "Ba"
        Me.Banda.Name = "Banda"
        Me.Banda.ReadOnly = True
        Me.Banda.Width = 30
        '
        'Columna
        '
        Me.Columna.DataPropertyName = "racd_co"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Columna.DefaultCellStyle = DataGridViewCellStyle3
        Me.Columna.HeaderText = "Co"
        Me.Columna.Name = "Columna"
        Me.Columna.ReadOnly = True
        Me.Columna.Width = 30
        '
        'Piso
        '
        Me.Piso.DataPropertyName = "racd_pi"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Piso.DefaultCellStyle = DataGridViewCellStyle4
        Me.Piso.HeaderText = "Pi"
        Me.Piso.Name = "Piso"
        Me.Piso.ReadOnly = True
        Me.Piso.Width = 30
        '
        'Nivel
        '
        Me.Nivel.DataPropertyName = "racd_ni"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Nivel.DefaultCellStyle = DataGridViewCellStyle5
        Me.Nivel.HeaderText = "Ni"
        Me.Nivel.Name = "Nivel"
        Me.Nivel.ReadOnly = True
        Me.Nivel.Width = 30
        '
        'CambioPallet
        '
        Me.CambioPallet.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.CambioPallet.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CAMBIARPALLET})
        Me.CambioPallet.Name = "ContextMenuStrip1"
        Me.CambioPallet.Size = New System.Drawing.Size(180, 26)
        '
        'CAMBIARPALLET
        '
        Me.CAMBIARPALLET.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CAMBIARPALLET.Name = "CAMBIARPALLET"
        Me.CAMBIARPALLET.Size = New System.Drawing.Size(179, 22)
        Me.CAMBIARPALLET.Text = "CAMBIAR PALLET"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(-1, -1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(824, 37)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "DETALLE DEL PEDIDO"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtkilos
        '
        Me.txtkilos.Location = New System.Drawing.Point(715, 400)
        Me.txtkilos.Name = "txtkilos"
        Me.txtkilos.ReadOnly = True
        Me.txtkilos.Size = New System.Drawing.Size(94, 21)
        Me.txtkilos.TabIndex = 7
        '
        'txtcajas
        '
        Me.txtcajas.Location = New System.Drawing.Point(584, 400)
        Me.txtcajas.Name = "txtcajas"
        Me.txtcajas.ReadOnly = True
        Me.txtcajas.Size = New System.Drawing.Size(73, 21)
        Me.txtcajas.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(534, 403)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "CAJAS"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(669, 404)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "KILOS"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(398, 403)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "SOPORTANTES"
        '
        'txtsopo
        '
        Me.txtsopo.Location = New System.Drawing.Point(492, 400)
        Me.txtsopo.Name = "txtsopo"
        Me.txtsopo.ReadOnly = True
        Me.txtsopo.Size = New System.Drawing.Size(36, 21)
        Me.txtsopo.TabIndex = 12
        '
        'Frm_PedidosDetalle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(822, 424)
        Me.Controls.Add(Me.txtsopo)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtcajas)
        Me.Controls.Add(Me.txtkilos)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Frm_PedidosDetalle"
        Me.Padding = New System.Windows.Forms.Padding(0, 30, 0, 30)
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DETALLE PEDIDO POR EL CLIENTE"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CambioPallet.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Folio_precisa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Folio_Cliente As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Producto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Unidades As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Kilos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Camara As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Banda As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Columna As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Piso As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nivel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtkilos As System.Windows.Forms.TextBox
    Friend WithEvents txtcajas As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtsopo As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents CambioPallet As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CAMBIARPALLET As System.Windows.Forms.ToolStripMenuItem
End Class
