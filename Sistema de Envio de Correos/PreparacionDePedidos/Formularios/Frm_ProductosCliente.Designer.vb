<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_ProductosCliente
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Rut_Cliente = New System.Windows.Forms.Label()
        Me.Data_Productos = New System.Windows.Forms.DataGridView()
        Me.Folio_precisa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FOLIO_CLIENTE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PRODUCTO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cajas = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Kilos = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CAMARA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BANDA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COLUMNA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PISO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NIVEL = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Txt_Soportante = New System.Windows.Forms.TextBox()
        Me.Cmbo_Producto = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CAMBIARPORESTEPALLETToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataGridViewImageColumn1 = New System.Windows.Forms.DataGridViewImageColumn()
        Me.OK = New System.Windows.Forms.DataGridViewImageColumn()
        Me.Btn_Todos = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.Data_Productos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Panel1.Controls.Add(Me.Rut_Cliente)
        Me.Panel1.Controls.Add(Me.Data_Productos)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1087, 525)
        Me.Panel1.TabIndex = 0
        '
        'Rut_Cliente
        '
        Me.Rut_Cliente.AutoSize = True
        Me.Rut_Cliente.Location = New System.Drawing.Point(775, 81)
        Me.Rut_Cliente.Name = "Rut_Cliente"
        Me.Rut_Cliente.Size = New System.Drawing.Size(14, 13)
        Me.Rut_Cliente.TabIndex = 4
        Me.Rut_Cliente.Text = "0"
        Me.Rut_Cliente.Visible = False
        '
        'Data_Productos
        '
        Me.Data_Productos.AllowUserToAddRows = False
        Me.Data_Productos.AllowUserToDeleteRows = False
        Me.Data_Productos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Data_Productos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Folio_precisa, Me.FOLIO_CLIENTE, Me.PRODUCTO, Me.Cajas, Me.Kilos, Me.CAMARA, Me.BANDA, Me.COLUMNA, Me.PISO, Me.NIVEL, Me.OK})
        Me.Data_Productos.Location = New System.Drawing.Point(14, 111)
        Me.Data_Productos.Name = "Data_Productos"
        Me.Data_Productos.ReadOnly = True
        Me.Data_Productos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.Data_Productos.Size = New System.Drawing.Size(1054, 402)
        Me.Data_Productos.TabIndex = 0
        '
        'Folio_precisa
        '
        Me.Folio_precisa.DataPropertyName = "racd_codi"
        Me.Folio_precisa.HeaderText = "FOLIO PRECISA"
        Me.Folio_precisa.Name = "Folio_precisa"
        Me.Folio_precisa.ReadOnly = True
        '
        'FOLIO_CLIENTE
        '
        Me.FOLIO_CLIENTE.DataPropertyName = "drec_sopocli"
        Me.FOLIO_CLIENTE.HeaderText = "FOLIO CLIENTE"
        Me.FOLIO_CLIENTE.Name = "FOLIO_CLIENTE"
        Me.FOLIO_CLIENTE.ReadOnly = True
        '
        'PRODUCTO
        '
        Me.PRODUCTO.DataPropertyName = "mae_descr"
        Me.PRODUCTO.HeaderText = "PRODUCTO"
        Me.PRODUCTO.Name = "PRODUCTO"
        Me.PRODUCTO.ReadOnly = True
        Me.PRODUCTO.Width = 250
        '
        'Cajas
        '
        Me.Cajas.DataPropertyName = "racd_unidades"
        Me.Cajas.HeaderText = "Cajas"
        Me.Cajas.Name = "Cajas"
        Me.Cajas.ReadOnly = True
        '
        'Kilos
        '
        Me.Kilos.DataPropertyName = "racd_peso"
        Me.Kilos.HeaderText = "Kilos"
        Me.Kilos.Name = "Kilos"
        Me.Kilos.ReadOnly = True
        '
        'CAMARA
        '
        Me.CAMARA.DataPropertyName = "racd_ca"
        Me.CAMARA.HeaderText = "CA"
        Me.CAMARA.Name = "CAMARA"
        Me.CAMARA.ReadOnly = True
        Me.CAMARA.Width = 60
        '
        'BANDA
        '
        Me.BANDA.DataPropertyName = "racd_ba"
        Me.BANDA.HeaderText = "BA"
        Me.BANDA.Name = "BANDA"
        Me.BANDA.ReadOnly = True
        Me.BANDA.Width = 60
        '
        'COLUMNA
        '
        Me.COLUMNA.DataPropertyName = "racd_co"
        Me.COLUMNA.HeaderText = "CO"
        Me.COLUMNA.Name = "COLUMNA"
        Me.COLUMNA.ReadOnly = True
        Me.COLUMNA.Width = 60
        '
        'PISO
        '
        Me.PISO.DataPropertyName = "racd_pi"
        Me.PISO.HeaderText = "PI"
        Me.PISO.Name = "PISO"
        Me.PISO.ReadOnly = True
        Me.PISO.Width = 60
        '
        'NIVEL
        '
        Me.NIVEL.DataPropertyName = "racd_ni"
        Me.NIVEL.HeaderText = "NI"
        Me.NIVEL.Name = "NIVEL"
        Me.NIVEL.ReadOnly = True
        Me.NIVEL.Width = 60
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Btn_Todos)
        Me.GroupBox1.Controls.Add(Me.Txt_Soportante)
        Me.GroupBox1.Controls.Add(Me.Cmbo_Producto)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(755, 84)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "FILTRAR"
        '
        'Txt_Soportante
        '
        Me.Txt_Soportante.Location = New System.Drawing.Point(109, 51)
        Me.Txt_Soportante.Name = "Txt_Soportante"
        Me.Txt_Soportante.Size = New System.Drawing.Size(163, 21)
        Me.Txt_Soportante.TabIndex = 3
        '
        'Cmbo_Producto
        '
        Me.Cmbo_Producto.FormattingEnabled = True
        Me.Cmbo_Producto.Location = New System.Drawing.Point(109, 20)
        Me.Cmbo_Producto.Name = "Cmbo_Producto"
        Me.Cmbo_Producto.Size = New System.Drawing.Size(414, 21)
        Me.Cmbo_Producto.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "SOPORTANTE"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "PRODUCTO"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CAMBIARPORESTEPALLETToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(224, 26)
        '
        'CAMBIARPORESTEPALLETToolStripMenuItem
        '
        Me.CAMBIARPORESTEPALLETToolStripMenuItem.Name = "CAMBIARPORESTEPALLETToolStripMenuItem"
        Me.CAMBIARPORESTEPALLETToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.CAMBIARPORESTEPALLETToolStripMenuItem.Text = "CAMBIAR POR ESTE PALLET"
        '
        'DataGridViewImageColumn1
        '
        Me.DataGridViewImageColumn1.HeaderText = "OK"
        Me.DataGridViewImageColumn1.Image = Global.SistemaEnvioCorreos.Web.My.Resources.Resources._1373576845_tick_16
        Me.DataGridViewImageColumn1.Name = "DataGridViewImageColumn1"
        Me.DataGridViewImageColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewImageColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewImageColumn1.Width = 40
        '
        'OK
        '
        Me.OK.HeaderText = "OK"
        Me.OK.Image = Global.SistemaEnvioCorreos.Web.My.Resources.Resources._1373576845_tick_16
        Me.OK.Name = "OK"
        Me.OK.ReadOnly = True
        Me.OK.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.OK.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.OK.Width = 40
        '
        'Btn_Todos
        '
        Me.Btn_Todos.Location = New System.Drawing.Point(566, 26)
        Me.Btn_Todos.Name = "Btn_Todos"
        Me.Btn_Todos.Size = New System.Drawing.Size(173, 37)
        Me.Btn_Todos.TabIndex = 4
        Me.Btn_Todos.Text = "Todos los Productos"
        Me.Btn_Todos.UseVisualStyleBackColor = True
        '
        'Frm_ProductosCliente
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1087, 525)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Frm_ProductosCliente"
        Me.ShowIcon = False
        Me.Text = "Listado de pallet por cliente"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.Data_Productos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Data_Productos As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Txt_Soportante As System.Windows.Forms.TextBox
    Friend WithEvents Cmbo_Producto As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Rut_Cliente As System.Windows.Forms.Label
    Friend WithEvents Btn_Todos As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CAMBIARPORESTEPALLETToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Folio_precisa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FOLIO_CLIENTE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PRODUCTO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Cajas As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Kilos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CAMARA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BANDA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COLUMNA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PISO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NIVEL As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OK As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents DataGridViewImageColumn1 As System.Windows.Forms.DataGridViewImageColumn
End Class
