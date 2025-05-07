<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SEI_SRV_VOXEL
    Inherits SBO_BASE.SBO_Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblEmpresa = New System.Windows.Forms.Label()
        Me.txtempresa = New System.Windows.Forms.TextBox()
        Me.lblUsuario = New System.Windows.Forms.Label()
        Me.txtUsuario = New System.Windows.Forms.TextBox()
        Me.lblProceso = New System.Windows.Forms.Label()
        Me.txtProceso = New System.Windows.Forms.TextBox()
        Me.btnEjecutar = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.lblmsg = New System.Windows.Forms.Label()
        Me.chkFacturacion = New System.Windows.Forms.CheckBox()
        Me.chkConfirmacions = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'lblEmpresa
        '
        Me.lblEmpresa.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpresa.Location = New System.Drawing.Point(9, 37)
        Me.lblEmpresa.Name = "lblEmpresa"
        Me.lblEmpresa.Size = New System.Drawing.Size(76, 17)
        Me.lblEmpresa.TabIndex = 2
        Me.lblEmpresa.Text = "Empresa"
        '
        'txtempresa
        '
        Me.txtempresa.Location = New System.Drawing.Point(91, 34)
        Me.txtempresa.Name = "txtempresa"
        Me.txtempresa.Size = New System.Drawing.Size(322, 20)
        Me.txtempresa.TabIndex = 3
        '
        'lblUsuario
        '
        Me.lblUsuario.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsuario.Location = New System.Drawing.Point(9, 63)
        Me.lblUsuario.Name = "lblUsuario"
        Me.lblUsuario.Size = New System.Drawing.Size(76, 17)
        Me.lblUsuario.TabIndex = 4
        Me.lblUsuario.Text = "Usuario"
        '
        'txtUsuario
        '
        Me.txtUsuario.Location = New System.Drawing.Point(91, 60)
        Me.txtUsuario.Name = "txtUsuario"
        Me.txtUsuario.Size = New System.Drawing.Size(322, 20)
        Me.txtUsuario.TabIndex = 5
        '
        'lblProceso
        '
        Me.lblProceso.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProceso.Location = New System.Drawing.Point(9, 89)
        Me.lblProceso.Name = "lblProceso"
        Me.lblProceso.Size = New System.Drawing.Size(76, 17)
        Me.lblProceso.TabIndex = 6
        Me.lblProceso.Text = "Proceso"
        '
        'txtProceso
        '
        Me.txtProceso.Location = New System.Drawing.Point(91, 86)
        Me.txtProceso.Name = "txtProceso"
        Me.txtProceso.Size = New System.Drawing.Size(322, 20)
        Me.txtProceso.TabIndex = 7
        '
        'btnEjecutar
        '
        Me.btnEjecutar.Location = New System.Drawing.Point(12, 284)
        Me.btnEjecutar.Name = "btnEjecutar"
        Me.btnEjecutar.Size = New System.Drawing.Size(114, 22)
        Me.btnEjecutar.TabIndex = 19
        Me.btnEjecutar.Text = "Ejecutar Procesos "
        Me.btnEjecutar.UseVisualStyleBackColor = True
        '
        'btnSalir
        '
        Me.btnSalir.Location = New System.Drawing.Point(132, 284)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(74, 22)
        Me.btnSalir.TabIndex = 20
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'lblmsg
        '
        Me.lblmsg.Location = New System.Drawing.Point(14, 207)
        Me.lblmsg.Name = "lblmsg"
        Me.lblmsg.Size = New System.Drawing.Size(585, 72)
        Me.lblmsg.TabIndex = 15
        '
        'chkFacturacion
        '
        Me.chkFacturacion.AutoSize = True
        Me.chkFacturacion.Location = New System.Drawing.Point(91, 133)
        Me.chkFacturacion.Name = "chkFacturacion"
        Me.chkFacturacion.Size = New System.Drawing.Size(180, 17)
        Me.chkFacturacion.TabIndex = 8
        Me.chkFacturacion.Text = "Facturación electronica/abonos "
        Me.chkFacturacion.UseVisualStyleBackColor = True
        '
        'chkConfirmacions
        '
        Me.chkConfirmacions.AutoSize = True
        Me.chkConfirmacions.Location = New System.Drawing.Point(91, 180)
        Me.chkConfirmacions.Name = "chkConfirmacions"
        Me.chkConfirmacions.Size = New System.Drawing.Size(165, 17)
        Me.chkConfirmacions.TabIndex = 21
        Me.chkConfirmacions.Text = "Leer confirmaciones facturas "
        Me.chkConfirmacions.UseVisualStyleBackColor = True
        '
        'SEI_SRV_VOXEL
        '
        Me.ClientSize = New System.Drawing.Size(615, 327)
        Me.Controls.Add(Me.chkConfirmacions)
        Me.Controls.Add(Me.chkFacturacion)
        Me.Controls.Add(Me.lblmsg)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnEjecutar)
        Me.Controls.Add(Me.lblProceso)
        Me.Controls.Add(Me.txtProceso)
        Me.Controls.Add(Me.lblUsuario)
        Me.Controls.Add(Me.txtUsuario)
        Me.Controls.Add(Me.lblEmpresa)
        Me.Controls.Add(Me.txtempresa)
        Me.Name = "SEI_SRV_VOXEL"
        Me.Text = "Facturación electronica - Voxel"
        Me.TopMost = False
        Me.Controls.SetChildIndex(Me.txtempresa, 0)
        Me.Controls.SetChildIndex(Me.lblEmpresa, 0)
        Me.Controls.SetChildIndex(Me.txtUsuario, 0)
        Me.Controls.SetChildIndex(Me.lblUsuario, 0)
        Me.Controls.SetChildIndex(Me.txtProceso, 0)
        Me.Controls.SetChildIndex(Me.lblProceso, 0)
        Me.Controls.SetChildIndex(Me.btnEjecutar, 0)
        Me.Controls.SetChildIndex(Me.btnSalir, 0)
        Me.Controls.SetChildIndex(Me.lblmsg, 0)
        Me.Controls.SetChildIndex(Me.chkFacturacion, 0)
        Me.Controls.SetChildIndex(Me.chkConfirmacions, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblEmpresa As System.Windows.Forms.Label
    Friend WithEvents txtempresa As System.Windows.Forms.TextBox
    Friend WithEvents lblUsuario As System.Windows.Forms.Label
    Friend WithEvents txtUsuario As System.Windows.Forms.TextBox
    Friend WithEvents lblProceso As System.Windows.Forms.Label
    Friend WithEvents txtProceso As System.Windows.Forms.TextBox
    Friend WithEvents btnEjecutar As System.Windows.Forms.Button
    Friend WithEvents btnSalir As System.Windows.Forms.Button
    Friend WithEvents lblmsg As System.Windows.Forms.Label
    Friend WithEvents chkFacturacion As System.Windows.Forms.CheckBox
    Friend WithEvents chkConfirmacions As System.Windows.Forms.CheckBox

End Class
