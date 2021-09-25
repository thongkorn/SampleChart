<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChartExportExcel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChartExportExcel))
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvData = New System.Windows.Forms.DataGridView()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnGenData = New System.Windows.Forms.Button()
        Me.btnChart = New System.Windows.Forms.Button()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.cmbChartType = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkShowLegend = New System.Windows.Forms.CheckBox()
        Me.btnCopyClipboard = New System.Windows.Forms.Button()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 14)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Label1"
        '
        'dgvData
        '
        Me.dgvData.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvData.Location = New System.Drawing.Point(1, 74)
        Me.dgvData.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.Size = New System.Drawing.Size(1052, 241)
        Me.dgvData.TabIndex = 7
        '
        'btnClose
        '
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Moccasin
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnClose.Image = CType(resources.GetObject("btnClose.Image"), System.Drawing.Image)
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.Location = New System.Drawing.Point(527, 12)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(124, 31)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close Program"
        Me.btnClose.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnGenData
        '
        Me.btnGenData.BackColor = System.Drawing.SystemColors.Control
        Me.btnGenData.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnGenData.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Moccasin
        Me.btnGenData.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGenData.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnGenData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGenData.Image = CType(resources.GetObject("btnGenData.Image"), System.Drawing.Image)
        Me.btnGenData.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGenData.Location = New System.Drawing.Point(6, 13)
        Me.btnGenData.Name = "btnGenData"
        Me.btnGenData.Size = New System.Drawing.Size(124, 31)
        Me.btnGenData.TabIndex = 0
        Me.btnGenData.Text = "Generate Data"
        Me.btnGenData.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnGenData.UseVisualStyleBackColor = False
        '
        'btnChart
        '
        Me.btnChart.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnChart.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Moccasin
        Me.btnChart.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChart.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnChart.Image = CType(resources.GetObject("btnChart.Image"), System.Drawing.Image)
        Me.btnChart.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnChart.Location = New System.Drawing.Point(136, 13)
        Me.btnChart.Name = "btnChart"
        Me.btnChart.Size = New System.Drawing.Size(124, 31)
        Me.btnChart.TabIndex = 1
        Me.btnChart.Text = "Chart Sample"
        Me.btnChart.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnChart.UseVisualStyleBackColor = True
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExportExcel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Moccasin
        Me.btnExportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExportExcel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnExportExcel.Image = CType(resources.GetObject("btnExportExcel.Image"), System.Drawing.Image)
        Me.btnExportExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExportExcel.Location = New System.Drawing.Point(267, 13)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(124, 31)
        Me.btnExportExcel.TabIndex = 2
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExportExcel.UseVisualStyleBackColor = True
        '
        'Chart1
        '
        Me.Chart1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Me.Chart1.Location = New System.Drawing.Point(1, 320)
        Me.Chart1.Name = "Chart1"
        Me.Chart1.Size = New System.Drawing.Size(1052, 318)
        Me.Chart1.TabIndex = 8
        Me.Chart1.Text = "Chart1"
        '
        'cmbChartType
        '
        Me.cmbChartType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbChartType.FormattingEnabled = True
        Me.cmbChartType.Location = New System.Drawing.Point(737, 17)
        Me.cmbChartType.Name = "cmbChartType"
        Me.cmbChartType.Size = New System.Drawing.Size(140, 22)
        Me.cmbChartType.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(663, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 14)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Chart Type"
        '
        'chkShowLegend
        '
        Me.chkShowLegend.AutoSize = True
        Me.chkShowLegend.Checked = True
        Me.chkShowLegend.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLegend.Location = New System.Drawing.Point(737, 50)
        Me.chkShowLegend.Name = "chkShowLegend"
        Me.chkShowLegend.Size = New System.Drawing.Size(107, 18)
        Me.chkShowLegend.TabIndex = 6
        Me.chkShowLegend.Text = "Show Legends"
        Me.chkShowLegend.UseVisualStyleBackColor = True
        '
        'btnCopyClipboard
        '
        Me.btnCopyClipboard.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCopyClipboard.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Moccasin
        Me.btnCopyClipboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCopyClipboard.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnCopyClipboard.Image = CType(resources.GetObject("btnCopyClipboard.Image"), System.Drawing.Image)
        Me.btnCopyClipboard.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCopyClipboard.Location = New System.Drawing.Point(397, 12)
        Me.btnCopyClipboard.Name = "btnCopyClipboard"
        Me.btnCopyClipboard.Size = New System.Drawing.Size(124, 31)
        Me.btnCopyClipboard.TabIndex = 3
        Me.btnCopyClipboard.Text = "Copy to Clipboard"
        Me.btnCopyClipboard.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCopyClipboard.UseVisualStyleBackColor = True
        '
        'frmChartExportExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1054, 639)
        Me.Controls.Add(Me.btnCopyClipboard)
        Me.Controls.Add(Me.chkShowLegend)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbChartType)
        Me.Controls.Add(Me.btnExportExcel)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.btnChart)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnGenData)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvData)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.KeyPreview = True
        Me.Name = "frmChartExportExcel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sample Chart & Export Data From DataGridView to Excel - coDe bY: Thongkorn Tubtim" & _
    "krob"
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnGenData As System.Windows.Forms.Button
    Friend WithEvents btnChart As System.Windows.Forms.Button
    Friend WithEvents btnExportExcel As System.Windows.Forms.Button
    Friend WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents cmbChartType As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkShowLegend As System.Windows.Forms.CheckBox
    Friend WithEvents btnCopyClipboard As System.Windows.Forms.Button

End Class
