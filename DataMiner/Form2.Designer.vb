<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.oFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.cmdLoad = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lblStatus1 = New System.Windows.Forms.Label()
        Me.lblStatus2 = New System.Windows.Forms.Label()
        Me.cmdLoadFolder = New System.Windows.Forms.Button()
        Me.oFolderDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.cbAfterAnalysis = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbIncludeCitation = New System.Windows.Forms.CheckBox()
        Me.cbPairwiseDel = New System.Windows.Forms.CheckBox()
        Me.cbListwiseDel = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'oFileDialog
        '
        Me.oFileDialog.Filter = "Excel File|*.xlsx|All Files|*.*"
        '
        'cmdLoad
        '
        Me.cmdLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdLoad.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLoad.Location = New System.Drawing.Point(12, 12)
        Me.cmdLoad.Name = "cmdLoad"
        Me.cmdLoad.Size = New System.Drawing.Size(98, 23)
        Me.cmdLoad.TabIndex = 3
        Me.cmdLoad.Text = "Load File"
        Me.cmdLoad.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.Location = New System.Drawing.Point(237, 12)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(98, 23)
        Me.cmdExit.TabIndex = 4
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'lblStatus1
        '
        Me.lblStatus1.AutoSize = True
        Me.lblStatus1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus1.Location = New System.Drawing.Point(13, 92)
        Me.lblStatus1.Name = "lblStatus1"
        Me.lblStatus1.Size = New System.Drawing.Size(90, 16)
        Me.lblStatus1.TabIndex = 5
        Me.lblStatus1.Text = "Current Status"
        '
        'lblStatus2
        '
        Me.lblStatus2.AutoSize = True
        Me.lblStatus2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblStatus2.Location = New System.Drawing.Point(284, 92)
        Me.lblStatus2.Name = "lblStatus2"
        Me.lblStatus2.Size = New System.Drawing.Size(55, 16)
        Me.lblStatus2.TabIndex = 6
        Me.lblStatus2.Text = "READY"
        '
        'cmdLoadFolder
        '
        Me.cmdLoadFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdLoadFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLoadFolder.Location = New System.Drawing.Point(116, 12)
        Me.cmdLoadFolder.Name = "cmdLoadFolder"
        Me.cmdLoadFolder.Size = New System.Drawing.Size(98, 23)
        Me.cmdLoadFolder.TabIndex = 7
        Me.cmdLoadFolder.Text = "Load Folder"
        Me.cmdLoadFolder.UseVisualStyleBackColor = True
        '
        'cbAfterAnalysis
        '
        Me.cbAfterAnalysis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbAfterAnalysis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAfterAnalysis.FormattingEnabled = True
        Me.cbAfterAnalysis.Items.AddRange(New Object() {"View Report", "Close Program"})
        Me.cbAfterAnalysis.Location = New System.Drawing.Point(116, 51)
        Me.cbAfterAnalysis.Name = "cbAfterAnalysis"
        Me.cbAfterAnalysis.Size = New System.Drawing.Size(146, 24)
        Me.cbAfterAnalysis.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "On Complete:"
        '
        'cbIncludeCitation
        '
        Me.cbIncludeCitation.AutoSize = True
        Me.cbIncludeCitation.Location = New System.Drawing.Point(464, 22)
        Me.cbIncludeCitation.Name = "cbIncludeCitation"
        Me.cbIncludeCitation.Size = New System.Drawing.Size(104, 17)
        Me.cbIncludeCitation.TabIndex = 14
        Me.cbIncludeCitation.Text = "Include Citations"
        Me.cbIncludeCitation.UseVisualStyleBackColor = True
        '
        'cbPairwiseDel
        '
        Me.cbPairwiseDel.AutoSize = True
        Me.cbPairwiseDel.Location = New System.Drawing.Point(618, 45)
        Me.cbPairwiseDel.Name = "cbPairwiseDel"
        Me.cbPairwiseDel.Size = New System.Drawing.Size(107, 17)
        Me.cbPairwiseDel.TabIndex = 15
        Me.cbPairwiseDel.Text = "Pairwise Deletion"
        Me.cbPairwiseDel.UseVisualStyleBackColor = True
        '
        'cbListwiseDel
        '
        Me.cbListwiseDel.AutoSize = True
        Me.cbListwiseDel.Location = New System.Drawing.Point(618, 68)
        Me.cbListwiseDel.Name = "cbListwiseDel"
        Me.cbListwiseDel.Size = New System.Drawing.Size(105, 17)
        Me.cbListwiseDel.TabIndex = 16
        Me.cbListwiseDel.Text = "Listwise Deletion"
        Me.cbListwiseDel.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(861, 121)
        Me.ControlBox = False
        Me.Controls.Add(Me.cbListwiseDel)
        Me.Controls.Add(Me.cbPairwiseDel)
        Me.Controls.Add(Me.cbIncludeCitation)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbAfterAnalysis)
        Me.Controls.Add(Me.cmdLoadFolder)
        Me.Controls.Add(Me.lblStatus2)
        Me.Controls.Add(Me.lblStatus1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdLoad)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DataMiner CTT"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents oFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdLoad As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents lblStatus1 As System.Windows.Forms.Label
    Friend WithEvents lblStatus2 As System.Windows.Forms.Label
    Friend WithEvents cmdLoadFolder As System.Windows.Forms.Button
    Friend WithEvents oFolderDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents cbAfterAnalysis As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbIncludeCitation As System.Windows.Forms.CheckBox
    Friend WithEvents cbPairwiseDel As System.Windows.Forms.CheckBox
    Friend WithEvents cbListwiseDel As System.Windows.Forms.CheckBox

End Class
