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
        Me.SuspendLayout()
        '
        'oFileDialog
        '
        Me.oFileDialog.Filter = "Excel File|*.xlsx|All Files|*.*"
        '
        'cmdLoad
        '
        Me.cmdLoad.Location = New System.Drawing.Point(12, 12)
        Me.cmdLoad.Name = "cmdLoad"
        Me.cmdLoad.Size = New System.Drawing.Size(75, 23)
        Me.cmdLoad.TabIndex = 3
        Me.cmdLoad.Text = "Load File"
        Me.cmdLoad.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(235, 12)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(75, 23)
        Me.cmdExit.TabIndex = 4
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'lblStatus1
        '
        Me.lblStatus1.AutoSize = True
        Me.lblStatus1.Location = New System.Drawing.Point(12, 55)
        Me.lblStatus1.Name = "lblStatus1"
        Me.lblStatus1.Size = New System.Drawing.Size(189, 13)
        Me.lblStatus1.TabIndex = 5
        Me.lblStatus1.Text = "Contact Chris Stefancik for assistance."
        '
        'lblStatus2
        '
        Me.lblStatus2.AutoSize = True
        Me.lblStatus2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblStatus2.Location = New System.Drawing.Point(266, 55)
        Me.lblStatus2.Name = "lblStatus2"
        Me.lblStatus2.Size = New System.Drawing.Size(44, 13)
        Me.lblStatus2.TabIndex = 6
        Me.lblStatus2.Text = "READY"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(322, 82)
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

End Class
