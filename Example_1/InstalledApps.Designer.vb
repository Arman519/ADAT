<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InstalledApps
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InstalledApps))
        Me.InstalledAppsRtb = New System.Windows.Forms.RichTextBox()
        Me.CloseButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'InstalledAppsRtb
        '
        Me.InstalledAppsRtb.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.InstalledAppsRtb.Location = New System.Drawing.Point(12, 12)
        Me.InstalledAppsRtb.Name = "InstalledAppsRtb"
        Me.InstalledAppsRtb.ReadOnly = True
        Me.InstalledAppsRtb.Size = New System.Drawing.Size(517, 600)
        Me.InstalledAppsRtb.TabIndex = 0
        Me.InstalledAppsRtb.Text = ""
        '
        'CloseButton
        '
        Me.CloseButton.Location = New System.Drawing.Point(232, 618)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(84, 29)
        Me.CloseButton.TabIndex = 1
        Me.CloseButton.Text = "Close"
        Me.CloseButton.UseVisualStyleBackColor = True
        '
        'InstalledApps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(541, 659)
        Me.Controls.Add(Me.CloseButton)
        Me.Controls.Add(Me.InstalledAppsRtb)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "InstalledApps"
        Me.Text = "Installed Applications"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents InstalledAppsRtb As RichTextBox
    Friend WithEvents CloseButton As Button
End Class
