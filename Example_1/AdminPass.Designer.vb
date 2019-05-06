<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AdminPass
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AdminPass))
        Me.CloseButton = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.AdminPassText = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'CloseButton
        '
        Me.CloseButton.Location = New System.Drawing.Point(139, 63)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(84, 40)
        Me.CloseButton.TabIndex = 2
        Me.CloseButton.Text = "Close"
        Me.CloseButton.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(46, 63)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(87, 40)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Copy to Clipboard"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'AdminPassText
        '
        Me.AdminPassText.Location = New System.Drawing.Point(46, 19)
        Me.AdminPassText.Name = "AdminPassText"
        Me.AdminPassText.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.AdminPassText.Size = New System.Drawing.Size(177, 29)
        Me.AdminPassText.TabIndex = 5
        Me.AdminPassText.Text = ""
        '
        'AdminPass
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(282, 127)
        Me.Controls.Add(Me.AdminPassText)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CloseButton)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AdminPass"
        Me.Text = "AdminPass"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CloseButton As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents AdminPassText As RichTextBox
End Class
