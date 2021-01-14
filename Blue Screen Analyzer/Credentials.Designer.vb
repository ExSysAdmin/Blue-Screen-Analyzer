<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Credentials
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.UserNameTextBox = New System.Windows.Forms.TextBox()
        Me.PasswordTextBox = New System.Windows.Forms.TextBox()
        Me.WindowsAuthCheckBox = New System.Windows.Forms.CheckBox()
        Me.DomainComboBox = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 164)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Ok"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(93, 164)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Username:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Domain:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Password:"
        '
        'UserNameTextBox
        '
        Me.UserNameTextBox.Location = New System.Drawing.Point(76, 21)
        Me.UserNameTextBox.Name = "UserNameTextBox"
        Me.UserNameTextBox.Size = New System.Drawing.Size(170, 20)
        Me.UserNameTextBox.TabIndex = 0
        '
        'PasswordTextBox
        '
        Me.PasswordTextBox.Location = New System.Drawing.Point(76, 56)
        Me.PasswordTextBox.Name = "PasswordTextBox"
        Me.PasswordTextBox.Size = New System.Drawing.Size(170, 20)
        Me.PasswordTextBox.TabIndex = 1
        Me.PasswordTextBox.UseSystemPasswordChar = True
        '
        'WindowsAuthCheckBox
        '
        Me.WindowsAuthCheckBox.AutoSize = True
        Me.WindowsAuthCheckBox.Location = New System.Drawing.Point(12, 131)
        Me.WindowsAuthCheckBox.Name = "WindowsAuthCheckBox"
        Me.WindowsAuthCheckBox.Size = New System.Drawing.Size(163, 17)
        Me.WindowsAuthCheckBox.TabIndex = 3
        Me.WindowsAuthCheckBox.Text = "Use Windows Authentication"
        Me.WindowsAuthCheckBox.UseVisualStyleBackColor = True
        '
        'DomainComboBox
        '
        Me.DomainComboBox.FormattingEnabled = True
        Me.DomainComboBox.Location = New System.Drawing.Point(76, 93)
        Me.DomainComboBox.Name = "DomainComboBox"
        Me.DomainComboBox.Size = New System.Drawing.Size(170, 21)
        Me.DomainComboBox.TabIndex = 2
        '
        'Credentials
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(266, 193)
        Me.Controls.Add(Me.DomainComboBox)
        Me.Controls.Add(Me.WindowsAuthCheckBox)
        Me.Controls.Add(Me.PasswordTextBox)
        Me.Controls.Add(Me.UserNameTextBox)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Credentials"
        Me.Text = "Credentials"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents UserNameTextBox As TextBox
    Friend WithEvents PasswordTextBox As TextBox
    Friend WithEvents WindowsAuthCheckBox As CheckBox
    Friend WithEvents DomainComboBox As ComboBox
End Class
