Public Class Credentials
    Public ObjCred As System.Net.NetworkCredential

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If WindowsAuthCheckBox.Checked = False Then


            ObjCred = New Net.NetworkCredential(UserNameTextBox.Text, PasswordTextBox.Text, DomainComboBox.Text)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            Me.DialogResult = DialogResult.No
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub WindowsAuthCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles WindowsAuthCheckBox.CheckedChanged
        If WindowsAuthCheckBox.Checked Then
            UserNameTextBox.Enabled = False
            PasswordTextBox.Enabled = False
            DomainComboBox.Enabled = False
        Else
            UserNameTextBox.Enabled = True
            PasswordTextBox.Enabled = True
            DomainComboBox.Enabled = True
        End If
    End Sub
End Class